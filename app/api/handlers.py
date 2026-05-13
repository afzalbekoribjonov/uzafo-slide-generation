from __future__ import annotations

from aiohttp import web
from app.api.auth import validate_init_data
from app.repositories.users import UsersRepository
from app.services.generations import GenerationAccessService
from app.services.generation_queue import GenerationQueueService
from app.services.presentation_templates import PresentationTemplateRegistry
import json
from pathlib import Path

template_registry = PresentationTemplateRegistry()

async def get_user_from_request(request: web.Request):
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return None
    
    init_data = auth_header[7:]
    settings = request.app['settings']
    user_info = validate_init_data(init_data, settings.bot_token)
    if not user_info:
        return None
    
    users_repo: UsersRepository = request.app['users_repo']
    user = await users_repo.get_by_telegram_id(user_info['id'])
    return user

async def handle_init(request: web.Request) -> web.Response:
    user = await get_user_from_request(request)
    if not user:
        return web.json_response({'error': 'Unauthorized'}, status=401)
    
    access_service: GenerationAccessService = request.app['generation_access_service']
    available = access_service.available_generations(user)
    
    templates = template_registry.list_templates()
    # Simplify templates for WebApp
    web_templates = []
    for t in templates:
        web_templates.append({
            'id': t['id'],
            'name': t.get('name', t['id']),
            'description': t.get('description', ''),
            'preview_url': f"/api/templates/preview/{t['id']}" # We'll add this route too
        })

    return web.json_response({
        'user': {
            'id': user['telegram_id'],
            'full_name': user['full_name'],
            'available_generations': available,
            'is_blocked': user.get('generation_access_blocked', False)
        },
        'templates': web_templates
    })

async def handle_create(request: web.Request) -> web.Response:
    user = await get_user_from_request(request)
    if not user:
        return web.json_response({'error': 'Unauthorized'}, status=401)
    
    if user.get('generation_access_blocked'):
        return web.json_response({'error': 'Access blocked'}, status=403)

    access_service: GenerationAccessService = request.app['generation_access_service']
    if not access_service.has_available_generation(user):
        return web.json_response({'error': 'No credits available'}, status=403)

    queue_service: GenerationQueueService = request.app['generation_queue_service']
    existing_job, _ = await queue_service.describe_existing_job(user['telegram_id'])
    if existing_job:
        return web.json_response({'error': 'Already in queue', 'job_id': str(existing_job['_id'])}, status=400)

    try:
        data = await request.json()
    except Exception:
        return web.json_response({'error': 'Invalid JSON'}, status=400)

    # Validate required fields
    required = ['topic', 'presenter_name', 'slide_count', 'template_id', 'language_code', 'wants_pdf']
    if not all(k in data for k in required):
        return web.json_response({'error': 'Missing required fields'}, status=400)

    users_repo: UsersRepository = request.app['users_repo']
    consumed_from = await access_service.consume_generation(users_repo, user['telegram_id'])
    if not consumed_from:
        return web.json_response({'error': 'Failed to consume credit'}, status=403)

    try:
        job, ahead_count, active_job = await queue_service.create_job(
            telegram_id=user['telegram_id'],
            full_name=user['full_name'],
            username=user.get('username'),
            payload=data,
            consumed_from=consumed_from,
            status_chat_id=None, # WebApp doesn't have a chat_id in this context
            status_message_id=None,
        )
        if active_job:
            await access_service.restore_consumed_generation(users_repo, user['telegram_id'], consumed_from)
            return web.json_response({'error': 'Already in queue', 'job_id': str(active_job['_id'])}, status=400)
            
        return web.json_response({
            'success': True,
            'job_id': str(job['_id']),
            'ahead_count': ahead_count
        })
    except Exception as e:
        await access_service.restore_consumed_generation(users_repo, user['telegram_id'], consumed_from)
        return web.json_response({'error': str(e)}, status=500)

async def handle_status(request: web.Request) -> web.Response:
    user = await get_user_from_request(request)
    if not user:
        return web.json_response({'error': 'Unauthorized'}, status=401)
    
    job_id = request.match_info.get('job_id')
    queue_service: GenerationQueueService = request.app['generation_queue_service']
    
    job = await queue_service.generations_repo.get_by_id(job_id)
    if not job:
        return web.json_response({'error': 'Job not found'}, status=404)
    
    if job['telegram_id'] != user['telegram_id']:
        return web.json_response({'error': 'Forbidden'}, status=403)

    return web.json_response({
        'status': job['status'],
        'step': job.get('step', 'pending'),
        'progress': job.get('progress', 0),
        'error': job.get('error_message')
    })

async def handle_template_preview(request: web.Request) -> web.Response:
    template_id = request.match_info.get('template_id')
    template = template_registry.get(template_id)
    if not template or 'preview_path' not in template:
        return web.json_response({'error': 'Not found'}, status=404)
    
    preview_path = Path(template['preview_path'])
    if not preview_path.exists():
        return web.json_response({'error': 'File not found'}, status=404)
    
    return web.FileResponse(preview_path)

async def handle_options(request: web.Request) -> web.Response:
    return web.Response(headers={
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    })
