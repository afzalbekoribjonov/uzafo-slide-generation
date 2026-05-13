from __future__ import annotations

import os
from pathlib import Path
from aiohttp import web
from app.api.handlers import (
    handle_init,
    handle_create,
    handle_status,
    handle_template_preview,
    handle_options
)

def setup_api_routes(app: web.Application):
    # API Routes
    app.router.add_post('/api/init', handle_init)
    app.router.add_post('/api/create', handle_create)
    app.router.add_get('/api/status/{job_id}', handle_status)
    app.router.add_get('/api/templates/preview/{template_id}', handle_template_preview)
    
    # CORS Preflight
    app.router.add_route('OPTIONS', '/api/{tail:.*}', handle_options)

def setup_static_routes(app: web.Application):
    # Serve static files from webapp/dist
    dist_path = Path(__file__).parent.parent.parent / 'webapp' / 'dist'
    
    if dist_path.exists():
        # 1. Serve /assets/ directly
        if (dist_path / 'assets').exists():
            app.router.add_static('/assets/', dist_path / 'assets', name='assets')
        
        # 2. Serve index.html for root and SPA paths
        async def serve_index(request):
            return web.FileResponse(dist_path / 'index.html')
            
        # 3. Serve specific static files in root if they exist
        # This is a bit manual but safer than a catch-all for static
        for item in dist_path.iterdir():
            if item.is_file() and item.name != 'index.html':
                # Map individual files like favicon.svg, icons.svg
                async def serve_file(request, file_path=item):
                    return web.FileResponse(file_path)
                app.router.add_get(f'/{item.name}', serve_file)
        
        # Root path
        app.router.add_get('/', serve_index)
        
        # 4. Catch-all for SPA (must be last)
        # Only if it doesn't match /api/ or /telegram/
        async def spa_handler(request):
            path = request.path
            if path.startswith('/api/') or path.startswith('/telegram/'):
                return web.Response(text="Not Found", status=404)
            return await serve_index(request)
            
        app.router.add_get('/{tail:.*}', spa_handler)
    else:
        async def no_dist(request):
            return web.Response(text="Frontend build (dist) not found.", status=404)
        app.router.add_get('/', no_dist)
        app.router.add_get('/{tail:.*}', no_dist)

@web.middleware
async def cors_middleware(request, handler):
    if request.method == "OPTIONS":
        return await handle_options(request)
    
    response = await handler(request)
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    return response
