export interface User {
  id: number;
  full_name: string;
  available_generations: number;
  is_blocked: boolean;
}

export interface Template {
  id: string;
  name: string;
  description: string;
  preview_url: string;
}

export interface InitData {
  user: User;
  templates: Template[];
}

export interface CreatePresentationData {
  topic: string;
  presenter_name: string;
  slide_count: number;
  template_id: string;
  language_code: string;
  wants_pdf: boolean;
}

export interface JobStatus {
  status: 'pending' | 'processing' | 'completed' | 'failed';
  step: string;
  progress: number;
  error?: string;
}
