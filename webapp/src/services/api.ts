import axios from 'axios';
import type { CreatePresentationData, InitData, JobStatus } from '../types';

const API_BASE_URL = import.meta.env.VITE_API_URL || '';

const api = axios.create({
  baseURL: API_BASE_URL,
});

api.interceptors.request.use((config) => {
  const initData = window.Telegram?.WebApp?.initData;
  if (initData) {
    config.headers.Authorization = `Bearer ${initData}`;
  }
  return config;
});

export const apiService = {
  init: async (): Promise<InitData> => {
    const response = await api.post('/api/init');
    return response.data;
  },
  
  create: async (data: CreatePresentationData): Promise<{ job_id: string; ahead_count: number }> => {
    const response = await api.post('/api/create', data);
    return response.data;
  },
  
  getStatus: async (jobId: string): Promise<JobStatus> => {
    const response = await api.get(`/api/status/${jobId}`);
    return response.data;
  }
};
