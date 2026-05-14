import React, { useState, useEffect } from 'react';
import { AnimatePresence, motion } from 'framer-motion';
import { Menu, X, Rocket, HelpCircle, MessageSquare, ChevronLeft, ShieldAlert, Award, UserPlus, Zap, Check } from 'lucide-react';
import { apiService } from './services/api';
import type { User, Template } from './types';
import { WizardView } from './components/WizardView';

const App: React.FC = () => {
  const [view, setView] = useState<'home' | 'wizard' | 'status' | 'how-to' | 'credits'>('home');
  const [currentJobId, setCurrentJobId] = useState<string | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [user, setUser] = useState<User | null>(null);
  const [templates, setTemplates] = useState<Template[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isExternalBrowser, setIsExternalBrowser] = useState(false);

  useEffect(() => {
    const init = async () => {
      try {
        const tg = (window as any).Telegram?.WebApp;
        if (tg?.initData && tg.initData !== "") {
          tg.ready();
          tg.expand();
          
          try {
            const data = await apiService.init();
            setUser(data.user);
            setTemplates(data.templates);
            
            const activeJob = await apiService.getActiveJob();
            if (activeJob && activeJob.job_id) {
                setCurrentJobId(activeJob.job_id);
                setView('status');
            }
          } catch (apiErr: any) {
            setError('Server bilan aloqa o‘rnatib bo‘lmadi.');
          }
          setLoading(false);
        } else {
            setLoading(false);
        }
      } catch (err: any) {
        setError('Kutilmagan xatolik yuz berdi.');
        setLoading(false);
      }
    };
    init();
  }, []);

  const toggleSidebar = () => setSidebarOpen(!sidebarOpen);

  if (loading) return null;

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white font-sans overflow-hidden relative flex flex-col">
      <header className="fixed top-0 left-0 right-0 h-16 flex items-center justify-between px-6 bg-[#0a0a0a]/80 backdrop-blur-md z-40 border-b border-white/5">
        <button onClick={toggleSidebar} className="p-2 hover:bg-white/5 rounded-full transition-colors"><Menu className="w-6 h-6" /></button>
        <h1 className="text-xl font-bold bg-gradient-to-r from-[#a78bfa] to-[#7c3aed] bg-clip-text text-transparent">Slide Generator</h1>
        <div className="w-10" />
      </header>

      <main className="pt-16 flex-1 flex flex-col relative">
        <AnimatePresence mode="wait">
          {view === 'home' && <HomeView key="home" user={user} onStart={() => user && user.available_generations > 0 ? setView('wizard') : setView('credits')} />}
          {view === 'wizard' && <WizardView key="wizard" templates={templates} onComplete={(id) => { setCurrentJobId(id); setView('status'); }} onCancel={() => setView('home')} />}
          {view === 'status' && <StatusView key="status" jobId={currentJobId || ''} onDone={() => { setCurrentJobId(null); setView('home'); }} />}
          {view === 'credits' && <CreditsView key="credits" user={user} onBack={() => setView('home')} />}
          {view === 'how-to' && <HowToView key="how-to" onBack={() => setView('home')} />}
        </AnimatePresence>
      </main>
    </div>
  );
};

const StatusView: React.FC<{ jobId: string, onDone: () => void }> = ({ jobId, onDone }) => {
  const [progress, setProgress] = useState(0);
  const [step, setStep] = useState<string>('queued');

  useEffect(() => {
    const fetchStatus = async () => {
      try {
        const data = await apiService.getStatus(jobId);
        setProgress(data.progress);
        setStep(data.step);
        if (data.status === 'completed') onDone();
      } catch (err) { console.error(err); }
    };
    fetchStatus();
    const interval = setInterval(fetchStatus, 3000);
    return () => clearInterval(interval);
  }, [jobId, onDone]);

  const steps = [
    { id: 'queued', label: 'So‘rov qabul qilindi' },
    { id: 'research', label: 'Ma’lumotlar yig‘ilmoqda' },
    { id: 'planning', label: 'Reja tuzilmoqda' },
    { id: 'rendering', label: 'Taqdimot yig‘ilmoqda' },
    { id: 'uploading', label: 'Fayl yuborilmoqda' },
    { id: 'generating_pdf', label: 'PDF tayyorlanmoqda' },
    { id: 'done', label: 'Taqdimot tayyor' }
  ];
  const currentStepIndex = steps.findIndex(s => s.id === step);

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex-1 flex flex-col p-8 space-y-8">
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold">Jarayon holati</h2>
        <div className="text-5xl font-black text-primary">{progress}%</div>
      </div>
      <div className="h-3 w-full bg-[#171717] rounded-full overflow-hidden">
        <motion.div className="h-full bg-primary" animate={{ width: `${progress}%` }} />
      </div>
      <div className="flex-1 space-y-3">
        {steps.map((s, i) => (
            <div key={s.id} className={`p-3 rounded-2xl flex items-center space-x-4 transition-all ${i <= currentStepIndex ? 'bg-[#7c3aed]/10 border border-[#7c3aed]/20' : 'bg-[#171717] border border-transparent'}`}>
                <div className={`w-8 h-8 rounded-full flex items-center justify-center ${i <= currentStepIndex ? 'bg-primary text-white' : 'bg-white/5 text-white/30'}`}>{i < currentStepIndex ? <Check size={16} /> : <span>{i + 1}</span>}</div>
                <span className={`text-sm ${i <= currentStepIndex ? 'text-white' : 'text-white/30'}`}>{s.label}</span>
            </div>
        ))}
      </div>
      {progress >= 100 && (
        <button onClick={() => window.location.href = 'https://t.me/uzafo_slide_bot'} className="w-full py-4 bg-primary text-white rounded-2xl font-bold transition-all hover:bg-primary/90">Botga qaytish</button>
      )}
    </motion.div>
  );
};

// ... (other components like CreditsView, HowToView, HomeView remain as defined before)

export default App;
