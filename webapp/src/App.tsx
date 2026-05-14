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
            setIsExternalBrowser(true);
            setLoading(false);
        }
      } catch (err: any) {
        setError('Kutilmagan xatolik yuz berdi.');
        setLoading(false);
      }
    };
    init();
  }, []);

  if (loading) return null;
  if (isExternalBrowser) return <ExternalBrowserView />;
  if (error) return <ErrorView error={error} />;

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white font-sans overflow-hidden relative flex flex-col">
      <header className="fixed top-0 left-0 right-0 h-16 flex items-center justify-between px-6 bg-[#0a0a0a]/80 backdrop-blur-md z-40 border-b border-white/5">
        <button onClick={() => setSidebarOpen(!sidebarOpen)} className="p-2 hover:bg-white/5 rounded-full transition-colors"><Menu className="w-6 h-6" /></button>
        <h1 className="text-xl font-bold bg-gradient-to-r from-[#a78bfa] to-[#7c3aed] bg-clip-text text-transparent">Slide Generator</h1>
        <div className="w-10" />
      </header>

      <AnimatePresence>
        {sidebarOpen && (
          <>
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setSidebarOpen(false)} className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50" />
            <motion.div initial={{ x: '-100%' }} animate={{ x: 0 }} exit={{ x: '-100%' }} className="fixed top-0 left-0 bottom-0 w-4/5 max-w-sm bg-[#171717] z-50 border-r border-white/10 p-6 flex flex-col">
              <div className="flex items-center justify-between mb-8">
                <span className="text-xl font-bold text-[#a78bfa]">Menyu</span>
                <button onClick={() => setSidebarOpen(false)} className="p-2 hover:bg-white/5 rounded-full"><X className="w-6 h-6" /></button>
              </div>
              <nav className="flex-1 space-y-4">
                <SidebarItem icon={<Rocket className="w-5 h-5" />} label="Imkoniyatlarim" onClick={() => { setView('credits'); setSidebarOpen(false); }} active={view === 'credits'} />
                <SidebarItem icon={<HelpCircle className="w-5 h-5" />} label="Qo‘llanma" onClick={() => { setView('how-to'); setSidebarOpen(false); }} active={view === 'how-to'} />
                <SidebarItem icon={<MessageSquare className="w-5 h-5" />} label="Savol yoki taklif?" onClick={() => { window.open('https://uzafo.site/en/discussions/free-slide-generator', '_blank'); setSidebarOpen(false); }} />
              </nav>
            </motion.div>
          </>
        )}
      </AnimatePresence>

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

const ExternalBrowserView = () => (
    <div className="min-h-screen flex flex-col items-center justify-center p-8 text-center bg-[#0a0a0a] text-white">
        <ShieldAlert size={64} className="text-red-500 mb-6" />
        <h2 className="text-2xl font-bold mb-4">Kirish taqiqlangan</h2>
        <p className="text-white/60 mb-8">Ushbu xizmatdan faqat Telegram bot ichida foydalanish mumkin.</p>
        <button onClick={() => window.open('https://t.me/slayd_pptxbot', '_blank')} className="px-8 py-4 bg-primary rounded-2xl font-bold">Botga o‘tish</button>
    </div>
);

const ErrorView = ({ error }: { error: string }) => (
    <div className="min-h-screen flex flex-col items-center justify-center p-8 text-center bg-[#0a0a0a] text-white">
        <X size={64} className="text-red-500 mb-6" />
        <p className="text-white/60 mb-8">{error}</p>
        <button onClick={() => window.location.reload()} className="px-8 py-4 bg-primary rounded-2xl font-bold">Qayta urinish</button>
    </div>
);

const HomeView: React.FC<{ user: User | null, onStart: () => void }> = ({ user, onStart }) => (
  <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -20 }} className="flex-1 flex flex-col items-center justify-center p-8 space-y-12">
    <div className="w-32 h-32 bg-gradient-to-br from-[#a78bfa] to-[#5b21b6] rounded-3xl flex items-center justify-center shadow-[0_0_50px_rgba(124,58,237,0.3)]">
        <Rocket className="w-16 h-16 text-white" />
    </div>
    <div className="text-center space-y-4">
      <h2 className="text-3xl font-bold">Taqdimot tayyorlashni hoziroq boshlang</h2>
    </div>
    <button onClick={onStart} className="px-12 py-5 bg-[#7c3aed] rounded-2xl font-bold text-xl">Boshlash</button>
    <div className="text-sm text-white/40">Imkoniyatlaringiz: <span className="text-[#a78bfa] font-bold">{user?.available_generations || 0} ta</span></div>
  </motion.div>
);

const CreditsView: React.FC<{ user: User | null, onBack: () => void }> = ({ user, onBack }) => (
  <motion.div initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }} className="flex-1 flex flex-col p-6 overflow-y-auto pb-20">
    <div className="flex items-center space-x-3 mb-8"><button onClick={onBack} className="p-2 bg-white/5 rounded-full"><ChevronLeft className="w-6 h-6" /></button><h2 className="text-2xl font-bold">Imkoniyatlarim</h2></div>
    
    <div className="bg-gradient-to-br from-[#7c3aed] to-[#5b21b6] p-6 rounded-3xl mb-6 relative overflow-hidden">
        <div className="absolute top-0 right-0 p-4 opacity-10"><Award size={120} /></div>
        <p className="text-white/80">Mavjud imkoniyatlar</p>
        <h3 className="text-4xl font-black">{user?.available_generations || 0} ta</h3>
    </div>

    <div className="space-y-4">
        <div className="bg-[#171717] p-5 rounded-3xl flex items-start space-x-4">
            <div className="p-3 bg-blue-500/10 rounded-2xl text-blue-400"><UserPlus size={24} /></div>
            <div><h5 className="font-bold">Do‘stlarni taklif qiling</h5><p className="text-sm text-white/50">Har bir taklif uchun +1 ta imkoniyat.</p></div>
        </div>
        <div className="bg-[#171717] p-5 rounded-3xl flex items-start space-x-4">
            <div className="p-3 bg-amber-500/10 rounded-2xl text-amber-400"><Zap size={24} /></div>
            <div><h5 className="font-bold">Maxsus paketlar</h5><p className="text-sm text-white/50">Admin bilan bog‘lanib paket oling.</p></div>
        </div>
    </div>
    <button onClick={() => window.open('https://t.me/uzafo', '_blank')} className="w-full py-4 mt-6 bg-white/5 rounded-2xl font-bold">Admin bilan bog‘lanish</button>
  </motion.div>
);

const HowToView: React.FC<{ onBack: () => void }> = ({ onBack }) => (
  <motion.div initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }} className="flex-1 flex flex-col p-6 overflow-y-auto pb-20">
    <div className="flex items-center space-x-3 mb-8"><button onClick={onBack} className="p-2 bg-white/5 rounded-full"><ChevronLeft className="w-6 h-6" /></button><h2 className="text-2xl font-bold">Qo‘llanma</h2></div>
    <div className="space-y-4">
      {[ {n: '01', t: 'Mavzu va Muallif', d: 'Taqdimot mavzusi va muallif ismini kiriting.' }, {n: '02', t: 'Dizayn', d: '8 ta dizayndan birini tanlang.' }, {n: '03', t: 'Yaratish', d: 'Tizim avtomatik tayyorlaydi.' }, {n: '04', t: 'Fayl', d: 'Tayyor fayllar botga keladi.' }].map(s => (
        <div key={s.n} className="bg-[#171717] p-5 rounded-3xl"><h5 className="font-bold">{s.t}</h5><p className="text-sm text-white/50">{s.d}</p></div>
      ))}
    </div>
  </motion.div>
);

const StatusView: React.FC<{ jobId: string, onDone: () => void }> = ({ jobId, onDone }) => {
  const [progress, setProgress] = useState(0);
  const [step, setStep] = useState<string>('queued');
  useEffect(() => {
    const fetch = async () => {
        try { const d = await apiService.getStatus(jobId); setProgress(d.progress); setStep(d.step); if (d.status === 'completed') onDone(); } catch(e) {}
    };
    fetch();
    const interval = setInterval(fetch, 3000);
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
  const idx = steps.findIndex(s => s.id === step);
  return (
    <div className="flex-1 flex flex-col p-8 space-y-8">
        <div className="text-center"><div className="text-5xl font-black text-primary">{progress}%</div></div>
        <div className="h-3 w-full bg-[#171717] rounded-full overflow-hidden"><motion.div className="h-full bg-primary" animate={{ width: `${progress}%` }} /></div>
        <div className="flex-1 space-y-3">{steps.map((s, i) => (
            <div key={s.id} className={`p-3 rounded-2xl flex items-center space-x-4 ${i <= idx ? 'bg-[#7c3aed]/10' : 'bg-[#171717]'}`}>
                <div className={`w-8 h-8 rounded-full flex items-center justify-center ${i <= idx ? 'bg-primary' : 'bg-white/5'}`}>{i < idx ? <Check size={16}/> : i+1}</div>
                <span>{s.label}</span>
            </div>
        ))}</div>
        {progress >= 100 && <button onClick={() => window.location.href = 'https://t.me/slayd_pptxbot'} className="w-full py-4 bg-primary rounded-2xl font-bold">Botga qaytish</button>}
    </div>
  );
};

const SidebarItem: React.FC<{ icon: React.ReactNode, label: string, onClick: () => void, active?: boolean }> = ({ icon, label, onClick, active }) => (
    <button onClick={onClick} className={`w-full flex items-center space-x-4 p-4 rounded-xl ${active ? 'bg-[#7c3aed]/20 text-[#a78bfa]' : 'hover:bg-white/5'}`}>{icon}<span className="font-medium">{label}</span></button>
);

export default App;
