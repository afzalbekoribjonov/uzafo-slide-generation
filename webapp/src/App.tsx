import React, { useState, useEffect } from 'react';
import { AnimatePresence, motion } from 'framer-motion';
import { Menu, X, Rocket, HelpCircle, MessageSquare, ChevronLeft, ShieldAlert, Award, UserPlus, Zap } from 'lucide-react';
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
      console.log('App initialization started...');
      try {
        const tg = (window as any).Telegram?.WebApp;
        
        // Telegram initData borligini tekshiramiz
        if (tg?.initData && tg.initData !== "") {
          console.log('Telegram WebApp detected with data');
          tg.ready();
          tg.expand();
          
          try {
            const data = await apiService.init();
            setUser(data.user);
            setTemplates(data.templates);
          } catch (apiErr: any) {
            console.error('API Init failed:', apiErr);
            setError('Server bilan aloqa o‘rnatib bo‘lmadi.');
          }
          setLoading(false);
        } else {
          // Telegramdan tashqarida kirilgan
          console.warn('Outside Telegram environment detected.');
          
          // Agar localhost bo'lsa yoki 127.0.0.1 bo'lsa, test uchun ruxsat beramiz
          const isLocal = window.location.hostname === 'localhost' || 
                          window.location.hostname === '127.0.0.1' || 
                          window.location.hostname.startsWith('192.168.');

          if (isLocal) {
            console.log('Local environment detected, enabling mock mode');
            setTimeout(() => {
              setUser({
                id: 12345,
                full_name: 'Test Foydalanuvchi',
                available_generations: 5,
                is_blocked: false
              });
              setTemplates([
                { id: 'aurora_clean', name: 'Aurora Clean', description: '', preview_url: 'https://placehold.co/400x300/6d28d9/white?text=Aurora' },
                { id: 'midnight_neon', name: 'Midnight Neon', description: '', preview_url: 'https://placehold.co/400x300/4c1d95/white?text=Midnight' },
                { id: 'royal_boardroom', name: 'Royal Boardroom', description: '', preview_url: 'https://placehold.co/400x300/1e1b4b/white?text=Royal' },
              ]);
              setLoading(false);
            }, 800);
          } else {
            // Tashqi brauzerda taqiqlash sahifasini ko'rsatamiz
            setIsExternalBrowser(true);
            setLoading(false);
          }
        }
      } catch (err: any) {
        console.error('General Init Error:', err);
        setError('Kutilmagan xatolik yuz berdi: ' + (err instanceof Error ? err.message : String(err)));
        setLoading(false);
      }
    };

    init();
  }, []);

  const toggleSidebar = () => {
    setSidebarOpen(!sidebarOpen);
  };

  // Inline CSS for the loading state to ensure it works even if Tailwind fails
  if (loading) {
    return (
      <div style={{ 
        backgroundColor: '#0a0a0a', 
        color: 'white', 
        display: 'flex', 
        flexDirection: 'column', 
        alignItems: 'center', 
        justifyContent: 'center', 
        minHeight: '100vh', 
        width: '100%',
        position: 'fixed',
        top: 0,
        left: 0,
        fontFamily: 'sans-serif' 
      }}>
        <div style={{ 
          width: '40px', 
          height: '40px', 
          border: '4px solid #7c3aed', 
          borderTopColor: 'transparent', 
          borderRadius: '50%', 
          animation: 'spin 1s linear infinite' 
        }} />
        <p style={{ marginTop: '20px', opacity: 0.5 }}>Yuklanmoqda...</p>
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  if (isExternalBrowser) {
    return (
      <div style={{ 
        backgroundColor: '#0a0a0a', 
        color: 'white', 
        display: 'flex', 
        flexDirection: 'column', 
        alignItems: 'center', 
        justifyContent: 'center', 
        minHeight: '100vh', 
        padding: '40px',
        textAlign: 'center',
        fontFamily: 'sans-serif' 
      }}>
        <div style={{ 
          width: '80px', 
          height: '80px', 
          backgroundColor: 'rgba(239, 68, 68, 0.1)', 
          borderRadius: '50%', 
          display: 'flex', 
          alignItems: 'center', 
          justifyContent: 'center', 
          marginBottom: '30px',
          border: '1px solid rgba(239, 68, 68, 0.2)'
        }}>
          <ShieldAlert size={48} color="#ef4444" />
        </div>
        <h2 style={{ fontSize: '28px', fontWeight: 'bold', marginBottom: '15px' }}>Kirish taqiqlangan</h2>
        <p style={{ opacity: 0.6, fontSize: '18px', maxWidth: '400px', lineHeight: '1.6', marginBottom: '30px' }}>
          Ushbu xizmatdan faqat <b>Slide Generator Bot</b> ichidagi WebApp orqali foydalanish mumkin.
        </p>
        <button 
          onClick={() => window.open('https://t.me/uzafo_slide_bot', '_blank')}
          style={{ 
            padding: '16px 32px', 
            backgroundColor: '#7c3aed', 
            borderRadius: '16px', 
            color: 'white', 
            fontWeight: 'bold', 
            fontSize: '18px',
            border: 'none',
            boxShadow: '0 10px 25px rgba(124, 58, 237, 0.3)',
            cursor: 'pointer'
          }}
        >
          Botga o‘tish
        </button>
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ 
        backgroundColor: '#0a0a0a', 
        color: 'white', 
        display: 'flex', 
        flexDirection: 'column', 
        alignItems: 'center', 
        justifyContent: 'center', 
        minHeight: '100vh', 
        padding: '20px', 
        textAlign: 'center', 
        fontFamily: 'sans-serif' 
      }}>
        <X size={64} color="#ef4444" style={{ marginBottom: '20px' }} />
        <h2 style={{ fontSize: '24px', fontWeight: 'bold', marginBottom: '10px' }}>Xatolik</h2>
        <p style={{ opacity: 0.6, marginBottom: '30px' }}>{error}</p>
        <button 
          onClick={() => window.location.reload()}
          style={{ padding: '12px 24px', backgroundColor: '#7c3aed', border: 'none', borderRadius: '12px', color: 'white', fontWeight: 'bold', cursor: 'pointer' }}
        >
          Qayta urinish
        </button>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white font-sans overflow-hidden relative flex flex-col">
      {/* Header */}
      <header className="fixed top-0 left-0 right-0 h-16 flex items-center justify-between px-6 bg-[#0a0a0a]/80 backdrop-blur-md z-40 border-b border-white/5">
        <button onClick={toggleSidebar} className="p-2 hover:bg-white/5 rounded-full transition-colors">
          <Menu className="w-6 h-6" />
        </button>
        <h1 className="text-xl font-bold bg-gradient-to-r from-[#a78bfa] to-[#7c3aed] bg-clip-text text-transparent">
          Slide Generator
        </h1>
        <div className="w-10" />
      </header>

      {/* Sidebar Overlay */}
      <AnimatePresence>
        {sidebarOpen && (
          <>
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSidebarOpen(false)}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50"
            />
            <motion.div
              initial={{ x: '-100%' }}
              animate={{ x: 0 }}
              exit={{ x: '-100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="fixed top-0 left-0 bottom-0 w-4/5 max-w-sm bg-[#171717] z-50 border-r border-white/10 p-6 flex flex-col"
            >
              <div className="flex items-center justify-between mb-8">
                <span className="text-xl font-bold text-[#a78bfa]">Menyu</span>
                <button onClick={() => setSidebarOpen(false)} className="p-2 hover:bg-white/5 rounded-full transition-colors">
                  <X className="w-6 h-6" />
                </button>
              </div>

              <nav className="flex-1 space-y-4">
                <SidebarItem 
                  icon={<Rocket className="w-5 h-5" />} 
                  label="Imkoniyatlarim" 
                  onClick={() => { setView('credits'); setSidebarOpen(false); }}
                  active={view === 'credits'}
                />
                <SidebarItem 
                  icon={<HelpCircle className="w-5 h-5" />} 
                  label="Qo‘llanma" 
                  onClick={() => { setView('how-to'); setSidebarOpen(false); }}
                  active={view === 'how-to'}
                />
                <SidebarItem 
                  icon={<MessageSquare className="w-5 h-5" />} 
                  label="Savol yoki taklif?" 
                  onClick={() => { window.open('https://uzafo.site/en/discussions/free-slide-generator', '_blank'); setSidebarOpen(false); }}
                />
              </nav>

              <div className="pt-6 border-t border-white/5 mt-auto">
                <p className="text-xs text-white/40 text-center">Talqin 1.0.0</p>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Main Content */}
      <main className="pt-16 flex-1 flex flex-col relative">
        <AnimatePresence mode="wait">
          {view === 'home' && (
            <HomeView 
              key="home"
              user={user} 
              onStart={() => {
                if (user && user.available_generations > 0) {
                  setView('wizard');
                } else {
                  setView('credits');
                }
              }} 
            />
          )}
          {view === 'wizard' && (
            <WizardView 
              key="wizard"
              templates={templates} 
              onComplete={(jobId) => {
                setCurrentJobId(jobId);
                setView('status');
              }}
              onCancel={() => setView('home')}
            />
          )}
          {view === 'status' && (
            <StatusView key="status" jobId={currentJobId || ''} onDone={() => { setCurrentJobId(null); setView('home'); }} />
          )}
          {view === 'credits' && (
             <CreditsView 
              key="credits" 
              user={user}
              onBack={() => setView('home')} 
             />
          )}
          {view === 'how-to' && (
             <HowToView 
              key="how-to" 
              onBack={() => setView('home')} 
             />
          )}
        </AnimatePresence>
      </main>
    </div>
  );
};

const CreditsView: React.FC<{ user: User | null, onBack: () => void }> = ({ user, onBack }) => (
  <motion.div 
    initial={{ opacity: 0, x: 20 }}
    animate={{ opacity: 1, x: 0 }}
    exit={{ opacity: 0, x: -20 }}
    className="flex-1 flex flex-col p-6 overflow-y-auto pb-20"
  >
    <div className="flex items-center space-x-3 mb-8">
        <button onClick={onBack} className="p-2 bg-white/5 rounded-full">
            <ChevronLeft className="w-6 h-6" />
        </button>
        <h2 className="text-2xl font-bold">Imkoniyatlarim</h2>
    </div>

    <div className="space-y-6">
      {/* Current Balance Card */}
      <div className="bg-gradient-to-br from-[#7c3aed] to-[#5b21b6] p-6 rounded-3xl shadow-lg shadow-[#7c3aed]/20 relative overflow-hidden">
        <div className="absolute top-0 right-0 p-4 opacity-10">
          <Award size={120} />
        </div>
        <p className="text-white/80 font-medium mb-1">Mavjud imkoniyatlar</p>
        <h3 className="text-4xl font-black">{user?.available_generations || 0} ta</h3>
        <p className="text-xs text-white/60 mt-4">Har bir so‘rov uchun 1 ta imkoniyat sarflanadi</p>
      </div>

      <h4 className="text-lg font-bold px-2">Qanday qilib ko‘paytirish mumkin?</h4>
      
      {/* Referral Option */}
      <div className="bg-[#171717] border border-white/5 p-5 rounded-3xl flex items-start space-x-4">
        <div className="p-3 bg-blue-500/10 rounded-2xl text-blue-400">
          <UserPlus size={24} />
        </div>
        <div>
          <h5 className="font-bold mb-1">Do‘stlarni taklif qiling</h5>
          <p className="text-sm text-white/50 leading-relaxed">
            Har bir taklif qilingan va botdan ro‘yxatdan o‘tgan do‘stingiz uchun <b>+1 ta</b> imkoniyatga ega bo‘ling.
          </p>
        </div>
      </div>

      {/* Admin Option */}
      <div className="bg-[#171717] border border-white/5 p-5 rounded-3xl flex items-start space-x-4">
        <div className="p-3 bg-amber-500/10 rounded-2xl text-amber-400">
          <Zap size={24} />
        </div>
        <div>
          <h5 className="font-bold mb-1">Maxsus paketlar</h5>
          <p className="text-sm text-white/50 leading-relaxed">
            Ko‘p miqdorda taqdimotlar tayyorlash uchun admin bilan bog‘lanib, maxsus tariflarni faollashtiring.
          </p>
        </div>
      </div>

      <button 
        onClick={() => window.open('https://t.me/uzafo', '_blank')}
        className="w-full py-4 bg-white/5 hover:bg-white/10 border border-white/10 rounded-2xl font-bold transition-all text-white/80"
      >
        Admin bilan bog‘lanish
      </button>
    </div>
  </motion.div>
);

const HowToView: React.FC<{ onBack: () => void }> = ({ onBack }) => (
  <motion.div 
    initial={{ opacity: 0, x: 20 }}
    animate={{ opacity: 1, x: 0 }}
    exit={{ opacity: 0, x: -20 }}
    className="flex-1 flex flex-col p-6 overflow-y-auto pb-20"
  >
    <div className="flex items-center space-x-3 mb-8">
        <button onClick={onBack} className="p-2 bg-white/5 rounded-full">
            <ChevronLeft className="w-6 h-6" />
        </button>
        <h2 className="text-2xl font-bold">Qo‘llanma</h2>
    </div>

    <div className="space-y-4">
      <StepCard 
        number="01" 
        title="Mavzu va Muallif" 
        description="Taqdimot mavzusini aniq kiriting va slaydning muallif ismini yozing."
      />
      <StepCard 
        number="02" 
        title="Dizayn tanlash" 
        description="O‘zingizga ma’qul kelgan 8 ta dizayndan birini va slaydlar sonini tanlang."
      />
      <StepCard 
        number="03" 
        title="Avtomatik yaratish" 
        description="Tizim kiritilgan ma'lumotlar asosida taqdimot strukturasi va mazmunini avtomatik tayyorlaydi."
      />
      <StepCard 
        number="04" 
        title="Faylni olish" 
        description="Taqdimot tayyor bo‘lgach, bot sizga .pptx va .pdf formatdagi fayllarni yuboradi."
      />

      <div className="bg-primary/10 border border-primary/20 p-5 rounded-3xl mt-4">
        <p className="text-sm text-primary-light leading-relaxed text-center">
          <b>Maslahat:</b> Mavzu qanchalik aniq bo‘lsa (masalan: "O‘zbekiston iqtisodiyoti 2024"), slaydlar shunchalik mazmunli chiqadi.
        </p>
      </div>
    </div>
  </motion.div>
);

const StepCard: React.FC<{ number: string, title: string, description: string }> = ({ number, title, description }) => (
  <div className="bg-[#171717] border border-white/5 p-5 rounded-3xl relative overflow-hidden group">
    <div className="absolute -right-2 -top-2 text-6xl font-black text-white/5 group-hover:text-white/10 transition-colors">
      {number}
    </div>
    <h5 className="font-bold text-lg mb-2 flex items-center">
      <span className="w-8 h-8 bg-primary/20 text-primary-light rounded-lg flex items-center justify-center text-sm mr-3">
        {number}
      </span>
      {title}
    </h5>
    <p className="text-sm text-white/50 leading-relaxed pr-8">
      {description}
    </p>
  </div>
);

const StatusView: React.FC<{ jobId: string, onDone: () => void }> = ({ jobId, onDone }) => {
  const [status, setStatus] = useState({ progress: 0, step: 'pending' });

  useEffect(() => {
    const fetchStatus = async () => {
      try {
        const data = await apiService.getStatus(jobId);
        setStatus({ progress: data.progress, step: data.step });
        if (data.status === 'completed') {
            onDone();
        }
      } catch (err) {
        console.error('Status polling error', err);
      }
    };

    fetchStatus();
    const interval = setInterval(fetchStatus, 5000); // 5 soniyada bir yangilash
    return () => clearInterval(interval);
  }, [jobId, onDone]);

  const getStepText = (step: string) => {
    switch(step) {
      case 'queued': return 'So‘rovingiz navbatga qo‘shildi';
      case 'research': return 'Ma’lumotlar yig‘ilmoqda...';
      case 'planning': return 'Reja tuzilmoqda...';
      case 'rendering': return 'Taqdimot yig‘ilmoqda...';
      case 'uploading': return 'Fayl tayyorlanmoqda...';
      default: return 'Jarayon davom etmoqda...';
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, scale: 0.9 }}
      animate={{ opacity: 1, scale: 1 }}
      exit={{ opacity: 0, scale: 0.9 }}
      className="flex-1 flex flex-col items-center justify-center p-8 text-center space-y-8"
    >
      <div className="relative w-24 h-24">
          <motion.div 
              animate={{ rotate: 360 }}
              transition={{ repeat: Infinity, duration: 2, ease: "linear" }}
              className="absolute inset-0 border-4 border-[#7c3aed] border-t-transparent rounded-full"
          />
          <div className="absolute inset-0 flex items-center justify-center font-bold">
            {status.progress}%
          </div>
      </div>
      <div className="space-y-4">
          <h2 className="text-2xl font-bold">Taqdimot tayyorlanmoqda</h2>
          <p className="text-white/60">{getStepText(status.step)}</p>
      </div>
      <p className="text-white/40 text-sm max-w-[250px]">Jarayon yakunlangach, bot sizga faylni yuboradi. Siz sahifadan chiqishingiz mumkin.</p>
    </motion.div>
  );
};

const SidebarItem: React.FC<{ icon: React.ReactNode, label: string, onClick: () => void, active?: boolean }> = ({ icon, label, onClick, active }) => (
  <button 
    onClick={onClick}
    className={`w-full flex items-center space-x-4 p-4 rounded-xl transition-all ${active ? 'bg-[#7c3aed]/20 text-[#a78bfa] border border-[#7c3aed]/20' : 'hover:bg-white/5 text-white/70'}`}
  >
    {icon}
    <span className="font-medium">{label}</span>
  </button>
);

const HomeView: React.FC<{ user: User | null, onStart: () => void }> = ({ user, onStart }) => (
  <motion.div 
    initial={{ opacity: 0, y: 20 }}
    animate={{ opacity: 1, y: 0 }}
    exit={{ opacity: 0, y: -20 }}
    className="flex-1 flex flex-col items-center justify-center p-8 space-y-12"
  >
    <div className="relative">
      <div className="absolute inset-0 bg-[#7c3aed] blur-[100px] opacity-20 rounded-full" />
      <motion.div 
        animate={{ scale: [1, 1.05, 1] }}
        transition={{ repeat: Infinity, duration: 3 }}
        className="w-32 h-32 bg-gradient-to-br from-[#a78bfa] to-[#5b21b6] rounded-3xl flex items-center justify-center shadow-[0_0_50px_rgba(124,58,237,0.3)]"
      >
        <Rocket className="w-16 h-16 text-white" />
      </motion.div>
    </div>

    <div className="text-center space-y-4">
      <h2 className="text-3xl font-bold">Taqdimot tayyorlashni hoziroq boshlang</h2>
      <p className="text-white/60 max-w-xs mx-auto">
        Sun'iy intellekt yordamida bir necha soniya ichida professional slaydlarga ega bo'ling.
      </p>
    </div>

    <motion.button
      whileHover={{ scale: 1.05 }}
      whileTap={{ scale: 0.95 }}
      onClick={onStart}
      className="group relative px-12 py-5 bg-[#7c3aed] rounded-2xl font-bold text-xl shadow-lg shadow-[#7c3aed]/20 overflow-hidden"
    >
      <div className="absolute inset-0 bg-gradient-to-r from-white/0 via-white/20 to-white/0 -translate-x-full group-hover:translate-x-full transition-transform duration-1000" />
      <span className="relative">Boshlash</span>
    </motion.button>

    <div className="flex items-center space-x-2 text-sm text-white/40">
      <span>Imkoniyatlaringiz:</span>
      <span className="text-[#a78bfa] font-bold">{user?.available_generations || 0} ta</span>
    </div>
  </motion.div>
);

export default App;
