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
            
            // Check for active job
            try {
                const activeJob = await apiService.getActiveJob();
                if (activeJob && activeJob.job_id) {
                    setCurrentJobId(activeJob.job_id);
                    setView('status');
                }
            } catch (e) {
                console.error("No active job found or API error", e);
            }
          } catch (apiErr: any) {
            console.error('API Init failed:', apiErr);
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

const StatusView: React.FC<{ jobId: string, onDone: () => void }> = ({ jobId, onDone }) => {
  const [progress, setProgress] = useState(0);
  const [step, setStep] = useState<string>('queued');

  useEffect(() => {
    const fetchStatus = async () => {
      try {
        const data = await apiService.getStatus(jobId);
        setProgress(data.progress);
        setStep(data.step);
        
        if (data.status === 'completed') {
            onDone();
        }
      } catch (err) {
        console.error('Status polling error', err);
      }
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
    <motion.div 
      initial={{ opacity: 0, scale: 0.9 }}
      animate={{ opacity: 1, scale: 1 }}
      exit={{ opacity: 0, scale: 0.9 }}
      className="flex-1 flex flex-col p-8 space-y-8"
    >
      <div className="text-center space-y-2">
          <h2 className="text-2xl font-bold">Jarayon holati</h2>
          <div className="text-5xl font-black text-primary">{progress}%</div>
      </div>
      
      {/* Progress Bar */}
      <div className="h-3 w-full bg-[#171717] rounded-full overflow-hidden">
        <motion.div 
            className="h-full bg-primary"
            initial={{ width: 0 }}
            animate={{ width: `${progress}%` }}
            transition={{ duration: 0.5 }}
        />
      </div>

      {/* Steps List */}
      <div className="flex-1 space-y-3">
        {steps.map((s, i) => (
            <div key={s.id} className={`p-3 rounded-2xl flex items-center space-x-4 transition-all ${i <= currentStepIndex ? 'bg-[#7c3aed]/10 border border-[#7c3aed]/20' : 'bg-[#171717] border border-transparent'}`}>
                <div className={`w-8 h-8 rounded-full flex items-center justify-center ${i <= currentStepIndex ? 'bg-primary text-white' : 'bg-white/5 text-white/30'}`}>
                    {i < currentStepIndex ? <Check size={16} /> : <span>{i + 1}</span>}
                </div>
                <span className={`text-sm ${i <= currentStepIndex ? 'text-white font-medium' : 'text-white/30'}`}>{s.label}</span>
            </div>
        ))}
      </div>

      {progress >= 100 ? (
        <button 
            onClick={() => window.location.href = 'https://t.me/uzafo_slide_bot'}
            className="w-full py-4 bg-primary text-white rounded-2xl font-bold transition-all hover:bg-primary/90 shadow-lg shadow-primary/20"
        >
            Botga qaytish
        </button>
      ) : (
        <p className="text-white/40 text-center text-sm">Jarayon yakunlangach, bot sizga faylni yuboradi. Siz sahifadan chiqishingiz mumkin.</p>
      )}
    </motion.div>
  );
};
