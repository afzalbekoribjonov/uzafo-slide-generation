import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { ChevronRight, Check, Layout, Type, Users, Hash, Globe, FileText, Send } from 'lucide-react';
import type { Template, CreatePresentationData } from '../types';
import { apiService } from '../services/api';

interface WizardViewProps {
  templates: Template[];
  onComplete: (jobId: string) => void;
  onCancel: () => void;
}

const steps = [
  { id: 'topic', label: 'Mavzu', icon: <Type className="w-5 h-5" /> },
  { id: 'presenter', label: 'Muallif', icon: <Users className="w-5 h-5" /> },
  { id: 'slides', label: 'Slaydlar', icon: <Hash className="w-5 h-5" /> },
  { id: 'template', label: 'Dizayn', icon: <Layout className="w-5 h-5" /> },
  { id: 'language', label: 'Til', icon: <Globe className="w-5 h-5" /> },
  { id: 'pdf', label: 'PDF', icon: <FileText className="w-5 h-5" /> },
];

export const WizardView: React.FC<WizardViewProps> = ({ templates, onComplete, onCancel }) => {
  const [currentStep, setCurrentStep] = useState(0);
  const [formData, setFormData] = useState<Partial<CreatePresentationData>>({
    slide_count: 8,
    language_code: 'uz',
    wants_pdf: true,
  });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const prevStep = () => {
    if (currentStep > 0) {
      setCurrentStep(currentStep - 1);
    } else {
      onCancel();
    }
  };

  useEffect(() => {
    const tg = (window as any).Telegram?.WebApp;
    if (tg?.BackButton) {
      if (currentStep > 0) {
        tg.BackButton.show();
        tg.BackButton.onClick(prevStep);
      } else {
        tg.BackButton.hide();
      }
    }
    return () => {
        const tg = (window as any).Telegram?.WebApp;
        if (tg?.BackButton) {
            tg.BackButton.offClick(prevStep);
        }
    }
  }, [currentStep]);

  const nextStep = () => {
    if (currentStep < steps.length - 1) {
      setCurrentStep(currentStep + 1);
      const tg = (window as any).Telegram?.WebApp;
      if (tg?.HapticFeedback) {
        tg.HapticFeedback.selectionChanged();
      }
    } else {
      handleSubmit();
    }
  };

  const updateData = (data: Partial<CreatePresentationData>) => {
    setFormData(prev => ({ ...prev, ...data }));
  };

  const handleSubmit = async () => {
    setSubmitting(true);
    setError(null);
    try {
      const tg = (window as any).Telegram?.WebApp;
      if (tg?.HapticFeedback) {
        tg.HapticFeedback.notificationOccurred('success');
      }
      const result = await apiService.create(formData as CreatePresentationData);
      onComplete(result.job_id);
    } catch (err: any) {
      setError(err.response?.data?.error || 'Xatolik yuz berdi. Iltimos qayta urinib ko‘ring.');
      const tg = (window as any).Telegram?.WebApp;
      if (tg?.HapticFeedback) {
        tg.HapticFeedback.notificationOccurred('error');
      }
    } finally {
      setSubmitting(false);
    }
  };

  const isStepValid = () => {
    switch (steps[currentStep].id) {
      case 'topic': return (formData.topic?.length || 0) >= 3;
      case 'presenter': return (formData.presenter_name?.length || 0) >= 2;
      case 'template': return !!formData.template_id;
      default: return true;
    }
  };

  return (
    <div className="flex-1 flex flex-col p-6 overflow-hidden">
      {/* Progress Bar */}
      <div className="flex justify-between items-center mb-8 px-2">
        {steps.map((_, idx) => (
          <div key={idx} className="flex flex-col items-center space-y-2">
            <div 
              className={`w-2.5 h-2.5 rounded-full transition-all duration-300 ${
                idx <= currentStep ? 'bg-primary scale-125' : 'bg-white/10'
              }`} 
            />
          </div>
        ))}
      </div>

      {/* Step Content */}
      <div className="flex-1 relative overflow-hidden">
        <AnimatePresence mode="wait">
          <motion.div
            key={currentStep}
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -20 }}
            className="h-full flex flex-col"
          >
            <div className="flex items-center space-x-3 mb-6">
              <div className="p-2 bg-primary/20 rounded-lg text-primary-light">
                {steps[currentStep].icon}
              </div>
              <h3 className="text-xl font-bold">{steps[currentStep].label}</h3>
            </div>

            {currentStep === 0 && (
              <div className="space-y-4">
                <p className="text-white/60">Taqdimot mavzusini kiriting:</p>
                <textarea
                  autoFocus
                  value={formData.topic || ''}
                  onChange={(e) => updateData({ topic: e.target.value })}
                  className="w-full h-32 bg-surface border border-white/10 rounded-2xl p-4 focus:border-primary focus:ring-1 focus:ring-primary outline-none transition-all resize-none text-lg"
                  placeholder="Masalan: Sun'iy intellekt tarixi va kelajagi"
                />
              </div>
            )}

            {currentStep === 1 && (
              <div className="space-y-4">
                <p className="text-white/60">Taqdimot muallifi ismini kiriting:</p>
                <input
                  autoFocus
                  type="text"
                  value={formData.presenter_name || ''}
                  onChange={(e) => updateData({ presenter_name: e.target.value })}
                  className="w-full bg-surface border border-white/10 rounded-2xl p-4 focus:border-primary focus:ring-1 focus:ring-primary outline-none transition-all text-lg"
                  placeholder="Ism Familiyangiz"
                />
              </div>
            )}

            {currentStep === 2 && (
              <div className="space-y-6">
                <p className="text-white/60">Slaydlar sonini tanlang:</p>
                <div className="flex flex-wrap gap-3">
                  {[6, 7, 8, 9, 10, 11, 12].map(num => (
                    <button
                      key={num}
                      onClick={() => updateData({ slide_count: num })}
                      className={`flex-1 min-w-[60px] py-4 rounded-xl font-bold text-lg border transition-all ${
                        formData.slide_count === num 
                        ? 'bg-primary border-primary text-white scale-105' 
                        : 'bg-surface border-white/10 text-white/60'
                      }`}
                    >
                      {num}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {currentStep === 3 && (
              <div className="space-y-4 flex-1 overflow-y-auto pr-2 pb-20">
                <p className="text-white/60">Dizayn shablonini tanlang (jami 8 ta):</p>
                <div className="grid grid-cols-2 gap-4">
                  {templates.map(tpl => (
                    <button
                      key={tpl.id}
                      onClick={() => updateData({ template_id: tpl.id })}
                      className={`group relative aspect-[4/3] rounded-2xl overflow-hidden border-2 transition-all ${
                        formData.template_id === tpl.id 
                        ? 'border-primary' 
                        : 'border-white/5 grayscale-[0.3] hover:grayscale-0'
                      }`}
                    >
                      <img 
                        src={tpl.preview_url} 
                        alt={tpl.name}
                        className="w-full h-full object-cover"
                      />
                      <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-3">
                        <span className="text-[10px] font-bold truncate text-white/90">{tpl.name}</span>
                      </div>
                      {formData.template_id === tpl.id && (
                        <div className="absolute top-2 right-2 bg-primary text-white p-1 rounded-full shadow-lg">
                          <Check className="w-3 h-3" />
                        </div>
                      )}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {currentStep === 4 && (
              <div className="space-y-4">
                <p className="text-white/60">Taqdimot tilini tanlang:</p>
                <div className="space-y-3">
                  {[
                    { id: 'uz', label: 'O‘zbek' },
                    { id: 'ru', label: 'Русский' },
                    { id: 'en', label: 'English' }
                  ].map(lang => (
                    <button
                      key={lang.id}
                      onClick={() => updateData({ language_code: lang.id })}
                      className={`w-full p-4 rounded-2xl border flex items-center justify-between transition-all ${
                        formData.language_code === lang.id 
                        ? 'bg-primary/20 border-primary text-primary-light' 
                        : 'bg-surface border-white/10 text-white/60'
                      }`}
                    >
                      <span className="font-bold text-lg">{lang.label}</span>
                      {formData.language_code === lang.id && <Check className="w-5 h-5" />}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {currentStep === 5 && (
              <div className="space-y-4">
                <p className="text-white/60">PDF formatida ham yuklansinmi?</p>
                <div className="flex gap-4">
                  {[
                    { val: true, label: 'Ha' },
                    { val: false, label: 'Yo‘q' }
                  ].map(opt => (
                    <button
                      key={opt.label}
                      onClick={() => updateData({ wants_pdf: opt.val })}
                      className={`flex-1 p-6 rounded-2xl border font-bold text-xl transition-all ${
                        formData.wants_pdf === opt.val 
                        ? 'bg-primary border-primary text-white' 
                        : 'bg-surface border-white/10 text-white/60'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </motion.div>
        </AnimatePresence>
      </div>

      {/* Footer Actions */}
      <div className="pt-6 mt-auto">
        {error && (
            <p className="text-red-500 text-sm mb-4 text-center">{error}</p>
        )}
        <button
          disabled={!isStepValid() || submitting}
          onClick={nextStep}
          className={`w-full py-5 rounded-2xl font-bold text-xl flex items-center justify-center space-x-3 transition-all ${
            isStepValid() && !submitting 
            ? 'bg-primary text-white shadow-lg shadow-primary/20' 
            : 'bg-white/5 text-white/20'
          }`}
        >
          {submitting ? (
            <motion.div
                animate={{ rotate: 360 }}
                transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
                className="w-6 h-6 border-2 border-white border-t-transparent rounded-full"
            />
          ) : (
            <>
              <span>{currentStep === steps.length - 1 ? 'Yuborish' : 'Keyingisi'}</span>
              {currentStep === steps.length - 1 ? <Send className="w-5 h-5" /> : <ChevronRight className="w-6 h-6" />}
            </>
          )}
        </button>
      </div>
    </div>
  );
};
