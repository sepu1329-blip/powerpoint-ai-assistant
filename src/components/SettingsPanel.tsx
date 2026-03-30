import React, { useState } from 'react';
import type { AppSettings, GeminiModel, CustomPrompt } from '../types';
import { validateApiKey } from '../services/gemini';

interface SettingsPanelProps {
  settings: AppSettings;
  onSave: (updates: Partial<AppSettings>) => void;
  onBack: () => void;
}

const MODELS: { value: GeminiModel; label: string }[] = [
  { value: 'gemini-2.0-flash', label: 'Gemini 2.0 Flash (빠름)' },
  { value: 'gemini-2.5-pro-exp-03-25', label: 'Gemini 2.5 Pro (최고 성능)' },
  { value: 'gemini-1.5-pro', label: 'Gemini 1.5 Pro (안정)' },
  { value: 'gemini-1.5-flash', label: 'Gemini 1.5 Flash (경량)' },
];

export const SettingsPanel: React.FC<SettingsPanelProps> = ({
  settings,
  onSave,
  onBack,
}) => {
  const [apiKey, setApiKey] = useState(settings.apiKey || '');
  const [model, setModel] = useState<GeminiModel>(settings.model || 'gemini-2.0-flash');
  const [autoContext, setAutoContext] = useState(settings.autoContext ?? true);
  const [customPrompts, setCustomPrompts] = useState<CustomPrompt[]>(settings.customPrompts || []);
  
  const [newPromptName, setNewPromptName] = useState('');
  const [newPromptContent, setNewPromptContent] = useState('');

  const [isValidating, setIsValidating] = useState(false);
  const [validationStatus, setValidationStatus] = useState<'idle' | 'success' | 'error'>('idle');

  const handleValidate = async () => {
    if (!apiKey.trim()) return;
    setIsValidating(true);
    setValidationStatus('idle');
    try {
      const valid = await validateApiKey(apiKey.trim(), model);
      setValidationStatus(valid ? 'success' : 'error');
    } catch {
      setValidationStatus('error');
    } finally {
      setIsValidating(false);
    }
  };

  const addPrompt = () => {
    if (!newPromptName.trim() || !newPromptContent.trim()) return;
    const newPrompt: CustomPrompt = {
      id: Date.now().toString(),
      name: newPromptName.trim(),
      content: newPromptContent.trim()
    };
    setCustomPrompts([...customPrompts, newPrompt]);
    setNewPromptName('');
    setNewPromptContent('');
  };

  const deletePrompt = (id: string) => {
    setCustomPrompts(customPrompts.filter(p => p.id !== id));
  };

  const handleSave = () => {
    onSave({ apiKey: apiKey.trim(), model, autoContext, customPrompts });
    onBack();
  };

  return (
    <div className="settings-panel" style={{ backgroundColor: '#f7f9fc', height: '100%', display: 'flex', flexDirection: 'column' }}>


      <div className="settings-content" style={{ flex: 1, overflowY: 'auto', padding: '20px', display: 'flex', flexDirection: 'column', gap: '24px' }}>
        {/* Gemini API Key */}
        <div className="settings-group">
          <label className="settings-title" style={{ color: '#64748b', display: 'flex', alignItems: 'center', gap: '6px' }}>
            🔑 Gemini API Key
          </label>
          <input
            type="password"
            className="settings-input"
            style={{ height: '48px', borderRadius: '10px' }}
            value={apiKey}
            onChange={(e) => {
              setApiKey(e.target.value);
              setValidationStatus('idle');
            }}
            placeholder="AIza... (Google AI Studio)"
          />
          <button
            className="secondary-btn"
            onClick={handleValidate}
            disabled={!apiKey.trim() || isValidating}
            style={{ marginTop: '8px', fontSize: '12px', padding: '10px', borderRadius: '8px' }}
          >
            {isValidating ? '연결 테스트 중...' : 
             validationStatus === 'success' ? '✅ 연결 성공' : 
             validationStatus === 'error' ? '❌ 연결 실패' : '연결 테스트'}
          </button>
        </div>

        {/* AI 모델 */}
        <div className="settings-group">
          <label className="settings-title" style={{ color: '#64748b', display: 'flex', alignItems: 'center', gap: '6px' }}>
            🧠 모델 선택 (Model)
          </label>
          <select 
            className="styled-select" 
            value={model} 
            onChange={(e) => setModel(e.target.value as GeminiModel)}
            style={{ height: '48px', fontSize: '14px', background: 'white', borderRadius: '10px' }}
          >
            {MODELS.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
          </select>
        </div>

        <div style={{ height: '1px', background: '#e2e8f0', margin: '4px 0' }} />

        {/* 프롬프트 관리 섹션 */}
        <div className="settings-group">
          <label className="settings-title" style={{ color: '#64748b', display: 'flex', alignItems: 'center', gap: '6px' }}>
            📝 자주 쓰는 프롬프트 관리
          </label>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <input 
              className="settings-input" 
              style={{ background: 'white', border: '1px solid #e2e8f0', height: '48px', borderRadius: '10px' }}
              placeholder="프롬프트 이름" 
              value={newPromptName}
              onChange={e => setNewPromptName(e.target.value)}
            />
            <textarea 
              className="settings-input" 
              style={{ background: 'white', border: '1px solid #e2e8f0', height: '60px', resize: 'none', borderRadius: '10px' }}
              placeholder="프롬프트 내용" 
              value={newPromptContent}
              onChange={e => setNewPromptContent(e.target.value)}
            />
            <button 
              className="secondary-btn" 
              style={{ height: '48px', background: 'white', border: '1px solid #e2e8f0', color: '#1e293b', fontWeight: 600, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', borderRadius: '8px' }} 
              onClick={addPrompt}
            >
              <span style={{ fontSize: '18px', fontWeight: 700 }}>+</span> 추가
            </button>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '16px' }}>
            {customPrompts.length > 0 ? customPrompts.map(p => (
              <div key={p.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <span style={{ fontWeight: 700, fontSize: '14px', color: '#1e293b' }}>{p.name}</span>
                  <span style={{ fontSize: '12px', color: '#64748b' }}>{p.content}</span>
                </div>
                <button 
                  onClick={() => deletePrompt(p.id)}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '4px' }}
                >
                  <span style={{ fontSize: '20px' }}>❌</span>
                </button>
              </div>
            )) : (
              <div style={{ textAlign: 'center', padding: '20px', color: '#94a3b8', fontSize: '12px', background: 'white', borderRadius: '12px', border: '1px dashed #e2e8f0' }}>
                등록된 프롬프트가 없습니다.
              </div>
            )}
          </div>
        </div>
      </div>

      <div style={{ padding: '16px', background: 'transparent' }}>
        <button className="primary-btn" style={{ height: '52px', width: '100%', borderRadius: '12px', fontSize: '16px' }} onClick={handleSave}>
          저장 및 시작
        </button>
        <div style={{ textAlign: 'center', fontSize: '11px', color: '#94a3b8', marginTop: '12px' }}>
          Press AI v1.0.3 | PowerPoint Edition
        </div>
      </div>
    </div>
  );
};
