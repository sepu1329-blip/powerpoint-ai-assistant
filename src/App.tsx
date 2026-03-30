import React, { useState, useEffect } from 'react';
import { ChatPanel } from './components/ChatPanel';
import { SettingsPanel } from './components/SettingsPanel';
import { useSettings } from './hooks/useSettings';
import { useChat } from './hooks/useChat';
import type { ViewMode } from './types';
import './index.css';

function App() {
  const { settings, updateSettings, hasApiKey } = useSettings();
  const {
    messages,
    isLoading,
    statusText,
    error,
    context,
    sendUserMessage,
    markActionsApplied,
    clearChat,
    refreshContext,
    stopGeneration,
  } = useChat();
  
  const [view, setView] = useState<ViewMode>('chat');
  const [isOfficeReady, setIsOfficeReady] = useState(false);

  useEffect(() => {
    // Office.js 초기화
    if (typeof Office !== 'undefined') {
      Office.onReady((info) => {
        setIsOfficeReady(true);
      });
    } else {
      setIsOfficeReady(true);
    }

    if (!hasApiKey) {
      setView('settings');
    }
  }, [hasApiKey]);

  if (!isOfficeReady) {
    return (
      <div className="loading-screen">
        <p className="loading-text">PowerPoint AI 로딩 중...</p>
      </div>
    );
  }

  return (
    <div className="app" style={{ height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <header className="app-header">
        <div className="header-left">
          <img 
            src="/powerpoint-ai-assistant/assets/PressAI_Re.png" 
            alt="Press AI Logo" 
            className="header-logo-img" 
          />
          <span className="header-title-text">Press AI</span>
        </div>
        <div className="header-right">
          {view === 'chat' ? (
            <>
              <button className="header-btn" onClick={clearChat} title="대화 초기화">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"></path>
                </svg>
              </button>
              <button className="header-btn" onClick={() => setView('settings')} title="설정">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <circle cx="12" cy="12" r="3"></circle>
                  <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
                </svg>
              </button>
            </>
          ) : (
            <button className="header-btn" onClick={() => setView('chat')} title="홈으로">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"></path>
                <polyline points="9 22 9 12 15 12 15 22"></polyline>
              </svg>
            </button>
          )}
        </div>
      </header>

      <main className="app-main">
        {view === 'chat' ? (
          <ChatPanel 
            settings={settings} 
            messages={messages}
            isLoading={isLoading}
            statusText={statusText}
            error={error}
            context={context}
            sendUserMessage={sendUserMessage}
            markActionsApplied={markActionsApplied}
            refreshContext={refreshContext}
            stopGeneration={stopGeneration}
            clearChat={clearChat}
          />
        ) : (
          <SettingsPanel
            settings={settings}
            onSave={updateSettings}
            onBack={() => setView('chat')}
          />
        )}
      </main>
    </div>
  );
}

export default App;
