import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { translations } from '../translations';

interface UploadScreenProps {
  onFileSelected: (file: File, workbook: any) => void;
  onLoadMockData: () => void;
  theme: string;
  onToggleTheme: () => void;
  lang: 'vi' | 'en' | 'ko';
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
}

export const UploadScreen: React.FC<UploadScreenProps> = ({
  onFileSelected,
  onLoadMockData,
  theme,
  onToggleTheme,
  lang,
  setLang,
}) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const [showError, setShowError] = useState(false);
  const [errorText, setErrorText] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const t = translations[lang];

  const isValidExt = (name: string) => /\.(xlsx|xls)$/i.test(name);

  const triggerError = (msg: string) => {
    setErrorText(msg);
    setShowError(false);
    setTimeout(() => setShowError(true), 10);
    setTimeout(() => setShowError(false), 4000);
  };

  const processFile = (file: File) => {
    if (!isValidExt(file.name)) {
      triggerError(t.invalidFileError);
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        onFileSelected(file, workbook);
      } catch (err: any) {
        const msg = (err?.message || '').toLowerCase();
        if (msg.includes('password') || msg.includes('encrypt')) {
          triggerError(lang === 'vi' ? 'Tệp được bảo vệ bằng mật khẩu. Vui lòng bỏ bảo vệ rồi thử lại.' : lang === 'en' ? 'The file is password protected. Please unprotect and try again.' : '파일이 비밀번호로 보호되어 있습니다. 보호를 해제하고 다시 시도하십시오.');
        } else {
          triggerError(t.loadingError);
        }
      }
    };
    reader.onerror = () => triggerError(t.loadingError);
    reader.readAsArrayBuffer(file);
  };



  return (
    <>
      <header className="topbar" style={{ borderBottom: 'none', background: 'transparent', backdropFilter: 'none', position: 'absolute', top: 0, right: 0, left: 0, pointerEvents: 'none' }}>
        <div className="topbar-left" style={{ visibility: 'hidden' }}>
          <div className="brandmark">
            <div className="brandmark-dot">📊</div>
            <span className="brandmark-text">Marketing Insights</span>
          </div>
        </div>
        <div className="topbar-right" style={{ pointerEvents: 'auto', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <select value={lang} onChange={e => setLang(e.target.value as any)} className="lang-select">
            <option value="vi">Tiếng Việt</option>
            <option value="en">English</option>
            <option value="ko">한국어</option>
          </select>
          <button
            className="theme-toggle"
            onClick={onToggleTheme}
            title={theme === 'dark' ? 'Chuyển sang sáng' : 'Chuyển sang tối'}
          >
            {theme === 'dark' ? '☀️' : '🌙'}
          </button>
        </div>
      </header>

      <main className="upload-screen">
        <div className="hero">
          <h1>{t.mainTitleUpload}</h1>
          <div className="header-line"></div>
        </div>

        <div
          className={`dropzone ${isDragOver ? 'dragover' : ''} ${showError ? 'error-state' : ''}`}
          tabIndex={0}
          role="button"
          aria-label={t.selectExcelBtn}
          onClick={() => fileInputRef.current?.click()}
          onKeyDown={e => {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              fileInputRef.current?.click();
            }
          }}
          onDragEnter={e => { e.preventDefault(); setIsDragOver(true); }}
          onDragOver={e => { e.preventDefault(); setIsDragOver(true); }}
          onDragLeave={e => { e.preventDefault(); setIsDragOver(false); }}
          onDrop={e => {
            e.preventDefault();
            setIsDragOver(false);
            const file = e.dataTransfer.files[0];
            if (file) processFile(file);
          }}
        >
          <div className="dz-icon">📁</div>
          <h2>{t.dragDrop}</h2>
          <p className="dz-sub">{t.orClickBrowse}</p>

          <div className="upload-actions" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', justifyContent: 'center' }}>
            <button
              type="button"
              className="btn btn-primary btn-lg"
              onClick={e => { e.stopPropagation(); fileInputRef.current?.click(); }}
            >
              📂 {t.selectExcelBtn}
            </button>

          </div>

          <input
            type="file"
            ref={fileInputRef}
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={e => {
              const file = e.target.files?.[0];
              if (file) processFile(file);
              e.target.value = '';
            }}
          />

          <p className="dz-hint">{t.formatsHint} · {lang === 'vi' ? 'Xử lý hoàn toàn trên trình duyệt' : lang === 'en' ? 'Processed entirely in the browser' : '브라우저에서 완전히 처리됨'}</p>
          <div className={`error-msg ${showError ? 'show' : ''}`}>{errorText}</div>
        </div>


      </main>
    </>
  );
};
