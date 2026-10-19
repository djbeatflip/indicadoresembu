import React, { useState, useMemo, useEffect, useRef } from 'react';
import { UploadCloud, Users, CheckCircle, AlertTriangle, AlertOctagon, Award, FileText, Search, Filter, X, BarChart2, FolderOpen, LogIn, Layout, Settings, Shield, UserPlus, LogOut, ArrowUpDown, ChevronUp, ChevronDown, Trash2, Menu, ChevronLeft } from 'lucide-react';
import { BarChart, Bar, LineChart, Line, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LabelList, Legend } from 'recharts';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

// --- IMPORTAÇÕES DO FIREBASE ---
import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, sendPasswordResetEmail } from 'firebase/auth';
import { getFirestore, collection, doc, setDoc, getDoc, getDocs, deleteDoc, writeBatch } from 'firebase/firestore';

// --- CONFIGURAÇÃO DO FIREBASE ---
const firebaseConfig = {
  apiKey: "AIzaSyBwepUe-Cz4-7jlGE6LeYRxi75azTJYO7A",
  authDomain: "painel-siaps-embu.firebaseapp.com",
  projectId: "painel-siaps-embu",
  storageBucket: "painel-siaps-embu.firebasestorage.app",
  messagingSenderId: "512775097820",
  appId: "1:512775097820:web:ace3c91a288b46833c5319",
  measurementId: "G-E1GWDGC1R2"
};

// Inicialização Segura do Firebase (Evita o "Crash" por Duplicidade)
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
const secondaryApp = getApps().find(a => a.name === "Secondary") || initializeApp(firebaseConfig, "Secondary");

const auth = getAuth(app);
const db = getFirestore(app);
const secondaryAuth = getAuth(secondaryApp);

// --- TIPAGENS ---
interface ImportedFile { id: string; name: string; lastModified?: number; data: any[]; }
interface User { id: string; email: string; role: string; name: string; }

// --- UTILITÁRIOS MATEMÁTICOS ---
const monthOrder: Record<string, number> = { 'Janeiro':1, 'Fevereiro':2, 'Março':3, 'Abril':4, 'Maio':5, 'Junho':6, 'Julho':7, 'Agosto':8, 'Setembro':9, 'Outubro':10, 'Novembro':11, 'Dezembro':12 };
const S_COLORS: Record<string, string> = { 'Ótimo': '#3B82F6', 'Bom': '#10B981', 'Suficiente': '#F59E0B', 'Regular': '#EF4444' };
const S_WEIGHTS: Record<string, number> = { 'Regular': 1, 'Suficiente': 2, 'Bom': 3, 'Ótimo': 4 };
const W_STATUS: Record<number, string> = { 1: 'Regular', 2: 'Suficiente', 3: 'Bom', 4: 'Ótimo' };
const C_COLORS = ['#118DFF', '#10B981', '#F59E0B', '#8B5CF6', '#EF4444', '#14B8A6', '#E66C37'];

const parseNum = (v: any) => { if (typeof v === 'number' && !isNaN(v)) return v; if (typeof v !== 'string') return null; let s = v.trim().replace(/%/g, '').replace(/R\$/g, '').trim(); if (!s) return null; if (s.includes('.') && s.includes(',')) s = s.replace(/\./g, '').replace(',', '.'); else if (s.includes(',')) s = s.replace(',', '.'); const n = Number(s); return isNaN(n) ? null : n; };
const formatVal = (v: number) => Number.isInteger(v) ? v.toLocaleString('pt-BR') : v.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
const normTxt = (v: any) => String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
const normIne = (v: any) => String(v || '').replace(/\D/g, '').replace(/^0+/, '');
const normStat = (v: any): string | null => { const s = String(v || '').toLowerCase().trim(); if (!s) return null; if (s.includes('ótimo') || s.includes('otimo')) return 'Ótimo'; if (['bom', 'boa'].includes(s) || s.startsWith('bom ')) return 'Bom'; if (s === 'suficiente' || (s.includes('suficiente') && !s.includes('insuficiente'))) return 'Suficiente'; if (['regular', 'ruim', 'crítico', 'critico', 'insuficiente'].includes(s) || s.includes('regular')) return 'Regular'; return null; };
const extMes = (f: string) => { const u = f.toUpperCase().trim(); for (const m of Object.keys(monthOrder)) if (u.startsWith(m.toUpperCase().substring(0,3))) return m; return f.split('-')[0].replace(/[0-9]/g, '').trim() || f; };
const extInd = (f: string) => f.split('-').length > 1 ? f.split('-').slice(1).join('-').trim() : f.trim();

// --- MOTOR DE LEITURA DE PLANILHAS ---
const buildRows = (rows: any[][]) => {
  let hIdx = rows.findIndex(r => r.join(' ').toUpperCase().includes('CNES'));
  if (hIdx === -1) hIdx = 0;
  const h1 = rows[hIdx] || [], h2 = rows[hIdx + 1] || [];
  const cnesIdx = h1.findIndex(c => String(c).toUpperCase().includes('CNES'));
  const cnesVal = String(h2[cnesIdx] || '').trim();
  const isSub = h2.length > 0 && (!cnesVal || !cnesVal.replace(/\D/g, '') || cnesVal.toUpperCase().includes('CNES'));
  
  const headers: string[] = [];
  for (let c = 0; c < Math.max(h1.length, isSub ? h2.length : 0); c++) {
    const v1 = String(h1[c] || '').trim(), v2 = isSub ? String(h2[c] || '').trim() : '';
    let ch = (v2 && v2.toUpperCase() !== v1.toUpperCase()) ? (v1.length <= 15 && ['NOME', 'CÓDIGO', 'CODIGO'].includes(v2.toUpperCase()) ? `${v1} - ${v2}` : v2) : v1;
    headers.push(headers.includes(ch) && ch ? `${ch} (${c + 1})` : ch);
  }

  return rows.slice(isSub ? hIdx + 2 : hIdx + 1).filter(r => {
    const f = r.slice(0, 4).join(' ').toUpperCase();
    return f && !f.includes('TOTAL') && !f.includes('FONTE:') && !f.includes('LEGENDA');
  }).map(r => {
    const obj: any = {}; let has = false;
    headers.forEach((col, c) => { if (col && !col.toUpperCase().includes('EMPTY')) { obj[col] = r[c] ?? ''; if (String(obj[col]).trim() !== '') has = true; } });
    return has ? obj : null;
  }).filter(Boolean);
};

// ==========================================
// APLICAÇÃO PRINCIPAL (FRONT-END)
// ==========================================
export default function App() {
  const [isLoaded, setIsLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [activeModule, setActiveModule] = useState('dashboard');
  
  // ESTADO PARA RECOLHER/EXPANDIR O MENU LATERAL
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  
  const [files, setFiles] = useState<ImportedFile[]>([]);
  const [usersList, setUsersList] = useState<User[]>([]);
  const [sysTheme, setSysTheme] = useState('blue');

  // Monitora o Estado de Login via Firebase Auth
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        try {
          const userDoc = await getDoc(doc(db, 'users', user.uid));
          if (userDoc.exists()) {
            setCurrentUser({ id: user.uid, email: user.email!, ...userDoc.data() } as User);
          } else {
            setCurrentUser({ id: user.uid, email: user.email!, role: 'Administrador', name: 'Administrador' });
          }

          const setDocSnap = await getDoc(doc(db, 'settings', 'layout'));
          if (setDocSnap.exists()) setSysTheme(setDocSnap.data().themeColor);

          // Busca Arquivos da Nuvem
          const filesSnap = await getDocs(collection(db, 'dados_siaps'));
          const agrupados = new Map<string, ImportedFile>();
          filesSnap.docs.forEach(d => {
             const row = d.data();
             const fileName = row._arquivoOrigem + '.xlsx';
             if(!agrupados.has(fileName)) {
                agrupados.set(fileName, { id: fileName, name: fileName, data: [] });
             }
             agrupados.get(fileName)!.data.push(row);
          });
          setFiles(Array.from(agrupados.values()));

          const usersSnap = await getDocs(collection(db, 'users'));
          setUsersList(usersSnap.docs.map(d => ({ id: d.id, ...d.data() } as User)));
        } catch (e) {
          console.error("Erro ao carregar dados:", e);
        }
      } else {
        setCurrentUser(null);
      }
      setIsLoaded(true);
    });
    return () => unsubscribe();
  }, []);

  const themes: Record<string, string> = { blue: 'bg-blue-600', emerald: 'bg-emerald-600', violet: 'bg-violet-600', slate: 'bg-slate-700' };
  const bgClass = themes[sysTheme] || themes.blue;

  if (!isLoaded) return <div className="min-h-screen flex items-center justify-center bg-slate-100 text-slate-500 font-medium">Conectando ao Firebase Nuvem...</div>;

  if (!currentUser) return <LoginScreen bgClass={bgClass} />;

  const handleLogout = async () => {
    await signOut(auth);
    setFiles([]);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex overflow-hidden">
      
      {/* MENU LATERAL (COM TRANSIÇÃO E LARGURA DINÂMICA) */}
      <aside className={`${isSidebarOpen ? 'w-64' : 'w-20'} bg-slate-900 text-slate-300 flex flex-col shadow-2xl z-20 transition-all duration-300 ease-in-out shrink-0`}>
        
        {/* CABEÇALHO DO MENU */}
        <div className={`p-4 ${bgClass} bg-opacity-90 flex items-center h-20 border-b border-slate-800 transition-all ${isSidebarOpen ? 'justify-between' : 'justify-center'}`}>
          {isSidebarOpen ? (
            <>
              <div className="flex items-center space-x-3 overflow-hidden">
                <div className="bg-white/20 p-2 rounded-lg shrink-0"><BarChart2 className="h-6 w-6 text-white" /></div>
                <div className="truncate">
                  <h1 className="font-bold text-white leading-tight truncate">Indicadores APS</h1>
                  <p className="text-[10px] text-white/80 uppercase tracking-wider truncate">Embu das Artes</p>
                </div>
              </div>
              <button onClick={() => setIsSidebarOpen(false)} className="text-white/70 hover:text-white shrink-0 bg-black/10 p-1.5 rounded-lg hover:bg-black/20 transition-colors" title="Recolher Menu">
                <ChevronLeft className="h-5 w-5" />
              </button>
            </>
          ) : (
            <button onClick={() => setIsSidebarOpen(true)} className="text-white hover:text-white shrink-0 p-2 bg-white/20 rounded-lg shadow-sm transition-colors" title="Expandir Menu">
              <Menu className="h-6 w-6" />
            </button>
          )}
        </div>
        
        {/* ITENS DO MENU */}
        <div className="p-4 flex-1 space-y-1.5 overflow-y-auto overflow-x-hidden scrollbar-thin">
          {isSidebarOpen && <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-2">Módulos</p>}
          
          <button onClick={() => setActiveModule('dashboard')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'dashboard' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Dashboard SIAPS" : ""}>
            <Layout className="h-5 w-5 shrink-0" />
            {isSidebarOpen && <span className="whitespace-nowrap">Dashboard SIAPS</span>}
          </button>
          
          {(currentUser.role === 'Administrador' || currentUser.role === 'Coordenador') && (
            <button onClick={() => setActiveModule('import')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'import' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Importar Dados" : ""}>
              <UploadCloud className="h-5 w-5 shrink-0" />
              {isSidebarOpen && <span className="whitespace-nowrap">Importar Dados</span>}
            </button>
          )}

          {(currentUser.role === 'Administrador') && (
            <>
              {isSidebarOpen ? (
                 <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-6">Administração</p>
              ) : (
                 <div className="h-px bg-slate-800 my-4 mx-2"></div>
              )}
              
              <button onClick={() => setActiveModule('users')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'users' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Usuários" : ""}>
                <Users className="h-5 w-5 shrink-0" />
                {isSidebarOpen && <span className="whitespace-nowrap">Usuários</span>}
              </button>
              <button onClick={() => setActiveModule('access')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'access' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Acessos" : ""}>
                <Shield className="h-5 w-5 shrink-0" />
                {isSidebarOpen && <span className="whitespace-nowrap">Acessos</span>}
              </button>
              <button onClick={() => setActiveModule('settings')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'settings' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Configurações" : ""}>
                <Settings className="h-5 w-5 shrink-0" />
                {isSidebarOpen && <span className="whitespace-nowrap">Configurações</span>}
              </button>
            </>
          )}
        </div>

        {/* RODAPÉ DO MENU (USUÁRIO E SAIR) */}
        <div className="p-4 border-t border-slate-800 bg-slate-950">
          {isSidebarOpen ? (
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3 overflow-hidden">
                <div className={`h-8 w-8 rounded-full shrink-0 ${bgClass} flex items-center justify-center text-white font-bold text-xs`} title={currentUser.name}>
                  {currentUser.name.charAt(0)}
                </div>
                <div className="truncate w-28">
                  <p className="text-xs font-bold text-slate-200 truncate">{currentUser.name}</p>
                  <p className="text-[10px] text-slate-500 truncate">{currentUser.role}</p>
                </div>
              </div>
              <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400 bg-slate-900 rounded-lg transition-colors" title="Sair do Sistema">
                <LogOut className="h-4 w-4 shrink-0" />
              </button>
            </div>
          ) : (
            <div className="flex flex-col items-center space-y-4">
              <div className={`h-8 w-8 rounded-full shrink-0 ${bgClass} flex items-center justify-center text-white font-bold text-xs`} title={currentUser.name}>
                {currentUser.name.charAt(0)}
              </div>
              <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400 bg-slate-900 rounded-lg transition-colors" title="Sair do Sistema">
                <LogOut className="h-4 w-4 shrink-0" />
              </button>
            </div>
          )}
        </div>
      </aside>

      <main className="flex-1 overflow-y-auto p-6">
        {activeModule === 'dashboard' && <ModuleDashboard files={files} bgClass={bgClass} />}
        {activeModule === 'import' && <ModuleImport files={files} setFiles={setFiles} bgClass={bgClass} />}
        {activeModule === 'users' && <ModuleUsers usersList={usersList} setUsersList={setUsersList} bgClass={bgClass} />}
        {activeModule === 'access' && <ModuleAccess bgClass={bgClass} />}
        {activeModule === 'settings' && <ModuleSettings sysTheme={sysTheme} setSysTheme={setSysTheme} bgClass={bgClass} />}
      </main>
    </div>
  );
}

// ==========================================
// TELA DE LOGIN & RECUPERAÇÃO DE SENHA
// ==========================================
function LoginScreen({ bgClass }: any) {
  const [isRegister, setIsRegister] = useState(false);
  const [isForgot, setIsForgot] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [successMsg, setSuccessMsg] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    const frm = e.target as any;
    const email = frm.email.value;
    const pass = frm.pass.value;

    try {
      if (isRegister) {
        const nome = frm.nome.value;
        const userCred = await createUserWithEmailAndPassword(auth, email, pass);
        await setDoc(doc(db, 'users', userCred.user.uid), {
          email: email,
          name: nome,
          role: 'Administrador'
        });
      } else {
        await signInWithEmailAndPassword(auth, email, pass);
      }
    } catch (err: any) {
      setErrorMsg(err.message.includes('user-not-found') || err.message.includes('invalid-credential') ? 'Credenciais inválidas ou e-mail não cadastrado.' : err.message);
    }
  };

  const handleForgotPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    setSuccessMsg('');
    const frm = e.target as any;
    const email = frm.email.value;

    try {
      await sendPasswordResetEmail(auth, email);
      setSuccessMsg('E-mail de recuperação enviado! Verifique a sua caixa de entrada e spam.');
      setTimeout(() => { setIsForgot(false); setSuccessMsg(''); }, 5000);
    } catch (err: any) {
      setErrorMsg(err.message.includes('user-not-found') ? 'Este e-mail não está cadastrado.' : 'Erro ao tentar recuperar a senha.');
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-100">
      <div className="bg-white p-8 rounded-xl shadow-lg max-w-sm w-full border border-slate-200">
        <div className="flex justify-center mb-6"><div className={`p-4 rounded-full ${bgClass} bg-opacity-10`}><Shield className={`h-10 w-10 text-slate-800`} /></div></div>
        <h2 className="text-2xl font-bold text-center text-slate-800 mb-2 leading-tight">Indicadores APS<br/><span className="text-lg font-medium text-slate-500">Embu das Artes</span></h2>
        
        {isForgot ? (
          <>
            <p className="text-center text-slate-500 text-sm mb-6">Recuperar Palavra-passe</p>
            {errorMsg && <div className="bg-red-50 text-red-600 p-3 rounded-lg text-xs mb-4 text-center border border-red-200">{errorMsg}</div>}
            {successMsg && <div className="bg-emerald-50 text-emerald-600 p-3 rounded-lg text-xs mb-4 text-center border border-emerald-200">{successMsg}</div>}
            
            <form onSubmit={handleForgotPassword} className="space-y-4">
              <div><label className="block text-xs font-semibold text-slate-600 mb-1">E-mail Cadastrado</label><input name="email" type="email" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>
              <button type="submit" className={`w-full ${bgClass} hover:opacity-90 text-white font-bold py-2.5 rounded-lg transition-opacity flex justify-center items-center`}>
                Enviar link de recuperação
              </button>
            </form>
            
            <div className="mt-6 text-center">
              <button onClick={() => {setIsForgot(false); setErrorMsg(''); setSuccessMsg('');}} className="text-xs text-slate-500 hover:underline">
                Voltar ao Login
              </button>
            </div>
          </>
        ) : (
          <>
            <p className="text-center text-slate-500 text-sm mb-6">{isRegister ? 'Cadastro de Administrador Master' : 'Acesso ao Sistema'}</p>
            
            {errorMsg && <div className="bg-red-50 text-red-600 p-3 rounded-lg text-xs mb-4 text-center border border-red-200">{errorMsg}</div>}

            <form onSubmit={handleSubmit} className="space-y-4">
              {isRegister && (
                <div><label className="block text-xs font-semibold text-slate-600 mb-1">Nome Completo</label><input name="nome" type="text" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>
              )}
              <div><label className="block text-xs font-semibold text-slate-600 mb-1">E-mail</label><input name="email" type="email" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>
              <div>
                <div className="flex justify-between items-center mb-1">
                  <label className="block text-xs font-semibold text-slate-600">Senha (Mín. 6)</label>
                  {!isRegister && <button type="button" onClick={() => {setIsForgot(true); setErrorMsg('');}} className="text-[10px] text-blue-600 hover:underline">Esqueceu?</button>}
                </div>
                <input name="pass" type="password" required minLength={6} className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" />
              </div>
              
              <button type="submit" className={`w-full ${bgClass} hover:opacity-90 text-white font-bold py-2.5 rounded-lg transition-opacity flex justify-center items-center`}>
                {isRegister ? 'Criar Conta Master' : <><LogIn className="h-4 w-4 mr-2" /> Entrar no Sistema</>}
              </button>
            </form>

            <div className="mt-6 text-center">
              <button onClick={() => {setIsRegister(!isRegister); setErrorMsg('');}} className="text-xs text-blue-600 hover:underline">
                {isRegister ? 'Já tenho uma conta. Fazer Login' : 'Primeiro Acesso? Criar Conta Master'}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ==========================================
// MÓDULO: IMPORTAÇÃO PARA O FIRESTORE
// ==========================================
function ModuleImport({ files, setFiles, bgClass }: any) {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (selectedFiles.length === 0) return;
    
    setIsProcessing(true);
    const newFiles: ImportedFile[] = [];

    for (const file of selectedFiles) {
      if (file.name.startsWith('~$') || file.name.startsWith('.')) continue;
      const ext = file.name.split('.').pop()?.toLowerCase();
      if (!['csv', 'xlsx', 'xls'].includes(ext || '')) continue;
      
      setProgressMsg(`Lendo ${file.name}...`);

      const fileData = await new Promise<any[]>((resolve) => {
        if (ext === 'csv') {
          Papa.parse(file, { header: false, skipEmptyLines: true, complete: (res) => resolve(buildRows(res.data as any[][])) });
        } else {
          const reader = new FileReader();
          reader.onload = (e) => {
            try {
              const wb = XLSX.read(e.target?.result, { type: 'binary' });
              resolve(buildRows(XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' })));
            } catch { resolve([]); }
          };
          reader.readAsBinaryString(file);
        }
      });

      if (fileData.length > 0) {
        const cleanName = file.name.replace(/\.[^/.]+$/, "");
        const properMonth = extMes(cleanName);
        const properIndicator = extInd(cleanName);
        
        const enrichedData = fileData.map(r => ({ ...r, _arquivoOrigem: cleanName, _mesOrigem: properMonth, _indicadorOrigem: properIndicator }));
        newFiles.push({ id: Date.now() + '-' + Math.random().toString(36).substring(2,7), name: file.name, data: enrichedData });
        
        setProgressMsg(`Salvando ${file.name} na Nuvem...`);
        try {
          let batch = writeBatch(db);
          let opCount = 0;
          
          for (let i = 0; i < enrichedData.length; i++) {
             const row = enrichedData[i];
             const rowId = `${cleanName}_${i}_${Date.now()}`;
             const docRef = doc(db, 'dados_siaps', rowId);
             batch.set(docRef, row);
             opCount++;
             
             if (opCount >= 400 || i === enrichedData.length - 1) {
                await batch.commit();
                batch = writeBatch(db);
                opCount = 0;
             }
          }
        } catch (err: any) {
           console.error("Erro ao salvar arquivo:", err);
           alert(`Erro ao salvar ${file.name}. Verifique as regras do Firestore.`);
        }
      }
    }

    setFiles((prev: ImportedFile[]) => {
      const map = new Map<string, ImportedFile>();
      prev.forEach(f => map.set(f.name, f));
      newFiles.forEach(f => map.set(f.name, f));
      return Array.from(map.values());
    });
    
    setIsProcessing(false);
    setProgressMsg('');
    if (fileInputRef.current) fileInputRef.current.value = '';
    alert('Upload concluído! Dados salvos e disponíveis para todos.');
  };

  const handleClearCloud = async () => {
    if(!window.confirm('CUIDADO: Isso apagará TODOS os relatórios do banco de dados na nuvem para todos os usuários! Tem certeza?')) return;
    setIsProcessing(true);
    setProgressMsg('Apagando nuvem...');
    
    try {
      const qs = await getDocs(collection(db, 'dados_siaps'));
      let batch = writeBatch(db);
      let opCount = 0;
      
      for(let i = 0; i < qs.docs.length; i++) {
         batch.delete(qs.docs[i].ref);
         opCount++;
         if(opCount >= 400 || i === qs.docs.length - 1) {
            await batch.commit();
            batch = writeBatch(db);
            opCount = 0;
         }
      }
      setFiles([]);
    } catch(e) {
      console.error(e);
      alert('Erro ao apagar. Verifique regras do Firestore.');
    }
    
    setIsProcessing(false);
    setProgressMsg('');
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className={`p-10 rounded-2xl border-2 border-dashed text-center bg-white ${isProcessing ? 'border-slate-300 opacity-50' : 'border-slate-300 hover:border-slate-400'}`}>
        <div className={`mx-auto w-16 h-16 ${bgClass} bg-opacity-10 rounded-full flex items-center justify-center mb-4`}><UploadCloud className={`h-8 w-8 text-slate-700`} /></div>
        <h3 className="font-bold text-xl text-slate-800">Conexão de Dados na Nuvem</h3>
        <p className="text-slate-500 mb-6 mt-2 max-w-lg mx-auto">Selecione as planilhas do SIAPS do seu computador. O sistema vai fatiar as planilhas e enviar para o banco de dados Firebase.</p>
        
        {isProcessing && <p className="text-blue-600 font-bold mb-4 animate-pulse">{progressMsg}</p>}

        <div className="flex justify-center gap-4">
          <label className={`cursor-pointer inline-flex items-center px-6 py-3 rounded-lg text-white font-bold shadow-md transition-all hover:opacity-90 ${bgClass}`}>
            <FolderOpen className="h-5 w-5 mr-2"/> {isProcessing ? 'Processando...' : 'Selecionar e Enviar para Nuvem'}
            <input type="file" multiple accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFileUpload} disabled={isProcessing} ref={fileInputRef} />
          </label>
        </div>
      </div>

      {files.length > 0 && (
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center mb-4">
            <h3 className="font-bold text-slate-800">Planilhas Hospedadas ({files.length})</h3>
            <button onClick={handleClearCloud} disabled={isProcessing} className="text-red-500 text-sm hover:bg-red-50 px-3 py-1.5 rounded-lg transition-colors flex items-center"><Trash2 className="h-4 w-4 mr-1"/> Apagar TUDO da Nuvem</button>
          </div>
          <div className="flex flex-wrap gap-2">
            {files.map((f: ImportedFile) => (
              <div key={f.id} className="bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-medium text-slate-700 flex items-center">
                <FileText className="h-3.5 w-3.5 mr-1.5 text-blue-500"/>{f.name}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ==========================================
// SUB-MÓDULOS DE ADMINISTRAÇÃO
// ==========================================
function ModuleUsers({ usersList, setUsersList, bgClass }: any) {
  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault(); 
    const frm = e.target as any;
    const email = frm.email.value;
    const pass = frm.pass.value;
    const name = frm.nome.value;
    const role = frm.role.value;

    try {
      const userCred = await createUserWithEmailAndPassword(secondaryAuth, email, pass);
      const newUid = userCred.user.uid;
      const newUser = { email, name, role };
      await setDoc(doc(db, 'users', newUid), newUser);
      
      setUsersList([...usersList, { id: newUid, ...newUser }]);
      frm.reset();
      alert('Usuário cadastrado com sucesso!');
    } catch (err: any) {
      alert('Erro ao cadastrar: ' + err.message);
    }
  };

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      <div className="md:col-span-1 bg-white p-5 rounded-xl border border-slate-200 shadow-sm h-fit">
        <h3 className="font-bold text-slate-800 mb-4 flex items-center"><UserPlus className="h-5 w-5 mr-2 text-blue-600"/> Novo Usuário</h3>
        <form onSubmit={handleAdd} className="space-y-3">
          <input name="nome" placeholder="Nome Completo" required className="w-full p-2 border rounded-lg outline-none text-sm"/>
          <input name="email" type="email" placeholder="E-mail de Login" required className="w-full p-2 border rounded-lg outline-none text-sm"/>
          <input name="pass" type="password" placeholder="Senha (Mín. 6)" required minLength={6} className="w-full p-2 border rounded-lg outline-none text-sm"/>
          <select name="role" className="w-full p-2 border rounded-lg outline-none text-sm bg-white"><option>Administrador</option><option>Coordenador</option><option>Leitor</option></select>
          <button type="submit" className={`w-full ${bgClass} text-white py-2 rounded-lg font-bold text-sm hover:opacity-90`}>Salvar na Nuvem</button>
        </form>
      </div>
      <div className="md:col-span-2 bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
        <h3 className="font-bold text-slate-800 mb-4">Usuários do Sistema</h3>
        <table className="w-full text-left text-sm"><thead className="bg-slate-100"><tr><th className="p-2">Nome</th><th className="p-2">E-mail</th><th className="p-2">Perfil</th></tr></thead>
        <tbody>
          {usersList.map((u:any) => <tr key={u.id} className="border-b"><td className="p-2 font-medium">{u.name}</td><td className="p-2 text-slate-500">{u.email}</td><td className="p-2"><span className={`px-2 py-1 rounded text-xs ${u.role==='Administrador'?'bg-red-100 text-red-700':u.role==='Coordenador'?'bg-amber-100 text-amber-700':'bg-emerald-100 text-emerald-700'}`}>{u.role}</span></td></tr>)}
        </tbody></table>
      </div>
    </div>
  );
}

function ModuleAccess({ bgClass }: any) {
  return (
    <div className="max-w-4xl bg-white rounded-xl shadow-sm border border-slate-200 p-6">
      <h2 className="text-lg font-bold flex items-center mb-6"><Shield className="mr-2 h-5 w-5 text-blue-600"/> Tabela de Permissões</h2>
      <div className="space-y-4">
        <div className="p-4 border rounded-lg bg-slate-50"><h3 className="font-bold text-red-700">Administrador</h3><p className="text-slate-600 text-sm">Acesso total. Gestão de usuários, layout, importação e Dashboard.</p></div>
        <div className="p-4 border rounded-lg bg-slate-50"><h3 className="font-bold text-amber-700">Coordenador</h3><p className="text-slate-600 text-sm">Apenas importação de relatórios para a nuvem e visualização do Dashboard.</p></div>
        <div className="p-4 border rounded-lg bg-slate-50"><h3 className="font-bold text-emerald-700">Leitor</h3><p className="text-slate-600 text-sm">Visualização exclusiva do Dashboard. Não pode editar nem enviar arquivos.</p></div>
      </div>
    </div>
  );
}

function ModuleSettings({ sysTheme, setSysTheme, bgClass }: any) {
  const handleSaveTheme = async (theme: string) => {
    setSysTheme(theme);
    await setDoc(doc(db, 'settings', 'layout'), { themeColor: theme });
  };
  return (
    <div className="max-w-2xl bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div className={`p-5 ${bgClass} text-white`}><h2 className="text-lg font-bold flex items-center"><Settings className="mr-2 h-5 w-5"/> Customizar Interface</h2></div>
      <div className="p-6">
        <h3 className="text-sm font-bold text-slate-800 mb-3">Cor Principal</h3>
        <div className="flex space-x-3">
          {['blue', 'emerald', 'violet', 'slate'].map(c => (
            <button key={c} onClick={() => handleSaveTheme(c)} className={`h-12 w-12 rounded-full border-4 transition-all ${sysTheme === c ? 'border-slate-800 scale-110 shadow-md' : 'border-transparent hover:scale-105'}`} style={{ backgroundColor: c==='blue'?'#2563eb':c==='emerald'?'#059669':c==='violet'?'#7c3aed':'#334155' }} />
          ))}
        </div>
      </div>
    </div>
  );
}

// ==========================================
// SUB-MÓDULO: DASHBOARD SIAPS (O VISUAL)
// ==========================================
function ModuleDashboard({ files, bgClass }: any) {
  // DECLARA TODOS OS HOOKS AQUI EM CIMA PARA O REACT NÃO TRAVAR
  const rawAllData = useMemo(() => files.flatMap((f: any) => f.data), [files]);
  const columns = useMemo(() => Array.from(new Set(rawAllData.flatMap(Object.keys))).filter(c => !c.startsWith('_') && !c.toUpperCase().includes('EMPTY') && c.trim() !== ''), [rawAllData]);

  const colUnidade = useMemo(() => columns.slice(0, 8).find(c => ['estabeleci', 'unidade', 'ubs', 'nome'].some(w => c.toLowerCase().includes(w)) && !c.toLowerCase().includes('cnes')) || null, [columns]);
  const colIne = useMemo(() => columns.slice(0, 8).find(c => ['ine', 'código ine', 'codigo ine'].includes(c.toLowerCase().trim()) || c.toLowerCase().endsWith(' ine')) || null, [columns]);
  const colEquipe = useMemo(() => {
    const blocks = ['nota', 'pontua', 'resultado', 'meta', 'total', 'perc', 'média', 'media', 'pondera', 'avalia', 'propor', 'razão', 'razao', 'numerador', 'denominador', 'cadastro', 'atendimento', 'cuidado', 'pessoa', 'gesta', 'diabet', 'hiperten', 'vacina', 'exame', 'consulta', 'visita', 'procedimento', 'classifica', 'conceito', 'qualidade', 'status', 'desempenho', 'ter pelo menos', 'realizar', 'solicita'];
    const candidates = columns.slice(0, 7).filter(c => c !== colUnidade && c !== colIne && !blocks.some(w => c.toLowerCase().includes(w)) && ['equipe', 'nome', 'identifica', 'descri'].some(w => c.toLowerCase().includes(w)) && !c.toLowerCase().includes('cnes') && !c.toLowerCase().includes('ibge'));
    return candidates.find(c => c.toLowerCase().includes('nome')) || candidates.find(c => !c.toLowerCase().includes('tipo') && !c.toLowerCase().includes('sigla')) || candidates[0] || null;
  }, [columns, colUnidade, colIne]);
  const colMes = '_mesOrigem';
  const colIndicador = useMemo(() => columns.slice(0, 8).find(c => ['indicador', 'condição', 'condicao'].includes(c.toLowerCase())) || '_indicadorOrigem', [columns]);

  const allData = useMemo(() => {
    const groups = new Map<string, any[]>();
    rawAllData.forEach(row => {
      const ubsKey = colUnidade ? normTxt(row[colUnidade]) : 'UBS';
      const gKey = `${row._arquivoOrigem || 'ARQ'}__${ubsKey}`;
      if (!groups.has(gKey)) groups.set(gKey, []);
      groups.get(gKey)!.push(row);
    });

    const enriched: any[] = [];
    groups.forEach(rows => {
      const sorted = [...rows].sort((a, b) => {
        const eqA = colEquipe ? normTxt(a[colEquipe]) : '';
        const eqB = colEquipe ? normTxt(b[colEquipe]) : '';
        if (eqA !== eqB) return eqA.localeCompare(eqB);
        return (colIne ? normIne(a[colIne]) : '').localeCompare(colIne ? normIne(b[colIne]) : '');
      });

      const cleanNames = sorted.map((r, idx) => {
        let n = colEquipe ? String(r[colEquipe] || '').replace(/\s+/g, ' ').trim() : '';
        n = n.replace(/\(\s*0*\d+\s*\)/g, '').trim();
        const ic = colIne ? normIne(r[colIne]) : '';
        if (n) return n.toUpperCase();
        return ic ? `EQUIPE INE ${ic}` : `EQUIPE ${idx + 1}`;
      });

      const nCnt: Record<string, number> = {}; cleanNames.forEach(n => nCnt[n] = (nCnt[n] || 0) + 1);
      const nOcc: Record<string, number> = {};

      sorted.forEach((r, idx) => {
        const ubs = colUnidade ? String(r[colUnidade] || 'Sem Unidade').replace(/\s+/g, ' ').trim().toUpperCase() : 'GERAL';
        let lbl = cleanNames[idx];
        if (nCnt[lbl] > 1) { nOcc[lbl] = (nOcc[lbl] || 0) + 1; lbl = `${lbl} (${nOcc[lbl]})`; }
        enriched.push({ ...r, _teamKey: `UBS_${normTxt(ubs)}__EQ_${normTxt(lbl)}`, _teamLabel: lbl, _ubsLabel: ubs });
      });
    });
    return enriched;
  }, [rawAllData, colUnidade, colEquipe, colIne]);

  const [selectedUnidade, setSelectedUnidade] = useState('');
  const [selectedEquipe, setSelectedEquipe] = useState('');
  const [selectedMes, setSelectedMes] = useState('');
  const [selectedIndicador, setSelectedIndicador] = useState('');
  const [selectedVariavel, setSelectedVariavel] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState<{ key: string, direction: 'asc' | 'desc' } | null>(null);
  const [tableMode, setTableMode] = useState<'consolidado' | 'historico'>('consolidado');

  const filteredData = useMemo(() => {
    let result = allData;
    if (selectedUnidade) result = result.filter(d => d._ubsLabel === selectedUnidade);
    if (selectedEquipe) result = result.filter(d => d._teamLabel === selectedEquipe);
    if (selectedMes) result = result.filter(d => String(d[colMes]) === selectedMes);
    if (selectedIndicador) result = result.filter(d => String(d[colIndicador]) === selectedIndicador);
    if (searchTerm) {
      const st = searchTerm.toLowerCase();
      result = result.filter(row => columns.some(col => String(row[col] || '').toLowerCase().includes(st)));
    }
    return result;
  }, [allData, columns, selectedUnidade, selectedEquipe, selectedMes, selectedIndicador, searchTerm, colMes, colIndicador]);

  const filteredDataForTrend = useMemo(() => {
    let result = allData;
    if (selectedUnidade) result = result.filter(d => d._ubsLabel === selectedUnidade);
    if (selectedEquipe) result = result.filter(d => d._teamLabel === selectedEquipe);
    if (selectedIndicador) result = result.filter(d => String(d[colIndicador]) === selectedIndicador);
    return result;
  }, [allData, selectedUnidade, selectedEquipe, selectedIndicador, colIndicador]);

  const variableColumns = useMemo(() => {
    const sampleRows = filteredData.length > 0 ? filteredData : allData;
    if (sampleRows.length === 0) return [];
    const blocks = ['cnes', 'ibge', 'ine', 'ano', 'mês', 'mes', 'competência', 'competencia', 'período', 'periodo', 'uf', 'município', 'municipio', 'cidade'];
    return columns.filter(col => {
      if ([colUnidade, colEquipe, colIne, colMes].includes(col)) return false;
      const lower = col.toLowerCase().trim();
      if (blocks.includes(lower) || lower.includes('cnes') || lower.includes('ibge')) return false;
      if ((lower.includes('código') || lower.includes('codigo') || lower.includes('id ')) && lower.length < 15) return false;
      return sampleRows.slice(0, 30).some(r => parseNum(r[col]) !== null);
    });
  }, [columns, filteredData, allData, colUnidade, colEquipe, colIne, colMes]);

  const activeVariable = useMemo(() => {
    if (selectedVariavel && variableColumns.includes(selectedVariavel)) return selectedVariavel;
    const scoreCol = variableColumns.find(c => ['nota', 'pontua', 'resultado', 'indicador', 'perc', 'total'].some(w => c.toLowerCase().includes(w)));
    return scoreCol || variableColumns[variableColumns.length - 1] || '';
  }, [selectedVariavel, variableColumns]);

  const consolidatedTeams = useMemo(() => {
    if (filteredData.length === 0) return [];
    const statusCol = columns.find(c => ['classifica', 'conceito', 'qualidade', 'status', 'desempenho'].some(w => c.toLowerCase().includes(w)));
    const teamMap = new Map<string, any>();
    
    filteredData.forEach(row => {
      const key = row._teamKey;
      if (!teamMap.has(key)) teamMap.set(key, { key, uni: row._ubsLabel, eq: row._teamLabel, rows: [] });
      teamMap.get(key).rows.push(row);
    });

    return Array.from(teamMap.values()).map(group => {
      const varValues: Record<string, number | null> = {};
      variableColumns.forEach(c => {
        let sum = 0, cnt = 0;
        group.rows.forEach((r: any) => { const n = parseNum(r[c]); if (n !== null) { sum += n; cnt++; } });
        varValues[c] = cnt > 0 ? sum / cnt : null;
      });

      let weightSum = 0, weightCount = 0;
      group.rows.forEach((r: any) => { 
        let st = statusCol ? normStat(r[statusCol]) : null;
        if (!st) for (let i = columns.length - 1; i >= 0; i--) if ((st = normStat(r[columns[i]]))) break;
        if (!st && activeVariable) {
          const v = parseNum(r[activeVariable]);
          if (v !== null) {
            const pct = v <= 10 && v > 0 ? v * 10 : v;
            st = pct > 75 ? 'Ótimo' : pct > 50 ? 'Bom' : pct > 25 ? 'Suficiente' : 'Regular';
          }
        }
        if (st) { weightSum += S_WEIGHTS[st]; weightCount++; } 
      });
      
      const avgWeight = weightCount > 0 ? Math.round(weightSum / weightCount) : null;
      const finalStatus = avgWeight ? W_STATUS[Math.min(4, Math.max(1, avgWeight))] : null;

      const baseRow = { ...group.rows[group.rows.length - 1] };
      variableColumns.forEach(c => { if (varValues[c] !== null) baseRow[c] = formatVal(Number(varValues[c]!.toFixed(2))); });
      if (statusCol && finalStatus) baseRow[statusCol] = finalStatus;

      const uniqueMonths = Array.from(new Set(group.rows.map((r: any) => r._mesOrigem)));
      return { key: group.key, unidade: group.uni, equipe: group.eq, varValues, status: finalStatus, monthsLabel: uniqueMonths.join(', '), consolidatedRow: baseRow };
    });
  }, [filteredData, columns, variableColumns, activeVariable]);

  const isSingleTeam = Boolean(selectedEquipe) || consolidatedTeams.length === 1;
  const totalEquipesUnicas = consolidatedTeams.length;

  const requestSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') direction = 'desc';
    setSortConfig({ key, direction });
  };

  const sortedTableData = useMemo(() => {
    let sortableData = tableMode === 'consolidado' ? consolidatedTeams.map(t => ({ ...t.consolidatedRow, _ms: t.monthsLabel, _eq: t.equipe })) : filteredData.map(r => ({ ...r, _ms: r._mesOrigem, _eq: r._teamLabel }));
    if (sortConfig !== null) {
      sortableData.sort((a, b) => {
        const valA = a[sortConfig.key], valB = b[sortConfig.key];
        const stA = normStat(valA), stB = normStat(valB);
        if (stA && stB) {
          const rankA = S_WEIGHTS[stA] || 0, rankB = S_WEIGHTS[stB] || 0;
          if (rankA !== rankB) return sortConfig.direction === 'asc' ? rankA - rankB : rankB - rankA;
        }
        const numA = parseNum(valA), numB = parseNum(valB);
        if (numA !== null && numB !== null) return sortConfig.direction === 'asc' ? numA - numB : numB - numA;
        const strA = String(valA || '').toLowerCase(), strB = String(valB || '').toLowerCase();
        if (strA < strB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (strA > strB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return sortableData;
  }, [tableMode, consolidatedTeams, filteredData, sortConfig]);

  const optM = Array.from(new Set(allData.map(d => String(d._mesOrigem || '')))).filter(Boolean).sort((a, b) => (monthOrder[a] || 99) - (monthOrder[b] || 99));
  const optU = Array.from(new Set(allData.map(d => String(d._ubsLabel || '')))).filter(Boolean).sort();
  const optE = Array.from(new Set(allData.filter(d => !selectedUnidade || d._ubsLabel === selectedUnidade).map(d => String(d._teamLabel || '')))).filter(Boolean).sort();
  const optI = Array.from(new Set(allData.map(d => String(d[colIndicador] || '')))).filter(Boolean).sort();

  const statusCounts = { 'Ótimo': 0, 'Bom': 0, 'Suficiente': 0, 'Regular': 0 };
  consolidatedTeams.forEach(t => { if (t.status && statusCounts[t.status as keyof typeof statusCounts] !== undefined) statusCounts[t.status as keyof typeof statusCounts]++; });

  const variablesChartData = variableColumns.map(col => {
    let sum = 0, count = 0;
    consolidatedTeams.forEach(team => { if (team.varValues[col] !== null && team.varValues[col] !== undefined) { sum += team.varValues[col]!; count++; } });
    return { name: col.length > 22 ? col.substring(0, 20) + '...' : col, fullName: col, valor: Number((count > 0 ? sum / count : 0).toFixed(2)), equipesContabilizadas: count };
  });

  const ubsChartData = (() => {
    const groups: Record<string, { sum: number; count: number }> = {};
    consolidatedTeams.forEach(t => {
      const label = (selectedUnidade || selectedEquipe ? t.equipe : t.unidade).replace('UBS JARDIM', 'UBS JD').replace('UBS NOSSA SENHORA', 'UBS NS');
      if (!groups[label]) groups[label] = { sum: 0, count: 0 };
      if (activeVariable) { if (t.varValues[activeVariable] !== null && t.varValues[activeVariable] !== undefined) { groups[label].sum += t.varValues[activeVariable]!; groups[label].count++; } } 
      else { groups[label].sum++; groups[label].count++; }
    });
    return Object.entries(groups).map(([name, g]) => ({ name, valor: Number((activeVariable ? (g.count > 0 ? g.sum / g.count : 0) : g.count).toFixed(2)) })).sort((a, b) => b.valor - a.valor).slice(0, 15);
  })();

  const lineChartData = (() => {
    const byMonth = new Map<string, any[]>();
    filteredDataForTrend.forEach(row => { const mes = String(row[colMes] || 'Geral'); if (!byMonth.has(mes)) byMonth.set(mes, []); byMonth.get(mes)!.push(row); });
    return Array.from(byMonth.entries()).map(([mes, rows]) => {
      const tMap = new Map<string, { sum: number; count: number }>();
      rows.forEach(r => {
        if (!tMap.has(r._teamKey)) tMap.set(r._teamKey, { sum: 0, count: 0 });
        if (activeVariable) { const n = parseNum(r[activeVariable]); if (n !== null) { tMap.get(r._teamKey)!.sum += n; tMap.get(r._teamKey)!.count++; } } 
        else { tMap.get(r._teamKey)!.sum++; tMap.get(r._teamKey)!.count++; }
      });
      let sT = 0, cT = 0;
      tMap.forEach(t => { if (t.count > 0) { sT += t.sum / t.count; cT++; } });
      return { name: mes, valor: Number((activeVariable ? (cT > 0 ? sT / cT : 0) : tMap.size).toFixed(2)) };
    }).sort((a, b) => (monthOrder[a.name] || 99) - (monthOrder[b.name] || 99));
  })();

  const summaryVariablesMap: Record<string, string> = {};
  variablesChartData.forEach(v => summaryVariablesMap[v.fullName] = formatVal(v.valor));

  const statusPieData = [
    { name: 'Ótimo', total: statusCounts['Ótimo'], color: '#3B82F6' },
    { name: 'Bom', total: statusCounts['Bom'], color: '#10B981' },
    { name: 'Suficiente', total: statusCounts['Suficiente'], color: '#F59E0B' },
    { name: 'Regular', total: statusCounts['Regular'], color: '#EF4444' }
  ].filter(d => d.total > 0);

  // --- TRAVA DE SEGURANÇA: SE NÃO HOUVER ARQUIVOS, MOSTRA ECRÃ VAZIO ---
  if (files.length === 0) return (
    <div className="h-[80vh] flex flex-col items-center justify-center text-slate-400">
      <BarChart2 className="h-20 w-20 mb-4 opacity-20"/>
      <h2 className="text-xl font-medium text-slate-600">Nenhum dado hospedado no Firebase.</h2>
      <p className="mt-2 text-sm">Vá ao módulo "Importar Dados" no menu lateral para iniciar.</p>
    </div>
  );

  return (
    <div className="space-y-4">
      <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center text-sm font-medium text-slate-700"><Filter className="h-4 w-4 mr-2 text-blue-600" /> Filtros e Variáveis:</div>
          <div className={`px-3 py-1 rounded-full text-xs font-semibold border ${isSingleTeam ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-blue-50 text-blue-800 border-blue-200'}`}>{isSingleTeam ? `🎯 Reais: ${consolidatedTeams[0]?.equipe || selectedEquipe}` : `📊 Média das ${totalEquipesUnicas} Equipes`}</div>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          <select value={selectedMes} onChange={e => setSelectedMes(e.target.value)} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50"><option value="">📅 Todos os Meses</option>{optM.map(o => <option key={o} value={o}>{o}</option>)}</select>
          <select value={selectedIndicador} onChange={e => { setSelectedIndicador(e.target.value); setSelectedVariavel(''); }} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50"><option value="">📊 Todos os Indicadores</option>{optI.map(o => <option key={o} value={o}>{o}</option>)}</select>
          <select value={activeVariable} onChange={e => setSelectedVariavel(e.target.value)} className="w-full p-2 border border-blue-300 rounded-lg text-sm bg-blue-50/60 font-medium text-blue-950">{!variableColumns.length && <option value="">Sem numéricas</option>}{variableColumns.map(c => <option key={c} value={c}>📈 Variável: {c}</option>)}</select>
          <select value={selectedUnidade} onChange={e => { setSelectedUnidade(e.target.value); setSelectedEquipe(''); }} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50"><option value="">🏥 Todas as Unidades</option>{optU.map(o => <option key={o} value={o}>{o}</option>)}</select>
          <select value={selectedEquipe} onChange={e => setSelectedEquipe(e.target.value)} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50"><option value="">👥 Todas as Equipes ({optE.length})</option>{optE.map(o => <option key={o} value={o}>{o}</option>)}</select>
          <div className="relative w-full"><Search className="absolute left-2.5 top-2.5 h-4 w-4 text-slate-400" /><input type="text" placeholder="Pesquisar..." value={searchTerm} onChange={e => setSearchTerm(e.target.value)} className="w-full pl-8 pr-3 py-2 border border-slate-200 rounded-lg text-sm bg-slate-50" /></div>
        </div>
      </section>

      <section className="grid grid-cols-2 lg:grid-cols-5 gap-3">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-3"><div className="bg-slate-100 p-2.5 rounded-full shrink-0"><Users className="h-5 w-5 text-slate-700" /></div><div><p className="text-xs text-slate-700 font-semibold uppercase">Equipes</p><p className="text-xl font-bold text-slate-800">{totalEquipesUnicas}</p></div></div>
        <div className="bg-white p-4 rounded-xl border border-blue-200 shadow-sm flex items-center space-x-3"><div className="bg-blue-100 p-2.5 rounded-full shrink-0"><Award className="h-5 w-5 text-blue-600" /></div><div><p className="text-xs text-blue-700 font-semibold uppercase">Ótimo</p><p className="text-xl font-bold text-blue-600">{statusCounts['Ótimo']}</p></div></div>
        <div className="bg-white p-4 rounded-xl border border-emerald-200 shadow-sm flex items-center space-x-3"><div className="bg-emerald-100 p-2.5 rounded-full shrink-0"><CheckCircle className="h-5 w-5 text-emerald-600" /></div><div><p className="text-xs text-emerald-700 font-semibold uppercase">Bom</p><p className="text-xl font-bold text-emerald-600">{statusCounts['Bom']}</p></div></div>
        <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-sm flex items-center space-x-3"><div className="bg-amber-100 p-2.5 rounded-full shrink-0"><AlertTriangle className="h-5 w-5 text-amber-600" /></div><div><p className="text-xs text-amber-700 font-semibold uppercase">Suficiente</p><p className="text-xl font-bold text-amber-600">{statusCounts['Suficiente']}</p></div></div>
        <div className="bg-white p-4 rounded-xl border border-red-200 shadow-sm flex items-center space-x-3 col-span-2 lg:col-span-1"><div className="bg-red-100 p-2.5 rounded-full shrink-0"><AlertOctagon className="h-5 w-5 text-red-600" /></div><div><p className="text-xs text-red-700 font-semibold uppercase">Regular</p><p className="text-xl font-bold text-red-600">{statusCounts['Regular']}</p></div></div>
      </section>

      {variablesChartData.length > 0 && (
        <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center mb-3"><div className="flex items-center space-x-2"><BarChart2 className="h-4 w-4 text-blue-600" /><h3 className="text-sm font-bold text-slate-800">{isSingleTeam ? `Reais (${consolidatedTeams[0]?.equipe || selectedEquipe})` : 'Resumo das Variáveis (Média)'}</h3></div><span className="text-xs text-slate-500">Clique para focar</span></div>
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2.5">
            {variablesChartData.map((v, i) => (
              <div key={i} onClick={() => setSelectedVariavel(v.fullName)} className={`p-3 rounded-lg border cursor-pointer transition-all ${v.fullName === activeVariable ? 'border-blue-600 bg-blue-50/70 shadow-sm' : 'border-slate-200 bg-slate-50/60 hover:bg-slate-100'}`}>
                <p className="text-[11px] text-slate-600 font-medium line-clamp-2 leading-tight" title={v.fullName}>{v.fullName}</p>
                <p className="text-lg font-bold text-slate-900 mt-1.5">{formatVal(v.valor)}</p>
                <p className="text-[10px] text-slate-500 mt-0.5">{isSingleTeam ? 'Valor real da equipe' : `Média de ${v.equipesContabilizadas} equipes`}</p>
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm lg:col-span-2">
          <h3 className="text-base font-bold text-slate-800 mb-1">{isSingleTeam ? 'Valores Reais' : 'Média das Equipes'}</h3>
          <div style={{ height: 290 }}>
            <ResponsiveContainer><BarChart data={variablesChartData} margin={{ top: 25, right: 15, left: -10, bottom: 55 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" /><XAxis dataKey="name" stroke="#475569" fontSize={10} tickLine={false} axisLine={false} angle={-30} textAnchor="end" interval={0} /><YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} /><Tooltip formatter={(val:number, _n:string, p:any) => [formatVal(val), p.payload.fullName]} /><Bar dataKey="valor" radius={[4, 4, 0, 0]} maxBarSize={48}>{variablesChartData.map((e, i) => <Cell key={i} fill={e.fullName === activeVariable ? '#1D4ED8' : C_COLORS[i % C_COLORS.length]} />)}<LabelList dataKey="valor" position="top" fontSize={11} fontWeight="bold" formatter={(v:number) => formatVal(v)} /></Bar></BarChart></ResponsiveContainer>
          </div>
        </div>
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-base font-bold text-slate-800 mb-1">Classificação Final</h3>
          <div style={{ height: 290 }}>
            {statusPieData.length ? (
              <ResponsiveContainer><PieChart><Pie data={statusPieData} cx="50%" cy="45%" outerRadius={82} dataKey="total" nameKey="name" label={({ name, value }) => `${name}: ${value}`}>{statusPieData.map((e, i) => <Cell key={i} fill={e.color} />)}</Pie><Tooltip /><Legend verticalAlign="bottom" height={36} iconType="circle" /></PieChart></ResponsiveContainer>
            ) : <div className="h-full flex items-center justify-center text-slate-400 text-sm">Sem classificação</div>}
          </div>
        </div>
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm lg:col-span-2">
          <h3 className="text-base font-bold text-slate-800 mb-1">{selectedUnidade || selectedEquipe ? `Equipes (${selectedUnidade})` : 'Por Unidade (UBS)'}</h3>
          <div style={{ height: 280 }}>
            <ResponsiveContainer><BarChart data={ubsChartData} margin={{ top: 25, right: 10, left: -10, bottom: 50 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" /><XAxis dataKey="name" stroke="#475569" fontSize={10} tickLine={false} axisLine={false} angle={-35} textAnchor="end" interval={0} /><YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} /><Tooltip formatter={(v:number) => [formatVal(v), activeVariable || 'Valor']} /><Bar dataKey="valor" fill="#118DFF" radius={[4, 4, 0, 0]} maxBarSize={45}><LabelList dataKey="valor" position="top" fontSize={11} fontWeight="bold" formatter={(v:number) => formatVal(v)} /></Bar></BarChart></ResponsiveContainer>
          </div>
        </div>
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-base font-bold text-slate-800 mb-1">Evolução Mensal</h3>
          <div style={{ height: 280 }}>
            <ResponsiveContainer><LineChart data={lineChartData} margin={{ top: 25, right: 25, left: -10, bottom: 15 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" /><XAxis dataKey="name" stroke="#475569" fontSize={11} tickLine={false} axisLine={false} padding={{ left: 20, right: 20 }} /><YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} /><Tooltip formatter={(v:number) => [formatVal(v), isSingleTeam ? 'Equipe' : 'Média']} /><Line type="monotone" dataKey="valor" stroke="#10B981" strokeWidth={3} dot={{ r: 6, fill: '#10B981', strokeWidth: 2, stroke: '#fff' }}><LabelList dataKey="valor" position="top" offset={12} fontSize={12} fontWeight="bold" formatter={(v:number) => formatVal(v)} /></Line></LineChart></ResponsiveContainer>
          </div>
        </div>
      </section>

      <section className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="flex justify-between items-center mb-3">
          <h3 className="text-base font-bold text-slate-800 flex items-center"><FileText className="h-5 w-5 mr-2 text-blue-600" /> Relatório Tabela</h3>
          <div className="flex space-x-1 bg-slate-100 p-1 rounded-lg text-xs">
            <button onClick={() => setTableMode('consolidado')} className={`px-3 py-1 rounded-md font-medium ${tableMode === 'consolidado' ? 'bg-blue-600 text-white' : 'text-slate-600'}`}>Por Equipe</button>
            <button onClick={() => setTableMode('historico')} className={`px-3 py-1 rounded-md font-medium ${tableMode === 'historico' ? 'bg-blue-600 text-white' : 'text-slate-600'}`}>Histórico Mês</button>
          </div>
        </div>
        
        <div className="overflow-x-auto max-h-96 border border-slate-200 rounded-lg">
          <table className="w-full text-xs text-left whitespace-nowrap">
            <thead className="text-slate-700 bg-slate-100 sticky top-0 shadow-sm z-10">
              <tr>
                <th onClick={() => requestSort('_ms')} className="px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group"><div className="flex items-center">Meses {sortConfig?.key === '_ms' ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100" />}</div></th>
                <th onClick={() => requestSort('_eq')} className="px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group"><div className="flex items-center">Equipe {sortConfig?.key === '_eq' ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100" />}</div></th>
                {columns.map(c => (
                  <th key={c} onClick={() => requestSort(c)} className={`px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group ${c === activeVariable ? 'bg-blue-100' : ''}`}><div className="flex items-center">{c} {sortConfig?.key === c ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100 shrink-0" />}</div></th>
                ))}
              </tr>
            </thead>
            <tbody>
              {!isSingleTeam && totalEquipesUnicas > 1 && (
                <tr className="bg-blue-50/90 border-b border-blue-200 font-bold"><td className="px-3 py-2">RESUMO</td><td className="px-3 py-2">MÉDIA DAS EQUIPES</td>{columns.map(c => <td key={c} className="px-3 py-2">{summaryVariablesMap[c] !== undefined ? summaryVariablesMap[c] : '-'}</td>)}</tr>
              )}
              {sortedTableData.slice(0, 150).map((r, i) => (
                <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-3 py-2 text-blue-800 bg-blue-50/20">{r._ms}</td>
                  <td className="px-3 py-2 font-semibold bg-blue-50/10">{r._eq}</td>
                  {columns.map(c => {
                    let st = null;
                    if (c.toLowerCase().includes('classifica') || c.toLowerCase().includes('conceito') || c.toLowerCase().includes('qualidade') || c.toLowerCase().includes('status')) {
                      st = normStat(r[c]);
                    }
                    if (!st && c === activeVariable) {
                      const v = parseNum(r[c]);
                      if (v !== null) {
                         const pct = v <= 10 && v > 0 ? v * 10 : v;
                         st = pct > 75 ? 'Ótimo' : pct > 50 ? 'Bom' : pct > 25 ? 'Suficiente' : 'Regular';
                      }
                    }
                    return <td key={c} className={`px-3 py-2 text-slate-600 ${c === activeVariable ? 'bg-blue-50/40 font-semibold' : ''}`}>{st ? <span className="px-2 py-0.5 rounded-full text-[11px] font-bold text-white" style={{ backgroundColor: S_COLORS[st] }}>{String(r[c])}</span> : String(r[c] ?? '')}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}