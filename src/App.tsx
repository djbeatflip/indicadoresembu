import React, { useState, useMemo, useEffect, useRef } from 'react';
import { UploadCloud, Users, CheckCircle, AlertTriangle, AlertOctagon, Award, FileText, Search, Filter, X, BarChart2, FolderOpen, LogIn, Layout, Settings, Shield, UserPlus, LogOut, ArrowUpDown, ChevronUp, ChevronDown, Trash2, Menu, ChevronLeft, Trophy } from 'lucide-react';
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

// Inicialização Segura do Firebase
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
const secondaryApp = getApps().find(a => a.name === "Secondary") || initializeApp(firebaseConfig, "Secondary");

const auth = getAuth(app);
const db = getFirestore(app);
const secondaryAuth = getAuth(secondaryApp);

// --- TIPAGENS ---
interface ImportedFile { id: string; name: string; lastModified?: number; data: any[]; }
interface User { id: string; email: string; role: string; name: string; }

// --- ESCUDO ANTI ECRÃ BRANCO ---
class ErrorBoundary extends React.Component<any, { hasError: boolean, error: Error | null, isClearing: boolean }> {
  constructor(props: any) { 
    super(props); 
    this.state = { hasError: false, error: null, isClearing: false }; 
  }
  
  static getDerivedStateFromError(error: Error) { return { hasError: true, error }; }

  handleEmergencyClear = async () => {
    this.setState({ isClearing: true });
    try {
      const qs = await getDocs(collection(db, 'dados_siaps'));
      let batch = writeBatch(db);
      let opCount = 0;
      for(let i = 0; i < qs.docs.length; i++) {
         batch.delete(qs.docs[i].ref);
         opCount++;
         if(opCount >= 400 || i === qs.docs.length - 1) { 
            await batch.commit(); batch = writeBatch(db); opCount = 0; 
         }
      }
      window.location.reload();
    } catch(e) { 
      alert('Erro ao apagar nuvem. Verifique a sua ligação.'); 
      this.setState({ isClearing: false }); 
    }
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex flex-col items-center justify-center min-h-screen bg-slate-100 p-6">
          <div className="bg-white p-8 rounded-2xl shadow-2xl max-w-lg w-full text-center border-t-4 border-red-500">
            <AlertTriangle className="h-16 w-16 text-red-500 mx-auto mb-4" />
            <h2 className="text-2xl font-bold text-slate-800 mb-2">Modo de Segurança Ativado</h2>
            <p className="text-sm text-slate-600 mb-6">O painel encontrou dados corrompidos e desligou a interface para evitar falhas. Utilize o botão abaixo para limpar a base de dados e recomeçar.</p>
            <div className="bg-slate-50 p-4 rounded-lg text-left mb-6 overflow-auto max-h-32 text-[10px] font-mono text-red-400 border border-red-100">{this.state.error?.toString()}</div>
            <button onClick={this.handleEmergencyClear} disabled={this.state.isClearing} className="w-full bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-4 rounded-xl transition-all flex justify-center items-center shadow-lg hover:shadow-red-500/30">
              {this.state.isClearing ? 'A Limpar a Nuvem...' : <><Trash2 className="mr-2 h-5 w-5"/> Forçar Limpeza da Nuvem</>}
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// --- UTILITÁRIOS MATEMÁTICOS & PADRONIZAÇÃO ---
const monthOrder: Record<string, number> = { 'Janeiro':1, 'Fevereiro':2, 'Março':3, 'Abril':4, 'Maio':5, 'Junho':6, 'Julho':7, 'Agosto':8, 'Setembro':9, 'Outubro':10, 'Novembro':11, 'Dezembro':12 };
const S_COLORS: Record<string, string> = { 'Ótimo': '#3B82F6', 'Bom': '#10B981', 'Satisfatório': '#F59E0B', 'Regular': '#EF4444' };
const S_WEIGHTS: Record<string, number> = { 'Regular': 1, 'Satisfatório': 2, 'Bom': 3, 'Ótimo': 4 };
const W_STATUS: Record<number, string> = { 1: 'Regular', 2: 'Satisfatório', 3: 'Bom', 4: 'Ótimo' };
const C_COLORS = ['#118DFF', '#10B981', '#F59E0B', '#8B5CF6', '#EF4444', '#14B8A6', '#E66C37'];

const normalizeStrForFind = (s: any) => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();

// Tabela de Lista Negra (Esconde colunas de identificação dos cartões)
const isMetadataColumn = (colName: string) => {
  const s = normalizeStrForFind(colName);
  const exactMatches = ['cnes', 'ine', 'uf', 'municipio', 'cidade', 'ano', 'mes', 'competencia', 'periodo', 'estabelecimento', 'unidade', 'ubs', 'equipe', 'tipo', 'sigla', 'condicao', 'id'];
  
  if (exactMatches.includes(s)) return true;
  if (s.startsWith('tipo de') || s.startsWith('tipo do')) return true;
  if (s.startsWith('nome da') || s.startsWith('nome do')) return true;
  if (s.startsWith('sigla da') || s.startsWith('sigla do')) return true;
  if (s.startsWith('codigo ine') || s.startsWith('codigo da') || s.startsWith('codigo do')) return true;
  if (s.includes('condicao das')) return true;
  
  return false;
};

// Transforma hífens em 0 matematicamente
const parseNum = (v: any) => { 
  if (v === '-' || v === '–' || v === '—') return 0;
  if (typeof v === 'number' && !isNaN(v)) return v; 
  if (typeof v !== 'string') return null; 
  let s = v.trim().replace(/%/g, '').replace(/R\$/g, '').trim(); 
  if (s === '-' || s === '–' || s === '—') return 0;
  if (!s) return null; 
  if (s.includes('.') && s.includes(',')) s = s.replace(/\./g, '').replace(',', '.'); 
  else if (s.includes(',')) s = s.replace(',', '.'); 
  const n = Number(s); 
  return isNaN(n) ? null : n; 
};

const formatVal = (v: number) => Number.isInteger(v) ? v.toLocaleString('pt-BR') : v.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
const normTxt = (v: any) => String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
const normIne = (v: any) => String(v || '').replace(/\D/g, '').replace(/^0+/, '');

const normStat = (v: any): string | null => { 
  const s = String(v || '').toLowerCase().trim(); 
  if (!s) return null; 
  if (s.includes('ótimo') || s.includes('otimo')) return 'Ótimo'; 
  if (['bom', 'boa'].includes(s) || s.startsWith('bom ')) return 'Bom'; 
  if (s === 'satisfatório' || s.includes('satisfat') || s === 'suficiente' || (s.includes('suficiente') && !s.includes('insuficiente'))) return 'Satisfatório'; 
  if (['regular', 'ruim', 'crítico', 'critico', 'insuficiente'].includes(s) || s.includes('regular')) return 'Regular'; 
  return null; 
};

const calcStatus = (v: number | null, ind: string): string | null => {
  if (v === null) return null;
  const i = String(ind || '').toLowerCase();
  if (i.includes('vinculo') || i.includes('vínculo')) {
    if (v < 5) return 'Regular';
    if (v < 7) return 'Satisfatório';
    if (v <= 8.5) return 'Bom';
    return 'Ótimo';
  }
  if (i.includes('c1')) {
    if (v <= 10 || v > 70) return 'Regular';
    if (v <= 30) return 'Satisfatório';
    if (v <= 50) return 'Bom';
    return 'Ótimo';
  }
  if (v <= 25) return 'Regular';
  if (v <= 50) return 'Satisfatório';
  if (v <= 75) return 'Bom';
  return 'Ótimo';
};

const extMes = (f: string) => { const u = String(f || '').toUpperCase().trim(); for (const m of Object.keys(monthOrder)) if (u.startsWith(m.toUpperCase().substring(0,3))) return m; return String(f || '').split('-')[0].replace(/[0-9]/g, '').trim() || String(f || ''); };
const extInd = (f: string) => String(f || '').split('-').length > 1 ? String(f || '').split('-').slice(1).join('-').trim() : String(f || '').trim();

const mapIndicatorName = (name: string) => {
  const n = String(name || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if(n.includes('acesso')) return 'C1 - Mais Acesso';
  if(n.includes('infantil')) return 'C2 - Desenvolvimento Infantil';
  if(n.includes('gesta') || n.includes('puerp')) return 'C3 - Cuidado na gestação e Puerpério';
  if(n.includes('diabet')) return 'C4 - Diabetes';
  if(n.includes('hiperten')) return 'C5 - Hipertensão';
  if(n.includes('idos')) return 'C6 - Idosos';
  if(n.includes('cancer') || n.includes('feminino') || n.includes('colo')) return 'C7 - Prevenção cancer feminino';
  if(n.includes('vinculo') || n.includes('territorial')) return 'Vínculo e Acompanhamento Territorial';
  return String(name || '').toUpperCase();
};

// ==========================================
// MOTOR DE LEITURA INFALÍVEL
// ==========================================
const buildRows = (rows: any[][]) => {
  if (!rows || !Array.isArray(rows)) return [];
  const validRows = rows.filter(r => Array.isArray(r) && r.length > 0);
  if (validRows.length === 0) return [];

  let hIdx = -1;
  for (let i = 0; i < validRows.length; i++) {
    const rowStr = validRows[i].map(c => String(c || '').replace(/\r?\n|\r/g, ' ').trim().toUpperCase());
    if (rowStr.some(c => c === 'CNES' || c === 'ESTABELECIMENTO' || c === 'INE' || c === 'NOME DA EQUIPE' || c === 'EQUIPE')) {
      hIdx = i; break;
    }
  }

  if (hIdx === -1) return [];

  const rawHeaders = validRows[hIdx] || [];
  const seen = new Set<string>();
  const headers = rawHeaders.map((c, i) => {
    let val = String(c || '').replace(/\r?\n|\r/g, ' ').trim();
    if (!val) val = `EMPTY_${i}`;
    let base = val; let count = 1;
    while (seen.has(base)) { base = `${val} (${count})`; count++; }
    seen.add(base); return base;
  });

  return validRows.slice(hIdx + 1).filter(r => {
    const firstCells = r.slice(0, 4).join(' ').toUpperCase();
    if (!firstCells.trim()) return false;
    if (firstCells.includes('TOTAL') || firstCells.includes('FONTE:') || firstCells.includes('LEGENDA')) return false;
    return true;
  }).map(r => {
    const obj: any = {}; let hasData = false;
    headers.forEach((col, c) => { 
      if (!col.startsWith('EMPTY_')) { obj[col] = r[c] ?? ''; if (String(obj[col]).trim() !== '') hasData = true; } 
    });
    return hasData ? obj : null;
  }).filter(Boolean);
};

// ==========================================
// APLICAÇÃO ENCAPSULADA (COM MODO PÚBLICO)
// ==========================================
function MainApp() {
  const [isLoaded, setIsLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [activeModule, setActiveModule] = useState('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [files, setFiles] = useState<ImportedFile[]>([]);
  const [usersList, setUsersList] = useState<User[]>([]);
  const [sysTheme, setSysTheme] = useState('blue');

  // Lógica de Leitura Mista (Carrega dados mesmo que visitante não esteja logado)
  useEffect(() => {
    const loadPublicData = async () => {
      try {
        const setDocSnap = await getDoc(doc(db, 'settings', 'layout'));
        if (setDocSnap.exists()) setSysTheme(setDocSnap.data().themeColor);

        const filesSnap = await getDocs(collection(db, 'dados_siaps'));
        const agrupados = new Map<string, ImportedFile>();
        filesSnap.docs.forEach(d => {
           const row = d.data();
           const fileName = row._arquivoOrigem + '.xlsx';
           if(!agrupados.has(fileName)) { agrupados.set(fileName, { id: fileName, name: fileName, data: [] }); }
           agrupados.get(fileName)!.data.push(row);
        });
        setFiles(Array.from(agrupados.values()));
      } catch (e) {
        console.error("Erro ao carregar dados. Verifique as regras do Firestore.", e);
      }
      setIsLoaded(true);
    };

    loadPublicData();

    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        try {
          const userDoc = await getDoc(doc(db, 'users', user.uid));
          if (userDoc.exists()) {
            setCurrentUser({ id: user.uid, email: user.email!, ...userDoc.data() } as User);
          } else {
            setCurrentUser({ id: user.uid, email: user.email!, role: 'Administrador', name: 'Administrador' });
          }
          const usersSnap = await getDocs(collection(db, 'users'));
          setUsersList(usersSnap.docs.map(d => ({ id: d.id, ...d.data() } as User)));
        } catch (e) { console.error(e); }
      } else {
        setCurrentUser(null);
      }
    });
    return () => unsubscribe();
  }, []);

  // Se logar com sucesso enquanto estava na aba de Login, muda logo para Dashboard
  useEffect(() => {
    if (currentUser && activeModule === 'login') setActiveModule('dashboard');
  }, [currentUser, activeModule]);

  const themes: Record<string, string> = { blue: 'bg-blue-600', emerald: 'bg-emerald-600', violet: 'bg-violet-600', slate: 'bg-slate-700' };
  const bgClass = themes[sysTheme] || themes.blue;

  if (!isLoaded) return <div className="min-h-screen flex items-center justify-center bg-slate-100 text-slate-500 font-medium">Conectando ao banco de dados público...</div>;

  const handleLogout = async () => { 
    await signOut(auth); 
    setActiveModule('dashboard'); // Quando faz logout, volta à vista pública
  };

  return (
    <div className="min-h-screen bg-slate-50 flex overflow-hidden">
      <aside className={`${isSidebarOpen ? 'w-64' : 'w-20'} bg-slate-900 text-slate-300 flex flex-col shadow-2xl z-20 transition-all duration-300 ease-in-out shrink-0`}>
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
              <button onClick={() => setIsSidebarOpen(false)} className="text-white/70 hover:text-white shrink-0 bg-black/10 p-1.5 rounded-lg hover:bg-black/20 transition-colors" title="Recolher Menu"><ChevronLeft className="h-5 w-5" /></button>
            </>
          ) : (
            <button onClick={() => setIsSidebarOpen(true)} className="text-white hover:text-white shrink-0 p-2 bg-white/20 rounded-lg shadow-sm transition-colors" title="Expandir Menu"><Menu className="h-6 w-6" /></button>
          )}
        </div>
        
        <div className="p-4 flex-1 space-y-1.5 overflow-y-auto overflow-x-hidden scrollbar-thin">
          {isSidebarOpen && <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-2">Público</p>}
          <button onClick={() => setActiveModule('dashboard')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'dashboard' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Dashboard SIAPS" : ""}>
            <Layout className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Dashboard SIAPS</span>}
          </button>
          
          <button onClick={() => setActiveModule('ranking')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'ranking' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Ranking" : ""}>
            <Trophy className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Ranking</span>}
          </button>

          {/* Menus Restritos a Utilizadores Logados */}
          {currentUser && (currentUser.role === 'Administrador' || currentUser.role === 'Coordenador') && (
            <>
              {isSidebarOpen ? ( <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-6">Administração</p> ) : ( <div className="h-px bg-slate-800 my-4 mx-2"></div> )}
              <button onClick={() => setActiveModule('import')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'import' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Importar Dados" : ""}>
                <UploadCloud className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Importar Dados</span>}
              </button>
            </>
          )}

          {currentUser && currentUser.role === 'Administrador' && (
            <>
              <button onClick={() => setActiveModule('users')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'users' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Usuários" : ""}>
                <Users className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Usuários</span>}
              </button>
              <button onClick={() => setActiveModule('access')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'access' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Acessos" : ""}>
                <Shield className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Acessos</span>}
              </button>
              <button onClick={() => setActiveModule('settings')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'settings' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Configurações" : ""}>
                <Settings className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Configurações</span>}
              </button>
            </>
          )}
        </div>

        {/* Rodapé da Sidebar (Botão de Login Público ou Info do User Logado) */}
        <div className="p-4 border-t border-slate-800 bg-slate-950">
          {currentUser ? (
             <div className={`flex items-center ${isSidebarOpen ? 'justify-between' : 'justify-center flex-col space-y-4'}`}>
              <div className="flex items-center space-x-3 overflow-hidden">
                <div className={`h-8 w-8 rounded-full shrink-0 ${bgClass} flex items-center justify-center text-white font-bold text-xs`} title={currentUser.name}>{(currentUser.name || 'U').charAt(0).toUpperCase()}</div>
                {isSidebarOpen && <div className="truncate w-28"><p className="text-xs font-bold text-slate-200 truncate">{currentUser.name}</p><p className="text-[10px] text-slate-500 truncate">{currentUser.role}</p></div>}
              </div>
              <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400 bg-slate-900 rounded-lg transition-colors" title="Sair da Conta"><LogOut className="h-4 w-4 shrink-0" /></button>
            </div>
          ) : (
            <button onClick={() => setActiveModule('login')} className={`w-full flex items-center ${isSidebarOpen ? 'justify-start px-2' : 'justify-center'} py-2 text-slate-400 hover:text-white transition-colors rounded-lg hover:bg-slate-800`} title="Acesso Restrito">
              <LogIn className="h-5 w-5 shrink-0" />
              {isSidebarOpen && <span className="ml-3 text-sm font-medium">Acesso Restrito</span>}
            </button>
          )}
        </div>
      </aside>

      {/* ÁREA PRINCIPAL (Renders baseados no activeModule) */}
      <main className="flex-1 overflow-y-auto p-6 bg-slate-50">
        {activeModule === 'dashboard' && <ModuleDashboard files={files} bgClass={bgClass} />}
        {activeModule === 'ranking' && <ModuleRanking files={files} bgClass={bgClass} />}
        {activeModule === 'login' && !currentUser && <LoginScreen bgClass={bgClass} onBack={() => setActiveModule('dashboard')} />}
        
        {currentUser && activeModule === 'import' && <ModuleImport files={files} setFiles={setFiles} bgClass={bgClass} />}
        {currentUser && activeModule === 'users' && <ModuleUsers usersList={usersList} setUsersList={setUsersList} bgClass={bgClass} />}
        {currentUser && activeModule === 'access' && <ModuleAccess bgClass={bgClass} />}
        {currentUser && activeModule === 'settings' && <ModuleSettings sysTheme={sysTheme} setSysTheme={setSysTheme} bgClass={bgClass} />}
      </main>
    </div>
  );
}

export default function AppWrapper() {
  return (
    <ErrorBoundary>
      <MainApp />
    </ErrorBoundary>
  );
}

// ==========================================
// TELA DE LOGIN & RECUPERAÇÃO DE SENHA (AGORA EMBUTIDA NO MAIN)
// ==========================================
function LoginScreen({ bgClass, onBack }: any) {
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
        await setDoc(doc(db, 'users', userCred.user.uid), { email: email, name: nome, role: 'Administrador' });
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
    <div className="flex flex-col items-center justify-center min-h-[80vh]">
      <div className="bg-white p-8 rounded-xl shadow-lg max-w-sm w-full border border-slate-200 relative">
        <button onClick={onBack} className="absolute top-4 left-4 p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors" title="Voltar ao Painel Público">
          <ChevronLeft className="h-5 w-5" />
        </button>
        
        <div className="flex justify-center mb-6 mt-4"><div className={`p-4 rounded-full ${bgClass} bg-opacity-10`}><Shield className={`h-10 w-10 text-slate-800`} /></div></div>
        <h2 className="text-2xl font-bold text-center text-slate-800 mb-2 leading-tight">Área Restrita<br/><span className="text-lg font-medium text-slate-500">Administração</span></h2>
        
        {isForgot ? (
          <>
            <p className="text-center text-slate-500 text-sm mb-6">Recuperar Palavra-passe</p>
            {errorMsg && <div className="bg-red-50 text-red-600 p-3 rounded-lg text-xs mb-4 text-center border border-red-200">{errorMsg}</div>}
            {successMsg && <div className="bg-emerald-50 text-emerald-600 p-3 rounded-lg text-xs mb-4 text-center border border-emerald-200">{successMsg}</div>}
            <form onSubmit={handleForgotPassword} className="space-y-4">
              <div><label className="block text-xs font-semibold text-slate-600 mb-1">E-mail Cadastrado</label><input name="email" type="email" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>
              <button type="submit" className={`w-full ${bgClass} hover:opacity-90 text-white font-bold py-2.5 rounded-lg transition-opacity flex justify-center items-center`}>Enviar link de recuperação</button>
            </form>
            <div className="mt-6 text-center"><button onClick={() => {setIsForgot(false); setErrorMsg(''); setSuccessMsg('');}} className="text-xs text-slate-500 hover:underline">Voltar ao Login</button></div>
          </>
        ) : (
          <>
            <p className="text-center text-slate-500 text-sm mb-6">{isRegister ? 'Registo de Administrador' : 'Insira as credenciais'}</p>
            {errorMsg && <div className="bg-red-50 text-red-600 p-3 rounded-lg text-xs mb-4 text-center border border-red-200">{errorMsg}</div>}
            <form onSubmit={handleSubmit} className="space-y-4">
              {isRegister && (<div><label className="block text-xs font-semibold text-slate-600 mb-1">Nome Completo</label><input name="nome" type="text" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>)}
              <div><label className="block text-xs font-semibold text-slate-600 mb-1">E-mail</label><input name="email" type="email" required className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" /></div>
              <div>
                <div className="flex justify-between items-center mb-1">
                  <label className="block text-xs font-semibold text-slate-600">Senha (Mín. 6)</label>
                  {!isRegister && <button type="button" onClick={() => {setIsForgot(true); setErrorMsg('');}} className="text-[10px] text-blue-600 hover:underline">Esqueceu?</button>}
                </div>
                <input name="pass" type="password" required minLength={6} className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" />
              </div>
              <button type="submit" className={`w-full ${bgClass} hover:opacity-90 text-white font-bold py-2.5 rounded-lg transition-opacity flex justify-center items-center`}>{isRegister ? 'Criar Conta' : <><LogIn className="h-4 w-4 mr-2" /> Entrar no Sistema</>}</button>
            </form>
            <div className="mt-6 text-center"><button onClick={() => {setIsRegister(!isRegister); setErrorMsg('');}} className="text-xs text-blue-600 hover:underline">{isRegister ? 'Já tenho uma conta. Fazer Login' : 'Primeiro Acesso? Criar Conta Master'}</button></div>
          </>
        )}
      </div>
    </div>
  );
}

// ==========================================
// MÓDULO DE IMPORTAÇÃO
// ==========================================
function ModuleImport({ files, setFiles, bgClass }: any) {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    try {
      const selectedFiles = Array.from(event.target.files || []);
      if (selectedFiles.length === 0) return;
      setIsProcessing(true);
      const newFiles: ImportedFile[] = [];
      let skippedFiles = [];

      for (const file of selectedFiles) {
        if (file.name.startsWith('~$') || file.name.startsWith('.')) continue;
        const ext = file.name.split('.').pop()?.toLowerCase();
        if (!['csv', 'xlsx', 'xls'].includes(ext || '')) continue;
        
        setProgressMsg(`A Processar ${file.name}...`);
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

        if (fileData && fileData.length > 0) {
          const cleanName = file.name.replace(/\.[^/.]+$/, "");
          const properMonth = extMes(cleanName);
          const properIndicator = extInd(cleanName);
          const enrichedData = fileData.map(r => ({ ...r, _arquivoOrigem: cleanName, _mesOrigem: properMonth, _indicadorOrigem: properIndicator }));
          newFiles.push({ id: Date.now() + '-' + Math.random().toString(36).substring(2,7), name: file.name, data: enrichedData });
          
          setProgressMsg(`A Enviar ${file.name} para a Nuvem...`);
          try {
            let batch = writeBatch(db);
            let opCount = 0;
            for (let i = 0; i < enrichedData.length; i++) {
               const row = enrichedData[i];
               const rowId = `${cleanName}_${i}_${Date.now()}`;
               const docRef = doc(db, 'dados_siaps', rowId);
               batch.set(docRef, row);
               opCount++;
               if (opCount >= 400 || i === enrichedData.length - 1) { await batch.commit(); batch = writeBatch(db); opCount = 0; }
            }
          } catch (err: any) { alert(`Erro ao salvar ${file.name}. Verifique as regras do Firestore.`); }
        } else {
          skippedFiles.push(file.name);
        }
      }

      if (newFiles.length > 0) {
        setFiles((prev: ImportedFile[]) => {
          const map = new Map<string, ImportedFile>();
          (prev || []).forEach(f => map.set(f.name, f));
          newFiles.forEach(f => map.set(f.name, f));
          return Array.from(map.values());
        });
        
        if (skippedFiles.length > 0) {
          alert(`Upload parcial. Os seguintes ficheiros falharam pois não foi encontrado o cabeçalho "CNES": \n${skippedFiles.join('\n')}`);
        } else {
          alert('Upload concluído com sucesso!');
        }
      } else if (skippedFiles.length > 0) {
         alert(`NENHUM arquivo foi importado. Não foi encontrado o cabeçalho "CNES" nas planilhas: \n${skippedFiles.join('\n')}`);
      }

    } catch (error) {
      console.error("Erro crítico na importação:", error);
      alert("Ocorreu um erro crítico ao processar o arquivo. Verifique se a planilha é válida.");
    } finally {
      setIsProcessing(false);
      setProgressMsg('');
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleClearCloud = async () => {
    if(!window.confirm('CUIDADO: Isso apagará TODOS os relatórios do banco de dados na nuvem! Tem certeza?')) return;
    setIsProcessing(true);
    setProgressMsg('A Apagar a Nuvem...');
    try {
      const qs = await getDocs(collection(db, 'dados_siaps'));
      let batch = writeBatch(db);
      let opCount = 0;
      for(let i = 0; i < qs.docs.length; i++) {
         batch.delete(qs.docs[i].ref);
         opCount++;
         if(opCount >= 400 || i === qs.docs.length - 1) { await batch.commit(); batch = writeBatch(db); opCount = 0; }
      }
      setFiles([]);
    } catch(e) { alert('Erro ao apagar. Verifique as regras do Firestore.'); }
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
            <FolderOpen className="h-5 w-5 mr-2"/> {isProcessing ? 'A Processar...' : 'Selecionar e Enviar para Nuvem'}
            <input type="file" multiple accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFileUpload} disabled={isProcessing} ref={fileInputRef} />
          </label>
        </div>
      </div>

      {files && files.length > 0 && (
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center mb-4">
            <h3 className="font-bold text-slate-800">Planilhas Hospedadas ({files.length})</h3>
            <button onClick={handleClearCloud} disabled={isProcessing} className="text-red-500 text-sm hover:bg-red-50 px-3 py-1.5 rounded-lg transition-colors flex items-center"><Trash2 className="h-4 w-4 mr-1"/> Apagar TUDO da Nuvem</button>
          </div>
          <div className="flex flex-wrap gap-2">
            {files.map((f: ImportedFile) => (<div key={f.id || Math.random()} className="bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-medium text-slate-700 flex items-center"><FileText className="h-3.5 w-3.5 mr-1.5 text-blue-500"/>{f.name || 'Arquivo sem nome'}</div>))}
          </div>
        </div>
      )}
    </div>
  );
}

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
    } catch (err: any) { alert('Erro ao cadastrar: ' + err.message); }
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
  const handleSaveTheme = async (theme: string) => { setSysTheme(theme); await setDoc(doc(db, 'settings', 'layout'), { themeColor: theme }); };
  return (
    <div className="max-w-2xl bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div className={`p-5 ${bgClass} text-white`}><h2 className="text-lg font-bold flex items-center"><Settings className="mr-2 h-5 w-5"/> Customizar Interface</h2></div>
      <div className="p-6">
        <h3 className="text-sm font-bold text-slate-800 mb-3">Cor Principal</h3>
        <div className="flex space-x-3">
          {['blue', 'emerald', 'violet', 'slate'].map(c => (<button key={c} onClick={() => handleSaveTheme(c)} className={`h-12 w-12 rounded-full border-4 transition-all ${sysTheme === c ? 'border-slate-800 scale-110 shadow-md' : 'border-transparent hover:scale-105'}`} style={{ backgroundColor: c==='blue'?'#2563eb':c==='emerald'?'#059669':c==='violet'?'#7c3aed':'#334155' }} />))}
        </div>
      </div>
    </div>
  );
}

// ==========================================
// MÓDULO DE DASHBOARD PRINCIPAL
// ==========================================
function ModuleDashboard({ files, bgClass }: any) {
  const rawAllData = useMemo(() => files.flatMap((f: any) => f.data), [files]);
  const columns = useMemo(() => Array.from(new Set(rawAllData.flatMap(Object.keys))).filter(c => c && !String(c).startsWith('_') && !String(c).toUpperCase().includes('EMPTY') && String(c).trim() !== ''), [rawAllData]);

  const colUnidade = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs');
  }) || null, [columns]);

  const colIne = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s === 'ine' || s.includes('codigo ine');
  }) || null, [columns]);

  const colTipoEquipe = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('tipo do ine') || s.includes('tipo de equipe') || s === 'tipo';
  }) || null, [columns]);

  const colEquipe = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('nome da equipe') || s === 'equipe' || s === 'nome';
  }) || null, [columns]);

  const colMes = '_mesOrigem';
  const colIndicador = useMemo(() => columns.slice(0, 8).find(c => ['indicador', 'condição', 'condicao'].includes(String(c).toLowerCase())) || '_indicadorOrigem', [columns]);

  const isCorrupted = rawAllData.length > 0 && (!colUnidade && !colEquipe);

  const allData = useMemo(() => {
    if (isCorrupted) return [];
    const enriched: any[] = [];
    rawAllData.forEach((r, idx) => {
      let n = colEquipe ? String(r[colEquipe] || '').replace(/\s+/g, ' ').trim() : '';
      const ic = colIne ? normIne(r[colIne]) : '';
      let lbl = n ? n.toUpperCase() : (ic ? `EQUIPE INE ${ic}` : `EQUIPE ${idx + 1}`);
      const ubs = colUnidade ? String(r[colUnidade] || 'Sem Unidade').replace(/\s+/g, ' ').trim().toUpperCase() : 'GERAL';
      
      let tipoStr = 'Outros';
      const typeVal = colTipoEquipe ? String(r[colTipoEquipe] || '').toUpperCase() : '';
      const nameVal = lbl.toUpperCase();

      if (typeVal.includes('ESF') || nameVal.includes('ESF') || typeVal.includes('FAMILIA') || nameVal.includes('FAMILIA')) {
          tipoStr = 'ESF';
      } else if (typeVal.includes('EAP') || nameVal.includes('EAP') || typeVal.includes('PRIMARIA') || nameVal.includes('PRIMARIA')) {
          tipoStr = 'EAP';
      }

      const indOrig = String(colIndicador !== '_indicadorOrigem' && r[colIndicador] ? r[colIndicador] : r._indicadorOrigem || '');
      const mappedInd = mapIndicatorName(indOrig);

      enriched.push({ 
        ...r, 
        _teamKey: `UBS_${normTxt(ubs)}__EQ_${normTxt(lbl)}`, 
        _teamLabel: lbl, 
        _ubsLabel: ubs,
        _tipoEquipe: tipoStr,
        _mappedIndicator: mappedInd
      });
    });
    return enriched;
  }, [rawAllData, colUnidade, colEquipe, colIne, colTipoEquipe, colIndicador, isCorrupted]);

  const [selectedUnidade, setSelectedUnidade] = useState('');
  const [selectedTipoEquipe, setSelectedTipoEquipe] = useState('');
  const [selectedEquipe, setSelectedEquipe] = useState('');
  const [selectedMes, setSelectedMes] = useState('');
  const [selectedIndicador, setSelectedIndicador] = useState('');
  const [sortConfig, setSortConfig] = useState<{ key: string, direction: 'asc' | 'desc' } | null>(null);
  const [tableMode, setTableMode] = useState<'consolidado' | 'historico'>('consolidado');
  
  const [activeModalStatus, setActiveModalStatus] = useState<string | null>(null);

  const optM = useMemo(() => Array.from(new Set(allData.map(d => String(d._mesOrigem || '')))).filter(Boolean).sort((a, b) => (monthOrder[a] || 99) - (monthOrder[b] || 99)), [allData]);
  const optI = useMemo(() => {
    return Array.from(new Set(allData.map(d => String(d._mappedIndicator || ''))))
      .filter(Boolean)
      .sort((a, b) => {
        if (a.toLowerCase().includes('vinculo') || a.toLowerCase().includes('vínculo')) return -1;
        if (b.toLowerCase().includes('vinculo') || b.toLowerCase().includes('vínculo')) return 1;
        return a.localeCompare(b);
      });
  }, [allData]);

  const handleClearFilters = () => {
    setSelectedUnidade('');
    setSelectedTipoEquipe('');
    setSelectedEquipe('');
    setSortConfig(null);
    if (optI.length > 0) setSelectedIndicador(optI[0]);
    if (optM.length > 0) {
      const currentMonthNum = new Date().getMonth() + 1;
      let closest = optM[optM.length - 1];
      let minDiff = Infinity;
      optM.forEach(m => {
        const mNum = monthOrder[m];
        if (mNum) {
          const diff = Math.abs(mNum - currentMonthNum);
          if (diff < minDiff) { minDiff = diff; closest = m; }
        }
      });
      setSelectedMes(closest);
    }
  };

  useEffect(() => {
    if (!selectedIndicador && optI.length > 0) setSelectedIndicador(optI[0]);
    if (!selectedMes && optM.length > 0) {
      const currentMonthNum = new Date().getMonth() + 1;
      let closest = optM[optM.length - 1];
      let minDiff = Infinity;
      optM.forEach(m => {
        const mNum = monthOrder[m];
        if (mNum) {
          const diff = Math.abs(mNum - currentMonthNum);
          if (diff < minDiff) { minDiff = diff; closest = m; }
        }
      });
      setSelectedMes(closest);
    }
  }, [optI, optM, selectedIndicador, selectedMes]);

  const baseFilterData = useMemo(() => {
    return allData.filter(d => 
       (!selectedMes || String(d[colMes]) === selectedMes) &&
       (!selectedIndicador || d._mappedIndicator === selectedIndicador)
    );
  }, [allData, selectedMes, selectedIndicador, colMes]);

  const currentIndicatorColumns = useMemo(() => {
    if (baseFilterData.length === 0) return [];
    const keys = new Set<string>();
    baseFilterData.forEach(r => {
      Object.keys(r).forEach(k => {
        if (!k.startsWith('_') && !k.toUpperCase().includes('EMPTY')) {
          keys.add(k);
        }
      });
    });
    return Array.from(keys);
  }, [baseFilterData]);

  const optU = useMemo(() => Array.from(new Set(baseFilterData
    .filter(d => (!selectedTipoEquipe || d._tipoEquipe === selectedTipoEquipe) && (!selectedEquipe || d._teamLabel === selectedEquipe))
    .map(d => String(d._ubsLabel || ''))
  )).filter(u => Boolean(u) && u !== 'SEM UNIDADE' && u !== 'GERAL').sort(), [baseFilterData, selectedTipoEquipe, selectedEquipe]);
  
  const optT = useMemo(() => Array.from(new Set(baseFilterData
    .filter(d => (!selectedUnidade || d._ubsLabel === selectedUnidade) && (!selectedEquipe || d._teamLabel === selectedEquipe))
    .map(d => String(d._tipoEquipe || 'Outros'))
  )).filter(t => Boolean(t) && t !== 'Outros').sort(), [baseFilterData, selectedUnidade, selectedEquipe]);

  const optE = useMemo(() => Array.from(new Set(baseFilterData
    .filter(d => (!selectedUnidade || d._ubsLabel === selectedUnidade) && (!selectedTipoEquipe || d._tipoEquipe === selectedTipoEquipe))
    .map(d => String(d._teamLabel || ''))
  )).filter(e => Boolean(e) && !/^EQUIPE \d+$/.test(e) && !/^EQUIPE INE/.test(e)).sort(), [baseFilterData, selectedUnidade, selectedTipoEquipe]);

  useEffect(() => {
    if (selectedEquipe) {
      const teamMatch = baseFilterData.find(d => d._teamLabel === selectedEquipe);
      if (teamMatch) {
        if (selectedUnidade !== teamMatch._ubsLabel && teamMatch._ubsLabel !== 'SEM UNIDADE' && teamMatch._ubsLabel !== 'GERAL') {
          setSelectedUnidade(teamMatch._ubsLabel);
        }
        if (selectedTipoEquipe !== teamMatch._tipoEquipe && teamMatch._tipoEquipe !== 'Outros') {
          setSelectedTipoEquipe(teamMatch._tipoEquipe);
        }
      }
    }
  }, [selectedEquipe, baseFilterData]);

  useEffect(() => {
    if (selectedUnidade && optU.length > 0 && !optU.includes(selectedUnidade)) setSelectedUnidade('');
    else if (optU.length === 1 && selectedUnidade !== optU[0]) setSelectedUnidade(optU[0]);
  }, [optU]);

  useEffect(() => {
    if (selectedTipoEquipe && optT.length > 0 && !optT.includes(selectedTipoEquipe)) setSelectedTipoEquipe('');
    else if (optT.length === 1 && selectedTipoEquipe !== optT[0]) setSelectedTipoEquipe(optT[0]);
  }, [optT]);

  useEffect(() => {
    if (selectedEquipe && optE.length > 0 && !optE.includes(selectedEquipe)) setSelectedEquipe('');
  }, [optE]);

  const filteredData = useMemo(() => {
    let result = baseFilterData;
    if (selectedUnidade) result = result.filter(d => d._ubsLabel === selectedUnidade);
    if (selectedTipoEquipe) result = result.filter(d => d._tipoEquipe === selectedTipoEquipe);
    if (selectedEquipe) result = result.filter(d => d._teamLabel === selectedEquipe);
    return result;
  }, [baseFilterData, selectedUnidade, selectedTipoEquipe, selectedEquipe]);

  const filteredDataForTrend = useMemo(() => {
    let result = allData;
    if (selectedUnidade) result = result.filter(d => d._ubsLabel === selectedUnidade);
    if (selectedTipoEquipe) result = result.filter(d => d._tipoEquipe === selectedTipoEquipe);
    if (selectedEquipe) result = result.filter(d => d._teamLabel === selectedEquipe);
    if (selectedIndicador) result = result.filter(d => d._mappedIndicator === selectedIndicador);
    return result;
  }, [allData, selectedUnidade, selectedTipoEquipe, selectedEquipe, selectedIndicador]);

  const variableColumns = useMemo(() => {
    if (currentIndicatorColumns.length === 0) return [];
    
    const paramIndex = currentIndicatorColumns.findIndex(c => normalizeStrForFind(c).includes('parametro'));
    let candidateCols = paramIndex !== -1 ? currentIndicatorColumns.slice(paramIndex) : currentIndicatorColumns;

    return candidateCols.filter(col => {
      if ([colUnidade, colEquipe, colIne, colTipoEquipe, colMes].includes(col)) return false;
      const lower = normalizeStrForFind(col);
      if (lower.includes('parametro')) return true;
      if (isMetadataColumn(col)) return false;
      return true; 
    });
  }, [currentIndicatorColumns, colUnidade, colEquipe, colIne, colTipoEquipe, colMes]);

  const activeVariable = useMemo(() => {
    if (!selectedIndicador || variableColumns.length === 0) return '';
    const indL = String(selectedIndicador).toLowerCase();
    
    const findCol = (keywords: string[]) => variableColumns.find(c => {
       const norm = normalizeStrForFind(c);
       return keywords.some(k => norm.includes(k));
    });

    let found;
    if (indL.includes('c2') || indL.includes('c4') || indL.includes('c5')) {
       found = findCol(['razao entre', 'razao']);
    } else if (indL.includes('c7')) {
       found = findCol(['somatorio', 'boa pratica']);
    } else {
       found = findCol(['pontuacao', 'pontua']);
    }

    if (!found) found = findCol(['nota', 'resultado', 'indicador', 'perc', 'total', 'proporcao']);
    return found ? String(found) : String(variableColumns[variableColumns.length - 1] || '');
  }, [selectedIndicador, variableColumns]);

  const displayedVariables = useMemo(() => activeVariable ? [activeVariable] : [], [activeVariable]);

  const consolidatedTeams = useMemo(() => {
    if (filteredData.length === 0) return [];
    const statusCol = currentIndicatorColumns.find(c => ['classifica', 'conceito', 'qualidade', 'status', 'desempenho'].some(w => String(c).toLowerCase().includes(w)));
    const teamMap = new Map<string, any>();
    filteredData.forEach(row => {
      const key = row._teamKey;
      if (!teamMap.has(key)) teamMap.set(key, { key, uni: row._ubsLabel, eq: row._teamLabel, rows: [] });
      teamMap.get(key).rows.push(row);
    });

    return Array.from(teamMap.values()).map(group => {
      const varValues: Record<string, number | null> = {};
      variableColumns.forEach(c => { let sum = 0, cnt = 0; group.rows.forEach((r: any) => { const n = parseNum(r[String(c)]); if (n !== null) { sum += n; cnt++; } }); varValues[String(c)] = cnt > 0 ? sum / cnt : null; });
      let weightSum = 0, weightCount = 0;
      group.rows.forEach((r: any) => { 
        let st = statusCol ? normStat(r[String(statusCol)]) : null;
        if (!st) for (let i = currentIndicatorColumns.length - 1; i >= 0; i--) if ((st = normStat(r[String(currentIndicatorColumns[i])]))) break;
        if (!st && activeVariable) {
          const v = parseNum(r[activeVariable]);
          st = calcStatus(v, r._mappedIndicator);
        }
        if (st) { weightSum += S_WEIGHTS[st]; weightCount++; } 
      });
      const avgWeight = weightCount > 0 ? Math.round(weightSum / weightCount) : null;
      const finalStatus = avgWeight ? W_STATUS[Math.min(4, Math.max(1, avgWeight))] : null;
      const baseRow = { ...group.rows[group.rows.length - 1] };
      variableColumns.forEach(c => { if (varValues[String(c)] !== null) baseRow[String(c)] = formatVal(Number(varValues[String(c)]!.toFixed(2))); });
      if (statusCol && finalStatus) baseRow[String(statusCol)] = finalStatus;
      const uniqueMonths = Array.from(new Set(group.rows.map((r: any) => r._mesOrigem)));
      return { key: group.key, unidade: group.uni, equipe: group.eq, varValues, status: finalStatus, monthsLabel: uniqueMonths.join(', '), consolidatedRow: baseRow };
    });
  }, [filteredData, currentIndicatorColumns, variableColumns, activeVariable]);

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
        if (stA && stB) { const rankA = S_WEIGHTS[stA] || 0, rankB = S_WEIGHTS[stB] || 0; if (rankA !== rankB) return sortConfig.direction === 'asc' ? rankA - rankB : rankB - rankA; }
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

  const statusCounts = { 'Ótimo': 0, 'Bom': 0, 'Satisfatório': 0, 'Regular': 0 };
  consolidatedTeams.forEach(t => { if (t.status && statusCounts[t.status as keyof typeof statusCounts] !== undefined) statusCounts[t.status as keyof typeof statusCounts]++; });

  const variablesChartData = variableColumns.map(col => {
    let sum = 0, count = 0;
    consolidatedTeams.forEach(team => { if (team.varValues[String(col)] !== null && team.varValues[String(col)] !== undefined) { sum += team.varValues[String(col)]!; count++; } });
    return { name: String(col).length > 22 ? String(col).substring(0, 20) + '...' : String(col), fullName: String(col), valor: Number((count > 0 ? sum / count : 0).toFixed(2)), equipesContabilizadas: count };
  });

  const teamsComparisonData = useMemo(() => {
    const map = new Map<string, {sum: number, count: number, ubs: string}>();
    filteredData.forEach(r => {
      const key = r._teamLabel;
      if (!map.has(key)) map.set(key, { sum: 0, count: 0, ubs: r._ubsLabel });
      if (activeVariable) {
        const v = parseNum(r[activeVariable]);
        if (v !== null) { map.get(key)!.sum += v; map.get(key)!.count++; }
      }
    });

    return Array.from(map.entries()).map(([name, val]) => ({
      name: String(name).replace('EQUIPE', 'EQ').substring(0, 15) + (String(name).length > 15 ? '...' : ''),
      fullName: `${name} (${val.ubs})`,
      valor: val.count > 0 ? Number((val.sum / val.count).toFixed(2)) : 0
    })).sort((a, b) => b.valor - a.valor);
  }, [filteredData, activeVariable]);

  const lineChartData = (() => {
    const byMonth = new Map<string, any[]>();
    optM.forEach(m => byMonth.set(m, []));
    
    filteredDataForTrend.forEach(row => { 
      const mes = String(row[colMes] || 'Geral'); 
      if (byMonth.has(mes)) byMonth.get(mes)!.push(row); 
    });
    
    return Array.from(byMonth.entries()).map(([mes, rows]) => {
      if (rows.length === 0) return { name: mes, valor: null };
      const tMap = new Map<string, { sum: number; count: number }>();
      rows.forEach(r => {
        if (!tMap.has(r._teamKey)) tMap.set(r._teamKey, { sum: 0, count: 0 });
        if (activeVariable) { const n = parseNum(r[activeVariable]); if (n !== null) { tMap.get(r._teamKey)!.sum += n; tMap.get(r._teamKey)!.count++; } } 
        else { tMap.get(r._teamKey)!.sum++; tMap.get(r._teamKey)!.count++; }
      });
      let sT = 0, cT = 0;
      tMap.forEach(t => { if (t.count > 0) { sT += t.sum / t.count; cT++; } });
      return { name: mes, valor: Number((activeVariable ? (cT > 0 ? sT / cT : 0) : tMap.size).toFixed(2)) };
    });
  })();

  const displayColumns = variableColumns;

  const summaryVariablesMap: Record<string, string> = {};
  displayColumns.forEach(col => {
    let sum = 0, count = 0;
    consolidatedTeams.forEach(team => { if (team.varValues[String(col)] !== null && team.varValues[String(col)] !== undefined) { sum += team.varValues[String(col)]!; count++; } });
    summaryVariablesMap[String(col)] = formatVal(Number((count > 0 ? sum / count : 0).toFixed(2)));
  });

  const statusPieData = [
    { name: 'Ótimo', total: statusCounts['Ótimo'], color: '#3B82F6' },
    { name: 'Bom', total: statusCounts['Bom'], color: '#10B981' },
    { name: 'Satisfatório', total: statusCounts['Satisfatório'], color: '#F59E0B' },
    { name: 'Regular', total: statusCounts['Regular'], color: '#EF4444' }
  ].filter(d => d.total > 0);

  if (files.length === 0) return (
    <div className="h-[80vh] flex flex-col items-center justify-center text-slate-400">
      <BarChart2 className="h-20 w-20 mb-4 opacity-20"/>
      <h2 className="text-xl font-medium text-slate-600">Nenhum dado hospedado no Firebase.</h2>
      <p className="mt-2 text-sm">Faça login e vá ao módulo "Importar Dados" para iniciar.</p>
    </div>
  );

  if (isCorrupted) {
    return (
      <div className="flex flex-col items-center justify-center h-full min-h-[60vh] p-8 bg-white rounded-xl shadow-sm border border-red-200 m-6">
        <AlertOctagon className="h-16 w-16 text-red-500 mb-4" />
        <h2 className="text-xl font-bold text-slate-800 mb-2">Dados Corrompidos na Nuvem Detetados</h2>
        <p className="text-sm text-slate-500 mb-6 max-w-md text-center">O sistema detetou que as planilhas guardadas atualmente estão corrompidas, devido a cabeçalhos mal formatados na importação anterior.</p>
        <div className="bg-slate-50 p-5 rounded-lg border border-slate-200 text-sm text-slate-700 space-y-3 w-full max-w-md">
          <p className="font-bold text-slate-800">Para repor o funcionamento normal do painel:</p>
          <p>1. Faça o Login Restrito.</p>
          <p>2. Vá ao menu esquerdo e clique em <b>"Importar Dados"</b>.</p>
          <p>3. Clique no botão vermelho <b>"Apagar TUDO da Nuvem"</b>.</p>
          <p>4. Faça o upload das planilhas originais novamente.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4 relative">
      <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center text-sm font-medium text-slate-700">
            <Filter className="h-4 w-4 mr-2 text-blue-600" /> Filtros e Variáveis:
            {(selectedUnidade || selectedEquipe || selectedTipoEquipe) && (
              <button onClick={handleClearFilters} title="Restaurar todos os filtros" className="ml-4 flex items-center px-2.5 py-1 bg-slate-100 hover:bg-red-50 text-slate-500 hover:text-red-600 text-[10px] uppercase font-bold rounded border border-slate-200 hover:border-red-200 transition-colors">
                <X className="h-3 w-3 mr-1" /> Limpar Filtros
              </button>
            )}
          </div>
          <div className={`px-3 py-1 rounded-full text-xs font-semibold border ${isSingleTeam ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-blue-50 text-blue-800 border-blue-200'}`}>{isSingleTeam ? `🎯 Reais: ${consolidatedTeams[0]?.equipe || selectedEquipe}` : `📊 Média das ${totalEquipesUnicas} Equipes`}</div>
        </div>
        
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          <select value={selectedUnidade} onChange={e => { setSelectedUnidade(e.target.value); setSelectedEquipe(''); }} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50">
            <option value="">🏥 Todas as Unidades</option>
            {optU.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
          <select value={selectedEquipe} onChange={e => setSelectedEquipe(e.target.value)} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50">
            <option value="">👥 Todas as Equipes</option>
            {optE.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
          <select value={selectedTipoEquipe} onChange={e => { setSelectedTipoEquipe(e.target.value); setSelectedEquipe(''); }} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50">
            <option value="">🏷️ Todos tipos de equipe</option>
            {optT.map(t => <option key={t} value={t}>Apenas {t}</option>)}
          </select>

          <select value={selectedIndicador} onChange={e => setSelectedIndicador(e.target.value)} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50 font-medium text-slate-900">
            {optI.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
          <select value={selectedMes} onChange={e => setSelectedMes(e.target.value)} className="w-full p-2 border border-slate-200 rounded-lg text-sm bg-slate-50 font-medium text-slate-900">
            {optM.map(o => <option key={o} value={o}>📅 {o}</option>)}
          </select>
          <div className="w-full px-3 py-1.5 border border-blue-300 rounded-lg text-sm bg-blue-50/60 text-blue-950 flex flex-col justify-center overflow-hidden h-[42px]">
            <span className="text-[9px] text-blue-600 font-bold uppercase tracking-wider leading-none mb-0.5">Variável Automática Selecionada</span>
            <span className="truncate text-xs font-semibold" title={activeVariable || 'Nenhuma variável'}>{activeVariable || 'Aguardando seleção...'}</span>
          </div>
        </div>
      </section>

      <section className="grid grid-cols-2 lg:grid-cols-5 gap-3">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-3">
          <div className="bg-slate-100 p-2.5 rounded-full shrink-0"><Users className="h-5 w-5 text-slate-700" /></div>
          <div><p className="text-xs text-slate-700 font-semibold uppercase">Equipes</p><p className="text-xl font-bold text-slate-800">{totalEquipesUnicas}</p></div>
        </div>
        <div onClick={() => statusCounts['Ótimo'] > 0 && setActiveModalStatus('Ótimo')} className={`bg-white p-4 rounded-xl border border-blue-200 shadow-sm flex items-center space-x-3 ${statusCounts['Ótimo'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-blue-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-blue-100 p-2.5 rounded-full shrink-0"><Award className="h-5 w-5 text-blue-600" /></div>
          <div><p className="text-xs text-blue-700 font-semibold uppercase">Ótimo</p><p className="text-xl font-bold text-blue-600">{statusCounts['Ótimo']}</p></div>
        </div>
        <div onClick={() => statusCounts['Bom'] > 0 && setActiveModalStatus('Bom')} className={`bg-white p-4 rounded-xl border border-emerald-200 shadow-sm flex items-center space-x-3 ${statusCounts['Bom'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-emerald-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-emerald-100 p-2.5 rounded-full shrink-0"><CheckCircle className="h-5 w-5 text-emerald-600" /></div>
          <div><p className="text-xs text-emerald-700 font-semibold uppercase">Bom</p><p className="text-xl font-bold text-emerald-600">{statusCounts['Bom']}</p></div>
        </div>
        <div onClick={() => statusCounts['Satisfatório'] > 0 && setActiveModalStatus('Satisfatório')} className={`bg-white p-4 rounded-xl border border-amber-200 shadow-sm flex items-center space-x-3 ${statusCounts['Satisfatório'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-amber-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-amber-100 p-2.5 rounded-full shrink-0"><AlertTriangle className="h-5 w-5 text-amber-600" /></div>
          <div><p className="text-xs text-amber-700 font-semibold uppercase">Satisfatório</p><p className="text-xl font-bold text-amber-600">{statusCounts['Satisfatório']}</p></div>
        </div>
        <div onClick={() => statusCounts['Regular'] > 0 && setActiveModalStatus('Regular')} className={`bg-white p-4 rounded-xl border border-red-200 shadow-sm flex items-center space-x-3 col-span-2 lg:col-span-1 ${statusCounts['Regular'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-red-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-red-100 p-2.5 rounded-full shrink-0"><AlertOctagon className="h-5 w-5 text-red-600" /></div>
          <div><p className="text-xs text-red-700 font-semibold uppercase">Regular</p><p className="text-xl font-bold text-red-600">{statusCounts['Regular']}</p></div>
        </div>
      </section>

      {variablesChartData.length > 0 && (
        <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center mb-3"><div className="flex items-center space-x-2"><BarChart2 className="h-4 w-4 text-blue-600" /><h3 className="text-sm font-bold text-slate-800">{isSingleTeam ? `Valores Reais (${consolidatedTeams[0]?.equipe || selectedEquipe})` : 'Resumo das Variáveis (Média)'}</h3></div></div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
            {variablesChartData.map((v, i) => (
              <div key={i} className={`p-4 rounded-lg border transition-all ${v.fullName === activeVariable ? 'border-blue-600 bg-blue-50/70 shadow-sm' : 'border-slate-200 bg-slate-50/60'}`}>
                <p className="text-xs text-slate-600 font-medium leading-tight" title={v.fullName}>{v.fullName}</p>
                <p className={`text-2xl font-bold mt-2 ${v.fullName === activeVariable ? 'text-blue-900' : 'text-slate-800'}`}>{formatVal(v.valor)}</p>
                <p className="text-[10px] text-slate-500 mt-1">{isSingleTeam ? 'Valor real da equipe' : `Média de ${v.equipesContabilizadas} equipes`}</p>
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm lg:col-span-2">
          <h3 className="text-base font-bold text-slate-800 mb-1">Comparativo por Equipe</h3>
          <div style={{ height: 290 }}>
            <ResponsiveContainer>
              <BarChart data={teamsComparisonData} margin={{ top: 30, right: 20, left: 0, bottom: 80 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" stroke="#475569" fontSize={9} tickLine={false} axisLine={false} angle={-45} textAnchor="end" interval={0} />
                <YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} />
                <Tooltip formatter={(val:number, _n:string, p:any) => [formatVal(val), p.payload.fullName || activeVariable]} />
                <Bar dataKey="valor" radius={[4, 4, 0, 0]} maxBarSize={40}>
                  {teamsComparisonData.map((e, i) => <Cell key={i} fill={'#1D4ED8'} />)}
                  <LabelList dataKey="valor" position="top" fontSize={10} fontWeight="bold" formatter={(v:number) => formatVal(v)} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-base font-bold text-slate-800 mb-1">Classificação Final</h3>
          <div style={{ height: 290 }}>
            {statusPieData.length ? (
              <ResponsiveContainer>
                <PieChart margin={{ top: 20, right: 30, left: 30, bottom: 20 }}>
                  <Pie data={statusPieData} cx="50%" cy="50%" outerRadius={65} dataKey="total" nameKey="name" label={({ name, value }) => `${name}: ${value}`}>
                    {statusPieData.map((e, i) => <Cell key={i} fill={e.color} />)}
                  </Pie>
                  <Tooltip />
                  <Legend verticalAlign="bottom" height={36} iconType="circle" />
                </PieChart>
              </ResponsiveContainer>
            ) : <div className="h-full flex items-center justify-center text-slate-400 text-sm">Sem classificação</div>}
          </div>
        </div>
        
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm lg:col-span-3">
          <h3 className="text-base font-bold text-slate-800 mb-1">Evolução Mensal ({activeVariable || 'Variável Focada'})</h3>
          <div style={{ height: 280 }}>
            <ResponsiveContainer>
              <LineChart data={lineChartData} margin={{ top: 30, right: 30, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" stroke="#475569" fontSize={11} tickLine={false} axisLine={false} padding={{ left: 20, right: 20 }} />
                <YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} />
                <Tooltip formatter={(v:number) => [formatVal(v), isSingleTeam ? 'Equipe' : 'Média']} />
                <Line type="monotone" dataKey="valor" stroke="#10B981" strokeWidth={3} dot={{ r: 6, fill: '#10B981', strokeWidth: 2, stroke: '#fff' }} connectNulls={false}>
                  <LabelList dataKey="valor" position="top" offset={12} fontSize={12} fontWeight="bold" formatter={(v:number) => formatVal(v)} />
                </Line>
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      <section className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="flex justify-between items-center mb-3">
          <h3 className="text-base font-bold text-slate-800 flex items-center"><FileText className="h-5 w-5 mr-2 text-blue-600" /> Relatório Completo (Planilha)</h3>
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
                {displayColumns.map(c => (
                  <th key={String(c)} onClick={() => requestSort(String(c))} className={`px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group ${String(c) === activeVariable ? 'bg-blue-100' : ''}`}><div className="flex items-center">{String(c)} {sortConfig?.key === String(c) ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100 shrink-0" />}</div></th>
                ))}
              </tr>
            </thead>
            <tbody>
              {!isSingleTeam && totalEquipesUnicas > 1 && (
                <tr className="bg-blue-50/90 border-b border-blue-200 font-bold"><td className="px-3 py-2">RESUMO</td><td className="px-3 py-2">MÉDIA DAS EQUIPES</td>{displayColumns.map(c => <td key={String(c)} className="px-3 py-2">{summaryVariablesMap[String(c)] !== undefined ? summaryVariablesMap[String(c)] : '-'}</td>)}</tr>
              )}
              {sortedTableData.slice(0, 150).map((r, i) => (
                <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-3 py-2 text-blue-800 bg-blue-50/20">{r._ms}</td>
                  <td className="px-3 py-2 font-semibold bg-blue-50/10">{r._eq}</td>
                  {displayColumns.map(c => {
                    let rawVal = String(r[String(c)] ?? '').trim();
                    if (rawVal === '-' || rawVal === '–' || rawVal === '—') rawVal = '0'; 

                    let st = null;
                    if (String(c).toLowerCase().includes('classifica') || String(c).toLowerCase().includes('conceito') || String(c).toLowerCase().includes('qualidade') || String(c).toLowerCase().includes('status')) { st = normStat(rawVal); }
                    if (!st && String(c) === activeVariable) { const v = parseNum(rawVal); st = calcStatus(v, r._mappedIndicator); }
                    return <td key={String(c)} className={`px-3 py-2 text-slate-600 ${String(c) === activeVariable ? 'bg-blue-50/40 font-semibold' : ''}`}>{st ? <span className="px-2 py-0.5 rounded-full text-[11px] font-bold text-white" style={{ backgroundColor: S_COLORS[st] }}>{rawVal}</span> : rawVal}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* MODAL DE DETALHES DOS STATUS */}
      {activeModalStatus && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[85vh]">
            <div className={`p-4 flex justify-between items-center text-white ${activeModalStatus === 'Ótimo' ? 'bg-blue-600' : activeModalStatus === 'Bom' ? 'bg-emerald-600' : activeModalStatus === 'Satisfatório' ? 'bg-amber-500' : 'bg-red-600'}`}>
              <div className="flex items-center space-x-2">
                {activeModalStatus === 'Ótimo' && <Award className="h-5 w-5" />}
                {activeModalStatus === 'Bom' && <CheckCircle className="h-5 w-5" />}
                {activeModalStatus === 'Satisfatório' && <AlertTriangle className="h-5 w-5" />}
                {activeModalStatus === 'Regular' && <AlertOctagon className="h-5 w-5" />}
                <h3 className="font-bold text-lg">Equipes em Status {activeModalStatus} ({statusCounts[activeModalStatus as keyof typeof statusCounts]})</h3>
              </div>
              <button onClick={() => setActiveModalStatus(null)} className="p-1 hover:bg-white/20 rounded-lg transition-colors"><X className="h-5 w-5" /></button>
            </div>
            <div className="p-0 overflow-y-auto flex-1 bg-slate-50">
              <ul className="divide-y divide-slate-200">
                {consolidatedTeams
                  .filter(t => t.status === activeModalStatus)
                  .sort((a, b) => {
                    const valA = parseNum(a.consolidatedRow[activeVariable]) ?? 0;
                    const valB = parseNum(b.consolidatedRow[activeVariable]) ?? 0;
                    return valB - valA;
                  })
                  .map((t, idx) => (
                  <li key={idx} className="p-4 hover:bg-slate-100 transition-colors flex items-center justify-between">
                    <div>
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-0.5 flex items-center"><Layout className="h-3 w-3 mr-1"/> {t.unidade}</p>
                      <p className="text-sm font-semibold text-slate-800 flex items-center"><Users className="h-4 w-4 mr-1.5 text-slate-400"/> {t.equipe}</p>
                    </div>
                    <div className="text-right ml-4 shrink-0">
                      <span className={`inline-block px-2.5 py-1 rounded text-xs font-bold text-white shadow-sm ${
                        activeModalStatus === 'Ótimo' ? 'bg-blue-500' : 
                        activeModalStatus === 'Bom' ? 'bg-emerald-500' : 
                        activeModalStatus === 'Satisfatório' ? 'bg-amber-500' : 
                        'bg-red-500'
                      }`}>
                        {t.consolidatedRow[activeVariable] ?? '-'}
                      </span>
                      <p className="text-[8px] text-slate-400 uppercase font-semibold mt-1 max-w-[80px] truncate" title={activeVariable}>
                        {activeVariable || 'Valor'}
                      </p>
                    </div>
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

// ==========================================
// MÓDULO: RANKING GERAL
// ==========================================
function ModuleRanking({ files, bgClass }: any) {
  const rawAllData = useMemo(() => files.flatMap((f: any) => f.data), [files]);
  const columns = useMemo(() => Array.from(new Set(rawAllData.flatMap(Object.keys))).filter(c => c && !String(c).startsWith('_') && !String(c).toUpperCase().includes('EMPTY') && String(c).trim() !== ''), [rawAllData]);

  const colUnidade = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs');
  }) || null, [columns]);

  const colIne = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s === 'ine' || s.includes('codigo ine');
  }) || null, [columns]);

  const colTipoEquipe = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('tipo do ine') || s.includes('tipo de equipe') || s === 'tipo';
  }) || null, [columns]);

  const colEquipe = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c); return s.includes('nome da equipe') || s === 'equipe' || s === 'nome';
  }) || null, [columns]);

  const colMes = '_mesOrigem';
  const colIndicador = useMemo(() => columns.slice(0, 8).find(c => ['indicador', 'condição', 'condicao'].includes(String(c).toLowerCase())) || '_indicadorOrigem', [columns]);

  const isCorrupted = rawAllData.length > 0 && (!colUnidade && !colEquipe);

  const allData = useMemo(() => {
    if (isCorrupted) return [];
    const enriched: any[] = [];
    rawAllData.forEach((r, idx) => {
      let n = colEquipe ? String(r[colEquipe] || '').replace(/\s+/g, ' ').trim() : '';
      const ic = colIne ? normIne(r[colIne]) : '';
      let lbl = n ? n.toUpperCase() : (ic ? `EQUIPE INE ${ic}` : `EQUIPE ${idx + 1}`);
      const ubs = colUnidade ? String(r[colUnidade] || 'Sem Unidade').replace(/\s+/g, ' ').trim().toUpperCase() : 'GERAL';
      
      let tipoStr = 'Outros';
      const typeVal = colTipoEquipe ? String(r[colTipoEquipe] || '').toUpperCase() : '';
      const nameVal = lbl.toUpperCase();

      if (typeVal.includes('ESF') || nameVal.includes('ESF') || typeVal.includes('FAMILIA') || nameVal.includes('FAMILIA')) {
          tipoStr = 'ESF';
      } else if (typeVal.includes('EAP') || nameVal.includes('EAP') || typeVal.includes('PRIMARIA') || nameVal.includes('PRIMARIA')) {
          tipoStr = 'EAP';
      }

      const indOrig = String(colIndicador !== '_indicadorOrigem' && r[colIndicador] ? r[colIndicador] : r._indicadorOrigem || '');
      const mappedInd = mapIndicatorName(indOrig);

      enriched.push({ 
        ...r, 
        _teamKey: `UBS_${normTxt(ubs)}__EQ_${normTxt(lbl)}`, 
        _teamLabel: lbl, 
        _ubsLabel: ubs,
        _tipoEquipe: tipoStr,
        _mappedIndicator: mappedInd
      });
    });
    return enriched;
  }, [rawAllData, colUnidade, colEquipe, colIne, colTipoEquipe, colIndicador, isCorrupted]);

  const optM = useMemo(() => Array.from(new Set(allData.map(d => String(d._mesOrigem || '')))).filter(Boolean).sort((a, b) => (monthOrder[a] || 99) - (monthOrder[b] || 99)), [allData]);
  const optI = useMemo(() => {
    return Array.from(new Set(allData.map(d => String(d._mappedIndicator || ''))))
      .filter(Boolean)
      .sort((a, b) => {
        if (a.toLowerCase().includes('vinculo') || a.toLowerCase().includes('vínculo')) return -1;
        if (b.toLowerCase().includes('vinculo') || b.toLowerCase().includes('vínculo')) return 1;
        return a.localeCompare(b);
      });
  }, [allData]);
  
  const [selectedMes, setSelectedMes] = useState('');
  const [sortConfigInd, setSortConfigInd] = useState<{ key: string, direction: 'asc' | 'desc' } | null>(null);

  useEffect(() => {
    if (!selectedMes && optM.length > 0) {
      setSelectedMes(optM[optM.length - 1]);
    }
  }, [optM, selectedMes]);

  const filteredData = useMemo(() => {
    return allData.filter(d => String(d[colMes]) === selectedMes);
  }, [allData, selectedMes, colMes]);

  const indicatorVariablesLookup = useMemo(() => {
    const lookup: Record<string, string> = {};
    optI.forEach(ind => {
      const indRows = allData.filter(d => d._mappedIndicator === ind);
      if (indRows.length > 0) {
        const keys = new Set<string>();
        indRows.slice(0, 10).forEach(r => Object.keys(r).forEach(k => {
           if (!k.startsWith('_') && !k.toUpperCase().includes('EMPTY')) keys.add(k);
        }));
        const indCols = Array.from(keys);
        
        const paramIndex = indCols.findIndex(c => normalizeStrForFind(c).includes('parametro'));
        const validCols = paramIndex !== -1 ? indCols.slice(paramIndex) : indCols;
        
        const indL = ind.toLowerCase();
        const findCol = (keywords: string[]) => validCols.find(c => {
           const norm = normalizeStrForFind(c);
           return keywords.some(k => norm.includes(k));
        });
        
        let found;
        if (indL.includes('c2') || indL.includes('c4') || indL.includes('c5')) {
           found = findCol(['razao entre', 'razao']);
        } else if (indL.includes('c7')) {
           found = findCol(['somatorio', 'boa pratica']);
        } else {
           found = findCol(['pontuacao', 'pontua']);
        }
        if (!found) found = findCol(['nota', 'resultado', 'indicador', 'perc', 'total', 'proporcao']);
        lookup[ind] = found || '';
      }
    });
    return lookup;
  }, [allData, optI]);

  const getActiveVarForIndicator = (ind: string) => {
    return indicatorVariablesLookup[ind] || '';
  };

  const getStatusForRow = (row: any, ind: string) => {
    const rowKeys = Object.keys(row || {}).filter(k => !String(k).startsWith('_'));
    const statusCol = rowKeys.find(c => ['classifica', 'conceito', 'qualidade', 'status', 'desempenho'].some(w => String(c).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').includes(w)));
    if (statusCol) {
      const st = normStat(row[statusCol]);
      if (st) return st;
    }
    const vCol = getActiveVarForIndicator(ind);
    if (vCol) return calcStatus(parseNum(row[String(vCol)]), ind);
    return null;
  };

  const { ubsRanking, indRanking } = useMemo(() => {
    const ubsScores = new Map<string, { sum: number, count: number, statusCounts: Record<string, number> }>();
    const indScores = new Map<string, { sum: number, count: number, statusCounts: Record<string, number> }>();

    optI.forEach(ind => {
      indScores.set(ind, { sum: 0, count: 0, statusCounts: {'Ótimo':0, 'Bom':0, 'Satisfatório':0, 'Regular':0} });
    });

    filteredData.forEach(row => {
      const ind = row._mappedIndicator;
      const st = getStatusForRow(row, ind);

      if (st && S_WEIGHTS[st]) {
        const w = S_WEIGHTS[st];
        const ubs = row._ubsLabel;

        if (ubs !== 'SEM UNIDADE' && ubs !== 'GERAL') {
            if (!ubsScores.has(ubs)) ubsScores.set(ubs, { sum: 0, count: 0, statusCounts: {'Ótimo':0, 'Bom':0, 'Satisfatório':0, 'Regular':0} });
            ubsScores.get(ubs)!.sum += w;
            ubsScores.get(ubs)!.count++;
            ubsScores.get(ubs)!.statusCounts[st]++;
        }

        if (indScores.has(ind)) {
           indScores.get(ind)!.sum += w;
           indScores.get(ind)!.count++;
           indScores.get(ind)!.statusCounts[st]++;
        }
      }
    });

    const sortScores = (map: Map<string, any>) => Array.from(map.entries()).map(([name, data]) => ({ name, avg: data.count > 0 ? data.sum / data.count : 0, ...data.statusCounts })).sort((a, b) => b.avg - a.avg);

    return { ubsRanking: sortScores(ubsScores), indRanking: sortScores(indScores) };
  }, [filteredData, optI]);

  const top10UBS = ubsRanking.slice(0, 10);
  const bottom10UBS = [...ubsRanking].reverse().slice(0, 10);

  const requestSortInd = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfigInd && sortConfigInd.key === key && sortConfigInd.direction === 'asc') direction = 'desc';
    setSortConfigInd({ key, direction });
  };

  const sortedIndRanking = useMemo(() => {
    let sortableData = [...indRanking];
    if (sortConfigInd !== null) {
      sortableData.sort((a: any, b: any) => {
        const valA = a[sortConfigInd.key];
        const valB = b[sortConfigInd.key];
        if (valA < valB) return sortConfigInd.direction === 'asc' ? -1 : 1;
        if (valA > valB) return sortConfigInd.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return sortableData;
  }, [indRanking, sortConfigInd]);

  if (files.length === 0) return (
    <div className="h-[80vh] flex flex-col items-center justify-center text-slate-400">
      <Trophy className="h-20 w-20 mb-4 opacity-20"/>
      <h2 className="text-xl font-medium text-slate-600">Nenhum dado para gerar Ranking.</h2>
      <p className="mt-2 text-sm">Importe as planilhas primeiro.</p>
    </div>
  );

  if (isCorrupted) {
    return (
      <div className="flex flex-col items-center justify-center h-full min-h-[60vh] p-8 bg-white rounded-xl shadow-sm border border-red-200 m-6">
        <AlertOctagon className="h-16 w-16 text-red-500 mb-4" />
        <h2 className="text-xl font-bold text-slate-800 mb-2">Dados Corrompidos na Nuvem Detetados</h2>
        <p className="text-sm text-slate-500 mb-6 max-w-md text-center">O sistema detetou que as planilhas guardadas atualmente estão corrompidas, devido a cabeçalhos mal formatados na importação anterior.</p>
        <div className="bg-slate-50 p-5 rounded-lg border border-slate-200 text-sm text-slate-700 space-y-3 w-full max-w-md">
          <p className="font-bold text-slate-800">Para repor o funcionamento normal do painel:</p>
          <p>1. Vá ao menu lateral esquerdo e clique em <b>"Importar Dados"</b>.</p>
          <p>2. Clique no botão vermelho <b>"Apagar TUDO da Nuvem"</b> (ao lado da lista de planilhas).</p>
          <p>3. Selecione e faça o upload das planilhas originais novamente.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-4">
        <div className="flex items-center text-lg font-bold text-slate-800">
          <div className={`p-2 rounded-lg mr-3 ${bgClass} bg-opacity-10`}><Trophy className={`h-6 w-6 text-blue-600`} /></div>
          Ranking de Desempenho Global
        </div>
        <div className="w-full sm:w-64">
          <select value={selectedMes} onChange={e => setSelectedMes(e.target.value)} className="w-full p-2 border border-slate-300 rounded-lg text-sm bg-white font-medium shadow-sm">
            {optM.map(o => <option key={o} value={o}>🏆 Mês: {o}</option>)}
          </select>
        </div>
      </section>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="bg-blue-50 p-4 border-b border-blue-100 flex items-center">
             <Award className="h-5 w-5 text-blue-600 mr-2" />
             <h3 className="font-bold text-blue-900">TOP 10 Melhores UBS</h3>
          </div>
          <div className="divide-y divide-slate-100">
             {top10UBS.map((ubs, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className={`font-black text-lg w-6 text-center ${i === 0 ? 'text-amber-400' : i === 1 ? 'text-slate-400' : i === 2 ? 'text-amber-600' : 'text-slate-300'}`}>#{i + 1}</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-1">{ubs.name}</p>
                     <div className="flex space-x-1.5">
                       {ubs['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{ubs['Ótimo']} Ótimo</span>}
                       {ubs['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{ubs['Bom']} Bom</span>}
                       {ubs['Satisfatório'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{ubs['Satisfatório']} Sat.</span>}
                       {ubs['Regular'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-red-100 text-red-700 font-bold">{ubs['Regular']} Reg.</span>}
                     </div>
                   </div>
                 </div>
                 <div className="text-right">
                   <p className="text-sm font-bold text-blue-700">{ubs.avg.toFixed(2)}</p>
                   <p className="text-[9px] text-slate-400 uppercase font-semibold">Média</p>
                 </div>
               </div>
             ))}
             {top10UBS.length === 0 && <div className="p-8 text-center text-slate-400 text-sm">Sem dados suficientes.</div>}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="bg-red-50 p-4 border-b border-red-100 flex items-center">
             <AlertOctagon className="h-5 w-5 text-red-600 mr-2" />
             <h3 className="font-bold text-red-900">TOP 10 Piores UBS</h3>
          </div>
          <div className="divide-y divide-slate-100">
             {bottom10UBS.map((ubs, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className="font-black text-lg w-6 text-center text-red-300">!</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-1">{ubs.name}</p>
                     <div className="flex space-x-1.5">
                       {ubs['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{ubs['Ótimo']} Ótimo</span>}
                       {ubs['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{ubs['Bom']} Bom</span>}
                       {ubs['Satisfatório'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{ubs['Satisfatório']} Sat.</span>}
                       {ubs['Regular'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-red-100 text-red-700 font-bold">{ubs['Regular']} Reg.</span>}
                     </div>
                   </div>
                 </div>
                 <div className="text-right">
                   <p className="text-sm font-bold text-red-600">{ubs.avg.toFixed(2)}</p>
                   <p className="text-[9px] text-slate-400 uppercase font-semibold">Média</p>
                 </div>
               </div>
             ))}
             {bottom10UBS.length === 0 && <div className="p-8 text-center text-slate-400 text-sm">Sem dados suficientes.</div>}
          </div>
        </div>
      </div>

      <section className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="bg-slate-50 p-4 border-b border-slate-200 flex items-center">
           <BarChart2 className="h-5 w-5 text-blue-600 mr-2" />
           <h3 className="font-bold text-slate-800">Ranking de Indicadores (Total de Avaliações)</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left whitespace-nowrap">
            <thead className="text-slate-700 bg-slate-100 border-b border-slate-200">
              <tr>
                <th onClick={() => requestSortInd('name')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center">Indicador {sortConfigInd?.key === 'name' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1 text-blue-600" /> : <ChevronDown className="h-4 w-4 ml-1 text-blue-600" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('Ótimo')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-blue-700">Ótimo {sortConfigInd?.key === 'Ótimo' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('Bom')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-emerald-700">Bom {sortConfigInd?.key === 'Bom' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('Satisfatório')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-amber-700">Satisfatório {sortConfigInd?.key === 'Satisfatório' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('Regular')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-red-700">Regular {sortConfigInd?.key === 'Regular' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sortedIndRanking.map((ind, i) => (
                <tr key={i} className="hover:bg-slate-50 transition-colors">
                  <td className="px-5 py-3.5 font-semibold text-slate-800">{ind.name}</td>
                  <td className="px-5 py-3.5 font-bold text-blue-600 bg-blue-50/30">{ind['Ótimo']}</td>
                  <td className="px-5 py-3.5 font-bold text-emerald-600 bg-emerald-50/30">{ind['Bom']}</td>
                  <td className="px-5 py-3.5 font-bold text-amber-600 bg-amber-50/30">{ind['Satisfatório']}</td>
                  <td className="px-5 py-3.5 font-bold text-red-600 bg-red-50/30">{ind['Regular']}</td>
                </tr>
              ))}
              {sortedIndRanking.length === 0 && <tr><td colSpan={5} className="px-5 py-8 text-center text-slate-400">Sem dados para este mês.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}