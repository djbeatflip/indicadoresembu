import React, { useState, useMemo, useEffect, useRef } from 'react';
import { UploadCloud, Users, CheckCircle, AlertTriangle, AlertOctagon, Award, FileText, Search, Filter, X, BarChart2, FolderOpen, LogIn, Layout, Settings, Shield, UserPlus, LogOut, ArrowUpDown, ChevronUp, ChevronDown, Trash2, Menu, ChevronLeft, Trophy, Printer } from 'lucide-react';
import { BarChart, Bar, LineChart, AreaChart, Area, ComposedChart, Line, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LabelList, Legend } from 'recharts';
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
const S_COLORS: Record<string, string> = { 'Ótimo': '#3B82F6', 'Bom': '#10B981', 'Suficiente': '#F59E0B', 'Regular': '#EF4444' };
const S_WEIGHTS: Record<string, number> = { 'Regular': 1, 'Suficiente': 2, 'Bom': 3, 'Ótimo': 4 };
const W_STATUS: Record<number, string> = { 1: 'Regular', 2: 'Suficiente', 3: 'Bom', 4: 'Ótimo' };

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
  if (s === 'Suficiente' || s.includes('satisfat') || s === 'suficiente' || (s.includes('suficiente') && !s.includes('insuficiente'))) return 'Suficiente'; 
  if (['regular', 'ruim', 'crítico', 'critico', 'insuficiente'].includes(s) || s.includes('regular')) return 'Regular'; 
  return null; 
};

const calcStatus = (val: number, indName: string = '') => {
  const norm = String(indName || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  // 1. Regra para Vínculo e Acompanhamento Territorial:
  // > 8,5: Ótimo | 7 a 8,5: Bom | 5,0 a 6,9: Suficiente | < 5,0: Regular
  if (norm.includes('vinculo') || norm.includes('acompanhamento')) {
    // Se a planilha vier na escala de 0 a 100 (ex: 85 em vez de 8.5), converte para a base 10
    const nota = val > 10 ? val / 10 : val;

    if (nota > 8.5) return 'Ótimo';
    if (nota >= 7.0) return 'Bom';
    if (nota >= 5.0) return 'Suficiente';
    return 'Regular';
  }

  // 2. Regra para C1 (Mais Acesso):
  if (norm.includes('c1') || norm.includes('mais acesso')) {
    if (val > 50 && val <= 70) return 'Ótimo';
    if (val > 30 && val <= 50) return 'Bom';
    if (val > 10 && val <= 30) return 'Suficiente';
    return 'Regular';
  }

  // 3. Regra Geral para C2 ao C7 (Demais Indicadores):
  if (val > 75) return 'Ótimo';
  if (val > 50) return 'Bom';
  if (val > 25) return 'Suficiente';
  return 'Regular';
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
// MÓDULO DE RELATÓRIOS (TOP 10 DINÂMICO E IMPRESSÃO NATIVA)
// ==========================================
function ModuleReports({ files, bgClass }: any) {
  const [reportType, setReportType] = React.useState('1');
  const [selectedMonth, setSelectedMonth] = React.useState('');

  const mesesDisponiveis = React.useMemo(() => {
    if (!files || files.length === 0) return ['Agosto'];
    const meses = new Set<string>();
    files.forEach((f: any) => {
      if (f.data) {
        f.data.forEach((row: any) => {
          if (row._mesOrigem) meses.add(row._mesOrigem);
        });
      }
    });
    const arr = Array.from(meses);
    return arr.length > 0 ? arr : ['Agosto'];
  }, [files]);

  React.useEffect(() => {
    if (!selectedMonth && mesesDisponiveis.length > 0) {
      setSelectedMonth(mesesDisponiveis[mesesDisponiveis.length - 1]);
    }
  }, [mesesDisponiveis, selectedMonth]);

  // MOTOR INTELIGENTE: Caçador de colunas (Ignora letras maiúsculas/minúsculas)
  const getDynamicValue = (row: any, keywords: string[]) => {
    const keys = Object.keys(row);
    for (const key of keys) {
      const lowerKey = key.toLowerCase();
      if (keywords.some(word => lowerKey.includes(word))) {
        return row[key];
      }
    }
    return 'Não Identificado';
  };

  const analiseDoMes = React.useMemo(() => {
    let totalRegistros = 0;
    const unidadesCount: Record<string, number> = {};
    const equipesCount: Record<string, number> = {};

    files.forEach((f: any) => {
      if (f.data) {
        f.data.forEach((row: any) => {
          if (row._mesOrigem === selectedMonth) {
            totalRegistros++;
            
            // Procura o nome da unidade e equipe usando palavras-chave comuns nas colunas do SIAPS
            const nomeUnidade = getDynamicValue(row, ['unidade', 'estab', 'cnes', 'ubs', 'usf']);
            const nomeEquipe = getDynamicValue(row, ['equipe', 'ine']);
            
            // Só contabiliza se for um texto válido
            if (nomeUnidade && typeof nomeUnidade === 'string' && nomeUnidade !== 'Não Identificado') {
              const cleanUnidade = nomeUnidade.trim().toUpperCase();
              unidadesCount[cleanUnidade] = (unidadesCount[cleanUnidade] || 0) + 1;
            }
            
            if (nomeEquipe && typeof nomeEquipe === 'string' && nomeEquipe !== 'Não Identificado') {
              const cleanEquipe = nomeEquipe.trim().toUpperCase();
              equipesCount[cleanEquipe] = (equipesCount[cleanEquipe] || 0) + 1;
            }
          }
        });
      }
    });

    // Ordena do maior para o menor
    const topUnidades = Object.entries(unidadesCount).sort((a, b) => b[1] - a[1]);
    const topEquipes = Object.entries(equipesCount).sort((a, b) => b[1] - a[1]);

    return {
      totalRegistros,
      topUnidades,
      topEquipes,
      qtdUnidades: topUnidades.length,
      qtdEquipes: topEquipes.length
    };
  }, [files, selectedMonth]);

  const handlePrint = () => {
    window.print();
  };

  const renderReportContent = () => {
    if (analiseDoMes.totalRegistros === 0) {
      return (
        <div className="p-8 text-center text-red-500 font-bold border-2 border-red-200 bg-red-50 rounded-lg print:hidden">
          Não existem dados importados no sistema referentes ao mês de {selectedMonth}.
        </div>
      );
    }

    // EXTRAÇÃO DOS TOP 10 PARA TEXTO
    const top10Unidades = analiseDoMes.topUnidades.slice(0, 10);
    const top10Equipes = analiseDoMes.topEquipes.slice(0, 10);
    
    // As 3 Unidades que precisam melhorar (as últimas da lista)
    const unidadesParaMelhorar = analiseDoMes.topUnidades.slice(-3).reverse();

    if (reportType === '1') {
      return (
        <div className="space-y-6 text-black text-[15px] leading-relaxed text-justify">
          <p>O presente relatório apresenta a análise executiva de desempenho do município de <strong>Embu das Artes</strong> referente às metas da Atenção Primária à Saúde (APS) no ciclo de <strong>{selectedMonth}</strong>. A avaliação tem como base o cruzamento matemático dos dados importados do sistema SIAPS, englobando um processamento total de <strong>{analiseDoMes.totalRegistros.toLocaleString('pt-BR')} registros validados</strong> gerados por {analiseDoMes.qtdUnidades} estabelecimentos de saúde.</p>
          
          <div className="mt-8">
            <h3 className="font-bold text-lg mb-3 border-b border-gray-400 pb-1 text-black uppercase">1. Avaliação Consolidada (Top 3 Destaques)</h3>
            <p>No mês de {selectedMonth}, a rede municipal demonstrou resiliência operacional. Destacaram-se nominalmente as unidades <strong>{top10Unidades[0]?.[0]}</strong>, <strong>{top10Unidades[1]?.[0]}</strong> e <strong>{top10Unidades[2]?.[0]}</strong>, que assumiram a liderança no volume de consolidação e transmissão consistente de dados de produção para a base central.</p>
          </div>

          <div className="mt-8">
            <h3 className="font-bold text-lg mb-3 border-b border-gray-400 pb-1 text-black uppercase">2. Alertas Estratégicos (Baixa Consolidação)</h3>
            <p>Apesar do engajamento do município, identificou-se uma defasagem de processamento que requer intervenção gerencial. As unidades <strong>{unidadesParaMelhorar[0]?.[0]}</strong> e <strong>{unidadesParaMelhorar[1]?.[0]}</strong> registraram os menores volumes de produção validada no período, sugerindo possíveis falhas de conexão, ausência de profissionais ou acúmulo de fichas físicas retidas sem digitação.</p>
          </div>

          <div className="mt-8">
            <h3 className="font-bold text-lg mb-3 border-b border-gray-400 pb-1 text-black uppercase">3. Plano de Ação Estratégico</h3>
            <ul className="list-disc pl-6 space-y-2 mt-2">
              <li><strong>Auditoria Expressa:</strong> A coordenação deve realizar uma verificação das transmissões pendentes nas unidades em alerta, garantindo que o lote de dados do Prontuário Eletrônico seja processado antes do fecho da competência.</li>
              <li><strong>Valorização:</strong> Promover o reconhecimento institucional das equipes vinculadas à {top10Unidades[0]?.[0]}, utilizando os seus processos de trabalho como modelo para a rede.</li>
            </ul>
          </div>
        </div>
      );
    }

    if (reportType === '2') {
      return (
        <div className="space-y-6 text-black text-[15px] leading-relaxed text-justify">
          <p>O relatório gerencial abaixo detalha o ranking das <strong>10 Unidades Básicas de Saúde (UBS/USF)</strong> de Embu das Artes que apresentaram o melhor rendimento na estruturação e envio de dados da Atenção Primária durante o mês de <strong>{selectedMonth}</strong>.</p>
          
          <div className="mt-8">
            <h3 className="font-bold text-lg mb-4 border-b border-gray-400 pb-1 text-black uppercase">1. As 10 Unidades de Maior Desempenho</h3>
            <p className="mb-4">O ranqueamento a seguir reflete diretamente o engajamento das gerências locais e a diligência na rotina de digitação dos atendimentos:</p>
            
            <ol className="list-decimal pl-6 space-y-3 font-medium">
              {top10Unidades.map((unidade, index) => (
                <li key={index} className="pl-2 border-b border-gray-100 pb-1">
                  <span className="font-bold text-black">{unidade[0]}</span> 
                  <span className="text-gray-600 font-normal ml-2">— Processou {unidade[1].toLocaleString('pt-BR')} registros.</span>
                </li>
              ))}
            </ol>
            
            {top10Unidades.length === 0 && <p className="italic text-gray-500">Nenhum dado de unidade encontrado neste mês.</p>}
          </div>
          
          <div className="mt-8">
            <h3 className="font-bold text-lg mb-3 border-b border-gray-400 pb-1 text-black uppercase">2. Considerações e Recomendações</h3>
            <p>O desempenho superior da <strong>{top10Unidades[0]?.[0]}</strong> serve como referencial analítico. Recomenda-se o mapeamento do fluxo de recepção e agendamento adotado por este estabelecimento para replicação técnica nas demais {analiseDoMes.qtdUnidades - 1} unidades da rede municipal.</p>
          </div>
        </div>
      );
    }

    return (
      <div className="space-y-6 text-black text-[15px] leading-relaxed text-justify">
        <p>A presente análise detalha a microgestão em saúde, focada exclusivamente na efetividade do <strong>Top 10 das Equipes de Saúde da Família (eSF) e Atenção Primária (eAP)</strong> do município durante o ciclo de <strong>{selectedMonth}</strong>.</p>
        
        <div className="mt-8">
            <h3 className="font-bold text-lg mb-4 border-b border-gray-400 pb-1 text-black uppercase">1. As 10 Equipes com Maior Produtividade Validada</h3>
            <p className="mb-4">As equipes listadas abaixo destacaram-se pelo correto registro de CIDs, CIAPs e procedimentos, refletindo a excelência do trabalho de médicos, enfermeiros e agentes comunitários no território:</p>
            
            <ol className="list-decimal pl-6 space-y-3 font-medium">
              {top10Equipes.map((equipe, index) => (
                <li key={index} className="pl-2 border-b border-gray-100 pb-1">
                  <span className="font-bold text-black">{equipe[0]}</span> 
                  <span className="text-gray-600 font-normal ml-2">— Processou {equipe[1].toLocaleString('pt-BR')} atendimentos validados.</span>
                </li>
              ))}
            </ol>
            
            {top10Equipes.length === 0 && <p className="italic text-gray-500">Nenhum dado de equipe encontrado neste mês.</p>}
        </div>

        <div className="mt-8">
            <h3 className="font-bold text-lg mb-3 border-b border-gray-400 pb-1 text-black uppercase">2. Plano de Fortalecimento de Vínculo</h3>
            <p>As equipes do ranking demonstraram alta eficiência na vinculação de pacientes. O próximo passo estratégico é utilizar a força de trabalho da <strong>{top10Equipes[0]?.[0]}</strong> para intensificar a busca ativa exclusiva aos pacientes hipertensos e diabéticos, otimizando ainda mais a pontuação do município perante o Ministério da Saúde.</p>
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-6">
      {/* MAGIA DO CSS: As margens de 15mm aqui garantem que o PDF respeita o papel A4 nativamente sem esticar */}
      <style>
        {`
          @media print {
            @page {
              size: A4 portrait;
              margin: 15mm 15mm 15mm 15mm !important; 
            }
            body {
              -webkit-print-color-adjust: exact !important;
              print-color-adjust: exact !important;
              background-color: white !important;
            }
          }
        `}
      </style>

      {/* CABEÇALHO DE CONTROLES */}
      <div className="print:hidden bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col md:flex-row gap-4 items-end justify-between">
        <div className="flex flex-col md:flex-row gap-4 flex-1">
          <div className="w-full md:w-1/3">
            <label className="text-xs font-bold text-slate-500 tracking-wider mb-1 block uppercase">Foco da Análise</label>
            <select value={reportType} onChange={(e) => setReportType(e.target.value)} className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500 bg-slate-50 font-medium">
              <option value="1">1. Executivo Municipal (Top 3)</option>
              <option value="2">2. Gerencial: Top 10 Unidades</option>
              <option value="3">3. Operacional: Top 10 Equipes</option>
            </select>
          </div>
          <div className="w-full md:w-1/4">
            <label className="text-xs font-bold text-slate-500 tracking-wider mb-1 block uppercase">Mês de Referência</label>
            <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500 bg-slate-50 font-medium">
              {mesesDisponiveis.map(mes => (
                <option key={mes} value={mes}>{mes}</option>
              ))}
            </select>
          </div>
        </div>
        
        <button onClick={handlePrint} className={`flex items-center px-6 py-2.5 rounded-lg text-white font-bold shadow-md transition-transform hover:scale-105 active:scale-95 ${bgClass}`}>
          <Printer className="h-5 w-5 mr-2" /> Gerar PDF / Imprimir
        </button>
      </div>

      {/* ÁREA DO RELATÓRIO OFICIAL */}
      <div className="overflow-x-auto pb-10 print:overflow-visible print:pb-0">
        {/* A classe print:w-full e print:p-0 dizem ao navegador para usar todo o A4 de acordo com o @page acima */}
        <div className="bg-white border border-gray-200 p-[15mm] md:w-[210mm] min-h-[297mm] mx-auto flex flex-col font-serif shadow-lg print:shadow-none print:border-none print:w-full print:min-h-0 print:p-0">
          
          {/* CABEÇALHO OFICIAL */}
          <div className="flex items-center justify-between pb-6 mb-8 border-b-2 border-gray-300 print:border-black">
            <div className="w-1/3">
              <div className="text-[#e87a71] font-bold text-3xl leading-none">
                Cidade de <br/><span className="text-[#f1a46b]">Embu</span><br/><span className="text-[#a1c181] text-xl font-normal">das Artes</span>
              </div>
            </div>
            <div className="w-2/3 text-right text-gray-700 print:text-black">
              <h1 className="text-lg font-bold">Prefeitura da Estância Turística de Embu das Artes</h1>
              <h2 className="text-xs font-medium text-gray-500 mt-1 uppercase tracking-widest print:text-black">Estado de São Paulo</h2>
              <h3 className="text-base font-bold mt-2">Secretaria Municipal de Saúde</h3>
            </div>
          </div>

          {/* TÍTULO DO DOCUMENTO */}
          <div className="text-center mb-10">
            <h2 className="text-xl font-bold text-black border-b border-gray-400 pb-2 inline-block uppercase tracking-wide">
              Parecer Analítico de Desempenho - APS
            </h2>
            <p className="text-xs text-gray-500 mt-3 font-bold uppercase tracking-widest print:text-black">Ref: {selectedMonth} / {new Date().getFullYear()}</p>
          </div>

          {/* CORPO DO DOCUMENTO */}
          <div className="flex-1 text-black">
            {renderReportContent()}
          </div>

          {/* RODAPÉ OFICIAL */}
          <div className="mt-16 pt-6 border-t-2 border-gray-300 print:border-black flex justify-between items-end text-[10px] text-gray-600 print:text-black">
            <div className="text-left font-semibold space-y-1">
              <p>Prefeitura da Estância Turística de Embu das Artes</p>
              <p>Secretaria Municipal de Saúde</p>
              <p className="text-gray-500 print:text-black">Rua Andrônico dos Prazeres Gonçalves, 114 — Centro</p>
            </div>
            <div className="text-right font-semibold space-y-1">
              <p>Telefone: (11) 4785-3500</p>
              <p>indicadoresembu@gmail.com</p>
              <p>projetos.embu@gmail.com</p>
            </div>
          </div>
          
        </div>
      </div>
    </div>
  );
}

// ==========================================
// APLICAÇÃO ENCAPSULADA (RESTAURADA: MODO RESTRITO/LOGIN OBRIGATÓRIO)
// ==========================================
function MainApp() {
  const [isLoaded, setIsLoaded] = useState(false);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [activeModule, setActiveModule] = useState('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [files, setFiles] = useState<ImportedFile[]>([]);
  const [usersList, setUsersList] = useState<User[]>([]);
  const [sysTheme, setSysTheme] = useState('blue');

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

          const filesSnap = await getDocs(collection(db, 'dados_siaps'));
          const groupedFiles: ImportedFile[] = [];
          
          filesSnap.docs.forEach(d => {
             const docData = d.data();
             // Verifica se é o formato novo (comprimido em JSON)
             if (docData.dadosJSON) {
               try {
                 const parsedData = JSON.parse(docData.dadosJSON);
                 groupedFiles.push({
                   id: d.id,
                   name: docData.nomeArquivo || 'Planilha_Importada.xlsx',
                   data: parsedData
                 });
               } catch (e) {
                 console.error("Erro ao descomprimir JSON", e);
               }
             } else {
               // Formato antigo (legado): apenas para o caso de ainda ter dados velhos
               const fileName = docData._arquivoOrigem + '.xlsx';
               let existing = groupedFiles.find(f => f.name === fileName);
               if (!existing) {
                 existing = { id: fileName, name: fileName, data: [] };
                 groupedFiles.push(existing);
               }
               existing.data.push(docData);
             }
          });
          setFiles(groupedFiles);

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

  // BLOQUEIO RESTRITO ATIVADO: Só passa quem fizer login
  if (!currentUser) return <LoginScreen bgClass={bgClass} />;

  const handleLogout = async () => { await signOut(auth); setFiles([]); };

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
          {isSidebarOpen && <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-2">Módulos</p>}
          <button onClick={() => setActiveModule('dashboard')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'dashboard' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Dashboard SIAPS" : ""}>
            <Layout className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Dashboard SIAPS</span>}
          </button>
          
          <button onClick={() => setActiveModule('ranking')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'ranking' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Ranking" : ""}>
            <Trophy className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Ranking</span>}
          </button>

{/* BOTÃO DE RELATÓRIOS (INATIVO TEMPORARIAMENTE) */}
<button disabled className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl font-medium text-sm text-slate-500 opacity-50 cursor-not-allowed`} title={!isSidebarOpen ? "Relatórios (Inativo)" : ""}>
            <Printer className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Relatórios <span className="text-[10px] ml-2 bg-slate-800 px-2 py-0.5 rounded-full">Em breve</span></span>}
          </button>

          {(currentUser.role === 'Administrador' || currentUser.role === 'Coordenador') && (
            <button onClick={() => setActiveModule('import')} className={`w-full flex items-center ${isSidebarOpen ? 'space-x-3 px-4' : 'justify-center'} py-3 rounded-xl transition-all font-medium text-sm ${activeModule === 'import' ? `${bgClass} text-white shadow-md` : 'text-slate-400 hover:bg-slate-800'}`} title={!isSidebarOpen ? "Importar Dados" : ""}>
              <UploadCloud className="h-5 w-5 shrink-0" />{isSidebarOpen && <span className="whitespace-nowrap">Importar Dados</span>}
            </button>
          )}

          {(currentUser.role === 'Administrador') && (
            <>
              {isSidebarOpen ? ( <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-3 ml-2 mt-6">Administração</p> ) : ( <div className="h-px bg-slate-800 my-4 mx-2"></div> )}
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

        <div className="p-4 border-t border-slate-800 bg-slate-950">
          {isSidebarOpen ? (
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3 overflow-hidden">
                <div className={`h-8 w-8 rounded-full shrink-0 ${bgClass} flex items-center justify-center text-white font-bold text-xs`} title={currentUser?.name || 'User'}>{(currentUser?.name || 'U').charAt(0).toUpperCase()}</div>
                <div className="truncate w-28"><p className="text-xs font-bold text-slate-200 truncate">{currentUser?.name || 'Usuário'}</p><p className="text-[10px] text-slate-500 truncate">{currentUser?.role || 'Acesso'}</p></div>
              </div>
              <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400 bg-slate-900 rounded-lg transition-colors" title="Sair do Sistema"><LogOut className="h-4 w-4 shrink-0" /></button>
            </div>
          ) : (
            <div className="flex flex-col items-center space-y-4">
              <div className={`h-8 w-8 rounded-full shrink-0 ${bgClass} flex items-center justify-center text-white font-bold text-xs`} title={currentUser?.name || 'User'}>{(currentUser?.name || 'U').charAt(0).toUpperCase()}</div>
              <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-400 bg-slate-900 rounded-lg transition-colors" title="Sair do Sistema"><LogOut className="h-4 w-4 shrink-0" /></button>
            </div>
          )}
        </div>
      </aside>

      <main className="flex-1 overflow-y-auto p-6 print:p-0 print:overflow-visible bg-slate-50 print:bg-white">
  {activeModule === 'dashboard' && <ModuleDashboard files={getFilteredFilesForUser(files, currentUser, usersList)} bgClass={bgClass} currentUser={currentUser} usersList={usersList} />}
  {activeModule === 'ranking' && <ModuleRanking files={getFilteredFilesForUser(files, currentUser, usersList)} bgClass={bgClass} />}
  {activeModule === 'import' && <ModuleImport files={files} setFiles={setFiles} bgClass={bgClass} />}
  {activeModule === 'users' && <ModuleUsers usersList={usersList} setUsersList={setUsersList} files={files} bgClass={bgClass} />}
  {activeModule === 'access' && <ModuleAccess usersList={usersList} setUsersList={setUsersList} files={files} bgClass={bgClass} />}
  {activeModule === 'settings' && <ModuleSettings sysTheme={sysTheme} setSysTheme={setSysTheme} bgClass={bgClass} />}
  {activeModule === 'reports' && <ModuleReports files={files} bgClass={bgClass} />}
</main>
    </div>
  );
}

// ==========================================
// FUNÇÕES AUXILIARES DE CPF
// ==========================================
const formatCPF = (value: string) => {
  return value
    .replace(/\D/g, '')
    .replace(/(\d{3})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d{1,2})/, '$1-$2')
    .replace(/(-\d{2})\d+?$/, '$1');
};

const isValidCPF = (cpf: string) => {
  const strCPF = cpf.replace(/\D/g, '');
  if (strCPF.length !== 11 || /^(\d)\1+$/.test(strCPF)) return false;
  
  let soma = 0;
  let resto;
  
  for (let i = 1; i <= 9; i++) soma = soma + parseInt(strCPF.substring(i - 1, i)) * (11 - i);
  resto = (soma * 10) % 11;
  if ((resto === 10) || (resto === 11)) resto = 0;
  if (resto !== parseInt(strCPF.substring(9, 10))) return false;
  
  soma = 0;
  for (let i = 1; i <= 10; i++) soma = soma + parseInt(strCPF.substring(i - 1, i)) * (12 - i);
  resto = (soma * 10) % 11;
  if ((resto === 10) || (resto === 11)) resto = 0;
  if (resto !== parseInt(strCPF.substring(10, 11))) return false;
  
  return true;
};

// ==========================================
// TELA DE LOGIN RESTRITO (ACESSO POR CPF)
// ==========================================
function LoginScreen({ bgClass }: any) {
  const [errorMsg, setErrorMsg] = useState('');
  const [cpfInput, setCpfInput] = useState('');

  const handleCpfChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCpfInput(formatCPF(e.target.value));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    const frm = e.target as any;
    
    const cleanCPF = cpfInput.replace(/\D/g, ''); 
    if (!isValidCPF(cleanCPF)) {
      setErrorMsg('CPF inválido. Verifique os números digitados.');
      return;
    }
    
    const systemEmail = `${cleanCPF}@embu.aps.br`; 
    const pass = frm.pass.value;

    try {
      await signInWithEmailAndPassword(auth, systemEmail, pass);
    } catch (err: any) {
      setErrorMsg(err.message.includes('user-not-found') || err.message.includes('invalid-credential') ? 'CPF ou senha incorretos.' : err.message);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-100">
      <div className="bg-white p-8 rounded-xl shadow-lg max-w-sm w-full border border-slate-200">
        <div className="flex justify-center mb-6">
          <div className={`p-4 rounded-full ${bgClass} bg-opacity-10`}>
            <Shield className={`h-10 w-10 text-slate-800`} />
          </div>
        </div>
        <h2 className="text-2xl font-bold text-center text-slate-800 mb-2 leading-tight">
          Indicadores APS<br/>
          <span className="text-lg font-medium text-slate-500">Embu das Artes</span>
        </h2>
        
        <p className="text-center text-slate-500 text-sm mb-6">Acesso ao Sistema</p>
        
        {errorMsg && <div className="bg-red-50 text-red-600 p-3 rounded-lg text-xs mb-4 text-center border border-red-200">{errorMsg}</div>}
        
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-600 mb-1">CPF</label>
            <input 
              name="cpf" 
              type="text" 
              value={cpfInput}
              onChange={handleCpfChange}
              placeholder="000.000.000-00" 
              required 
              className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500 font-medium tracking-wider" 
            />
          </div>
          
          <div>
            <label className="block text-xs font-semibold text-slate-600 mb-1">Senha</label>
            <input 
              name="pass" 
              type="password" 
              required 
              minLength={6} 
              className="w-full p-2.5 border border-slate-300 rounded-lg text-sm outline-none focus:border-slate-500" 
            />
          </div>
          
          <button type="submit" className={`w-full ${bgClass} hover:opacity-90 text-white font-bold py-2.5 rounded-lg transition-opacity flex justify-center items-center`}>
            <LogIn className="h-4 w-4 mr-2" /> Entrar
          </button>
        </form>
        
        <div className="mt-6 text-center text-[11px] text-slate-400 font-medium">
          Primeiro acesso ou esqueceu a senha?<br/>Solicite ao Administrador do Sistema.
        </div>
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
          
          setProgressMsg(`A Enviar ${file.name} (Compressão JSON)...`);
          try {
            // A MAGIA: Guarda a planilha inteira como uma STRING num único documento
            const docId = `planilha_${Date.now()}_${Math.random().toString(36).substring(2,7)}`;
            const docRef = doc(db, 'dados_siaps', docId);
            
            // JSON.stringify transforma tudo num texto otimizado
            await setDoc(docRef, {
              nomeArquivo: file.name,
              dadosJSON: JSON.stringify(enrichedData), 
              timestamp: Date.now()
            });

            newFiles.push({ id: file.name, name: file.name, data: enrichedData });
          } catch (err: any) { 
             alert(`Erro ao salvar ${file.name}. Se o erro for de limite, aguarde 24h ou mude para o plano Blaze.`); 
          }
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
      qs.docs.forEach(d => batch.delete(d.ref));
      await batch.commit();
      setFiles([]);
    } catch(e) { alert('Erro ao apagar. Verifique as regras do Firestore.'); }
    setIsProcessing(false);
    setProgressMsg('');
  };

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className={`p-10 rounded-2xl border-2 border-dashed text-center bg-white ${isProcessing ? 'border-slate-300 opacity-50' : 'border-slate-300 hover:border-slate-400'}`}>
        <div className={`mx-auto w-16 h-16 ${bgClass} bg-opacity-10 rounded-full flex items-center justify-center mb-4`}><UploadCloud className={`h-8 w-8 text-slate-700`} /></div>
        <h3 className="font-bold text-xl text-slate-800">Conexão de Dados na Nuvem (Otimizado)</h3>
        <p className="text-slate-500 mb-6 mt-2 max-w-lg mx-auto">As planilhas enviadas agora são comprimidas. O painel consumirá 99% menos do limite do Firestore!</p>
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

const DEFAULT_ROLES = [
  { name: 'Administrador', scope: 'geral', desc: 'Acesso total ao sistema, unidades, importação e configurações.', unidades: [], equipes: [] },
  { name: 'Coordenador', scope: 'geral', desc: 'Visualização de todas as unidades e importação de relatórios.', unidades: [], equipes: [] },
  { name: 'Gerente', scope: 'unidade', desc: 'Acesso restrito aos dados e equipes das Unidades selecionadas.', unidades: [], equipes: [] },
  { name: 'Enfermeiro', scope: 'equipe', desc: 'Acesso restrito exclusivamente às Equipes selecionadas.', unidades: [], equipes: [] },
  { name: 'Médico', scope: 'equipe', desc: 'Acesso restrito exclusivamente às Equipes selecionadas.', unidades: [], equipes: [] },
  { name: 'Leitor', scope: 'geral', desc: 'Visualização geral do painel sem permissão de edição.', unidades: [], equipes: [] }
];

function normAccessStr(v: any) {
  return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function extractUnidadesAndEquipes(files: any[]) {
  const raw = (files || []).flatMap((f: any) => f.data || []);
  const cols = Array.from(new Set(raw.flatMap(Object.keys))).filter(Boolean) as string[];

  const cUnid = cols.find(c => {
    const s = normAccessStr(c);
    if (s.includes('tipo') || s.includes('cnes') || s.includes('codigo')) return false;
    const isU = s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs');
    if (!isU) return false;
    const sample = raw.find((r: any) => r[c] && String(r[c]).trim() !== '');
    if (sample && String(sample[c]).toUpperCase().includes('CENTRO DE SAUDE')) return false;
    return true;
  });

  const cEq = cols.find(c => {
    const s = normAccessStr(c);
    return s.includes('nome da equipe') || s === 'equipe' || s === 'nome';
  });

  const uSet = new Set<string>();
  const eqMap: Record<string, Set<string>> = {};

  raw.forEach((r: any) => {
    const u = cUnid ? String(r[cUnid] || '').replace(/\s+/g, ' ').trim().toUpperCase() : '';
    const e = cEq ? String(r[cEq] || '').replace(/\s+/g, ' ').trim().toUpperCase() : '';
    if (u && u !== 'SEM UNIDADE' && u !== 'GERAL' && !u.includes('CENTRO DE SAUDE')) {
      uSet.add(u);
      if (!eqMap[u]) eqMap[u] = new Set<string>();
      if (e) eqMap[u].add(e);
    }
    if (e) {
      if (!eqMap['__TODAS__']) eqMap['__TODAS__'] = new Set<string>();
      eqMap['__TODAS__'].add(e);
    }
  });

  const outMap: Record<string, string[]> = {};
  Object.keys(eqMap).forEach(k => { outMap[k] = Array.from(eqMap[k]).sort(); });

  return { unidadesList: Array.from(uSet).sort(), equipesByUnidade: outMap, cUnid, cEq };
}

function getSystemRoles(usersList: any[]) {
  const cfg = (usersList || []).find((u: any) => u.id === '_roles_config_');
  return cfg && Array.isArray(cfg.roles) && cfg.roles.length > 0 ? cfg.roles : DEFAULT_ROLES;
}

function getFilteredFilesForUser(files: any[], currentUser: any, usersList: any[]) {
  if (!currentUser || !files || files.length === 0) return files;
  const roles = getSystemRoles(usersList);
  const userRec = (usersList || []).find((u: any) =>
    (u.email && currentUser.email && u.email.toLowerCase() === currentUser.email.toLowerCase()) ||
    u.id === currentUser.uid ||
    u.id === currentUser.id
  ) || currentUser;

  const roleObj = roles.find((r: any) => r.name === userRec?.role);
  const scope = userRec?.scope || roleObj?.scope || 'geral';

  if (userRec?.role === 'Administrador' || scope === 'geral') return files;

  const allowedUnidades: string[] = Array.from(new Set([
    ...(Array.isArray(userRec?.unidades) ? userRec.unidades : userRec?.unidade ? [userRec.unidade] : []),
    ...(Array.isArray(roleObj?.unidades) ? roleObj.unidades : [])
  ])).filter(Boolean);

  const allowedEquipes: string[] = Array.from(new Set([
    ...(Array.isArray(userRec?.equipes) ? userRec.equipes : userRec?.equipe ? [userRec.equipe] : []),
    ...(Array.isArray(roleObj?.equipes) ? roleObj.equipes : [])
  ])).filter(Boolean);

  if (allowedUnidades.length === 0 && allowedEquipes.length === 0) return files;

  const { cUnid, cEq } = extractUnidadesAndEquipes(files);

  return files.map((f: any) => ({
    ...f,
    data: (f.data || []).filter((r: any) => {
      const u = cUnid ? String(r[cUnid] || '').replace(/\s+/g, ' ').trim().toUpperCase() : '';
      const e = cEq ? String(r[cEq] || '').replace(/\s+/g, ' ').trim().toUpperCase() : '';
      const matchU = allowedUnidades.length === 0 || allowedUnidades.includes(u);
      const matchE = allowedEquipes.length === 0 || allowedEquipes.includes(e);
      return matchU && matchE;
    })
  })).filter((f: any) => f.data.length > 0);
}

function ModuleUsers({ usersList, setUsersList, files = [], bgClass }: any) {
  const roles = useMemo(() => getSystemRoles(usersList), [usersList]);
  const visibleUsers = useMemo(() => (usersList || []).filter((u: any) => u.id !== '_roles_config_' && !u.deleted), [usersList]);
  const { unidadesList, equipesByUnidade } = useMemo(() => extractUnidadesAndEquipes(files), [files]);

  const [editingUser, setEditingUser] = useState<any | null>(null);
  const [nome, setNome] = useState('');
  const [cpf, setCpf] = useState(''); // Estado do CPF visível
  const [pass, setPass] = useState('');
  const [selectedRole, setSelectedRole] = useState(roles[0]?.name || 'Administrador');
  const [checkedUnidades, setCheckedUnidades] = useState<string[]>([]);
  const [checkedEquipes, setCheckedEquipes] = useState<string[]>([]);

  const currentRoleObj = roles.find((r: any) => r.name === selectedRole) || { scope: 'geral', unidades: [], equipes: [] };

  // Função para formatar CPF no input
  const formatCPF = (value: string) => {
    return value
      .replace(/\D/g, '') 
      .replace(/(\d{3})(\d)/, '$1.$2') 
      .replace(/(\d{3})(\d)/, '$1.$2') 
      .replace(/(\d{3})(\d{1,2})/, '$1-$2') 
      .replace(/(-\d{2})\d+?$/, '$1'); 
  };

  const handleCpfChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCpf(formatCPF(e.target.value));
  };

  const availableEquipes = useMemo(() => {
    if (checkedUnidades.length > 0) {
      const set = new Set<string>();
      checkedUnidades.forEach(u => (equipesByUnidade[u] || []).forEach(eq => set.add(eq)));
      return Array.from(set).sort();
    }
    return equipesByUnidade['__TODAS__'] || [];
  }, [checkedUnidades, equipesByUnidade]);

  const toggleItem = (list: string[], setList: (v: string[]) => void, item: string) => {
    if (list.includes(item)) setList(list.filter(i => i !== item));
    else setList([...list, item]);
  };

  const handleRoleChange = (roleName: string) => {
    setSelectedRole(roleName);
    const rObj = roles.find((r: any) => r.name === roleName);
    setCheckedUnidades(Array.isArray(rObj?.unidades) ? rObj.unidades : []);
    setCheckedEquipes(Array.isArray(rObj?.equipes) ? rObj.equipes : []);
  };

  const startEdit = (u: any) => {
    setEditingUser(u);
    setNome(u.name || '');
    setCpf(u.cpf ? formatCPF(u.cpf) : ''); 
    setPass('');
    setSelectedRole(u.role || roles[0]?.name || 'Administrador');
    setCheckedUnidades(Array.isArray(u.unidades) ? u.unidades : u.unidade ? [u.unidade] : []);
    setCheckedEquipes(Array.isArray(u.equipes) ? u.equipes : u.equipe ? [u.equipe] : []);
  };

  const cancelEdit = () => {
    setEditingUser(null);
    setNome('');
    setCpf('');
    setPass('');
    setSelectedRole(roles[0]?.name || 'Administrador');
    setCheckedUnidades([]);
    setCheckedEquipes([]);
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    
    const cleanCPF = cpf.replace(/\D/g, ''); 
    if (!isValidCPF(cleanCPF)) {
      alert('CPF inválido! Por favor, insira um CPF real.');
      return;
    }
    const systemEmail = `${cleanCPF}@embu.aps.br`;

    const scope = currentRoleObj.scope || 'geral';
    // ... o resto da função continua igual ...
    const finalUnidades = (scope === 'unidade' || scope === 'equipe') ? checkedUnidades : [];
    const finalEquipes = scope === 'equipe' ? checkedEquipes : [];

    if (scope === 'unidade' && finalUnidades.length === 0) {
      alert('Marque pelo menos uma Unidade para este utilizador.');
      return;
    }
    if (scope === 'equipe' && finalEquipes.length === 0) {
      alert('Marque pelo menos uma Equipe para este utilizador.');
      return;
    }

    const userPayload = {
      name: nome,
      cpf: cleanCPF, // Salva o CPF limpo no banco
      emailSystem: systemEmail, // Salva a referência do login invisível
      role: selectedRole,
      scope,
      unidades: finalUnidades,
      equipes: finalEquipes,
      unidade: finalUnidades[0] || '',
      equipe: finalEquipes[0] || ''
    };

    try {
      if (editingUser) {
        const updatedUser = { ...editingUser, ...userPayload };
        await setDoc(doc(db, 'users', editingUser.id), updatedUser);
        setUsersList(usersList.map((u: any) => u.id === editingUser.id ? updatedUser : u));
        alert('Utilizador atualizado com sucesso!');
        cancelEdit();
      } else {
        const userCred = await createUserWithEmailAndPassword(secondaryAuth, systemEmail, pass);
        const newUid = userCred.user.uid;
        const newUser = { ...userPayload };
        await setDoc(doc(db, 'users', newUid), newUser);
        setUsersList([...usersList, { id: newUid, ...newUser }]);
        alert('Utilizador cadastrado com sucesso!');
        cancelEdit();
      }
    } catch (err: any) {
      alert(err.message.includes('email-already-in-use') ? 'Este CPF já está cadastrado no sistema!' : 'Erro ao guardar: ' + err.message);
    }
  };

  const handleDeleteUser = async (u: any) => {
    // TRAVA DE SEGURANÇA: Impede a exclusão do Administrador Master (que não tem CPF salvo)
    if (!u.cpf) {
      alert('Ação Negada: O Administrador Master original não pode ser excluído.');
      return;
    }

    if (!window.confirm(`Deseja remover o utilizador "${u.name}"?`)) return;
    
    try {
      const removed = { ...u, deleted: true };
      await setDoc(doc(db, 'users', u.id), removed);
      setUsersList(usersList.map((item: any) => item.id === u.id ? removed : item));
    } catch (err: any) {
      alert('Erro ao remover: ' + err.message);
    }
  };

  // Botão de "Reset de Senha" - Como o secondaryAuth está no FrontEnd, recriamos a credencial no Firebase Auth ou usamos o updatePassword (se permitido). 
  // A solução mais simples no Firebase Client sem Cloud Functions é alterar a senha via updatePassword se usarmos o secondaryAuth para relogar o usuário internamente.
  const handleResetPassword = async (u: any) => {
    if (!window.confirm(`Deseja definir a senha de "${u.name}" para mudar@123 ?`)) return;
    try {
      // Usa o secondaryAuth para autenticar o usuário alvo e trocar a senha
      const fakeEmail = u.emailSystem || `${u.cpf}@embu.aps.br`;
      alert('Em sistemas frontend (sem backend Node.js), a recuperação de senha por Admin exige deletar e recriar a conta no Firebase Auth, ou usar uma Cloud Function. Recomendação: Exclua este acesso e cadastre-o novamente com o mesmo CPF e a senha "mudar@123". Os dados e permissões podem ser recopiados facilmente.');
      // O código de reset ideal ficaria num backend (Admin SDK). O alerta avisa a melhor prática visual.
    } catch (err: any) {
      alert('Erro: ' + err.message);
    }
  };

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      <div className="md:col-span-1 bg-white p-5 rounded-xl border border-slate-200 shadow-sm h-fit">
        <h3 className="font-bold text-slate-800 mb-4 flex items-center">
          <UserPlus className="h-5 w-5 mr-2 text-blue-600" />
          {editingUser ? 'Editar Usuário' : 'Novo Usuário'}
        </h3>
        <form onSubmit={handleSave} className="space-y-3">
          <div>
            <label className="text-xs font-semibold text-slate-600">Nome Completo</label>
            <input value={nome} onChange={e => setNome(e.target.value)} placeholder="Nome Completo" required className="w-full p-2 border rounded-lg outline-none text-sm mt-1" />
          </div>

          <div>
            <label className="text-xs font-semibold text-slate-600">CPF (Login)</label>
            <input value={cpf} onChange={handleCpfChange} type="text" placeholder="000.000.000-00" required disabled={Boolean(editingUser)} className={`w-full p-2 border rounded-lg outline-none text-sm mt-1 font-medium tracking-wider ${editingUser ? 'bg-slate-100 text-slate-500 cursor-not-allowed' : ''}`} />
          </div>

          {!editingUser && (
            <div>
              <label className="text-xs font-semibold text-slate-600">Senha Inicial</label>
              <input value={pass} onChange={e => setPass(e.target.value)} type="password" placeholder="Senha (Mín. 6)" required minLength={6} className="w-full p-2 border rounded-lg outline-none text-sm mt-1" />
            </div>
          )}

          <div>
            <label className="text-xs font-semibold text-slate-600">Perfil de Acesso</label>
            <select value={selectedRole} onChange={e => handleRoleChange(e.target.value)} className="w-full p-2 border rounded-lg outline-none text-sm bg-white mt-1">
              {roles.map((r: any) => (
                <option key={r.name} value={r.name}>
                  {r.name} ({r.scope === 'unidade' ? 'Por Unidade' : r.scope === 'equipe' ? 'Por Equipe' : 'Acesso Geral'})
                </option>
              ))}
            </select>
          </div>

          {(currentRoleObj.scope === 'unidade' || currentRoleObj.scope === 'equipe') && (
            <div>
              <div className="flex justify-between items-center mb-1">
                <label className="text-xs font-semibold text-slate-600">
                  Unidades Permitidas ({checkedUnidades.length})
                </label>
                <button type="button" onClick={() => setCheckedUnidades(checkedUnidades.length === unidadesList.length ? [] : [...unidadesList])} className="text-[11px] text-blue-600 hover:underline">
                  {checkedUnidades.length === unidadesList.length ? 'Desmarcar todas' : 'Marcar todas'}
                </button>
              </div>
              <div className="max-h-40 overflow-y-auto border rounded-lg p-2 space-y-1.5 bg-slate-50">
                {unidadesList.map(u => (
                  <label key={u} className="flex items-center space-x-2 text-xs text-slate-700 cursor-pointer hover:bg-white p-1 rounded">
                    <input
                      type="checkbox"
                      checked={checkedUnidades.includes(u)}
                      onChange={() => toggleItem(checkedUnidades, setCheckedUnidades, u)}
                      className="rounded text-blue-600"
                    />
                    <span className="truncate">{u}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {currentRoleObj.scope === 'equipe' && (
            <div>
              <div className="flex justify-between items-center mb-1">
                <label className="text-xs font-semibold text-slate-600">
                  Equipes Permitidas ({checkedEquipes.length})
                </label>
                <button type="button" onClick={() => setCheckedEquipes(checkedEquipes.length === availableEquipes.length ? [] : [...availableEquipes])} className="text-[11px] text-blue-600 hover:underline">
                  {checkedEquipes.length === availableEquipes.length ? 'Desmarcar todas' : 'Marcar todas'}
                </button>
              </div>
              <div className="max-h-44 overflow-y-auto border rounded-lg p-2 space-y-1.5 bg-slate-50">
                {availableEquipes.map(eq => (
                  <label key={eq} className="flex items-center space-x-2 text-xs text-slate-700 cursor-pointer hover:bg-white p-1 rounded">
                    <input
                      type="checkbox"
                      checked={checkedEquipes.includes(eq)}
                      onChange={() => toggleItem(checkedEquipes, setCheckedEquipes, eq)}
                      className="rounded text-blue-600"
                    />
                    <span className="truncate">{eq}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          <div className="flex gap-2 pt-2">
            <button type="submit" className={`flex-1 ${bgClass} text-white py-2 rounded-lg font-bold text-sm hover:opacity-90`}>
              {editingUser ? 'Salvar Alterações' : 'Criar Acesso'}
            </button>
            {editingUser && (
              <button type="button" onClick={cancelEdit} className="px-3 py-2 bg-slate-200 text-slate-700 rounded-lg font-semibold text-sm hover:bg-slate-300">
                Cancelar
              </button>
            )}
          </div>
        </form>
      </div>

      <div className="md:col-span-2 bg-white p-5 rounded-xl border border-slate-200 shadow-sm overflow-x-auto">
        <h3 className="font-bold text-slate-800 mb-4">Usuários do Sistema</h3>
        <table className="w-full text-left text-sm">
          <thead className="bg-slate-100">
            <tr>
              <th className="p-2">Nome</th>
              <th className="p-2">CPF (Login)</th>
              <th className="p-2">Perfil</th>
              <th className="p-2">Acesso Liberado</th>
              <th className="p-2 text-right">Ações</th>
            </tr>
          </thead>
          <tbody>
            {visibleUsers.map((u: any) => {
              const uList = Array.isArray(u.unidades) ? u.unidades : u.unidade ? [u.unidade] : [];
              const eList = Array.isArray(u.equipes) ? u.equipes : u.equipe ? [u.equipe] : [];
              return (
                <tr key={u.id} className="border-b hover:bg-slate-50">
                  <td className="p-2 font-medium">{u.name}</td>
                  <td className="p-2 text-slate-500">{u.cpf ? formatCPF(u.cpf) : 'Migrar para CPF'}</td>
                  <td className="p-2">
                    <span className="px-2 py-1 rounded text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                      {u.role}
                    </span>
                  </td>
                  <td className="p-2 text-xs text-slate-600 max-w-xs">
                    {eList.length > 0 ? (
                      <div><strong>Equipes:</strong> {eList.join(', ')}</div>
                    ) : uList.length > 0 ? (
                      <div><strong>Unidades:</strong> {uList.join(', ')}</div>
                    ) : (
                      <span className="text-slate-400">Conforme Perfil / Geral</span>
                    )}
                  </td>
                  <td className="p-2 text-right space-x-2 whitespace-nowrap">
                    <button onClick={() => handleResetPassword(u)} className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 rounded text-[11px] font-bold" title="Resetar Senha">
                      Reset
                    </button>
                    <button onClick={() => startEdit(u)} className="px-2.5 py-1 bg-slate-100 hover:bg-blue-100 text-blue-700 rounded text-[11px] font-bold">
                      Editar
                    </button>
                    <button onClick={() => handleDeleteUser(u)} className="px-2.5 py-1 bg-slate-100 hover:bg-red-100 text-red-600 rounded text-[11px] font-bold">
                      X
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ModuleAccess({ usersList, setUsersList, files = [], bgClass }: any) {
  const roles = useMemo(() => getSystemRoles(usersList), [usersList]);
  const { unidadesList, equipesByUnidade } = useMemo(() => extractUnidadesAndEquipes(files), [files]);

  const [roleName, setRoleName] = useState('');
  const [roleScope, setRoleScope] = useState('unidade');
  const [roleDesc, setRoleDesc] = useState('');
  const [checkedUnidades, setCheckedUnidades] = useState<string[]>([]);
  const [checkedEquipes, setCheckedEquipes] = useState<string[]>([]);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);

  const availableEquipes = useMemo(() => {
    if (checkedUnidades.length > 0) {
      const set = new Set<string>();
      checkedUnidades.forEach(u => (equipesByUnidade[u] || []).forEach(eq => set.add(eq)));
      return Array.from(set).sort();
    }
    return equipesByUnidade['__TODAS__'] || [];
  }, [checkedUnidades, equipesByUnidade]);

  const toggleItem = (list: string[], setList: (v: string[]) => void, item: string) => {
    if (list.includes(item)) setList(list.filter(i => i !== item));
    else setList([...list, item]);
  };

  const saveRolesToCloud = async (newRoles: any[]) => {
    try {
      const cfgDoc = { id: '_roles_config_', isConfig: true, roles: newRoles };
      await setDoc(doc(db, 'users', '_roles_config_'), cfgDoc);
      const exists = (usersList || []).some((u: any) => u.id === '_roles_config_');
      if (exists) {
        setUsersList(usersList.map((u: any) => u.id === '_roles_config_' ? cfgDoc : u));
      } else {
        setUsersList([...(usersList || []), cfgDoc]);
      }
    } catch (err: any) {
      alert('Erro ao guardar perfis: ' + err.message);
    }
  };

  const handleSaveRole = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!roleName.trim()) return;

    const newRoleObj = {
      name: roleName.trim(),
      scope: roleScope,
      desc: roleDesc.trim() || (roleScope === 'unidade' ? 'Acesso restrito às Unidades marcadas.' : roleScope === 'equipe' ? 'Acesso restrito às Equipes marcadas.' : 'Acesso geral ao painel.'),
      unidades: (roleScope === 'unidade' || roleScope === 'equipe') ? checkedUnidades : [],
      equipes: roleScope === 'equipe' ? checkedEquipes : []
    };

    let updated = [...roles];
    if (editingIndex !== null) {
      updated[editingIndex] = newRoleObj;
    } else {
      if (updated.some(r => r.name.toLowerCase() === newRoleObj.name.toLowerCase())) {
        alert('Já existe um perfil com este nome!');
        return;
      }
      updated.push(newRoleObj);
    }

    await saveRolesToCloud(updated);
    setRoleName('');
    setRoleScope('unidade');
    setRoleDesc('');
    setCheckedUnidades([]);
    setCheckedEquipes([]);
    setEditingIndex(null);
  };

  const handleEditRole = (r: any, idx: number) => {
    setEditingIndex(idx);
    setRoleName(r.name);
    setRoleScope(r.scope || 'geral');
    setRoleDesc(r.desc || '');
    setCheckedUnidades(Array.isArray(r.unidades) ? r.unidades : []);
    setCheckedEquipes(Array.isArray(r.equipes) ? r.equipes : []);
  };

  const handleDeleteRole = async (idx: number) => {
    if (!window.confirm(`Deseja excluir o perfil "${roles[idx].name}"?`)) return;
    const updated = roles.filter((_: any, i: number) => i !== idx);
    await saveRolesToCloud(updated);
  };

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      <div className="md:col-span-1 bg-white p-5 rounded-xl border border-slate-200 shadow-sm h-fit">
        <h3 className="font-bold text-slate-800 mb-4 flex items-center">
          <Shield className="h-5 w-5 mr-2 text-blue-600" />
          {editingIndex !== null ? 'Editar Perfil' : 'Criar Novo Perfil'}
        </h3>
        <form onSubmit={handleSaveRole} className="space-y-3">
          <div>
            <label className="text-xs font-semibold text-slate-600">Nome do Perfil</label>
            <input
              value={roleName}
              onChange={e => setRoleName(e.target.value)}
              placeholder="Ex: Gerente UBS Centro, Equipe 1..."
              required
              className="w-full p-2 border rounded-lg outline-none text-sm mt-1"
            />
          </div>

          <div>
            <label className="text-xs font-semibold text-slate-600">Tipo de Restrição</label>
            <select
              value={roleScope}
              onChange={e => { setRoleScope(e.target.value); setCheckedUnidades([]); setCheckedEquipes([]); }}
              className="w-full p-2 border rounded-lg outline-none text-sm bg-white mt-1"
            >
              <option value="unidade">Restringir por Unidade(s)</option>
              <option value="equipe">Restringir por Equipe(s)</option>
              <option value="geral">Acesso Geral (Município Inteiro)</option>
            </select>
          </div>

          {(roleScope === 'unidade' || roleScope === 'equipe') && (
            <div>
              <div className="flex justify-between items-center mb-1">
                <label className="text-xs font-semibold text-slate-600">
                  Escolher Unidades ({checkedUnidades.length})
                </label>
                <button type="button" onClick={() => setCheckedUnidades(checkedUnidades.length === unidadesList.length ? [] : [...unidadesList])} className="text-[11px] text-blue-600 hover:underline">
                  {checkedUnidades.length === unidadesList.length ? 'Limpar' : 'Todas'}
                </button>
              </div>
              <div className="max-h-40 overflow-y-auto border rounded-lg p-2 space-y-1.5 bg-slate-50">
                {unidadesList.map(u => (
                  <label key={u} className="flex items-center space-x-2 text-xs text-slate-700 cursor-pointer hover:bg-white p-1 rounded">
                    <input
                      type="checkbox"
                      checked={checkedUnidades.includes(u)}
                      onChange={() => toggleItem(checkedUnidades, setCheckedUnidades, u)}
                      className="rounded text-blue-600"
                    />
                    <span className="truncate">{u}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {roleScope === 'equipe' && (
            <div>
              <div className="flex justify-between items-center mb-1">
                <label className="text-xs font-semibold text-slate-600">
                  Escolher Equipes ({checkedEquipes.length})
                </label>
                <button type="button" onClick={() => setCheckedEquipes(checkedEquipes.length === availableEquipes.length ? [] : [...availableEquipes])} className="text-[11px] text-blue-600 hover:underline">
                  {checkedEquipes.length === availableEquipes.length ? 'Limpar' : 'Todas'}
                </button>
              </div>
              <div className="max-h-44 overflow-y-auto border rounded-lg p-2 space-y-1.5 bg-slate-50">
                {availableEquipes.map(eq => (
                  <label key={eq} className="flex items-center space-x-2 text-xs text-slate-700 cursor-pointer hover:bg-white p-1 rounded">
                    <input
                      type="checkbox"
                      checked={checkedEquipes.includes(eq)}
                      onChange={() => toggleItem(checkedEquipes, setCheckedEquipes, eq)}
                      className="rounded text-blue-600"
                    />
                    <span className="truncate">{eq}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          <div>
            <label className="text-xs font-semibold text-slate-600">Descrição (Opcional)</label>
            <textarea
              value={roleDesc}
              onChange={e => setRoleDesc(e.target.value)}
              placeholder="Descrição do perfil..."
              rows={2}
              className="w-full p-2 border rounded-lg outline-none text-sm mt-1"
            />
          </div>

          <div className="flex gap-2 pt-2">
            <button type="submit" className={`flex-1 ${bgClass} text-white py-2 rounded-lg font-bold text-sm hover:opacity-90`}>
              {editingIndex !== null ? 'Salvar Alterações' : 'Criar Perfil'}
            </button>
            {editingIndex !== null && (
              <button
                type="button"
                onClick={() => { setEditingIndex(null); setRoleName(''); setRoleDesc(''); setCheckedUnidades([]); setCheckedEquipes([]); }}
                className="px-3 py-2 bg-slate-200 text-slate-700 rounded-lg font-semibold text-sm"
              >
                Cancelar
              </button>
            )}
          </div>
        </form>
      </div>

      <div className="md:col-span-2 bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <h2 className="text-lg font-bold flex items-center mb-4">
          <Shield className="mr-2 h-5 w-5 text-blue-600" /> Perfis Cadastrados
        </h2>
        <div className="space-y-3">
          {roles.map((r: any, idx: number) => (
            <div key={idx} className="p-4 border rounded-lg bg-slate-50 flex justify-between items-start gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="font-bold text-slate-800">{r.name}</h3>
                  <span className={`px-2 py-0.5 rounded-full text-[11px] font-bold ${
                    r.scope === 'unidade' ? 'bg-blue-100 text-blue-700' :
                    r.scope === 'equipe' ? 'bg-emerald-100 text-emerald-700' :
                    'bg-amber-100 text-amber-700'
                  }`}>
                    {r.scope === 'unidade' ? 'Por Unidade' : r.scope === 'equipe' ? 'Por Equipe' : 'Acesso Geral'}
                  </span>
                </div>
                <p className="text-slate-600 text-sm mt-1">{r.desc}</p>
                {Array.isArray(r.unidades) && r.unidades.length > 0 && (
                  <p className="text-xs text-blue-700 mt-1.5"><strong>Unidades marcadas:</strong> {r.unidades.join(', ')}</p>
                )}
                {Array.isArray(r.equipes) && r.equipes.length > 0 && (
                  <p className="text-xs text-emerald-700 mt-1"><strong>Equipes marcadas:</strong> {r.equipes.join(', ')}</p>
                )}
              </div>
              <div className="flex gap-2 shrink-0">
                <button onClick={() => handleEditRole(r, idx)} className="px-2.5 py-1 bg-white border hover:bg-blue-50 text-blue-600 rounded text-xs font-semibold">
                  Editar
                </button>
                <button onClick={() => handleDeleteRole(idx)} className="px-2.5 py-1 bg-white border hover:bg-red-50 text-red-600 rounded text-xs font-semibold">
                  Excluir
                </button>
              </div>
            </div>
          ))}
        </div>
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
function ModuleDashboard({ files, bgClass, currentUser, usersList }: any) {
  const rawAllData = useMemo(() => files.flatMap((f: any) => f.data), [files]);
  const columns = useMemo(() => Array.from(new Set(rawAllData.flatMap(Object.keys))).filter(c => c && !String(c).startsWith('_') && !String(c).toUpperCase().includes('EMPTY') && String(c).trim() !== ''), [rawAllData]);

  const colUnidade = useMemo(() => columns.find(c => {
    const s = normalizeStrForFind(c);
    if (s.includes('tipo') || s.includes('cnes') || s.includes('codigo')) return false;
    const isUnit = s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs');
    if (!isUnit) return false;
    const sample = rawAllData.find((r: any) => r[c] && String(r[c]).trim() !== '');
    if (sample && String(sample[c]).toUpperCase().includes('CENTRO DE SAUDE')) return false;
    return true;
  }) || null, [columns, rawAllData]);

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
  const [isResumoOpen, setIsResumoOpen] = useState(true);
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
  )).filter(u => Boolean(u) && u !== 'SEM UNIDADE' && u !== 'GERAL' && !u.toUpperCase().includes('CENTRO DE SAUDE')).sort(), [baseFilterData, selectedTipoEquipe, selectedEquipe]);

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
    if (!selectedIndicador || !allData) return [];
    
    const indRows = allData.filter(d => d._mappedIndicator === selectedIndicador);
    if (indRows.length === 0) return [];

    const keys = new Set<string>();
    indRows.forEach(r => {
      Object.keys(r).forEach(k => {
        if (
          !k.startsWith('_') && 
          !k.toUpperCase().includes('EMPTY') && 
          !isMetadataColumn(k)
        ) {
          keys.add(k);
        }
      });
    });
    return Array.from(keys);
  }, [selectedIndicador, allData]);

  const activeVariable = useMemo(() => {
    if (!selectedIndicador || !allData) return '';
    
    // 1. Analisa as colunas de todos os meses para não saltar
    const indRows = allData.filter(d => d._mappedIndicator === selectedIndicador);
    if (indRows.length === 0) return '';

    const keys = new Set<string>();
    indRows.forEach(r => {
      Object.keys(r).forEach(k => {
        if (!k.startsWith('_') && !k.toUpperCase().includes('EMPTY')) {
          keys.add(k);
        }
      });
    });
    const allIndCols = Array.from(keys);

    const normalizeStr = (str: string) => 
      str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

    const indL = String(selectedIndicador).toLowerCase();
    
    const checkInd = (num: string) => 
       indL.includes(`c${num}`) || indL.includes(`indicador ${num}`) || indL.includes(`ind ${num}`) || indL.includes(`0${num}`) || indL.includes(`- ${num}`);

    // Função que respeita a ordem de prioridade exata que passarmos
    const findCol = (keywordGroups: string[][]) => {
      for (const keywords of keywordGroups) {
        const found = allIndCols.find(col => {
          const normCol = normalizeStr(col);
          return keywords.some(kw => normCol.includes(kw));
        });
        if (found) return String(found);
      }
      return undefined;
    };

    let foundVar;

  // 2. Regra Exata: Gestantes (C3), C2, C4 e C5 buscam RAZÃO primeiro
  if (checkInd('2') || checkInd('3') || checkInd('4') || checkInd('5') || indL.includes('gesta')) {
    foundVar = findCol([ 
      ['razao entre', 'razao'], 
      ['pontuacao', 'pontua'] 
    ]);
 } 
 // 3. Regra Exata: C7 busca SOMATÓRIO primeiro
 else if (checkInd('7') || indL.includes('mulher') || indL.includes('cancer')) {
    foundVar = findCol([ 
      ['somatorio', 'boa pratica'], 
      ['pontuacao', 'pontua'] 
    ]);
 } 
 // 4. Regra Exata: Vínculo e C1 buscam PONTUAÇÃO, PARÂMETRO ou CADASTRO
 else {
    foundVar = findCol([ 
      ['pontuacao', 'pontua'],
      ['parametro populacional', 'parametro'],
      ['pessoas somente com cadastro']
    ]);
 }

    // 5. Segurança geral se nada bater certo
    if (!foundVar) {
        foundVar = findCol([ 
          ['nota', 'resultado', 'indicador', 'perc', 'total', 'proporcao'] 
        ]);
    }
    
    return foundVar ? foundVar : String(allIndCols[allIndCols.length - 1] || '');
  }, [selectedIndicador, allData]);

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

  const statusCounts = { 'Ótimo': 0, 'Bom': 0, 'Suficiente': 0, 'Regular': 0 };
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
    { name: 'Ótimo', value: statusCounts['Ótimo'], color: '#3B82F6' },
    { name: 'Bom', value: statusCounts['Bom'], color: '#10B981' },
    { name: 'Suficiente', value: statusCounts['Suficiente'], color: '#F59E0B' },
    { name: 'Regular', value: statusCounts['Regular'], color: '#EF4444' }
  ].filter(d => d.value > 0);

  if (files.length === 0) return (
    <div className="h-[80vh] flex flex-col items-center justify-center text-slate-400">
      <BarChart2 className="h-20 w-20 mb-4 opacity-20"/>
      <h2 className="text-xl font-medium text-slate-600">Nenhum dado hospedado no Firebase.</h2>
      <p className="mt-2 text-sm">Vá ao módulo "Importar Dados" no menu lateral para iniciar.</p>
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
        <div onClick={() => statusCounts['Suficiente'] > 0 && setActiveModalStatus('Suficiente')} className={`bg-white p-4 rounded-xl border border-amber-200 shadow-sm flex items-center space-x-3 ${statusCounts['Suficiente'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-amber-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-amber-100 p-2.5 rounded-full shrink-0"><AlertTriangle className="h-5 w-5 text-amber-600" /></div>
          <div><p className="text-xs text-amber-700 font-semibold uppercase">Suficiente</p><p className="text-xl font-bold text-amber-600">{statusCounts['Suficiente']}</p></div>
        </div>
        <div onClick={() => statusCounts['Regular'] > 0 && setActiveModalStatus('Regular')} className={`bg-white p-4 rounded-xl border border-red-200 shadow-sm flex items-center space-x-3 col-span-2 lg:col-span-1 ${statusCounts['Regular'] > 0 ? 'cursor-pointer hover:shadow-md hover:border-red-400 transition-all' : 'opacity-60'}`}>
          <div className="bg-red-100 p-2.5 rounded-full shrink-0"><AlertOctagon className="h-5 w-5 text-red-600" /></div>
          <div><p className="text-xs text-red-700 font-semibold uppercase">Regular</p><p className="text-xl font-bold text-red-600">{statusCounts['Regular']}</p></div>
        </div>
      </section>

      {variablesChartData.length > 0 && (
        <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <div 
            onClick={() => setIsResumoOpen(!isResumoOpen)} 
            className={`flex justify-between items-center cursor-pointer select-none group ${isResumoOpen ? 'mb-3' : ''}`}
          >
            <div className="flex items-center space-x-2">
              <BarChart2 className="h-4 w-4 text-blue-600" />
              <h3 className="text-sm font-bold text-slate-800">
                {isSingleTeam ? `Valores Reais (${consolidatedTeams[0]?.equipe || selectedEquipe})` : 'Resumo das Variáveis (Média)'}
              </h3>
            </div>
            <button className="p-1 rounded-lg text-slate-400 group-hover:text-slate-600 group-hover:bg-slate-100 transition-colors">
              {isResumoOpen ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
            </button>
          </div>
          
          {isResumoOpen && (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
              {variablesChartData.map((v, i) => (
                <div key={i} className={`p-4 rounded-lg border transition-all ${v.fullName === activeVariable ? 'border-blue-600 bg-blue-50/70 shadow-sm' : 'border-slate-200 bg-slate-50/60'}`}>
                  <p className="text-xs text-slate-600 font-medium leading-tight" title={v.fullName}>{v.fullName}</p>
                  <p className={`text-2xl font-bold mt-2 ${v.fullName === activeVariable ? 'text-blue-900' : 'text-slate-800'}`}>{formatVal(v.valor)}</p>
                  <p className="text-[10px] text-slate-500 mt-1">{isSingleTeam ? 'Valor real da equipe' : `Média de ${v.equipesContabilizadas} equipes`}</p>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

<section className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm lg:col-span-2 overflow-hidden">
          <h3 className="text-base font-bold text-slate-800 mb-1">Comparativo por Equipe</h3>
          <div className="w-full overflow-x-auto pb-2">
            <div style={{ minWidth: teamsComparisonData.length > 20 ? `${teamsComparisonData.length * 32}px` : '100%', height: 290 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={teamsComparisonData} margin={{ top: 30, right: 20, left: 0, bottom: 80 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis dataKey="name" stroke="#475569" fontSize={9} tickLine={false} axisLine={false} angle={-45} textAnchor="end" interval={0} />
                  <YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} />
                  <Tooltip formatter={(val: number, _n: string, p: any) => [formatVal(val), p.payload.fullName || activeVariable]} />
                  <Bar dataKey="valor" radius={[4, 4, 0, 0]} maxBarSize={40}>
                    {teamsComparisonData.map((_e, i) => <Cell key={i} fill="#1D4ED8" />)}
                    <LabelList dataKey="valor" position="top" fontSize={10} fontWeight="bold" formatter={(v: number) => formatVal(v)} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
        
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-base font-bold text-slate-800 mb-1">Classificação Final</h3>
          <div className="w-full" style={{ height: 290 }}>
            {statusPieData.length ? (
              <ResponsiveContainer width="100%" height={280}>
                <PieChart margin={{ top: 15, right: 35, left: 35, bottom: 5 }}>
                  <Pie
                    data={statusPieData.map((d: any) => ({
                      ...d,
                      value: Number(d.value ?? d.valor ?? d.count ?? 0)
                    }))}
                    cx="50%"
                    cy="45%"
                    outerRadius={58}
                    dataKey="value"
                    nameKey="name"
                    isAnimationActive={false}
                    label={({ value }) => value}
                    labelLine={true}
                  >
                    {statusPieData.map((entry: any, index: number) => {
                      const colorMap: Record<string, string> = {
                        'Ótimo': '#3B82F6',
                        'Bom': '#10B981',
                        'Suficiente': '#F59E0B',
                        'Satisfatório': '#F59E0B',
                        'Regular': '#EF4444'
                      };
                      return <Cell key={`cell-${index}`} fill={entry.color || entry.fill || colorMap[entry.name] || '#94A3B8'} />;
                    })}
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
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={lineChartData} margin={{ top: 30, right: 30, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" stroke="#475569" fontSize={11} tickLine={false} axisLine={false} padding={{ left: 20, right: 20 }} />
                <YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} />
                <Tooltip formatter={(v: number) => [formatVal(v), isSingleTeam ? 'Equipe' : 'Média']} />
                <Line type="monotone" dataKey="valor" stroke="#10B981" strokeWidth={3} dot={{ r: 6, fill: '#10B981', strokeWidth: 2, stroke: '#fff' }} connectNulls={false}>
                  <LabelList dataKey="valor" position="top" offset={12} fontSize={12} fontWeight="bold" formatter={(v: number) => formatVal(v)} />
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
                <th onClick={() => requestSort('_ubsLabel')} className="px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group"><div className="flex items-center">Unidade {sortConfig?.key === '_ubsLabel' ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100" />}</div></th>
                <th onClick={() => requestSort('_eq')} className="px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group"><div className="flex items-center">Equipe {sortConfig?.key === '_eq' ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100" />}</div></th>
                {displayColumns.map(c => (
                  <th key={String(c)} onClick={() => requestSort(String(c))} className={`px-3 py-2.5 cursor-pointer hover:bg-slate-200 select-none group ${String(c) === activeVariable ? 'bg-blue-100' : ''}`}><div className="flex items-center">{String(c)} {sortConfig?.key === String(c) ? (sortConfig.direction === 'asc' ? <ChevronUp className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" /> : <ChevronDown className="h-3.5 w-3.5 ml-1 text-blue-600 shrink-0" />) : <ArrowUpDown className="h-3 w-3 ml-1 text-slate-300 opacity-0 group-hover:opacity-100 shrink-0" />}</div></th>
                ))}
              </tr>
            </thead>
            <tbody>
              {!isSingleTeam && totalEquipesUnicas > 1 && (
                <tr className="bg-blue-50/90 border-b border-blue-200 font-bold">
                  <td className="px-3 py-2">RESUMO</td>
                  <td className="px-3 py-2">{selectedUnidade || 'TODAS AS UNIDADES'}</td>
                  <td className="px-3 py-2">MÉDIA DAS EQUIPES</td>
                  {displayColumns.map(c => (
                    <td key={String(c)} className="px-3 py-2">
                      {summaryVariablesMap[String(c)] !== undefined ? summaryVariablesMap[String(c)] : '-'}
                    </td>
                  ))}
                </tr>
              )}
              {sortedTableData.slice(0, 150).map((r, i) => (
                <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-3 py-2 text-blue-800 bg-blue-50/20">{r._ms}</td>
                  <td className="px-3 py-2 text-slate-600 bg-blue-50/10">{r._ubsLabel}</td>
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
            <div className={`p-4 flex justify-between items-center text-white ${activeModalStatus === 'Ótimo' ? 'bg-blue-600' : activeModalStatus === 'Bom' ? 'bg-emerald-600' : activeModalStatus === 'Suficiente' ? 'bg-amber-500' : 'bg-red-600'}`}>
              <div className="flex items-center space-x-2">
                {activeModalStatus === 'Ótimo' && <Award className="h-5 w-5" />}
                {activeModalStatus === 'Bom' && <CheckCircle className="h-5 w-5" />}
                {activeModalStatus === 'Suficiente' && <AlertTriangle className="h-5 w-5" />}
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
                        activeModalStatus === 'Suficiente' ? 'bg-amber-500' : 
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

  const colUnidade = useMemo(() => {
    const candidates = columns.filter(c => {
      const s = normalizeStrForFind(c);
      if (s.includes('tipo') || s.includes('cnes') || s.includes('codigo')) return false;
      return s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs') || s.includes('fantasia');
    });

    const validCol = candidates.find(c => {
      const sample = rawAllData.find(r => r[c] && String(r[c]).trim() !== '');
      if (!sample) return false;
      const val = String(sample[c]).toUpperCase();
      return !val.includes('CENTRO DE SAUDE') && !val.includes('UNIDADE BASICA');
    });

    return validCol || candidates[0] || null;
  }, [columns, rawAllData]);

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
      
      let ubs = 'GERAL';
      if (colUnidade && r[colUnidade]) {
        const val = String(r[colUnidade]).replace(/\s+/g, ' ').trim().toUpperCase();
        if (!val.includes('CENTRO DE SAUDE') && !val.includes('UNIDADE BASICA')) {
          ubs = val;
        }
      }
      if (ubs === 'GERAL') {
        const altKey = Object.keys(r).find(k => {
          const s = normalizeStrForFind(k);
          if (s.includes('tipo') || s.includes('cnes') || s.includes('codigo')) return false;
          if (!(s.includes('estabelecimento') || s.includes('unidade') || s.includes('ubs') || s.includes('fantasia'))) return false;
          const v = String(r[k] || '').toUpperCase();
          return v !== '' && !v.includes('CENTRO DE SAUDE') && !v.includes('UNIDADE BASICA');
        });
        if (altKey) ubs = String(r[altKey]).replace(/\s+/g, ' ').trim().toUpperCase();
      }
      
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

  // Estados de Minimizar
  const [showTopUbs, setShowTopUbs] = useState(true);
  const [showBottomUbs, setShowBottomUbs] = useState(true);
  const [showTopEq, setShowTopEq] = useState(true);
  const [showBottomEq, setShowBottomEq] = useState(true);
  const [showInd, setShowInd] = useState(true);
  const [showEvolution, setShowEvolution] = useState(true); // NOVO: Estado da Evolução

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
        
        const indL = normalizeStrForFind(ind);
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
    const indL = normalizeStrForFind(ind);

    if (indL.includes('vinculo') || indL.includes('acompanhamento')) {
      const rowKeys = Object.keys(row || {}).filter(k => !String(k).startsWith('_'));
      const notaCol = rowKeys.find(c => {
        const s = normalizeStrForFind(c);
        return s.includes('pontua') || s.includes('nota') || s.includes('resultado') || s.includes('valor');
      });
      if (notaCol && row[notaCol] !== undefined && String(row[notaCol]).trim() !== '') {
        return calcStatus(parseNum(row[notaCol]), ind);
      }
    }

    const rowKeys = Object.keys(row || {}).filter(k => !String(k).startsWith('_'));
    const statusCol = rowKeys.find(c => ['classifica', 'conceito', 'qualidade', 'status', 'desempenho'].some(w => String(c).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').includes(w)));
    if (statusCol) {
      const st = normStat(row[statusCol]);
      if (st) return st === 'Satisfatório' ? 'Suficiente' : st;
    }

    const vCol = getActiveVarForIndicator(ind);
    if (vCol && row[String(vCol)] !== undefined && String(row[String(vCol)]).trim() !== '') {
      return calcStatus(parseNum(row[String(vCol)]), ind);
    }

    return null;
  };

  const RANKING_WEIGHTS: Record<string, number> = useMemo(() => ({
    'Ótimo': 4,
    'Bom': 3,
    'Suficiente': 2,
    'Regular': 1 
  }), []);

  // NOVO: Cálculo da Evolução Mensal (Ignorando filtro de Mês)
  const evolutionData = useMemo(() => {
    return optM.map(m => {
      const monthRows = allData.filter(d => String(d._mesOrigem) === m);
      const result: any = { name: m };

      // Calcula a média ponderada de CADA indicador neste mês
      optI.forEach(ind => {
        let sum = 0;
        let count = 0;
        monthRows.forEach(row => {
          if (row._mappedIndicator === ind) {
            const st = getStatusForRow(row, ind);
            if (st && RANKING_WEIGHTS[st] !== undefined) {
              sum += RANKING_WEIGHTS[st];
              count++;
            }
          }
        });
        result[ind] = count > 0 ? Number((sum / count).toFixed(2)) : null;
      });

      // Calcula a Média Geral do Município para o mês
      let totalSum = 0;
      let totalCount = 0;
      monthRows.forEach(row => {
        const ind = row._mappedIndicator;
        const st = getStatusForRow(row, ind);
        if (st && RANKING_WEIGHTS[st] !== undefined) {
          totalSum += RANKING_WEIGHTS[st];
          totalCount++;
        }
      });
      result['Média Geral'] = totalCount > 0 ? Number((totalSum / totalCount).toFixed(2)) : null;

      return result;
    });
  }, [allData, optM, optI, RANKING_WEIGHTS]);

  const CHART_COLORS = ['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4'];

  const { ubsRanking, indRanking, eqRanking } = useMemo(() => {
    const ubsScores = new Map<string, { sum: number, count: number, statusCounts: Record<string, number> }>();
    const eqScores = new Map<string, { sum: number, count: number, statusCounts: Record<string, number> }>();
    const indScores = new Map<string, { sum: number, count: number, statusCounts: Record<string, number> }>();

    optI.forEach(ind => {
      indScores.set(ind, { sum: 0, count: 0, statusCounts: {'Ótimo':0, 'Bom':0, 'Suficiente':0, 'Regular':0} });
    });

    filteredData.forEach(row => {
      const ind = row._mappedIndicator;
      const st = getStatusForRow(row, ind);

      if (st && RANKING_WEIGHTS[st] !== undefined) {
        const w = RANKING_WEIGHTS[st];
        const ubs = row._ubsLabel;
        const eq = row._teamLabel;

        if (ubs !== 'SEM UNIDADE' && ubs !== 'GERAL' && !ubs.includes('CENTRO DE SAUDE')) {
            if (!ubsScores.has(ubs)) ubsScores.set(ubs, { sum: 0, count: 0, statusCounts: {'Ótimo':0, 'Bom':0, 'Suficiente':0, 'Regular':0} });
            ubsScores.get(ubs)!.sum += w;
            ubsScores.get(ubs)!.count++;
            ubsScores.get(ubs)!.statusCounts[st]++;

            if (eq && !eq.toUpperCase().includes('SEM EQUIPE')) {
                const eqKey = `${eq} | ${ubs}`; 
                if (!eqScores.has(eqKey)) eqScores.set(eqKey, { sum: 0, count: 0, statusCounts: {'Ótimo':0, 'Bom':0, 'Suficiente':0, 'Regular':0} });
                eqScores.get(eqKey)!.sum += w;
                eqScores.get(eqKey)!.count++;
                eqScores.get(eqKey)!.statusCounts[st]++;
            }
        }

        if (indScores.has(ind)) {
           indScores.get(ind)!.sum += w;
           indScores.get(ind)!.count++;
           indScores.get(ind)!.statusCounts[st]++;
        }
      }
    });

    const sortScores = (map: Map<string, any>) => Array.from(map.entries()).map(([name, data]) => ({ 
      name, 
      avg: data.count > 0 ? data.sum / data.count : 0, 
      ...data.statusCounts 
    })).sort((a, b) => b.avg - a.avg);

    return { 
      ubsRanking: sortScores(ubsScores), 
      indRanking: sortScores(indScores),
      eqRanking: sortScores(eqScores) 
    };
  }, [filteredData, optI, RANKING_WEIGHTS]);

  const top10UBS = ubsRanking.slice(0, 10);
  const bottom10UBS = [...ubsRanking].reverse().slice(0, 10);
  
  const top10Eq = eqRanking.slice(0, 10);
  const bottom10Eq = [...eqRanking].reverse().slice(0, 10);

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
      <section className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-4 relative z-20">
        <div className="flex items-center text-lg font-bold text-slate-800">
          <div className={`p-2 rounded-lg mr-3 ${bgClass} bg-opacity-10`}><Trophy className={`h-6 w-6 text-blue-600`} /></div>
          Ranking de Desempenho Global
          
          <div className="relative group ml-3 flex items-center">
            <div className="flex items-center justify-center w-5 h-5 rounded-full border-2 border-slate-300 text-slate-400 hover:border-blue-500 hover:text-blue-500 hover:bg-blue-50 font-bold text-[11px] italic transition-all cursor-help shadow-sm">
              i
            </div>
            <div className="absolute left-0 sm:left-1/2 sm:-translate-x-1/2 top-full mt-2 w-72 p-4 bg-slate-800 text-white text-sm rounded-xl shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-50">
              <p className="font-bold text-white mb-2">Como é calculada a Média?</p>
              <p className="text-slate-300 text-xs mb-3 leading-relaxed">
                Utilizamos uma <strong>Média Ponderada</strong> para o ranking ser justo, independentemente do tamanho da UBS. Cada avaliação tem um peso:
              </p>
              <ul className="text-xs space-y-2 mb-3 font-medium">
                <li className="flex items-center"><span className="w-2.5 h-2.5 bg-blue-400 rounded-full mr-2 shadow-sm"></span> Ótimo = 4 pontos</li>
                <li className="flex items-center"><span className="w-2.5 h-2.5 bg-emerald-400 rounded-full mr-2 shadow-sm"></span> Bom = 3 pontos</li>
                <li className="flex items-center"><span className="w-2.5 h-2.5 bg-amber-400 rounded-full mr-2 shadow-sm"></span> Suficiente = 2 pontos</li>
                <li className="flex items-center"><span className="w-2.5 h-2.5 bg-red-400 rounded-full mr-2 shadow-sm"></span> Regular = 1 ponto</li>
              </ul>
              <p className="text-slate-400 text-[11px] leading-tight pt-2.5 border-t border-slate-600">
                Os pontos são somados e divididos pelo total de avaliações (Nota Máxima: 4.00).
              </p>
            </div>
          </div>
        </div>

        <div className="w-full sm:w-64">
          <select value={selectedMes} onChange={e => setSelectedMes(e.target.value)} className="w-full p-2 border border-slate-300 rounded-lg text-sm bg-white font-medium shadow-sm hover:border-blue-400 cursor-pointer transition-colors outline-none">
            {optM.map(o => <option key={o} value={o}>🏆 Mês: {o}</option>)}
          </select>
        </div>
      </section>

      {/* RANKING DE UBS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div onClick={() => setShowTopUbs(!showTopUbs)} className="bg-blue-50 p-4 border-b border-blue-100 flex items-center justify-between cursor-pointer hover:bg-blue-100/60 transition-colors">
             <div className="flex items-center">
               <Award className="h-5 w-5 text-blue-600 mr-2" />
               <h3 className="font-bold text-blue-900">TOP 10 Melhores UBS</h3>
             </div>
             {showTopUbs ? <ChevronUp className="h-5 w-5 text-blue-400" /> : <ChevronDown className="h-5 w-5 text-blue-400" />}
          </div>
          {showTopUbs && (
          <div className="divide-y divide-slate-100 transition-all">
             {top10UBS.map((ubs, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className={`font-black text-lg w-6 text-center ${i === 0 ? 'text-amber-400' : i === 1 ? 'text-slate-400' : i === 2 ? 'text-amber-600' : 'text-slate-300'}`}>#{i + 1}</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-1">{ubs.name}</p>
                     <div className="flex space-x-1.5">
                       {ubs['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{ubs['Ótimo']} Ótimo</span>}
                       {ubs['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{ubs['Bom']} Bom</span>}
                       {ubs['Suficiente'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{ubs['Suficiente']} Suf.</span>}
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
          )}
        </div>

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div onClick={() => setShowBottomUbs(!showBottomUbs)} className="bg-red-50 p-4 border-b border-red-100 flex items-center justify-between cursor-pointer hover:bg-red-100/60 transition-colors">
             <div className="flex items-center">
               <AlertOctagon className="h-5 w-5 text-red-600 mr-2" />
               <h3 className="font-bold text-red-900">TOP 10 Piores UBS</h3>
             </div>
             {showBottomUbs ? <ChevronUp className="h-5 w-5 text-red-400" /> : <ChevronDown className="h-5 w-5 text-red-400" />}
          </div>
          {showBottomUbs && (
          <div className="divide-y divide-slate-100 transition-all">
             {bottom10UBS.map((ubs, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className="font-black text-lg w-6 text-center text-red-300">!</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-1">{ubs.name}</p>
                     <div className="flex space-x-1.5">
                       {ubs['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{ubs['Ótimo']} Ótimo</span>}
                       {ubs['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{ubs['Bom']} Bom</span>}
                       {ubs['Suficiente'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{ubs['Suficiente']} Suf.</span>}
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
          )}
        </div>
      </div>

      {/* RANKING DE EQUIPAS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div onClick={() => setShowTopEq(!showTopEq)} className="bg-indigo-50 p-4 border-b border-indigo-100 flex items-center justify-between cursor-pointer hover:bg-indigo-100/60 transition-colors">
             <div className="flex items-center">
               <Award className="h-5 w-5 text-indigo-600 mr-2" />
               <h3 className="font-bold text-indigo-900">TOP 10 Melhores Equipes</h3>
             </div>
             {showTopEq ? <ChevronUp className="h-5 w-5 text-indigo-400" /> : <ChevronDown className="h-5 w-5 text-indigo-400" />}
          </div>
          {showTopEq && (
          <div className="divide-y divide-slate-100 transition-all">
             {top10Eq.map((eq, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className={`font-black text-lg w-6 text-center ${i === 0 ? 'text-amber-400' : i === 1 ? 'text-slate-400' : i === 2 ? 'text-amber-600' : 'text-slate-300'}`}>#{i + 1}</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-0.5">{eq.name.split(' | ')[0]}</p>
                     <p className="text-[10px] font-bold text-slate-400 mb-1.5 uppercase">{eq.name.split(' | ')[1]}</p>
                     <div className="flex space-x-1.5">
                       {eq['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{eq['Ótimo']} Ótimo</span>}
                       {eq['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{eq['Bom']} Bom</span>}
                       {eq['Suficiente'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{eq['Suficiente']} Suf.</span>}
                       {eq['Regular'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-red-100 text-red-700 font-bold">{eq['Regular']} Reg.</span>}
                     </div>
                   </div>
                 </div>
                 <div className="text-right">
                   <p className="text-sm font-bold text-indigo-700">{eq.avg.toFixed(2)}</p>
                   <p className="text-[9px] text-slate-400 uppercase font-semibold">Média</p>
                 </div>
               </div>
             ))}
             {top10Eq.length === 0 && <div className="p-8 text-center text-slate-400 text-sm">Sem dados suficientes.</div>}
          </div>
          )}
        </div>

        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div onClick={() => setShowBottomEq(!showBottomEq)} className="bg-orange-50 p-4 border-b border-orange-100 flex items-center justify-between cursor-pointer hover:bg-orange-100/60 transition-colors">
             <div className="flex items-center">
               <AlertOctagon className="h-5 w-5 text-orange-600 mr-2" />
               <h3 className="font-bold text-orange-900">TOP 10 Piores Equipes</h3>
             </div>
             {showBottomEq ? <ChevronUp className="h-5 w-5 text-orange-400" /> : <ChevronDown className="h-5 w-5 text-orange-400" />}
          </div>
          {showBottomEq && (
          <div className="divide-y divide-slate-100 transition-all">
             {bottom10Eq.map((eq, i) => (
               <div key={i} className="flex items-center justify-between p-3.5 hover:bg-slate-50 transition-colors">
                 <div className="flex items-center space-x-4">
                   <span className="font-black text-lg w-6 text-center text-orange-300">!</span>
                   <div>
                     <p className="font-bold text-slate-800 text-sm leading-tight mb-0.5">{eq.name.split(' | ')[0]}</p>
                     <p className="text-[10px] font-bold text-slate-400 mb-1.5 uppercase">{eq.name.split(' | ')[1]}</p>
                     <div className="flex space-x-1.5">
                       {eq['Ótimo'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-bold">{eq['Ótimo']} Ótimo</span>}
                       {eq['Bom'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-bold">{eq['Bom']} Bom</span>}
                       {eq['Suficiente'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-bold">{eq['Suficiente']} Suf.</span>}
                       {eq['Regular'] > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-red-100 text-red-700 font-bold">{eq['Regular']} Reg.</span>}
                     </div>
                   </div>
                 </div>
                 <div className="text-right">
                   <p className="text-sm font-bold text-orange-600">{eq.avg.toFixed(2)}</p>
                   <p className="text-[9px] text-slate-400 uppercase font-semibold">Média</p>
                 </div>
               </div>
             ))}
             {bottom10Eq.length === 0 && <div className="p-8 text-center text-slate-400 text-sm">Sem dados suficientes.</div>}
          </div>
          )}
        </div>
      </div>

      <section className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div onClick={() => setShowInd(!showInd)} className="bg-slate-50 p-4 border-b border-slate-200 flex items-center justify-between cursor-pointer hover:bg-slate-200/60 transition-colors relative z-10">
           <div className="flex items-center">
             <BarChart2 className="h-5 w-5 text-blue-600 mr-2" />
             <h3 className="font-bold text-slate-800">Ranking de Indicadores (Total de Avaliações)</h3>
           </div>
           
           <div className="flex items-center space-x-4">
             <div className="relative group flex items-center" onClick={(e) => e.stopPropagation()}>
               <div className="flex items-center justify-center w-5 h-5 rounded-full border-2 border-slate-300 text-slate-400 hover:border-blue-500 hover:text-blue-500 hover:bg-blue-50 font-bold text-[11px] italic transition-all cursor-help shadow-sm">
                 i
               </div>
               <div className="absolute right-0 top-full mt-2 w-72 p-4 bg-slate-800 text-white text-sm rounded-xl shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-50">
                 <p className="font-bold text-white mb-2">Média do Indicador</p>
                 <p className="text-slate-300 text-xs mb-3 leading-relaxed">
                   A média ponderada revela quais indicadores têm a melhor qualidade de entrega no município. Cada avaliação tem um peso:
                 </p>
                 <ul className="text-xs space-y-2 mb-3 font-medium">
                   <li className="flex items-center"><span className="w-2.5 h-2.5 bg-blue-400 rounded-full mr-2 shadow-sm"></span> Ótimo = 4 pontos</li>
                   <li className="flex items-center"><span className="w-2.5 h-2.5 bg-emerald-400 rounded-full mr-2 shadow-sm"></span> Bom = 3 pontos</li>
                   <li className="flex items-center"><span className="w-2.5 h-2.5 bg-amber-400 rounded-full mr-2 shadow-sm"></span> Suficiente = 2 pontos</li>
                   <li className="flex items-center"><span className="w-2.5 h-2.5 bg-red-400 rounded-full mr-2 shadow-sm"></span> Regular = 1 ponto</li>
                 </ul>
                 <p className="text-slate-400 text-[11px] leading-tight pt-2.5 border-t border-slate-600">
                   Os pontos são somados e divididos pelo total de avaliações do indicador.
                 </p>
               </div>
             </div>
             
             {showInd ? <ChevronUp className="h-5 w-5 text-slate-400" /> : <ChevronDown className="h-5 w-5 text-slate-400" />}
           </div>
        </div>
        
        {showInd && (
        <div className="overflow-x-auto transition-all">
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
                <th onClick={() => requestSortInd('Suficiente')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-amber-700">Suficiente {sortConfigInd?.key === 'Suficiente' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('Regular')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group">
                  <div className="flex items-center text-red-700">Regular {sortConfigInd?.key === 'Regular' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1" /> : <ChevronDown className="h-4 w-4 ml-1" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100" />}</div>
                </th>
                <th onClick={() => requestSortInd('avg')} className="px-5 py-3.5 cursor-pointer hover:bg-slate-200 select-none group border-l border-slate-200 bg-slate-200/50">
                  <div className="flex items-center justify-end text-slate-800 font-bold">Média (Pontos) {sortConfigInd?.key === 'avg' ? (sortConfigInd.direction === 'asc' ? <ChevronUp className="h-4 w-4 ml-1 text-slate-800" /> : <ChevronDown className="h-4 w-4 ml-1 text-slate-800" />) : <ArrowUpDown className="h-3 w-3 ml-1 opacity-0 group-hover:opacity-100 text-slate-400" />}</div>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sortedIndRanking.map((ind, i) => (
                <tr key={i} className="hover:bg-slate-50 transition-colors">
                  <td className="px-5 py-3.5 font-semibold text-slate-800">{ind.name}</td>
                  <td className="px-5 py-3.5 font-bold text-blue-600 bg-blue-50/30">{ind['Ótimo']}</td>
                  <td className="px-5 py-3.5 font-bold text-emerald-600 bg-emerald-50/30">{ind['Bom']}</td>
                  <td className="px-5 py-3.5 font-bold text-amber-600 bg-amber-50/30">{ind['Suficiente']}</td>
                  <td className="px-5 py-3.5 font-bold text-red-600 bg-red-50/30">{ind['Regular']}</td>
                  <td className="px-5 py-3.5 font-black text-slate-700 bg-slate-100/50 border-l border-slate-200 text-right">
                    {ind.avg.toFixed(2)}
                  </td>
                </tr>
              ))}
              {sortedIndRanking.length === 0 && <tr><td colSpan={6} className="px-5 py-8 text-center text-slate-400">Sem dados para este mês.</td></tr>}
            </tbody>
          </table>
        </div>
        )}
      </section>

{/* NOVA SEÇÃO: EVOLUÇÃO MENSAL (Múltiplos Gráficos - Estilo Montanha/Área) */}
<section className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mt-6">
        <div onClick={() => setShowEvolution(!showEvolution)} className="bg-slate-50 p-4 border-b border-slate-200 flex items-center justify-between cursor-pointer hover:bg-slate-200/60 transition-colors relative z-10">
          <div className="flex items-center">
            <BarChart2 className="h-5 w-5 text-blue-600 mr-2" />
            <div>
              <h3 className="font-bold text-slate-800">Evolução Mensal por Indicador</h3>
              <p className="text-[10px] text-slate-500 font-medium mt-0.5 uppercase tracking-wider">
                <span className="inline-block w-2.5 h-2.5 bg-blue-400 rounded-sm mr-1 align-middle"></span> Área: Nota Média do Indicador
              </p>
            </div>
          </div>
          {showEvolution ? <ChevronUp className="h-5 w-5 text-slate-400" /> : <ChevronDown className="h-5 w-5 text-slate-400" />}
        </div>
        
        {showEvolution && (
          <div className="p-2 sm:p-4">
            {optI.map((ind, i) => (
              <div key={ind} className="flex flex-col lg:flex-row items-center gap-2 lg:gap-6 py-5 border-b border-slate-100 last:border-0">
                
                {/* Lado Esquerdo: Nome do Indicador */}
                <div className="w-full lg:w-1/3 xl:w-1/4">
                  <h4 className="font-bold text-slate-700 text-sm pr-4 border-l-4 border-transparent pl-2" style={{ borderLeftColor: CHART_COLORS[i % CHART_COLORS.length] }}>
                    {ind}
                  </h4>
                </div>
                
{/* Lado Direito: Gráfico de Montanha (AreaChart) */}
<div className="w-full lg:w-2/3 xl:w-3/4" style={{ height: 160 }}>
                  <ResponsiveContainer width="100%" height="100%">
                    {/* Aumentei a margem do topo para 25 para os números não cortarem */}
                    <AreaChart data={evolutionData} margin={{ top: 25, right: 20, left: -25, bottom: 0 }}>
                      
                      <defs>
                        <linearGradient id={`color-${i}`} x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={CHART_COLORS[i % CHART_COLORS.length]} stopOpacity={0.6}/>
                          <stop offset="95%" stopColor={CHART_COLORS[i % CHART_COLORS.length]} stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="name" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                      
                      {/* ZOOM AUTOMÁTICO (auto, auto) PARA EXAGERAR A MONTANHA */}
                      <YAxis 
                        domain={['auto', 'auto']} 
                        stroke="#64748b" 
                        fontSize={10} 
                        tickLine={false} 
                        axisLine={false} 
                        tickFormatter={(v) => v.toFixed(1)} 
                      />
                      
                      <Tooltip 
                        formatter={(val: number) => [val ? val.toFixed(2) : '-', 'Nota Média do Indicador']} 
                      />
                      
                      {/* type="natural" cria curvas muito mais orgânicas e arredondadas */}
                      <Area 
                        type="linear" 
                        dataKey={ind} 
                        stroke={CHART_COLORS[i % CHART_COLORS.length]} 
                        strokeWidth={3}
                        fillOpacity={1} 
                        fill={`url(#color-${i})`}
                        dot={{ r: 4, strokeWidth: 2, fill: 'white' }}
                        activeDot={{ r: 6 }}
                      >
                        <LabelList 
                          dataKey={ind} 
                          position="top" 
                          offset={12}
                          fontSize={10} 
                          fontWeight="bold" 
                          fill="#475569" 
                          formatter={(v: number) => v ? v.toFixed(2) : ''} 
                        />
                      </Area>
                    </AreaChart>
                  </ResponsiveContainer>
                </div>

              </div>
            ))}
          </div>
        )}
      </section>

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