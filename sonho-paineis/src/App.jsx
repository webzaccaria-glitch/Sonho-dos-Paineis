import React, { useState, useEffect, useMemo, useRef } from 'react';
import { dbGet, dbSet } from './supabase.js';
import { Plus, Trash2, Package, Calculator, TrendingUp, X, Check, AlertCircle,
         ChevronDown, ChevronLeft, ChevronRight, Lock, Unlock, Eye, EyeOff, KeyRound,
         Pencil, PackagePlus, FileDown, FileSpreadsheet, Calendar, Search,
         ArrowUp, ArrowDown, GripVertical } from 'lucide-react';
import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────
// ERROR BOUNDARY — captura erros de render
// ─────────────────────────────────────────────
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(e) { return { error: e }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{margin:'16px',padding:'20px',background:'#FCE8E6',borderRadius:'16px',color:'#B5302B'}}>
          <strong style={{display:'block',marginBottom:'8px'}}>⚠️ Erro no componente Exportar:</strong>
          <code style={{fontSize:'12px',wordBreak:'break-all'}}>{this.state.error.message}</code>
          <button onClick={()=>this.setState({error:null})} style={{marginTop:'12px',display:'block',padding:'8px 16px',background:'#B5302B',color:'white',border:'none',borderRadius:'8px',cursor:'pointer'}}>
            Tentar novamente
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}



// ─────────────────────────────────────────────
// CONSTANTES
// ─────────────────────────────────────────────
const KITS_PADRAO = [
  { id:'kit01', nome:'Kit 01 — painel + suporte', metragem:+(10*.55/3).toFixed(4), suporte:15, tecido:'helanca' , ativo:true },
  { id:'kit02', nome:'Kit 02 — 5 painéis',        metragem:+(5*.55/3).toFixed(4),  suporte:0,  tecido:'helanca' , ativo:true },
  { id:'kit03', nome:'Kit 03 — unitário',          metragem:+(1*.55/3).toFixed(4),  suporte:0,  tecido:'helanca' , ativo:true },
  { id:'kit04', nome:'Kit 04 — Painel redondo',    metragem:1.60,                   suporte:0,  tecido:'helanca' , ativo:true },
  { id:'kit05', nome:'Kit 05 — kit + portal',      metragem:6.00,                   suporte:0,  tecido:'helanca' , ativo:true },
  { id:'kit06', nome:'Kit 06 — Portal',            metragem:2.10,                   suporte:0,  tecido:'helanca' , ativo:true },
  { id:'kit07', nome:'Kit 07 — Shopee',            metragem:+(10*.55/3).toFixed(4), suporte:0,  tecido:'helanca' , ativo:true },
];
const PRECOS_PAD  = {kit01:100,kit02:60,kit03:20,kit04:50,kit05:190,kit06:60,kit07:60};
const CONSUMO_PAD = {kit01:{el:0,li:5},kit02:{el:0,li:3},kit03:{el:0,li:1},kit04:{el:5,li:4},kit05:{el:12,li:15},kit06:{el:4,li:5},kit07:{el:0,li:5}};
const PREM_PAD    = {tt:460,tr:80, htt:460,htr:80, hrt:460,hrr:80, pt:520,pr:300,it:1080,ir:1000,cal:3.5,et:18,er:100,lt:50,lr:500,sup:15,mft:45,mfr:10};

const PAG_CFG = {
  pendente:{ label:'Pendente', color:'#B5302B', bg:'#FCE8E6', factor:0   },
  meio:    { label:'50% pago', color:'#8B6B1A', bg:'#FFF4D6', factor:.5  },
  integral:{ label:'Pago',     color:'#3F6E3A', bg:'#E1EFDB', factor:1   },
};
const PAG_KEYS = ['pendente','meio','integral'];

// ─────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────
const brl    = n  => (isFinite(n)?n:0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const pct    = n  => (isFinite(n)?n*100:0).toFixed(1)+'%';
const uid    = () => 'k'+Date.now().toString(36)+Math.random().toString(36).slice(2,5);
// Usa data LOCAL do dispositivo (evita bug de UTC mudar o dia às 21h no Brasil)
const hoje   = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; };
const mesHj  = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; };
// pFactor: retorna fração paga (0-1). Usa valorPago se definido, senão usa o status.
const pFactor = (p, total) => {
  if (p.valorPago != null && total != null && total > 0)
    return Math.min(Math.max(p.valorPago / total, 0), 1);
  return PAG_CFG[p?.pagamento||'integral']?.factor ?? 1;
};
const recebido = (p, precos) => { const t=pedTotal(p,precos); return t*pFactor(p,t); };
const getItens = p   => p.itens || (p.kitId ? [{kitId:p.kitId, qtd:p.qtd||1}] : []);
const pedTotal = (p,prc) => getItens(p).reduce((s,it)=>s+(prc[it.kitId]||0)*(it.qtd||0),0);

const fmtLong    = iso => new Date(iso).toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
const fmtShort   = iso => new Date(iso+'T12:00:00').toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric'});
// Converte ISO UTC para data local (YYYY-MM-DD) — resolve problema de fuso horário
const localDate  = iso => { const d=new Date(iso); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; };
const localMonth = iso => localDate(iso).slice(0,7);

function semanaRange(ds) {
  const d=new Date(ds+'T12:00:00'), dow=d.getDay();
  const s=new Date(d); s.setDate(d.getDate()-((dow+6)%7)); s.setHours(0,0,0,0);
  const e=new Date(s); e.setDate(s.getDate()+6); e.setHours(23,59,59,999);
  return {s,e};
}

// ─────────────────────────────────────────────
// STORAGE
// ─────────────────────────────────────────────
// dbGet e dbSet importados de supabase.js

async function sha256(txt) {
  try {
    if(crypto?.subtle){
      const b=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(txt));
      return [...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join('');
    }
  } catch{}
  let h=0; for(const c of txt){h=((h<<5)-h)+c.charCodeAt(0);h|=0;} return 'fb'+h.toString(36);
}

// ─────────────────────────────────────────────
// ROOT
// ─────────────────────────────────────────────
export default function App() {
  const [tab,      setTab]  = useState('pedidos');
  const [ready,    setRdy]  = useState(false);
  const [kits,     setKits] = useState(KITS_PADRAO);
  const [pedidos,  setPeds] = useState([]);
  const [precos,   setPrc]  = useState(PRECOS_PAD);
  const [prem,     setPrem] = useState(PREM_PAD);
  const [consumo,  setCons] = useState(CONSUMO_PAD);
  const [pwdHash,  setPwd]  = useState(null);
  const [isAdmin,  setAdm]  = useState(false);
  const [authOpen, setAuth] = useState(false);

  // ── load ──
  useEffect(()=>{ (async()=>{
    // ── Migração automática: localStorage → Supabase (executa 1x por dispositivo) ──
    const migrated = localStorage.getItem('sdp:cloud_migrated');
    if (!migrated) {
      const KEYS = ['sdp:kits','sdp:peds','sdp:precos','sdp:prem','sdp:cons','sdp:pwdh'];
      for (const k of KEYS) {
        const local = localStorage.getItem(k);
        if (local) {
          // Só migra se Supabase ainda não tiver o dado
          const cloud = await dbGet(k, null);
          if (cloud === null) await dbSet(k, JSON.parse(local));
        }
      }
      localStorage.setItem('sdp:cloud_migrated', 'true');
    }

    // ── Carregar dados do Supabase ──
    setKits(await dbGet('sdp:kits',   KITS_PADRAO));
    setPeds(await dbGet('sdp:peds',   []));
    setPrc( await dbGet('sdp:precos', PRECOS_PAD));
    setPrem(await dbGet('sdp:prem',   PREM_PAD));
    setCons(await dbGet('sdp:cons',   CONSUMO_PAD));
    setPwd( await dbGet('sdp:pwdh',   null));
    setRdy(true);
  })(); },[]);

  // ── save ──
  useEffect(()=>{ if(ready) dbSet('sdp:kits',  kits);   },[kits,  ready]);
  useEffect(()=>{ if(ready) dbSet('sdp:peds',  pedidos);},[pedidos,ready]);
  useEffect(()=>{ if(ready) dbSet('sdp:precos',precos); },[precos, ready]);
  useEffect(()=>{ if(ready) dbSet('sdp:prem',  prem);   },[prem,   ready]);
  useEffect(()=>{ if(ready) dbSet('sdp:cons',  consumo);},[consumo,ready]);
  useEffect(()=>{ if(ready&&pwdHash!==null) dbSet('sdp:pwdh',pwdHash); },[pwdHash,ready]);

  // redirecionar se perdeu admin
  useEffect(()=>{ if(!isAdmin&&tab!=='pedidos') setTab('pedidos'); },[isAdmin]);

  // ── custos por kit ──
  const custosKits = useMemo(()=>{
    const cHelanca  = prem.tt/prem.tr;            // Helanca original
    const cTubular  = prem.htt/prem.htr;          // Helanca Tubular 1,20m
    const cRamada   = prem.hrt/prem.hrr;          // Helanca Ramada 1,80m
    const cMicro    = prem.mft/prem.mfr;          // Microfibra
    const cP=prem.pt/prem.pr, cI=prem.it/prem.ir;
    const cE=prem.et/prem.er, cL=prem.lt/prem.lr, cC=prem.cal;
    return kits.map(k=>{
      const tec = k.tecido||'helanca';
      const cT  = tec==='microfibra'       ? cMicro  :
                  tec==='helanca_tubular'   ? cTubular :
                  tec==='helanca_ramada'    ? cRamada  :
                  cHelanca;
      const cn=consumo[k.id]||{el:0,li:0};
      const ct=k.metragem*cT, cp=k.metragem*cP, cc=k.metragem*cC,
            ci=k.metragem*cI, ce=cn.el*cE, cl=cn.li*cL, cs=k.suporte||0;
      const tot=ct+cp+cc+ci+ce+cl+cs, pr=precos[k.id]||0;
      return {...k, det:{ct,cp,cc,ci,ce,cl,cs,tot}, pr, mg:pr-tot, mgPct:pr>0?(pr-tot)/pr:0};
    });
  },[kits,prem,consumo,precos]);

  // ── totais pedidos ──
  const totais = useMemo(()=>{
    let prev=0,rec=0;
    pedidos.forEach(p=>{ const v=pedTotal(p,precos); prev+=v; rec+=recebido(p,precos); });
    return {prev,rec,pend:prev-rec,n:pedidos.length};
  },[pedidos,precos]);

  // ── resumo ──
  const resumo = useMemo(()=>{
    const pk=kits.map(k=>{
      const pu=precos[k.id]||0, cu=custosKits.find(c=>c.id===k.id)?.det.tot||0;
      // agregar itens desse kit em todos os pedidos (suporta multi-itens)
      const pares=pedidos.flatMap(p=>{ const t=pedTotal(p,precos); return getItens(p).filter(it=>it.kitId===k.id).map(it=>({qtd:it.qtd||0,fac:pFactor(p,t)})); });
      const qtd=pares.reduce((s,x)=>s+x.qtd,0);
      const rP=qtd*pu;
      const rR=pares.reduce((s,x)=>s+pu*x.qtd*x.fac,0);
      return {...k,qtd,pu,cu,cT:qtd*cu,rP,rR,lP:rP-qtd*cu,lR:rR-qtd*cu};
    });
    const tRP=pk.reduce((s,x)=>s+x.rP,0), tRR=pk.reduce((s,x)=>s+x.rR,0), tC=pk.reduce((s,x)=>s+x.cT,0);
    return {pk,tRP,tRR,tC,lP:tRP-tC,lR:tRR-tC,pend:tRP-tRR};
  },[kits,pedidos,custosKits,precos]);

  // ── admin auth ──
  const unlock = async (senha)=>{
    if(!pwdHash){ const h=await sha256(senha); setPwd(h); setAdm(true); setAuth(false); return {ok:true}; }
    const h=await sha256(senha);
    if(h===pwdHash){ setAdm(true); setAuth(false); return {ok:true}; }
    return {ok:false,msg:'Senha incorreta'};
  };
  const changePwd = async (a,n)=>{
    if(await sha256(a)!==pwdHash) return {ok:false,msg:'Senha atual incorreta'};
    setPwd(await sha256(n)); return {ok:true};
  };
  const lock = ()=>{ setAdm(false); setTab('pedidos'); };

  // ── gestão kits ──
  const addKit    = d=>{ const id=uid(); setKits(p=>[...p,{...d,id}]); setPrc(p=>({...p,[id]:d.pr||0})); setCons(p=>({...p,[id]:{el:0,li:0}})); };
  const updateKit = (id,d)=>{ setKits(p=>p.map(k=>k.id===id?{...k,nome:d.nome,metragem:d.metragem,suporte:d.suporte,tecido:d.tecido||k.tecido||'helanca'}:k)); setPrc(p=>({...p,[id]:d.pr||0})); };
  const removeKit = id=>setKits(p=>p.filter(k=>k.id!==id));
  const moveKit       = (id,dir)=>setKits(prev=>{
    const i=prev.findIndex(k=>k.id===id);
    if(dir==='up'&&i===0||dir==='down'&&i===prev.length-1) return prev;
    const n=[...prev], j=dir==='up'?i-1:i+1;
    [n[i],n[j]]=[n[j],n[i]]; return n;
  });
  const reorderKit    = (fromId,toIdx)=>setKits(prev=>{
    const fi=prev.findIndex(k=>k.id===fromId);
    if(fi<0||fi===toIdx) return prev;
    const n=[...prev]; const [item]=n.splice(fi,1); n.splice(toIdx,0,item); return n;
  });
  const toggleKitAtivo= id=>setKits(p=>p.map(k=>k.id===id?{...k,ativo:k.ativo===false?true:false}:k));

  if(!ready) return <div className="min-h-screen flex items-center justify-center" style={{background:'#FAF7F2'}}><p className="text-stone-400">Carregando…</p></div>;

  const ADMIN_TABS=[
    {id:'pedidos',  l:'Pedidos',   I:Package},
    {id:'custos',   l:'Custos',    I:Calculator},
    {id:'resumo',   l:'Resumo',    I:TrendingUp},
    {id:'exportar', l:'Exportar',  I:FileSpreadsheet},
  ];
  const PUBLIC_TABS=[{id:'pedidos',l:'Pedidos',I:Package}];
  const tabs = isAdmin ? ADMIN_TABS : PUBLIC_TABS;

  return (
    <div className="min-h-screen" style={{background:'#FAF7F2'}}>


      {/* ── HEADER ── */}
      <header className="sticky top-0 z-20 border-b border-stone-200/70 bg-white/70 backdrop-blur">
        <div className="max-w-4xl mx-auto px-4 pt-4 flex items-center justify-between">
          <div>
            <h1 className="fd font-semibold text-stone-900 leading-tight" style={{fontSize:'clamp(1rem,4vw,1.4rem)',letterSpacing:'.03em'}}>
              SONHO DOS <span className="italic" style={{color:'#C65D3C'}}>PAINÉIS</span>
            </h1>
            <p className="text-[11px] text-stone-500 tracking-widest uppercase mt-0.5">controle de entrada</p>
          </div>
          <button onClick={()=>isAdmin?lock():setAuth(true)}
            className={`w-10 h-10 rounded-full flex items-center justify-center transition-all ${isAdmin?'bg-stone-900 text-white':'bg-stone-100 text-stone-500 hover:bg-stone-200'}`}
            title={isAdmin?'Bloquear admin':'Área administrativa'}>
            {isAdmin?<Unlock size={15}/>:<Lock size={15}/>}
          </button>
        </div>
        <nav className="max-w-4xl mx-auto px-4 flex overflow-x-auto">
          {tabs.map(({id,l,I})=>(
            <button key={id} onClick={()=>setTab(id)}
              className={`relative px-4 py-3 text-sm font-medium flex items-center gap-1.5 whitespace-nowrap transition-colors ${tab===id?'text-stone-900':'text-stone-500 hover:text-stone-700'}`}>
              <I size={15}/>{l}
              {tab===id&&<span className="absolute bottom-0 left-3 right-3 h-0.5 rounded-full" style={{background:'#C65D3C'}}/>}
            </button>
          ))}
        </nav>
      </header>

      {/* ── MAIN ── */}
      <main className="max-w-4xl mx-auto px-4 py-5 pb-24">
        {tab==='pedidos' &&
          <TabPedidos kits={kits} pedidos={pedidos} setPedidos={setPeds}
            precos={precos} totais={totais} isAdmin={isAdmin}/>}
        {tab==='custos' && isAdmin &&
          <TabCustos kits={kits} custosKits={custosKits}
            prem={prem} setPrem={setPrem}
            consumo={consumo} setCons={setCons}
            precos={precos} setPrecos={setPrc}
            pedidos={pedidos}
            onAdd={addKit} onUpdate={updateKit} onRemove={removeKit} onMove={moveKit}
            onReorder={reorderKit} onToggleAtivo={toggleKitAtivo}
            onChangePwd={changePwd}/>}
        {tab==='resumo' && isAdmin &&
          <TabResumo custosKits={custosKits} pedidos={pedidos} precos={precos} kits={kits}/>}
        {tab==='exportar' && isAdmin &&
          <ErrorBoundary><TabExportar pedidos={pedidos} kits={kits} precos={precos}/></ErrorBoundary>}
      </main>

      {authOpen&&<ModalAuth hasHash={!!pwdHash} onClose={()=>setAuth(false)} onSubmit={unlock}/>}
    </div>
  );
}

// ─────────────────────────────────────────────
// TAB PEDIDOS
// ─────────────────────────────────────────────
function TabPedidos({kits,pedidos,setPedidos,precos,totais,isAdmin}) {
  const [viewDate,    setVD]  = useState(hoje());
  const [searchMode,  setSM]  = useState(false);
  const [searchQuery, setSQ]  = useState('');
  const [novoCli,     setNC]  = useState('');
  const [novoPag,     setNP]  = useState('integral');
  const [novoItens,   setNI]  = useState([{kitId:'',qtd:1}]);
  const [novoVP,      setNVP] = useState('');   // valor pago customizado
  const [form,        setFm]  = useState(false);
  const [filtro,      setFi]  = useState('todos');
  const [verCob,      setVC]  = useState(false);
  const [openKit,     setOK]  = useState(null); // qual seletor de kit está aberto
  const [editando,    setEd]  = useState(null); // pedido sendo editado

  const prevDay = () => { const d=new Date(viewDate+'T12:00:00'); d.setDate(d.getDate()-1); setVD(d.toISOString().slice(0,10)); };
  const nextDay = () => { const d=new Date(viewDate+'T12:00:00'); d.setDate(d.getDate()+1); setVD(d.toISOString().slice(0,10)); };

  const addItem     = ()        => setNI(p=>[...p,{kitId:'',qtd:1}]);
  const removeItem  = idx       => setNI(p=>p.filter((_,i)=>i!==idx));
  const updateItem  = (idx,f,v) => setNI(p=>p.map((it,i)=>i===idx?{...it,[f]:v}:it));
  const novoTotal   = novoItens.reduce((s,it)=>s+(precos[it.kitId]||0)*(it.qtd||0),0);
  const novoRecebido= novoVP!==''&&+novoVP>=0 ? Math.min(+novoVP,novoTotal) : novoTotal*(PAG_CFG[novoPag]?.factor??1);
  const canSubmit   = novoCli.trim() && novoItens.some(it=>it.kitId&&+it.qtd>0);

  const add = ()=>{
    if(!canSubmit) return;
    const validos = novoItens.filter(it=>it.kitId&&+it.qtd>0);
    const id = pedidos.length ? Math.max(...pedidos.map(p=>p.id))+1 : 1;
    const vp = novoVP!==''&&+novoVP>=0 ? +novoVP : null;
    setPedidos(p=>[...p,{id, cliente:novoCli.trim(), itens:validos, pagamento:novoPag, valorPago:vp, data:new Date().toISOString()}]);
    setNC(''); setNI([{kitId:'',qtd:1}]); setNP('integral'); setNVP(''); setFm(false); setOK(null);
  };
  const del   = id => setPedidos(p=>p.filter(x=>x.id!==id));
  const setPg = (id,v) => setPedidos(p=>p.map(x=>x.id===id?{...x,pagamento:v,valorPago:null}:x));
  const setVP = (id,v) => setPedidos(p=>p.map(x=>x.id===id?{...x,valorPago:v!=null?+v:null}:x));

  const updatePedido = (id, dados) => {
    setPedidos(p=>p.map(x=>x.id===id ? {...x,...dados} : x));
    setEd(null);
  };

  // Cobranças: pedidos com valor em aberto (qualquer data)
  const cobrancas = useMemo(()=>
    [...pedidos].filter(p=>{ const t=pedTotal(p,precos); return t>0 && recebido(p,precos)<t; })
      .sort((a,b)=>new Date(a.data)-new Date(b.data))
  ,[pedidos,precos]);
  const totalCob = useMemo(()=>cobrancas.reduce((s,p)=>s+pedTotal(p,precos)-recebido(p,precos),0),[cobrancas,precos]);

  // Lista filtrada por data OU busca
  const lista = useMemo(()=>{
    const sorted=[...pedidos].sort((a,b)=>new Date(b.data)-new Date(a.data));
    if(searchMode && searchQuery.trim())
      return sorted.filter(p=>p.cliente?.toLowerCase().includes(searchQuery.toLowerCase()));
    const base = sorted.filter(p=>localDate(p.data)===viewDate);
    if(filtro==='todos') return base;
    return base.filter(p=>(p.pagamento||'integral')===filtro);
  },[pedidos,viewDate,searchMode,searchQuery,filtro]);

  const viewTotals = useMemo(()=>{
    let prev=0,rec=0;
    lista.forEach(p=>{ prev+=pedTotal(p,precos); rec+=recebido(p,precos); });
    return {prev,rec,pend:prev-rec};
  },[lista,precos]);

  const cnt=useMemo(()=>({
    todos:pedidos.filter(p=>localDate(p.data)===viewDate).length,
    pendente:pedidos.filter(p=>localDate(p.data)===viewDate&&(p.pagamento||'integral')==='pendente').length,
    meio:pedidos.filter(p=>localDate(p.data)===viewDate&&p.pagamento==='meio').length,
    integral:pedidos.filter(p=>localDate(p.data)===viewDate&&(p.pagamento||'integral')==='integral').length,
  }),[pedidos,viewDate]);

  const isHoje = viewDate===hoje();

  return (
    <div className="space-y-3">

      {/* ── BANNER cobranças (admin, fora da view cobranças) ── */}
      {!verCob && !searchMode && cobrancas.length>0 && (
        <button onClick={()=>{setVC(true);setFm(false);}}
          className="w-full rounded-2xl p-4 flex items-center justify-between border-2 hover:opacity-90 transition-all ai"
          style={{borderColor:'#C65D3C',background:'#FFF8F5'}}>
          <div className="text-left">
            <div className="font-semibold text-stone-900">
              💰 {cobrancas.length} cobrança{cobrancas.length!==1?'s':''} pendente{cobrancas.length!==1?'s':''}
            </div>
            <div className="text-sm text-stone-500 mt-0.5">{brl(totalCob)} a receber</div>
          </div>
          <ChevronRight size={18} className="text-stone-400 flex-shrink-0"/>
        </button>
      )}

      {/* ════════ VIEW COBRANÇAS ════════ */}
      {verCob ? (
        <div className="ai space-y-3">
          <div className="flex items-center gap-3">
            <button onClick={()=>setVC(false)}
              className="w-9 h-9 rounded-xl bg-white border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-50 flex-shrink-0">
              <ChevronLeft size={16}/>
            </button>
            <div>
              <div className="font-semibold text-stone-900">Cobranças pendentes</div>
              <div className="text-xs text-stone-500">{cobrancas.length} pedido{cobrancas.length!==1?'s':''} · {brl(totalCob)} a receber</div>
            </div>
          </div>
          {cobrancas.length===0 ? (
            <div className="text-center py-12">
              <div className="inline-flex items-center justify-center w-14 h-14 rounded-full bg-stone-100 mb-3">
                <Check size={22} className="text-stone-400"/>
              </div>
              <p className="text-stone-400 text-sm">Nenhuma cobrança pendente</p>
            </div>
          ) : cobrancas.map(p=>{
            const tot=pedTotal(p,precos);
            const rec=recebido(p,precos);
            const deve=tot-rec;
            const itens=getItens(p);
            return (
              <div key={p.id} className="bg-white rounded-2xl border-2 p-4" style={{borderColor:'#F5D4C8'}}>
                <div className="flex items-start gap-3 mb-3">
                  <div className="w-10 h-10 rounded-xl flex-shrink-0 flex items-center justify-center fd text-sm font-semibold text-white" style={{background:'#2A2420'}}>
                    #{p.id}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="font-semibold text-stone-900">{p.cliente}</div>
                    <div className="text-[11px] text-stone-400">{new Date(p.data).toLocaleDateString('pt-BR')}</div>
                    <div className="mt-1 space-y-0.5">
                      {itens.map((it,i)=>{
                        const kit=kits.find(k=>k.id===it.kitId);
                        return <div key={i} className="text-xs text-stone-500 flex justify-between"><span className="truncate mr-2">{kit?.nome||'—'}</span><span>{it.qtd}un · {brl((precos[it.kitId]||0)*it.qtd)}</span></div>;
                      })}
                    </div>
                  </div>
                  <div className="text-right flex-shrink-0">
                    <div className="text-[10px] text-stone-400">Total</div>
                    <div className="fd text-base font-semibold text-stone-900">{brl(tot)}</div>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-2 mb-3">
                  <div className="rounded-xl p-2.5 text-center" style={{background:'#E1EFDB'}}>
                    <div className="text-[10px] uppercase tracking-wider" style={{color:'#3F6E3A'}}>Já recebido</div>
                    <div className="fd text-base font-semibold" style={{color:'#3F6E3A'}}>{brl(rec)}</div>
                  </div>
                  <div className="rounded-xl p-2.5 text-center" style={{background:'#FCE8E6'}}>
                    <div className="text-[10px] uppercase tracking-wider" style={{color:'#B5302B'}}>A cobrar</div>
                    <div className="fd text-base font-semibold" style={{color:'#B5302B'}}>{brl(deve)}</div>
                  </div>
                </div>
                <button onClick={()=>setPg(p.id,'integral')}
                  className="w-full rounded-xl py-2.5 text-sm font-semibold text-white flex items-center justify-center gap-2 hover:opacity-90"
                  style={{background:'#3F6E3A'}}>
                  <Check size={15}/> Marcar como Pago total
                </button>
              </div>
            );
          })}
        </div>

      ) : ( /* ════════ VIEW NORMAL ════════ */
      <div className="space-y-3">

        {/* Seletor de data ou busca */}
        {!searchMode ? (
          <div className="flex items-center gap-2">
            <button onClick={prevDay} className="w-9 h-9 rounded-xl bg-white border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-50 flex-shrink-0"><ChevronLeft size={16}/></button>
            <div className="flex-1 relative">
              <input type="date" value={viewDate} onChange={e=>setVD(e.target.value)}
                className="w-full bg-white border border-stone-200 rounded-xl px-4 py-2.5 text-stone-900 text-sm focus:outline-none focus:border-stone-400"/>
              {isHoje&&<span className="absolute right-3 top-1/2 -translate-y-1/2 text-[10px] font-medium px-1.5 py-0.5 rounded-full" style={{background:'#E1EFDB',color:'#3F6E3A'}}>Hoje</span>}
            </div>
            {!isHoje&&<button onClick={()=>setVD(hoje())} className="px-3 py-2.5 rounded-xl bg-white border border-stone-200 text-xs font-medium text-stone-600 hover:bg-stone-50 flex-shrink-0">Hoje</button>}
            <button onClick={()=>{setSM(true);setSQ('');setFm(false);}} className="w-9 h-9 rounded-xl bg-white border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-50 flex-shrink-0" title="Buscar cliente">
              <Search size={16}/>
            </button>
            <button onClick={nextDay} className="w-9 h-9 rounded-xl bg-white border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-50 flex-shrink-0"><ChevronRight size={16}/></button>
          </div>
        ) : (
          <div className="flex items-center gap-2 ai">
            <div className="flex-1 relative">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none"/>
              <input type="text" value={searchQuery} onChange={e=>setSQ(e.target.value)} placeholder="Buscar por nome do cliente…" autoFocus
                className="w-full bg-white border border-stone-200 rounded-xl pl-9 pr-4 py-2.5 text-stone-900 text-sm focus:outline-none focus:border-stone-400"/>
            </div>
            <button onClick={()=>{setSM(false);setSQ('');}} className="w-9 h-9 rounded-xl bg-white border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-50"><X size={16}/></button>
          </div>
        )}

        {/* Cards admin */}
        {isAdmin && !searchMode && (
          <div className="grid grid-cols-3 gap-2">
            <MiniCard l="Pedidos"   v={lista.length} num/>
            <MiniCard l="Recebido"  v={brl(viewTotals.rec)}  cor="#3F6E3A"/>
            <MiniCard l="A receber" v={brl(viewTotals.pend)} cor={viewTotals.pend>0?'#B5302B':'#78716C'}/>
          </div>
        )}

        {/* Formulário novo pedido */}
        {!form ? (
          <button onClick={()=>setFm(true)}
            className="w-full bg-stone-900 text-white rounded-2xl py-4 font-medium flex items-center justify-center gap-2 hover:bg-stone-800 transition-colors">
            <Plus size={17}/> Novo pedido
          </button>
        ) : (
          <div className="bg-white rounded-2xl p-5 border border-stone-200/80 ai">
            <div className="flex items-center justify-between mb-4">
              <h3 className="fd text-lg text-stone-900">Novo pedido</h3>
              <button onClick={()=>setFm(false)} className="text-stone-400"><X size={19}/></button>
            </div>
            <div className="space-y-4">
              <Fld l="Cliente">
                <input type="text" value={novoCli} onChange={e=>setNC(e.target.value)} placeholder="Nome do cliente"
                  className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/>
              </Fld>
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs uppercase tracking-wider text-stone-500">Itens do pedido</label>
                  <button onClick={addItem} className="flex items-center gap-1 text-xs font-medium px-3 py-1.5 rounded-full" style={{background:'#F5EBE5',color:'#C65D3C'}}>
                    <Plus size={12}/> Adicionar item
                  </button>
                </div>
                <div className="space-y-2">
                  {novoItens.map((it,idx)=>{
                    const isOpen   = openKit===idx;
                    const selKit   = kits.find(k=>k.id===it.kitId);
                    return (
                      <div key={idx} className="space-y-1.5">
                        {/* Seletor — fechado por padrão, abre ao tocar */}
                        <div className="relative">
                          <button onClick={()=>setOK(isOpen?null:idx)}
                            className={`w-full flex items-center justify-between px-4 py-3.5 rounded-xl border-2 transition-all ${selKit?'border-transparent text-white':'border-stone-200 bg-stone-50 text-stone-400'}`}
                            style={selKit?{background:'#C65D3C'}:{}}>
                            <span className="text-sm font-medium truncate pr-2">
                              {selKit ? selKit.nome : 'Toque para escolher o kit…'}
                            </span>
                            <div className="flex items-center gap-2 flex-shrink-0">
                              {selKit&&<span className="text-sm tabular-nums opacity-75">{brl(precos[it.kitId]||0)}</span>}
                              <ChevronDown size={15} className={`transition-transform duration-200 ${isOpen?'rotate-180':''} ${selKit?'opacity-60':'text-stone-400'}`}/>
                            </div>
                          </button>
                          {isOpen&&(
                            <div className="absolute top-full left-0 right-0 z-30 mt-1 bg-white rounded-2xl border border-stone-200 shadow-xl overflow-hidden">
                              <div className="max-h-60 overflow-y-auto">
                                {kits.filter(k=>k.ativo!==false).map(k=>{
                                  const sel=it.kitId===k.id;
                                  return (
                                    <button key={k.id}
                                      onClick={()=>{updateItem(idx,'kitId',k.id);setOK(null);}}
                                      className={`w-full flex items-center justify-between px-4 py-3.5 border-b border-stone-100 last:border-0 transition-colors text-left ${sel?'text-white':'hover:bg-stone-50 text-stone-700'}`}
                                      style={sel?{background:'#C65D3C'}:{}}>
                                      <span className="text-sm font-medium">{k.nome}</span>
                                      <span className={`text-sm font-bold tabular-nums flex-shrink-0 ml-3 ${sel?'opacity-75':'text-stone-500'}`}>{brl(precos[k.id]||0)}</span>
                                    </button>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                        </div>
                        {/* Qtd + total + remover (só após selecionar) */}
                        {it.kitId&&(
                          <div className="flex items-center gap-2 px-1">
                            <div className="flex items-center gap-1.5">
                              <button onClick={()=>updateItem(idx,'qtd',Math.max(1,(it.qtd||1)-1))} className="w-8 h-8 rounded-xl bg-stone-100 flex items-center justify-center text-stone-600 hover:bg-stone-200 font-bold text-lg leading-none">−</button>
                              <input type="number" inputMode="numeric" value={it.qtd}
                                onChange={e=>updateItem(idx,'qtd',+e.target.value||1)}
                                className="w-12 bg-stone-50 border border-stone-200 rounded-xl py-1.5 text-sm text-stone-900 focus:outline-none text-center font-bold"/>
                              <button onClick={()=>updateItem(idx,'qtd',(it.qtd||1)+1)} className="w-8 h-8 rounded-xl bg-stone-100 flex items-center justify-center text-stone-600 hover:bg-stone-200 font-bold text-lg leading-none">+</button>
                            </div>
                            <span className="text-sm font-bold tabular-nums" style={{color:'#C65D3C'}}>{brl((precos[it.kitId]||0)*(it.qtd||1))}</span>
                            {novoItens.length>1&&<button onClick={()=>removeItem(idx)} className="ml-auto text-stone-300 hover:text-red-500 transition-colors p-1"><X size={14}/></button>}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              <Fld l="Status do pagamento">
                <div className="grid grid-cols-3 gap-2">
                  {PAG_KEYS.map(k=>{ const s=PAG_CFG[k]; const sel=novoPag===k; return (
                    <button key={k} onClick={()=>setNP(k)}
                      className={`rounded-xl py-2.5 text-sm font-medium border transition-all ${sel?'border-transparent':'border-stone-200 bg-stone-50 text-stone-600'}`}
                      style={sel?{background:s.bg,color:s.color}:{}}>{s.label}</button>
                  ); })}
                </div>
              </Fld>
              {/* Campo valor pago */}
              <Fld l="Valor recebido (opcional)">
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 text-sm">R$</span>
                  <input type="number" inputMode="decimal" step="0.01" value={novoVP}
                    onChange={e=>setNVP(e.target.value)}
                    placeholder={novoTotal>0 ? (novoTotal*(PAG_CFG[novoPag]?.factor??1)).toFixed(2) : '0,00'}
                    className="w-full bg-stone-50 border border-stone-200 rounded-xl pl-9 pr-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums"/>
                </div>
                <p className="text-[11px] text-stone-400 mt-1">Deixe em branco para usar o percentual do status acima. Preencha se o valor for diferente.</p>
              </Fld>
              {novoTotal>0&&(
                <div className="bg-stone-50 rounded-xl px-4 py-3 space-y-1.5">
                  <div className="flex justify-between items-center">
                    <span className="text-sm text-stone-600">Total do pedido</span>
                    <span className="fd text-xl text-stone-900">{brl(novoTotal)}</span>
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-stone-500">Recebido</span>
                    <span className="font-medium" style={{color:'#3F6E3A'}}>{brl(novoRecebido)}</span>
                  </div>
                  {novoRecebido<novoTotal&&(
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-stone-500">A receber</span>
                      <span className="font-medium" style={{color:'#B5302B'}}>{brl(novoTotal-novoRecebido)}</span>
                    </div>
                  )}
                </div>
              )}
              <button onClick={add} disabled={!canSubmit}
                className="w-full text-white rounded-xl py-3 font-medium flex items-center justify-center gap-2 disabled:opacity-40"
                style={{background:'#C65D3C'}}>
                <Check size={17}/> Adicionar pedido
              </button>
            </div>
          </div>
        )}

        {/* Filtros */}
        {lista.length>0 && !searchMode && (
          <div className="flex gap-2 overflow-x-auto pb-1">
            {[['todos','Todos',cnt.todos],['pendente','Pendentes',cnt.pendente],['meio','50% pagos',cnt.meio],['integral','Pagos',cnt.integral]]
              .map(([id,l,c])=>(
                <button key={id} onClick={()=>setFi(id)}
                  className={`px-3.5 py-1.5 rounded-full text-sm font-medium whitespace-nowrap border transition-all ${filtro===id?'bg-stone-900 text-white border-transparent':'bg-white text-stone-600 border-stone-200'}`}>
                  {l} <span className={filtro===id?'opacity-50':'text-stone-400'}>({c})</span>
                </button>
              ))}
          </div>
        )}

        {/* Botão imprimir etiquetas */}
        {lista.length>0 && !searchMode && (
          <button onClick={()=>imprimirEtiquetas(lista, kits)}
            className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-medium border border-stone-200 bg-white text-stone-600 hover:bg-stone-50 transition-colors">
            🖨️ Imprimir etiquetas do dia ({lista.length})
          </button>
        )}
        {lista.length>0 && searchMode && searchQuery && (
          <button onClick={()=>imprimirEtiquetas(lista, kits)}
            className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-medium border border-stone-200 bg-white text-stone-600 hover:bg-stone-50 transition-colors">
            🖨️ Imprimir etiquetas da busca ({lista.length})
          </button>
        )}

        {/* Lista */}
        {lista.length===0 ? (
          <div className="text-center py-14">
            <div className="inline-flex items-center justify-center w-14 h-14 rounded-full bg-stone-100 mb-3">
              {searchMode ? <Search size={22} className="text-stone-400"/> : <Package size={22} className="text-stone-400"/>}
            </div>
            <p className="text-stone-400 text-sm">
              {searchMode && searchQuery ? `Nenhum pedido para "${searchQuery}"` :
               searchMode ? 'Digite o nome do cliente para buscar' :
               `Nenhum pedido em ${new Date(viewDate+'T12:00:00').toLocaleDateString('pt-BR')}`}
            </p>
          </div>
        ) : (
          <div className="space-y-2">
            {searchMode && searchQuery && <p className="text-xs text-stone-500 px-1">{lista.length} pedido{lista.length!==1?'s':''} encontrado{lista.length!==1?'s':''}</p>}
            {lista.map(p=>{
              const itens=getItens(p);
              const tot=pedTotal(p,precos);
              const rec=recebido(p,precos);
              const deve=tot-rec;
              const pk=p.pagamento||'integral';
              const temVP=p.valorPago!=null;
              return (
                <div key={p.id} className="bg-white rounded-2xl border border-stone-200/80 p-4">
                  <div className="flex items-start gap-3">
                    <div className="w-10 h-10 rounded-xl flex-shrink-0 flex items-center justify-center fd text-sm font-semibold text-white" style={{background:'#2A2420'}}>#{p.id}</div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <div className="font-medium text-stone-900">{p.cliente}</div>
                        {searchMode&&<div className="text-[10px] text-stone-400">{new Date(p.data).toLocaleDateString('pt-BR')}</div>}
                      </div>
                      <div className="mt-1 space-y-0.5">
                        {itens.map((it,i)=>{
                          const kit=kits.find(k=>k.id===it.kitId);
                          const pr=precos[it.kitId]||0;
                          return <div key={i} className="flex items-center justify-between text-xs text-stone-500"><span className="truncate mr-2">{kit?.nome||<em>kit removido</em>}</span><span className="flex-shrink-0 tabular-nums">{it.qtd}un × {brl(pr)}</span></div>;
                        })}
                      </div>
                    </div>
                    <div className="text-right flex-shrink-0">
                      <div className="fd text-base text-stone-900">{brl(tot)}</div>
                      {itens.length>1&&<div className="text-[10px] text-stone-400">{itens.length} itens</div>}
                    </div>
                    <button onClick={()=>{setEd(p);setFm(false);setSM(false);}} className="w-8 h-8 rounded-lg flex-shrink-0 flex items-center justify-center text-stone-300 hover:text-stone-700 hover:bg-stone-100 transition-colors" title="Editar pedido"><Pencil size={14}/></button>
                  <button onClick={()=>del(p.id)} className="w-8 h-8 rounded-lg flex-shrink-0 flex items-center justify-center text-stone-300 hover:text-red-500 hover:bg-red-50 transition-colors" title="Remover pedido"><Trash2 size={15}/></button>
                  </div>

                  {/* Pagamento */}
                  <div className="mt-3 pt-3 border-t border-stone-100">
                    <div className="flex items-center justify-between gap-2 mb-2">
                      <span className="text-[11px] text-stone-500 uppercase tracking-wider">Pagamento</span>
                      <div className="flex gap-1">
                        {PAG_KEYS.map(k=>{ const s=PAG_CFG[k]; const sel=pk===k&&!temVP; return (
                          <button key={k} onClick={()=>setPg(p.id,k)}
                            className={`px-3 py-1 rounded-full text-xs font-medium transition-all ${sel?'shadow-sm':'bg-stone-100 text-stone-500 hover:bg-stone-200'}`}
                            style={sel?{background:s.bg,color:s.color}:{}}>{s.label}</button>
                        ); })}
                      </div>
                    </div>
                    {/* Valor pago customizado */}
                    <div className="flex items-center gap-2">
                      <span className="text-[11px] text-stone-400 flex-shrink-0">Valor recebido:</span>
                      <div className="relative flex-1">
                        <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 text-xs">R$</span>
                        <input type="number" inputMode="decimal" step="0.01"
                          value={p.valorPago!=null ? p.valorPago : ''}
                          onChange={e=>setVP(p.id, e.target.value===''?null:e.target.value)}
                          placeholder={brl(tot*(PAG_CFG[pk]?.factor??1)).replace('R$','').trim()}
                          className="w-full bg-stone-50 border border-stone-200 rounded-lg pl-7 pr-3 py-1.5 text-xs text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums"/>
                      </div>
                      <div className="text-right flex-shrink-0 min-w-[80px]">
                        <div className="text-[10px] text-stone-400">A receber</div>
                        <div className="text-sm font-semibold tabular-nums" style={{color:deve>0?'#B5302B':'#3F6E3A'}}>{brl(deve)}</div>
                      </div>
                    </div>
                    {temVP&&<p className="text-[10px] text-stone-400 mt-1">Valor personalizado · <button onClick={()=>setVP(p.id,null)} className="underline hover:text-stone-600">Limpar</button></p>}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
      )} {/* fim verCob ternary */}

      {/* Modal de edição de pedido */}
      {editando && (
        <ModalEditarPedido
          pedido={editando}
          kits={kits}
          precos={precos}
          onSave={(dados)=>updatePedido(editando.id, dados)}
          onClose={()=>setEd(null)}
        />
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// MODAL EDITAR PEDIDO
// ─────────────────────────────────────────────
function ModalEditarPedido({pedido, kits, precos, onSave, onClose}) {
  const [cliente, setC]  = useState(pedido.cliente || '');
  const [itens,   setI]  = useState(getItens(pedido).length ? getItens(pedido) : [{kitId:'',qtd:1}]);
  const [pag,     setP]  = useState(pedido.pagamento || 'integral');
  const [vp,      setV]  = useState(pedido.valorPago != null ? String(pedido.valorPago) : '');
  const [err,     setEr] = useState('');
  const [openKit, setOK] = useState(null);

  const addItem    = ()         => setI(p=>[...p,{kitId:'',qtd:1}]);
  const removeItem = idx        => setI(p=>p.filter((_,i)=>i!==idx));
  const updateItem = (idx,f,v)  => setI(p=>p.map((it,i)=>i===idx?{...it,[f]:v}:it));

  const total    = itens.reduce((s,it)=>s+(precos[it.kitId]||0)*(it.qtd||0),0);
  const recbVal  = vp!==''&&+vp>=0 ? Math.min(+vp,total) : total*(PAG_CFG[pag]?.factor??1);
  const canSave  = cliente.trim() && itens.some(it=>it.kitId&&+it.qtd>0);

  const save = () => {
    if(!canSave){setEr('Informe o cliente e ao menos 1 item');return;}
    const validos = itens.filter(it=>it.kitId&&+it.qtd>0);
    onSave({
      cliente:   cliente.trim(),
      itens:     validos,
      pagamento: pag,
      valorPago: vp!==''&&+vp>=0 ? +vp : null,
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 bg-stone-900/50 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-lg shadow-2xl max-h-[92vh] flex flex-col ai">

        {/* Cabeçalho */}
        <div className="flex items-center justify-between px-5 py-4 border-b border-stone-100 flex-shrink-0">
          <div>
            <h3 className="fd text-xl text-stone-900">Editar pedido #{pedido.id}</h3>
            <p className="text-xs text-stone-400 mt-0.5">
              Criado em {new Date(pedido.data).toLocaleDateString('pt-BR')}
            </p>
          </div>
          <button onClick={onClose} className="w-9 h-9 rounded-xl flex items-center justify-center text-stone-400 hover:bg-stone-100">
            <X size={19}/>
          </button>
        </div>

        {/* Formulário com scroll */}
        <div className="overflow-y-auto flex-1 px-5 py-4 space-y-4">

          {/* Cliente */}
          <Fld l="Cliente">
            <input type="text" value={cliente} onChange={e=>setC(e.target.value)}
              placeholder="Nome do cliente"
              className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/>
          </Fld>

          {/* Itens */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="text-xs uppercase tracking-wider text-stone-500">Itens do pedido</label>
              <button onClick={addItem}
                className="flex items-center gap-1 text-xs font-medium px-3 py-1.5 rounded-full"
                style={{background:'#F5EBE5',color:'#C65D3C'}}>
                <Plus size={12}/> Adicionar item
              </button>
            </div>
            <div className="space-y-2">
              {itens.map((it,idx)=>{
                const isOpen = openKit===idx;
                const selKit = kits.find(k=>k.id===it.kitId);
                return (
                  <div key={idx} className="space-y-1.5">
                    <div className="relative">
                      <button onClick={()=>setOK(isOpen?null:idx)}
                        className={`w-full flex items-center justify-between px-4 py-3.5 rounded-xl border-2 transition-all ${selKit?'border-transparent text-white':'border-stone-200 bg-stone-50 text-stone-400'}`}
                        style={selKit?{background:'#C65D3C'}:{}}>
                        <span className="text-sm font-medium truncate pr-2">
                          {selKit ? selKit.nome : 'Toque para escolher o kit…'}
                        </span>
                        <div className="flex items-center gap-2 flex-shrink-0">
                          {selKit&&<span className="text-sm tabular-nums opacity-75">{brl(precos[it.kitId]||0)}</span>}
                          <ChevronDown size={15} className={`transition-transform duration-200 ${isOpen?'rotate-180':''} ${selKit?'opacity-60':'text-stone-400'}`}/>
                        </div>
                      </button>
                      {isOpen&&(
                        <div className="absolute top-full left-0 right-0 z-30 mt-1 bg-white rounded-2xl border border-stone-200 shadow-xl overflow-hidden">
                          <div className="max-h-60 overflow-y-auto">
                            {kits.filter(k=>k.ativo!==false).map(k=>{
                              const sel=it.kitId===k.id;
                              return (
                                <button key={k.id}
                                  onClick={()=>{updateItem(idx,'kitId',k.id);setOK(null);}}
                                  className={`w-full flex items-center justify-between px-4 py-3.5 border-b border-stone-100 last:border-0 transition-colors text-left ${sel?'text-white':'hover:bg-stone-50 text-stone-700'}`}
                                  style={sel?{background:'#C65D3C'}:{}}>
                                  <span className="text-sm font-medium">{k.nome}</span>
                                  <span className={`text-sm font-bold tabular-nums flex-shrink-0 ml-3 ${sel?'opacity-75':'text-stone-500'}`}>{brl(precos[k.id]||0)}</span>
                                </button>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                    {it.kitId&&(
                      <div className="flex items-center gap-2 px-1">
                        <div className="flex items-center gap-1.5">
                          <button onClick={()=>updateItem(idx,'qtd',Math.max(1,(it.qtd||1)-1))} className="w-8 h-8 rounded-xl bg-stone-100 flex items-center justify-center text-stone-600 hover:bg-stone-200 font-bold text-lg leading-none">−</button>
                          <input type="number" inputMode="numeric" value={it.qtd}
                            onChange={e=>updateItem(idx,'qtd',+e.target.value||1)}
                            className="w-12 bg-stone-50 border border-stone-200 rounded-xl py-1.5 text-sm text-stone-900 focus:outline-none text-center font-bold"/>
                          <button onClick={()=>updateItem(idx,'qtd',(it.qtd||1)+1)} className="w-8 h-8 rounded-xl bg-stone-100 flex items-center justify-center text-stone-600 hover:bg-stone-200 font-bold text-lg leading-none">+</button>
                        </div>
                        <span className="text-sm font-bold tabular-nums" style={{color:'#C65D3C'}}>{brl((precos[it.kitId]||0)*(it.qtd||1))}</span>
                        {itens.length>1&&<button onClick={()=>removeItem(idx)} className="ml-auto text-stone-300 hover:text-red-500 transition-colors p-1"><X size={14}/></button>}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* Pagamento */}
          <Fld l="Status do pagamento">
            <div className="grid grid-cols-3 gap-2">
              {PAG_KEYS.map(k=>{ const s=PAG_CFG[k]; const sel=pag===k; return (
                <button key={k} onClick={()=>setP(k)}
                  className={`rounded-xl py-2.5 text-sm font-medium border transition-all ${sel?'border-transparent':'border-stone-200 bg-stone-50 text-stone-600'}`}
                  style={sel?{background:s.bg,color:s.color}:{}}>{s.label}</button>
              ); })}
            </div>
          </Fld>

          {/* Valor pago */}
          <Fld l="Valor recebido (opcional)">
            <div className="relative">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 text-sm">R$</span>
              <input type="number" inputMode="decimal" step="0.01" value={vp}
                onChange={e=>setV(e.target.value)}
                placeholder={total>0 ? (total*(PAG_CFG[pag]?.factor??1)).toFixed(2) : '0,00'}
                className="w-full bg-stone-50 border border-stone-200 rounded-xl pl-9 pr-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums"/>
            </div>
            <p className="text-[11px] text-stone-400 mt-1">
              Deixe em branco para usar o percentual do status. Preencha se o valor for diferente.
            </p>
          </Fld>

          {/* Resumo de valores */}
          {total>0&&(
            <div className="bg-stone-50 rounded-xl px-4 py-3 space-y-1.5">
              <div className="flex justify-between items-center">
                <span className="text-sm text-stone-600">Total do pedido</span>
                <span className="fd text-xl text-stone-900">{brl(total)}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-stone-500">Recebido</span>
                <span className="font-medium tabular-nums" style={{color:'#3F6E3A'}}>{brl(recbVal)}</span>
              </div>
              {recbVal<total&&(
                <div className="flex justify-between text-sm">
                  <span className="text-stone-500">A receber</span>
                  <span className="font-medium tabular-nums" style={{color:'#B5302B'}}>{brl(total-recbVal)}</span>
                </div>
              )}
            </div>
          )}

          {err&&<ErrBox msg={err}/>}
        </div>

        {/* Botões fixos na base */}
        <div className="px-5 py-4 border-t border-stone-100 flex gap-3 flex-shrink-0">
          <button onClick={onClose}
            className="flex-1 rounded-xl py-3 font-medium border border-stone-200 text-stone-600 hover:bg-stone-50">
            Cancelar
          </button>
          <button onClick={save} disabled={!canSave}
            className="flex-1 text-white rounded-xl py-3 font-medium flex items-center justify-center gap-1.5 disabled:opacity-40"
            style={{background:'#C65D3C'}}>
            <Check size={16}/> Salvar alterações
          </button>
        </div>
      </div>
    </div>
  );
}

function TabCustos({kits,custosKits,prem,setPrem,consumo,setCons,precos,setPrecos,pedidos,onAdd,onUpdate,onRemove,onMove,onReorder,onToggleAtivo,onChangePwd}) {
  const [sec,   setSec]  = useState('kits');
  const [chPwd, setChPwd]= useState(false);

  const SECS=[{id:'kits',l:'Custo p/ Kit'},{id:'itens',l:'✦ Gerenciar Itens'},{id:'insumos',l:'Insumos'},{id:'consumo',l:'Consumo Extra'}];

  return (
    <div>
      <div className="flex gap-2 mb-4 overflow-x-auto pb-1">
        {SECS.map(s=>(
          <button key={s.id} onClick={()=>setSec(s.id)}
            className={`px-4 py-2 rounded-full text-sm font-medium whitespace-nowrap transition-all ${sec===s.id?'bg-stone-900 text-white':'bg-white text-stone-600 border border-stone-200 hover:border-stone-300'}`}>
            {s.l}
          </button>
        ))}
      </div>

      {sec==='kits'&&(
        <div className="space-y-3">
          {custosKits.map(k=>(
            <div key={k.id} className="bg-white rounded-2xl p-5 border border-stone-200/80">
              <div className="flex items-start justify-between mb-3">
                <div>
                  <div className="fd text-lg text-stone-900 leading-tight">{k.nome}</div>
                  <div className="text-xs text-stone-500 mt-0.5">{+k.metragem.toFixed(2)}m linear · <span className="font-medium" style={{color:(k.tecido||'helanca')==='microfibra'?'#1E5A8A':'#A04A2E'}}>{(k.tecido||'helanca')==='microfibra'?'Microfibra':'Helanca'}</span> · extra {brl(k.suporte||0)}</div>
                </div>
                <div className="text-right">
                  <div className="text-[10px] uppercase tracking-wider text-stone-500">Margem</div>
                  <div className="fd text-xl font-semibold" style={{color:k.mgPct>=.5?'#3F6E3A':k.mgPct>=.3?'#C65D3C':'#B5302B'}}>{pct(k.mgPct)}</div>
                </div>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <KCard l="Custo" v={brl(k.det.tot)}/>
                <KCard l="Venda" v={brl(k.pr)}/>
                <KCard l="Lucro" v={brl(k.mg)} acc/>
              </div>
              <details className="mt-3">
                <summary className="text-xs text-stone-500 cursor-pointer select-none">Ver detalhe</summary>
                <div className="mt-2 space-y-1 text-sm">
                  {[['Tecido',k.det.ct],['Papel',k.det.cp],['Calandragem',k.det.cc],['Tinta',k.det.ci],
                    ['Elástico',k.det.ce],['Linha',k.det.cl],...(k.det.cs>0?[['Suporte',k.det.cs]]:[])
                  ].map(([l,v])=>(
                    <div key={l} className="flex justify-between py-0.5 text-stone-600"><span>{l}</span><span className="tabular-nums">{brl(v)}</span></div>
                  ))}
                </div>
              </details>
            </div>
          ))}
        </div>
      )}

      {sec==='itens'&&<GerenciarItens kits={kits} precos={precos} pedidos={pedidos} onAdd={onAdd} onUpdate={onUpdate} onRemove={onRemove} onMove={onMove} onReorder={onReorder} onToggleAtivo={onToggleAtivo}/>}

      {sec==='insumos'&&(
        <div className="space-y-3">
          <InCard t="Helanca Original"     d="Custo total / metros lineares"
            v1={{l:'Valor (R$)',v:prem.tt,f:v=>setPrem({...prem,tt:v})}}
            v2={{l:'Rendimento (m)',v:prem.tr,f:v=>setPrem({...prem,tr:v})}}
            un={`${brl(prem.tt/prem.tr)} / m`}/>
          <InCard t="Helanca Tubular 1,20m" d="Tubular · 0,50m rende 4 painéis"
            v1={{l:'Valor (R$)',v:prem.htt,f:v=>setPrem({...prem,htt:v})}}
            v2={{l:'Rendimento (m)',v:prem.htr,f:v=>setPrem({...prem,htr:v})}}
            un={`${brl(prem.htt/prem.htr)} / m`}/>
          <InCard t="Helanca Ramada 1,80m"  d="Ramada · metragem linear"
            v1={{l:'Valor (R$)',v:prem.hrt,f:v=>setPrem({...prem,hrt:v})}}
            v2={{l:'Rendimento (m)',v:prem.hrr,f:v=>setPrem({...prem,hrr:v})}}
            un={`${brl(prem.hrt/prem.hrr)} / m`}/>
          <InCard t="Papel Sublimático" d="Rolo / metros"
            v1={{l:'Valor (R$)',v:prem.pt,f:v=>setPrem({...prem,pt:v})}}
            v2={{l:'Rendimento (m)',v:prem.pr,f:v=>setPrem({...prem,pr:v})}}
            un={`${brl(prem.pt/prem.pr)} / m`}/>
          <InCard t="Tinta CMYK"       d="4 litros / metros estimados"
            v1={{l:'Valor (R$)',v:prem.it,f:v=>setPrem({...prem,it:v})}}
            v2={{l:'Rendimento (m)',v:prem.ir,f:v=>setPrem({...prem,ir:v})}}
            un={`${brl(prem.it/prem.ir)} / m`}
            warn="Rendimento estimado — ajuste conforme impressões reais"/>
          <SimpleInCard t="Calandragem" d="Por metro linear" v={prem.cal} f={v=>setPrem({...prem,cal:v})} un="/ m"/>
          <InCard t="Tecido Microfibra" d="Custo total / metros lineares"
            v1={{l:'Valor (R$)',v:prem.mft,f:v=>setPrem({...prem,mft:v})}}
            v2={{l:'Rendimento (m)',v:prem.mfr,f:v=>setPrem({...prem,mfr:v})}}
            un={`${brl(prem.mft/prem.mfr)} / m`}/>
          <InCard t="Elástico"         d="Rolo / metros"
            v1={{l:'Valor (R$)',v:prem.et,f:v=>setPrem({...prem,et:v})}}
            v2={{l:'Rendimento (m)',v:prem.er,f:v=>setPrem({...prem,er:v})}}
            un={`${brl(prem.et/prem.er)} / m`}/>
          <InCard t="Linha / Fio"      d="Valor / jardas"
            v1={{l:'Valor (R$)',v:prem.lt,f:v=>setPrem({...prem,lt:v})}}
            v2={{l:'Rendimento (jd)',v:prem.lr,f:v=>setPrem({...prem,lr:v})}}
            un={`${brl(prem.lt/prem.lr)} / jd`}/>
        </div>
      )}

      {sec==='consumo'&&(
        <div className="bg-white rounded-2xl p-5 border border-stone-200/80">
          <div className="fd text-lg text-stone-900 mb-1">Consumo por kit</div>
          <p className="text-sm text-stone-500 mb-4">Elástico e linha por unidade produzida</p>
          <div className="space-y-4">
            {kits.map(k=>(
              <div key={k.id}>
                <div className="text-sm font-medium text-stone-700 mb-2">{k.nome}</div>
                <div className="grid grid-cols-2 gap-3">
                  <Fld l="Elástico (m)"><NIn v={consumo[k.id]?.el||0} f={v=>setCons(c=>({...c,[k.id]:{...c[k.id],el:v}}))} /></Fld>
                  <Fld l="Linha (jardas)"><NIn v={consumo[k.id]?.li||0} f={v=>setCons(c=>({...c,[k.id]:{...c[k.id],li:v}}))} /></Fld>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="mt-6 text-center">
        <button onClick={()=>setChPwd(true)} className="text-xs text-stone-400 hover:text-stone-600 hover:underline underline-offset-2">
          Trocar senha de administrador
        </button>
      </div>
      {chPwd&&<ModalTrocarSenha onClose={()=>setChPwd(false)} onSubmit={onChangePwd}/>}
    </div>
  );
}

// ─────────────────────────────────────────────
// GERENCIAR ITENS
// ─────────────────────────────────────────────
function GerenciarItens({kits,precos,pedidos,onAdd,onUpdate,onRemove,onMove,onReorder,onToggleAtivo}) {
  const [editing,  setEditing]  = useState(null);
  const [confDel,  setConfDel]  = useState(null);
  const [dragFrom, setDragFrom] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  // ── Drag handlers (desktop + mobile via pointer events) ──
  const onDragStart = (e, idx) => {
    setDragFrom(idx);
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', String(idx));
  };
  const onDragOver = (e, idx) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if(dragOver !== idx) setDragOver(idx);
  };
  const onDrop = (e, idx) => {
    e.preventDefault();
    if(dragFrom !== null && dragFrom !== idx) onReorder(kits[dragFrom].id, idx);
    setDragFrom(null); setDragOver(null);
  };
  const onDragEnd = () => { setDragFrom(null); setDragOver(null); };

  // ── Touch drag (mobile) ──
  const touchRef = React.useRef({active:false, fromIdx:null, startY:0});
  const itemRefs = React.useRef([]);

  const onTouchStart = (e, idx) => {
    touchRef.current = {active:true, fromIdx:idx, startY:e.touches[0].clientY};
    setDragFrom(idx);
  };
  const onTouchMove = (e) => {
    if(!touchRef.current.active) return;
    e.preventDefault();
    const y = e.touches[0].clientY;
    for(let i=0;i<itemRefs.current.length;i++){
      const el=itemRefs.current[i];
      if(!el) continue;
      const r=el.getBoundingClientRect();
      if(y>=r.top&&y<=r.bottom){ setDragOver(i); break; }
    }
  };
  const onTouchEnd = () => {
    if(!touchRef.current.active) return;
    const {fromIdx}=touchRef.current;
    touchRef.current.active=false;
    if(dragOver!==null&&dragOver!==fromIdx) onReorder(kits[fromIdx].id, dragOver);
    setDragFrom(null); setDragOver(null);
  };

  const np = id => pedidos.filter(p=>getItens(p).some(it=>it.kitId===id)).length;

  return (
    <div className="space-y-3">
      {/* Lista de kits */}
      <div className="bg-white rounded-2xl border border-stone-200/80 overflow-hidden">
        <div className="px-5 py-3.5 border-b border-stone-100 flex items-center justify-between">
          <p className="text-sm font-medium text-stone-700">Arraste ≡ para reordenar</p>
          <span className="text-xs text-stone-400">{kits.length} kits</span>
        </div>
        <div className="divide-y divide-stone-100">
          {kits.map((k, idx) => {
            const ativo   = k.ativo !== false;
            const isDragging = dragFrom === idx;
            const isOver     = dragOver === idx;
            return (
              <div key={k.id}
                ref={el=>itemRefs.current[idx]=el}
                draggable
                onDragStart={e=>onDragStart(e,idx)}
                onDragOver={e=>onDragOver(e,idx)}
                onDrop={e=>onDrop(e,idx)}
                onDragEnd={onDragEnd}
                className={`flex items-center gap-2 px-4 py-3.5 transition-all select-none
                  ${isDragging ? 'opacity-40 bg-stone-50' : ''}
                  ${isOver && !isDragging ? 'border-t-2 border-t-orange-400' : ''}
                `}>

                {/* Handle de drag */}
                <div
                  onTouchStart={e=>onTouchStart(e,idx)}
                  onTouchMove={onTouchMove}
                  onTouchEnd={onTouchEnd}
                  className="text-stone-300 hover:text-stone-500 cursor-grab active:cursor-grabbing flex-shrink-0 touch-none p-1"
                  title="Arrastar para reordenar">
                  <GripVertical size={18}/>
                </div>

                {/* Infos do kit */}
                <div className={`flex-1 min-w-0 ${!ativo?'opacity-40':''}`}>
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-stone-900 truncate">{k.nome}</span>
                    {!ativo && <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-stone-200 text-stone-500 font-medium flex-shrink-0">Fechado</span>}
                  </div>
                  <div className="text-xs text-stone-500 flex flex-wrap gap-x-2 mt-0.5 items-center">
                    <span>{brl(precos[k.id]||0)}</span>
                    <span>·</span><span>{+parseFloat(k.metragem).toFixed(2)}m</span>
                    <span>·</span>
                    <span className="px-1.5 py-0.5 rounded-full text-[10px] font-medium"
                      style={{
                        background:(k.tecido||'helanca')==='microfibra'?'#E0F0FF':(k.tecido||'helanca')==='helanca_tubular'?'#E0F5E0':(k.tecido||'helanca')==='helanca_ramada'?'#FFF3E0':'#F5EBE5',
                        color:(k.tecido||'helanca')==='microfibra'?'#1E5A8A':(k.tecido||'helanca')==='helanca_tubular'?'#1E6A1E':(k.tecido||'helanca')==='helanca_ramada'?'#8A5A1E':'#A04A2E'
                      }}>
                      {(k.tecido||'helanca')==='microfibra'?'Microfibra':(k.tecido||'helanca')==='helanca_tubular'?'Tubular 1,20m':(k.tecido||'helanca')==='helanca_ramada'?'Ramada 1,80m':'Helanca'}
                    </span>
                    {np(k.id)>0&&<><span>·</span><span className="text-stone-400">{np(k.id)} pedido{np(k.id)!==1?'s':''}</span></>}
                  </div>
                </div>

                {/* Ações */}
                <div className="flex items-center gap-1 flex-shrink-0">
                  {/* Toggle ativo/fechado */}
                  <button onClick={()=>onToggleAtivo(k.id)}
                    className={`w-8 h-8 rounded-lg flex items-center justify-center transition-colors ${ativo?'text-stone-300 hover:text-stone-600 hover:bg-stone-100':'text-orange-400 hover:text-orange-600 hover:bg-orange-50'}`}
                    title={ativo?'Fechar kit (ocultar dos pedidos)':'Abrir kit (exibir nos pedidos)'}>
                    {ativo ? <Eye size={15}/> : <EyeOff size={15}/>}
                  </button>
                  {/* Editar */}
                  <button onClick={()=>{setEditing(k);setConfDel(null);}}
                    className="w-8 h-8 rounded-lg flex items-center justify-center text-stone-400 hover:text-stone-700 hover:bg-stone-100 transition-colors">
                    <Pencil size={14}/>
                  </button>
                  {/* Remover */}
                  {confDel===k.id ? (
                    <div className="flex items-center gap-1">
                      <button onClick={()=>{onRemove(k.id);setConfDel(null);}}
                        className="text-[11px] px-2 py-1 rounded-lg bg-red-500 text-white font-medium">Confirmar</button>
                      <button onClick={()=>setConfDel(null)}
                        className="text-[11px] px-2 py-1 rounded-lg bg-stone-100 text-stone-600">Cancelar</button>
                    </div>
                  ) : (
                    <button onClick={()=>{setConfDel(k.id);setEditing(null);}}
                      className="w-8 h-8 rounded-lg flex items-center justify-center text-stone-400 hover:text-red-500 hover:bg-red-50 transition-colors">
                      <Trash2 size={14}/>
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Adicionar kit */}
      {!editing && (
        <button onClick={()=>setEditing({})}
          className="w-full flex items-center justify-center gap-2 py-3.5 rounded-2xl border-2 border-dashed border-stone-200 text-stone-400 hover:border-stone-400 hover:text-stone-600 transition-colors text-sm font-medium">
          <Plus size={16}/> Adicionar novo kit
        </button>
      )}

      {/* Modal de edição */}
      {editing && (
        <ModalKit
          title={editing.id ? 'Editar item' : 'Novo item'}
          ini={{
            nome: editing.nome||'',
            met:  editing.metragem||'',
            sup:  editing.suporte||0,
            pr:   precos[editing.id]||0,
            tecido: editing.tecido||'helanca',
          }}
          onSave={d=>{
            if(editing.id) { onUpdate(editing.id,d); }
            else { onAdd(d); }
            setEditing(null);
          }}
          onClose={()=>setEditing(null)}
        />
      )}
    </div>
  );
}

function ModalKit({title,ini,onSave,onClose}) {
  const [nome,setNm]=useState(ini.nome);
  const [met, setMt]=useState(ini.met);
  const [sup, setSp]=useState(ini.sup);
  const [pr,  setPr]=useState(ini.pr);
  const [tec, setTc]=useState(ini.tecido||'helanca');
  const [err, setEr]=useState('');

  const save=()=>{
    if(!nome.trim()){setEr('Nome obrigatório');return;}
    const m=parseFloat(met);
    if(!m||m<=0){setEr('Metros de tecido deve ser > 0');return;}
    setEr(''); onSave({nome:nome.trim(),metragem:m,suporte:parseFloat(sup)||0,pr:parseFloat(pr)||0,tecido:tec});
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-2xl p-6 w-full max-w-sm ai shadow-2xl overflow-y-auto max-h-[90vh]">
        <div className="flex items-center justify-between mb-4">
          <h3 className="fd text-xl text-stone-900">{title}</h3>
          <button onClick={onClose} className="text-stone-400"><X size={19}/></button>
        </div>
        <div className="space-y-4">
          <Fld l="Nome do item">
            <input type="text" value={nome} onChange={e=>setNm(e.target.value)} placeholder="Ex: Kit 01 — painel + suporte"
              className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/>
          </Fld>
          <Fld l="Metros de tecido (linear)">
            <input type="number" inputMode="decimal" step="0.01" value={met} onChange={e=>setMt(e.target.value)} placeholder="Ex: 1.83"
              className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/>
            <p className="text-[10px] text-stone-400 mt-1 leading-relaxed">
              Painéis 55×55 cm: <b>qtd ÷ 3 × 0,55</b> (ex: 10 painéis = 1,83 m) · Peça única: altura em metros
            </p>
          </Fld>
          <Fld l="Tipo de tecido">
            <div className="grid grid-cols-2 gap-2">
              {[
                ['helanca','Helanca Original'],
                ['helanca_tubular','Tubular 1,20m'],
                ['helanca_ramada','Ramada 1,80m'],
                ['microfibra','Microfibra'],
              ].map(([id,l])=>(
                <button key={id} onClick={()=>setTc(id)}
                  className={`rounded-xl py-2.5 text-sm font-medium border transition-all ${tec===id?'border-transparent text-white':'border-stone-200 bg-stone-50 text-stone-600'}`}
                  style={tec===id?{background:'#C65D3C'}:{}}>
                  {l}
                </button>
              ))}
            </div>
            <p className="text-[11px] text-stone-400 mt-1.5">Tubular: 0,50m → 4 painéis (larg. 1,20m) · Ramada: metragem linear (larg. 1,80m)</p>
          </Fld>
          <Fld l="Custo extra / suporte (R$)">
            <RInput v={sup} f={v=>setSp(v)}/>
            <p className="text-[10px] text-stone-400 mt-1">Itens físicos fixos por unidade (hastes, suportes). Use 0 se não houver.</p>
          </Fld>
          <Fld l="Preço de venda (R$)">
            <RInput v={pr} f={v=>setPr(v)}/>
          </Fld>
          {err&&<ErrBox msg={err}/>}
          <div className="grid grid-cols-2 gap-3 pt-1">
            <button onClick={onClose} className="rounded-xl py-3 font-medium border border-stone-200 text-stone-600 hover:bg-stone-50">Cancelar</button>
            <button onClick={save} className="text-white rounded-xl py-3 font-medium flex items-center justify-center gap-1.5" style={{background:'#C65D3C'}}>
              <Check size={15}/> Salvar
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// TAB RESUMO (admin)
// ─────────────────────────────────────────────
function TabResumo({resumo, custosKits, pedidos, precos}) {

  // ── Filtro de período único — controla TUDO na aba ──
  const [tipo, setTipo] = useState('diario');
  const [dia,  setDia]  = useState(hoje());
  const [mes,  setMes]  = useState(mesHj());
  const [sem,  setSem]  = useState(hoje());

  const pedFiltrados = useMemo(()=>{
    const sorted=[...pedidos].sort((a,b)=>new Date(a.data)-new Date(b.data));
    if(tipo==='todos')   return sorted.filter(byKit);
    if(tipo==='diario')  return sorted.filter(p=>localDate(p.data)===dia);
    if(tipo==='mensal')  return sorted.filter(p=>localMonth(p.data)===mes);
    if(tipo==='semanal'){const{s,e}=semanaRange(sem);return sorted.filter(p=>{const d=new Date(p.data);return d>=s&&d<=e;});}
    return sorted;
  },[pedidos,tipo,dia,mes,sem]);

  const labelPeriodo = useMemo(()=>{
    const fmt=d=>new Date(d+'T12:00:00').toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric'});
    if(tipo==='todos')   return 'Todos os pedidos';
    if(tipo==='diario')  return `Dia ${fmt(dia)}`;
    if(tipo==='mensal'){const[y,m]=mes.split('-');const nm=new Date(+y,+m-1,15).toLocaleString('pt-BR',{month:'long'});return `${nm[0].toUpperCase()+nm.slice(1)} de ${y}`;}
    if(tipo==='semanal'){const{s,e}=semanaRange(sem);return `Semana ${s.toLocaleDateString('pt-BR')} a ${e.toLocaleDateString('pt-BR')}`;}
    return '';
  },[tipo,dia,mes,sem]);

  // ── Métricas financeiras do período ──
  const fin = useMemo(()=>{
    const pk = custosKits.map(k=>{
      const pu=precos[k.id]||0, cu=k.det?.tot||0;
      const pairs=pedFiltrados.flatMap(p=>{
        const t=pedTotal(p,precos);
        return getItens(p).filter(it=>it.kitId===k.id).map(it=>({qtd:it.qtd||0,fac:pFactor(p,t)}));
      });
      const qtd=pairs.reduce((s,x)=>s+x.qtd,0);
      const rP=qtd*pu, rR=pairs.reduce((s,x)=>s+pu*x.qtd*x.fac,0);
      const cT=qtd*cu;
      return {...k,qtd,pu,cu,cT,rP,rR,lP:rP-cT,lR:rR-cT};
    }).filter(k=>k.qtd>0);

    const tRP=pk.reduce((s,k)=>s+k.rP,0);
    const tRR=pk.reduce((s,k)=>s+k.rR,0);
    const tC =pk.reduce((s,k)=>s+k.cT,0);
    const lR=tRR-tC, lP=tRP-tC, pend=tRP-tRR;
    const mR=tRR>0?lR/tRR:0;
    return {tRP,tRR,tC,lR,lP,pend,mR,pk};
  },[pedFiltrados,custosKits,precos]);

  // ── Custos por insumo do período ──
  const custosProd = useMemo(()=>{
    let ct=0,cp=0,cc=0,ci=0,ce=0,cl=0,cs=0;
    const kq={};
    pedFiltrados.forEach(p=>getItens(p).forEach(it=>{kq[it.kitId]=(kq[it.kitId]||0)+(it.qtd||0);}));
    Object.entries(kq).forEach(([id,qtd])=>{
      const ck=custosKits.find(c=>c.id===id);
      if(!ck) return;
      ct+=ck.det.ct*qtd;cp+=ck.det.cp*qtd;cc+=ck.det.cc*qtd;
      ci+=ck.det.ci*qtd;ce+=ck.det.ce*qtd;cl+=ck.det.cl*qtd;cs+=ck.det.cs*qtd;
    });
    return{ct,cp,cc,ci,ce,cl,cs,total:ct+cp+cc+ci+ce+cl+cs};
  },[pedFiltrados,custosKits]);

  return (
    <div className="space-y-4">

      {/* ── Seletor de período (controla tudo) ── */}
      <div className="bg-white rounded-2xl border border-stone-200/80 p-4 space-y-3">
        <div className="flex gap-2 flex-wrap">
          {[['diario','Diário'],['semanal','Semanal'],['mensal','Mensal'],['todos','Todos']].map(([id,l])=>(
            <button key={id} onClick={()=>setTipo(id)}
              className={`px-4 py-2 rounded-full text-sm font-medium border transition-all ${tipo===id?'bg-stone-900 text-white border-transparent':'bg-stone-50 text-stone-600 border-stone-200 hover:border-stone-400'}`}>
              {l}
            </button>
          ))}
        </div>
        {tipo==='diario'&&(
          <input type="date" value={dia} onChange={e=>setDia(e.target.value)}
            className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:border-stone-400"/>
        )}
        {tipo==='mensal'&&(
          <input type="month" value={mes} onChange={e=>setMes(e.target.value)}
            className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:border-stone-400"/>
        )}
        {tipo==='semanal'&&(
          <div>
            <input type="date" value={sem} onChange={e=>setSem(e.target.value)}
              className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:border-stone-400"/>
            {sem&&(()=>{const{s,e}=semanaRange(sem);return <p className="text-xs text-stone-500 mt-1">Semana: <b>{s.toLocaleDateString('pt-BR')} a {e.toLocaleDateString('pt-BR')}</b></p>;})()}
          </div>
        )}
        <div className="flex items-center justify-between text-xs text-stone-500">
          <span className="font-medium">{labelPeriodo}</span>
          <span>{pedFiltrados.length} pedido{pedFiltrados.length!==1?'s':''}</span>
        </div>
      </div>

      {/* ── Lucro Realizado ── */}
      <div>
        <p className="text-[11px] uppercase tracking-wider text-stone-500 mb-2 px-1">Realizado (já recebido)</p>
        <div className="grid grid-cols-2 gap-2">
          <MiniCard l="Receita"  v={brl(fin.tRR)}/>
          <MiniCard l="Custo"    v={brl(fin.tC)} cor="#78716C"/>
          <div className="col-span-2 rounded-2xl p-4 text-white" style={{background:'linear-gradient(135deg,#3F6E3A,#2E5429)'}}>
            <div className="flex items-end justify-between">
              <div>
                <div className="text-[11px] uppercase tracking-wider opacity-70">Lucro realizado</div>
                <div className="fd text-3xl mt-0.5">{brl(fin.lR)}</div>
              </div>
              <div className="text-right">
                <div className="text-[11px] uppercase tracking-wider opacity-70">Margem</div>
                <div className="fd text-3xl">{pct(fin.mR)}</div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* ── Projetado ── */}
      <div>
        <p className="text-[11px] uppercase tracking-wider text-stone-500 mb-2 px-1">Projetado (100% pagos)</p>
        <div className="grid grid-cols-3 gap-2">
          <MiniCard l="Total"       v={brl(fin.tRP)} sm/>
          <MiniCard l="A receber"   v={brl(fin.pend)} cor="#B5302B" sm/>
          <MiniCard l="Lucro proj." v={brl(fin.lP)}  cor="#3F6E3A"  sm/>
        </div>
      </div>

      {/* ── Por kit (período) ── */}
      <div className="bg-white rounded-2xl border border-stone-200/80 overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-100">
          <h3 className="fd text-lg text-stone-900">Por kit</h3>
          <p className="text-xs text-stone-500 mt-0.5">{labelPeriodo}{labelKit}</p>
        </div>
        {fin.pk.length===0
          ? <p className="text-center py-10 text-stone-400 text-sm">Nenhum pedido no período</p>
          : <div className="divide-y divide-stone-100">
              {fin.pk.sort((a,b)=>b.lP-a.lP).map(k=>{
                const mp=k.pu>0?(k.pu-k.cu)/k.pu:0;
                return (
                  <div key={k.id} className="px-5 py-4">
                    <div className="flex items-center justify-between mb-2">
                      <div className="text-sm font-medium text-stone-900 truncate pr-2">{k.nome}</div>
                      <div className="text-xs px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 flex-shrink-0">{k.qtd}un</div>
                    </div>
                    <div className="grid grid-cols-3 gap-2 text-xs mb-2">
                      <div><div className="text-stone-500">Recebido</div><div className="fd text-base" style={{color:'#3F6E3A'}}>{brl(k.rR)}</div></div>
                      <div><div className="text-stone-500">A receber</div><div className="fd text-base" style={{color:k.rP-k.rR>0?'#B5302B':'#78716C'}}>{brl(k.rP-k.rR)}</div></div>
                      <div><div className="text-stone-500">Lucro proj.</div><div className="fd text-base font-semibold text-stone-900">{brl(k.lP)}</div></div>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="flex-1 h-1.5 bg-stone-100 rounded-full overflow-hidden">
                        <div className="h-full rounded-full" style={{width:`${Math.max(0,Math.min(100,mp*100))}%`,background:'#C65D3C'}}/>
                      </div>
                      <span className="text-[11px] text-stone-500">{pct(mp)}</span>
                    </div>
                  </div>
                );
              })}
            </div>
        }
      </div>

      {/* ── Margem unitária (sempre geral — é sobre precificação, não período) ── */}
      <div className="bg-white rounded-2xl border border-stone-200/80 overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-100">
          <h3 className="fd text-lg text-stone-900">Margem unitária de referência</h3>
          <p className="text-xs text-stone-500 mt-0.5">Custo fixo por unidade vs preço de venda</p>
        </div>
        <div className="divide-y divide-stone-100">
          {custosKits.map(k=>(
            <div key={k.id} className="px-5 py-3 flex items-center gap-2">
              <div className="flex-1 text-sm text-stone-700 truncate">{k.nome}</div>
              <span className="text-sm text-stone-500 tabular-nums">{brl(k.det.tot)}</span>
              <span className="text-stone-300">→</span>
              <span className="text-sm font-medium text-stone-900 tabular-nums">{brl(k.pr)}</span>
              <span className="text-xs px-2 py-0.5 rounded-full tabular-nums" style={{background:'#F5EBE5',color:'#A04A2E'}}>{pct(k.mgPct)}</span>
            </div>
          ))}
        </div>
      </div>

      {/* ── Custos de produção por insumo (mesmo período) ── */}
      <div className="bg-white rounded-2xl border border-stone-200/80 overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-100">
          <h3 className="fd text-lg text-stone-900">Custos de produção por insumo</h3>
          <p className="text-xs text-stone-500 mt-0.5">{labelPeriodo} · {pedFiltrados.length} pedido{pedFiltrados.length!==1?'s':''}</p>
        </div>
        <div className="divide-y divide-stone-100">
          {[
            ['Tecido',           custosProd.ct, '#A04A2E'],
            ['Papel sublimático',custosProd.cp, '#1E5A8A'],
            ['Calandragem',      custosProd.cc, '#5A3E8A'],
            ['Tinta',            custosProd.ci, '#1E7A4A'],
            ['Elástico',         custosProd.ce, '#7A5A1E'],
            ['Linha / Fio',      custosProd.cl, '#5A1E1E'],
            ['Suporte / Extras', custosProd.cs, '#3A5A2E'],
          ].filter(([,v])=>v>0).map(([l,v,cor])=>(
            <div key={l} className="px-5 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="w-2 h-2 rounded-full flex-shrink-0" style={{background:cor}}/>
                <span className="text-sm text-stone-700">{l}</span>
              </div>
              <span className="text-sm font-medium text-stone-900 tabular-nums">{brl(v)}</span>
            </div>
          ))}
          {custosProd.total===0&&(
            <div className="px-5 py-8 text-center text-stone-400 text-sm">Nenhum pedido no período selecionado</div>
          )}
          {custosProd.total>0&&(
            <>
              <div className="px-5 py-4 flex items-center justify-between" style={{background:'#FAF7F2'}}>
                <span className="font-semibold text-stone-900">Total de custos</span>
                <span className="fd text-xl font-semibold text-stone-900 tabular-nums">{brl(custosProd.total)}</span>
              </div>
              {fin.tRR>0&&(
                <div className="px-5 py-3 flex items-center justify-between bg-white">
                  <span className="text-sm text-stone-600">Lucro após custos (período)</span>
                  <span className="fd text-base font-semibold tabular-nums" style={{color:fin.tRR-custosProd.total>=0?'#3F6E3A':'#B5302B'}}>
                    {brl(fin.tRR-custosProd.total)}
                  </span>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function TabExportar({pedidos,kits,precos}) {
  const [tipo,  setTipo] = useState('diario');
  const [dataDia,  setDD] = useState(hoje());
  const [dataSem,  setDS] = useState(hoje());
  const [dataMes,  setDM] = useState(mesHj());
  const [cIni,     setCI] = useState(hoje());
  const [cFim,     setCF] = useState(hoje());
  const [kitFiltro, setKF] = useState('todos'); // filtro por kit específico
  const [busy,  setBusy] = useState(false);
  const [ok,    setOk]   = useState(false);
  const [erro,  setErro] = useState('');

  useEffect(()=>{ setOk(false); setErro(''); },[tipo,dataDia,dataSem,dataMes,cIni,cFim]);

  // ── Pedidos filtrados por período, ordenados por data ──
  const filtrados = useMemo(()=>{
    // Aplica filtro de kit se selecionado
    const byKit = p => kitFiltro==='todos' || getItens(p).some(it=>it.kitId===kitFiltro);
    const sorted=[...pedidos].sort((a,b)=>new Date(a.data)-new Date(b.data));
    if(tipo==='todos')   return sorted;
    if(tipo==='diario')  return sorted.filter(p=>localDate(p.data)===dataDia).filter(byKit);
    if(tipo==='semanal'){ const{s,e}=semanaRange(dataSem); return sorted.filter(p=>{const d=new Date(p.data);return d>=s&&d<=e;}).filter(byKit); }
    if(tipo==='mensal')  return sorted.filter(p=>localMonth(p.data)===dataMes).filter(byKit);
    if(tipo==='periodo'){ const s=new Date(cIni+'T00:00:00'),e=new Date(cFim+'T23:59:59'); return sorted.filter(p=>{const d=new Date(p.data);return d>=s&&d<=e;}).filter(byKit); }
    return sorted;
  },[pedidos,tipo,dataDia,dataSem,dataMes,cIni,cFim]);

  // ── Label legível do período selecionado ──
  const labelPeriodo = useMemo(()=>{
    const fmt = d => new Date(d).toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric'});
    if(tipo==='todos')    return 'Todos os pedidos';
    if(tipo==='diario')   return `Dia ${fmt(dataDia+'T12:00:00')}`;
    if(tipo==='semanal'){
      const{s,e}=semanaRange(dataSem);
      return `Semana ${s.toLocaleDateString('pt-BR')} a ${e.toLocaleDateString('pt-BR')}`;
    }
    if(tipo==='mensal'){
      const[y,m]=dataMes.split('-');
      const nm=new Date(+y,+m-1,15).toLocaleString('pt-BR',{month:'long'});
      return `${nm[0].toUpperCase()+nm.slice(1)} de ${y}`;
    }
    if(tipo==='periodo')  return `${fmt(cIni+'T12:00:00')} a ${fmt(cFim+'T12:00:00')}`;
    return '';
  },[tipo,dataDia,dataSem,dataMes,cIni,cFim]);

  const labelKit = useMemo(()=>{
    if(kitFiltro==='todos') return '';
    const k=kits.find(k=>k.id===kitFiltro);
    return k ? ` — ${k.nome}` : '';
  },[kitFiltro,kits]);

  // ── Resumo por kit para pré-visualização (suporta multi-itens) ──
  const preview = useMemo(()=>{
    const ids=[...new Set(filtrados.flatMap(p=>getItens(p).map(it=>it.kitId)).filter(Boolean))];
    return ids.map(id=>({
      id,
      nome: kits.find(k=>k.id===id)?.nome||`ID:${id}`,
      peds: filtrados.filter(p=>getItens(p).some(it=>it.kitId===id)).length,
      qtd:  filtrados.flatMap(p=>getItens(p).filter(it=>it.kitId===id)).reduce((s,it)=>s+(it.qtd||0),0),
      tot:  filtrados.flatMap(p=>getItens(p).filter(it=>it.kitId===id).map(it=>({qtd:it.qtd||0}))).reduce((s,it)=>s+(precos[id]||0)*it.qtd,0),
    }));
  },[filtrados,kits,precos]);

  const gerar = async ()=>{
    if(!filtrados.length) return;
    setBusy(true); setErro(''); setOk(false);
    try {
      exportXlsx(filtrados,kits,precos,labelPeriodo+labelKit);
      setOk(true);
      setTimeout(()=>setOk(false), 4000);
    } catch(e){
      console.error(e);
      setErro('Erro ao gerar o arquivo. Tente novamente.');
    }
    setBusy(false);
  };

  const TIPOS=[['diario','Diário'],['semanal','Semanal'],['mensal','Mensal'],['periodo','Período'],['todos','Todos']];

  return (
    <div className="space-y-4">

      {/* Cabeçalho */}
      <div className="bg-white rounded-2xl p-5 border border-stone-200/80 flex items-center gap-3">
        <div className="w-11 h-11 rounded-xl flex items-center justify-center flex-shrink-0" style={{background:'#F5EBE5',color:'#C65D3C'}}>
          <FileSpreadsheet size={22}/>
        </div>
        <div>
          <h2 className="fd text-xl text-stone-900">Exportar Pedidos</h2>
          <p className="text-xs text-stone-500">Gera planilha Excel separada por kit, ordenada por data</p>
        </div>
      </div>

      {/* Tipo de período */}
      <div className="bg-white rounded-2xl p-5 border border-stone-200/80">
        <p className="text-xs uppercase tracking-wider text-stone-500 mb-3">Período</p>
        <div className="flex gap-2 flex-wrap mb-4">
          {TIPOS.map(([id,l])=>(
            <button key={id} onClick={()=>setTipo(id)}
              className={`px-4 py-2 rounded-full text-sm font-medium border transition-all ${tipo===id?'bg-stone-900 text-white border-transparent':'bg-stone-50 text-stone-600 border-stone-200 hover:border-stone-400'}`}>
              {l}
            </button>
          ))}
        </div>

        {tipo==='diario'&&(
          <Fld l="Dia"><input type="date" value={dataDia} onChange={e=>setDD(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
        )}
        {tipo==='semanal'&&(
          <Fld l="Qualquer dia da semana">
            <input type="date" value={dataSem} onChange={e=>setDS(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/>
            {dataSem&&(()=>{const{s,e}=semanaRange(dataSem);return <p className="text-xs text-stone-500 mt-1.5">Semana: <b className="text-stone-700">{fmtShort(s.toISOString())} a {fmtShort(e.toISOString())}</b></p>;})()}
          </Fld>
        )}
        {tipo==='mensal'&&(
          <Fld l="Mês"><input type="month" value={dataMes} onChange={e=>setDM(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
        )}
        {tipo==='periodo'&&(
          <div className="grid grid-cols-2 gap-3">
            <Fld l="De"><input type="date" value={cIni} onChange={e=>setCI(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
            <Fld l="Até"><input type="date" value={cFim} onChange={e=>setCF(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
          </div>
        )}
        {tipo==='todos'&&(
          <div className="p-3 rounded-xl text-sm text-stone-600" style={{background:'#FFF8EB'}}>
            Todos os <b>{pedidos.length}</b> pedidos cadastrados serão exportados.
          </div>
        )}
      </div>

      {/* ── Filtro por kit (opcional) ── */}
      <div className="bg-white rounded-2xl border border-stone-200/80 p-4">
        <p className="text-xs font-semibold uppercase tracking-wider text-stone-500 mb-3">Filtrar por kit</p>
        <div className="flex flex-wrap gap-2">
          <button onClick={()=>setKF('todos')}
            className={`px-3.5 py-2 rounded-xl text-sm font-medium border-2 transition-all ${kitFiltro==='todos'?'border-transparent text-white':'border-stone-200 bg-stone-50 text-stone-600 hover:border-stone-300'}`}
            style={kitFiltro==='todos'?{background:'#2A2420'}:{}}>
            Todos os kits
          </button>
          {kits.map(k=>(
            <button key={k.id} onClick={()=>setKF(k.id)}
              className={`px-3.5 py-2 rounded-xl text-sm font-medium border-2 transition-all ${kitFiltro===k.id?'border-transparent text-white':'border-stone-200 bg-stone-50 text-stone-600 hover:border-stone-300'}`}
              style={kitFiltro===k.id?{background:'#C65D3C'}:{}}>
              {k.nome}
            </button>
          ))}
        </div>
        {kitFiltro!=='todos'&&(
          <p className="text-xs text-stone-400 mt-2.5">
            Exportando apenas pedidos que contêm <b className="text-stone-700">{kits.find(k=>k.id===kitFiltro)?.nome}</b>
          </p>
        )}
      </div>

      {/* Pré-visualização */}
      <div className="bg-white rounded-2xl border border-stone-200/80 overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-100 flex items-center justify-between">
          <div>
            <h3 className="fd text-lg text-stone-900">Pré-visualização</h3>
            <p className="text-xs text-stone-500 mt-0.5">{labelPeriodo}{labelKit}</p>
          </div>
          <div className="text-right">
            <p className="text-[10px] uppercase tracking-wider text-stone-500">Pedidos</p>
            <p className="fd text-2xl text-stone-900">{filtrados.length}</p>
          </div>
        </div>
        {filtrados.length===0
          ? <div className="py-10 text-center"><Calendar size={28} className="text-stone-300 mx-auto mb-2"/><p className="text-stone-400 text-sm">Nenhum pedido no período</p></div>
          : <>
              <div className="divide-y divide-stone-100">
                {preview.map(k=>(
                  <div key={k.id} className="px-5 py-3 flex items-center gap-3">
                    <div className="w-2 h-2 rounded-full flex-shrink-0" style={{background:'#C65D3C'}}/>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-stone-900 truncate">{k.nome}</div>
                      <div className="text-xs text-stone-500">{k.peds} pedido{k.peds!==1?'s':''} · {k.qtd} un</div>
                    </div>
                    <div className="text-sm font-medium text-stone-700 tabular-nums">{brl(k.tot)}</div>
                  </div>
                ))}
              </div>
              <div className="px-5 py-3 border-t border-stone-100 flex justify-between items-center bg-stone-50/50">
                <span className="text-xs uppercase tracking-wider text-stone-500">Total geral</span>
                <span className="fd text-lg font-semibold text-stone-900">
                  {brl(filtrados.reduce((s,p)=>s+pedTotal(p,precos),0))}
                </span>
              </div>
            </>
        }
      </div>

      {/* Estrutura do arquivo */}
      {filtrados.length>0&&(
        <div className="bg-white rounded-2xl p-5 border border-stone-200/80">
          <p className="text-xs uppercase tracking-wider text-stone-500 mb-3">Abas geradas no Excel</p>
          <div className="space-y-2">
            {[
              {n:'Resumo Geral',       d:'Totais por kit: receita, recebido e a receber'},
              {n:'Todos os Pedidos',   d:'Todos em ordem cronológica'},
              ...preview.map(k=>({n:k.nome.slice(0,28), d:`${k.peds} pedido${k.peds!==1?'s':''} · ${k.qtd} un · por data`})),
            ].map((ab,i)=>(
              <div key={i} className="flex items-start gap-2.5 text-sm">
                <div className="w-5 h-5 rounded flex-shrink-0 flex items-center justify-center text-[10px] font-bold text-white mt-0.5" style={{background:'#2E7D32'}}>{i+1}</div>
                <div>
                  <div className="font-medium text-stone-900">{ab.n}</div>
                  <div className="text-xs text-stone-500">{ab.d}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Botão exportar */}
      <button onClick={gerar} disabled={busy||filtrados.length===0}
        className="w-full rounded-2xl py-4 font-medium text-white flex items-center justify-center gap-2 transition-all disabled:opacity-40"
        style={{background: ok ? '#3F6E3A' : '#C65D3C'}}>
        {busy ? 'Gerando arquivo…' : ok ? <><Check size={18}/>Arquivo baixado!</> : <><FileDown size={18}/>Exportar para Excel</>}
      </button>

      {erro && <ErrBox msg={erro}/>}

      {/* Botão exportar Word */}
      <button onClick={()=>exportarWord(filtrados,kits,precos,labelPeriodo+labelKit)}
        disabled={filtrados.length===0}
        className="w-full rounded-2xl py-4 font-medium text-stone-800 flex items-center justify-center gap-2 border-2 border-stone-200 bg-white hover:bg-stone-50 transition-all disabled:opacity-40">
        📄 Exportar para Word (.doc)
      </button>

      {filtrados.length===0&&tipo!=='todos'&&(
        <p className="text-center text-xs text-stone-400">Nenhum pedido no período selecionado</p>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// FUNÇÃO EXCEL
// ─────────────────────────────────────────────
function exportXlsx(pedidos, kits, precos, label) {
  const wb  = XLSX.utils.book_new();
  const KN  = id => kits.find(k=>k.id===id)?.nome || `ID:${id}`;
  const PL  = {pendente:'Pendente',meio:'50% pago',integral:'Pago'};
  const PF  = pag => ({pendente:0,meio:.5,integral:1}[pag||'integral']);
  const F2  = n => +n.toFixed(2);
  const DT  = iso => new Date(iso).toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});

  const HDR = ['#','Nº Pedido','Data / Hora','Cliente','Kit','Qtd','Vlr Unit (R$)','Total (R$)','Pagamento','Recebido (R$)','A Receber (R$)'];
  const CW  = [{wch:4},{wch:9},{wch:18},{wch:26},{wch:32},{wch:5},{wch:12},{wch:12},{wch:12},{wch:13},{wch:13}];

  // Normalizar pedidos em linhas — filterKitId limita a um kit específico
  const toRows = (pedList, filterKitId = null) => {
    const rows = [];
    let seq = 0;
    pedList.forEach(p => {
      const all   = getItens(p);
      const itens = filterKitId ? all.filter(it => it.kitId === filterKitId) : all;
      if (!itens.length) return;
      const pedTot = getItens(p).reduce((s,it)=>s+(precos[it.kitId]||0)*(it.qtd||0),0);
      const fac   = pFactor(p, pedTot);
      itens.forEach((it, j) => {
        const pr  = precos[it.kitId] || 0;
        const tot = F2(pr * (it.qtd || 0));
        const rec = F2(tot * fac);
        rows.push([
          j === 0 ? ++seq : '',
          p.id,
          j === 0 ? DT(p.data) : '',
          j === 0 ? p.cliente  : '',
          KN(it.kitId),
          it.qtd || 0,
          F2(pr),
          tot,
          j === 0 ? PL[p.pagamento||'integral'] : '',
          rec,
          F2(tot - rec),
        ]);
      });
    });
    return rows;
  };

  // Kit IDs únicos nos pedidos
  const kitIds = [...new Set(
    pedidos.flatMap(p => getItens(p).map(it => it.kitId)).filter(Boolean)
  )];

  // ── Aba 1: Resumo Geral ──
  const rRows = [
    [`SONHO DOS PAINÉIS — Controle de Entrada`],
    [`Período: ${label}`],
    [`Gerado: ${DT(new Date().toISOString())}`],
    [],
    ['Kit / Produto','Pedidos','Itens','Receita (R$)','Recebido (R$)','A Receber (R$)'],
  ];
  let tPed=0, tItens=0, tTot=0, tRec=0;
  kitIds.forEach(id => {
    const pr    = precos[id] || 0;
    const pairs = pedidos.flatMap(p => { const t=pedTotal(p,precos); return getItens(p).filter(it => it.kitId === id).map(it => ({ qtd: it.qtd||0, fac: pFactor(p,t) })); });
    const nPeds  = pedidos.filter(p => getItens(p).some(it => it.kitId === id)).length;
    const qi     = pairs.reduce((s,x) => s + x.qtd, 0);
    const tot    = F2(pairs.reduce((s,x) => s + pr * x.qtd, 0));
    const rec    = F2(pairs.reduce((s,x) => s + pr * x.qtd * x.fac, 0));
    rRows.push([KN(id), nPeds, qi, tot, rec, F2(tot - rec)]);
    tPed += nPeds; tItens += qi; tTot += tot; tRec += rec;
  });
  rRows.push([], ['TOTAL', tPed, tItens, F2(tTot), F2(tRec), F2(tTot - tRec)]);
  const wsR = XLSX.utils.aoa_to_sheet(rRows);
  wsR['!cols'] = [{wch:34},{wch:10},{wch:8},{wch:15},{wch:15},{wch:15}];
  XLSX.utils.book_append_sheet(wb, wsR, 'Resumo Geral');

  // ── Aba 2: Todos os Pedidos ──
  const allDataRows = toRows(pedidos);
  const totQtd = allDataRows.filter(r => typeof r[5] === 'number').reduce((s,r) => s + r[5], 0);
  const totVal  = F2(allDataRows.filter(r => typeof r[7] === 'number').reduce((s,r) => s + r[7], 0));
  const totRec  = F2(allDataRows.filter(r => typeof r[9] === 'number').reduce((s,r) => s + r[9], 0));
  const allRows = [
    [`TODOS OS PEDIDOS — ${label}`], [],
    HDR,
    ...allDataRows,
    [],
    ['','','','', 'TOTAIS:', totQtd,'', totVal,'', totRec, F2(totVal - totRec)],
  ];
  const wsA = XLSX.utils.aoa_to_sheet(allRows);
  wsA['!cols'] = CW;
  XLSX.utils.book_append_sheet(wb, wsA, 'Todos os Pedidos');

  // ── Aba por kit ──
  kitIds.forEach(id => {
    try {
      const nome   = KN(id);
      const kitPeds = pedidos.filter(p => getItens(p).some(it => it.kitId === id));
      if (!kitPeds.length) return;

      const pr     = precos[id] || 0;
      const pairs  = kitPeds.flatMap(p => { const t=pedTotal(p,precos); return getItens(p).filter(it => it.kitId === id).map(it => ({ qtd: it.qtd||0, fac: pFactor(p,t) })); });
      const totK = F2(pairs.reduce((s,x) => s + pr * x.qtd, 0));
      const recK = F2(pairs.reduce((s,x) => s + pr * x.qtd * x.fac, 0));
      const qtdK = pairs.reduce((s,x) => s + x.qtd, 0);

      const kitRows = toRows(kitPeds, id);
      const rows = [
        [nome],
        [`Período: ${label}`],
        [`Pedidos: ${kitPeds.length}   |   Unidades: ${qtdK}`],
        [], HDR,
        ...kitRows,
        [],
        ['— Resumo por pagamento —'],
      ];
      [['integral','Pagos'],['meio','50% pagos'],['pendente','Pendentes']].forEach(([k,l]) => {
        const gs = kitPeds.filter(p => (p.pagamento||'integral') === k);
        if (!gs.length) return;
        const gPairs = gs.flatMap(p =>
          getItens(p).filter(it => it.kitId === id)
            .map(it => ({ qtd: it.qtd||0, fac: pFactor(p, pedTot) }))
        );
        const gt = F2(gPairs.reduce((s,x) => s + pr * x.qtd, 0));
        const gr = F2(gPairs.reduce((s,x) => s + pr * x.qtd * x.fac, 0));
        rows.push([l, `${gs.length} pedido(s)`, gPairs.reduce((s,x) => s+x.qtd, 0), gt, gr, F2(gt - gr)]);
      });
      rows.push([], ['TOTAL DO KIT','', qtdK,'', totK,'', recK,'', F2(totK - recK)]);

      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols'] = CW;
      const aba = nome.slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, aba);
    } catch(e) {
      console.error('Erro na aba do kit', id, e);
    }
  });

  // ── Download ──
  XLSX.writeFile(wb, `SonhoPaineis_${label.replace(/[\s\/]/g,'_')}.xlsx`);
}

// ─────────────────────────────────────────────
// MODAIS AUTH
// ─────────────────────────────────────────────
function ModalAuth({hasHash,onClose,onSubmit}) {
  const [pw,  setPw] = useState('');
  const [pw2, setPw2]= useState('');
  const [show,setSh] = useState(false);
  const [err, setEr] = useState('');
  const [busy,setBs] = useState(false);
  const ref=useRef(null);
  useEffect(()=>{setTimeout(()=>ref.current?.focus(),50);},[]);

  const go=async()=>{
    setEr('');
    if(!pw){setEr('Digite a senha');return;}
    if(!hasHash){if(pw.length<4){setEr('Mínimo 4 caracteres');return;} if(pw!==pw2){setEr('Senhas não coincidem');return;}}
    setBs(true);
    const r=await onSubmit(pw);
    setBs(false);
    if(r.ok){ onClose(); return; }
    setEr(r.msg||'Erro'); setPw(''); setPw2('');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-stone-900/50 backdrop-blur-sm">
      <div className="bg-white rounded-2xl p-6 w-full max-w-sm ai shadow-2xl">
        <div className="flex items-center justify-between mb-1">
          <div className="flex items-center gap-2">
            <div className="w-9 h-9 rounded-full flex items-center justify-center" style={{background:'#F5EBE5',color:'#C65D3C'}}><KeyRound size={16}/></div>
            <h3 className="fd text-xl text-stone-900">{hasHash?'Acesso administrativo':'Criar senha'}</h3>
          </div>
          <button onClick={onClose} className="text-stone-400"><X size={19}/></button>
        </div>
        <p className="text-sm text-stone-500 mb-5">{hasHash?'Senha necessária para ver custos, resumo e exportar.':'Crie uma senha para proteger a área administrativa. Guarde-a bem — sem recuperação.'}</p>
        <div className="space-y-3">
          <Fld l={hasHash?'Senha':'Nova senha'}>
            <div className="relative">
              <input ref={ref} type={show?'text':'password'} value={pw} onChange={e=>setPw(e.target.value)}
                onKeyDown={e=>e.key==='Enter'&&go()}
                className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 pr-11 text-stone-900 focus:outline-none focus:border-stone-400" placeholder="••••••"/>
              <button onClick={()=>setSh(!show)} className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 p-1">
                {show?<EyeOff size={15}/>:<Eye size={15}/>}
              </button>
            </div>
          </Fld>
          {!hasHash&&(
            <Fld l="Confirmar">
              <input type={show?'text':'password'} value={pw2} onChange={e=>setPw2(e.target.value)}
                onKeyDown={e=>e.key==='Enter'&&go()}
                className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400" placeholder="••••••"/>
            </Fld>
          )}
          {err&&<ErrBox msg={err}/>}
          <button onClick={go} disabled={busy}
            className="w-full text-white rounded-xl py-3 font-medium disabled:opacity-50"
            style={{background:'#C65D3C'}}>
            {busy?'Verificando…':hasHash?'Entrar':'Definir senha'}
          </button>
        </div>
      </div>
    </div>
  );
}

function ModalTrocarSenha({onClose,onSubmit}) {
  const [a,setA]=useState(''),[n,setN]=useState(''),[c,setC]=useState(''),
        [err,setEr]=useState(''), [ok,setOk]=useState(false);
  const go=async()=>{
    setEr('');
    if(!a||!n){setEr('Preencha todos os campos');return;}
    if(n.length<4){setEr('Mínimo 4 caracteres');return;}
    if(n!==c){setEr('Senhas não coincidem');return;}
    const r=await onSubmit(a,n);
    if(!r.ok){setEr(r.msg);return;}
    setOk(true); setTimeout(onClose,1200);
  };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-stone-900/50 backdrop-blur-sm">
      <div className="bg-white rounded-2xl p-6 w-full max-w-sm ai shadow-2xl">
        <div className="flex items-center justify-between mb-4">
          <h3 className="fd text-xl text-stone-900">Trocar senha</h3>
          <button onClick={onClose} className="text-stone-400"><X size={19}/></button>
        </div>
        {ok
          ?<div className="text-center py-6"><div className="inline-flex w-12 h-12 rounded-full items-center justify-center mb-3" style={{background:'#E1EFDB',color:'#3F6E3A'}}><Check size={20}/></div><p>Senha alterada!</p></div>
          :<div className="space-y-3">
            <Fld l="Senha atual"><input type="password" value={a} onChange={e=>setA(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
            <Fld l="Nova senha"><input type="password" value={n} onChange={e=>setN(e.target.value)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
            <Fld l="Confirmar"><input type="password" value={c} onChange={e=>setC(e.target.value)} onKeyDown={e=>e.key==='Enter'&&go()} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400"/></Fld>
            {err&&<ErrBox msg={err}/>}
            <button onClick={go} className="w-full text-white rounded-xl py-3 font-medium" style={{background:'#C65D3C'}}>Alterar</button>
          </div>
        }
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// IMPRIMIR ETIQUETAS
// ─────────────────────────────────────────────
function imprimirEtiquetas(lista, kits) {
  if (!lista.length) return;
  const PAG_LABEL = {pendente:'⚠ Pendente', meio:'50% pago', integral:'✓ Pago'};

  const etiquetas = lista.map(p => {
    const itens = getItens(p);
    const dataFmt = new Date(p.data).toLocaleDateString('pt-BR', {day:'2-digit',month:'2-digit',year:'numeric'});
    const itensList = itens.map(it => {
      const kit = kits.find(k => k.id === it.kitId);
      return `<div class="item"><span class="kit">${kit?.nome || '—'}</span><span class="qty">${it.qtd}un</span></div>`;
    }).join('');
    const pag = PAG_LABEL[p.pagamento||'integral'];
    return `
    <div class="etq">
      <div class="top">
        <span class="num">Pedido #${p.id}</span>
        <span class="data">${dataFmt}</span>
      </div>
      <div class="cliente">${p.cliente}</div>
      <div class="itens">${itensList}</div>
      <div class="pag">${pag}</div>
    </div>`;
  }).join('');

  const html = `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>Etiquetas — Sonho dos Painéis</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; background: #fff; padding: 10px; }
  h2 { font-size: 13px; color: #555; margin-bottom: 10px; text-align: center; }
  .grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
  .etq { border: 1.5px solid #333; border-radius: 6px; padding: 8px 10px; page-break-inside: avoid; min-height: 90px; display: flex; flex-direction: column; gap: 3px; }
  .top { display: flex; justify-content: space-between; align-items: center; }
  .num { font-size: 10px; font-weight: bold; color: #333; }
  .data { font-size: 9px; color: #666; }
  .cliente { font-size: 13px; font-weight: bold; color: #111; border-bottom: 1px dashed #ccc; padding-bottom: 3px; margin-bottom: 2px; }
  .itens { flex: 1; }
  .item { display: flex; justify-content: space-between; font-size: 10px; color: #333; padding: 1px 0; }
  .kit { flex: 1; margin-right: 4px; }
  .qty { font-weight: bold; white-space: nowrap; }
  .pag { font-size: 9px; color: #666; text-align: right; border-top: 1px dashed #ccc; padding-top: 2px; margin-top: 2px; }
  @media print {
    body { padding: 5px; }
    @page { margin: 10mm; }
  }
</style>
</head>
<body>
<h2>SONHO DOS PAINÉIS — Etiquetas de Pedidos (${lista.length} pedido${lista.length!==1?'s':''})</h2>
<div class="grid">${etiquetas}</div>
<script>
  window.onload = function() {
    setTimeout(function() { window.print(); }, 300);
  };
</script>
</body></html>`;

  const win = window.open('', '_blank', 'width=900,height=700');
  if (win) {
    win.document.write(html);
    win.document.close();
  } else {
    alert('Permite pop-ups para imprimir as etiquetas.');
  }
}

// ─────────────────────────────────────────────
// EXPORTAR WORD
// ─────────────────────────────────────────────
function exportarWord(pedidos, kits, precos, label) {
  if (!pedidos.length) return;
  const BRL = n => (n||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
  const DT  = iso => new Date(iso).toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric'});

  // Kits presentes nos pedidos (ordem de aparição)
  const kitIds = [...new Set(
    pedidos.flatMap(p => getItens(p).map(it => it.kitId)).filter(Boolean)
  )];

  const pages = kitIds.map((kitId, pageIdx) => {
    const kit     = kits.find(k => k.id === kitId);
    const kitNome = kit?.nome || `ID:${kitId}`;

    // Pedidos que contêm este kit, ordenados por data
    const kitOrders = pedidos
      .filter(p => getItens(p).some(it => it.kitId === kitId))
      .sort((a,b) => new Date(a.data) - new Date(b.data));

    const totalQtd = kitOrders.reduce((s,p) =>
      s + getItens(p).filter(it => it.kitId === kitId).reduce((q,it) => q+(it.qtd||0), 0), 0
    );

    const rows = kitOrders.map((p, rowIdx) => {
      const qtd  = getItens(p).filter(it => it.kitId === kitId).reduce((s,it) => s+(it.qtd||0), 0);
      const tot  = pedTotal(p, precos);
      const rec  = recebido(p, precos);
      const deve = +(tot - rec).toFixed(2);
      const bg   = rowIdx % 2 === 0 ? '#FFFFFF' : '#F9F9F9';

      const pagInfo = deve <= 0
        ? '<span style="color:#2D6A2D;font-weight:bold">&#10003; Pago &#8212; Quitado</span>'
        : `<span style="color:#B5302B;font-weight:bold">&#9888; D&eacute;bito: ${BRL(deve)}</span>`;

      return `<tr bgcolor="${bg}">
        <td style="padding:8px 12px;border:1px solid #ddd;font-size:10pt;color:#555">#${p.id}&nbsp;&middot;&nbsp;${DT(p.data)}</td>
        <td style="padding:8px 12px;border:1px solid #ddd;font-size:12pt;font-weight:bold">${p.cliente}</td>
        <td style="padding:8px 12px;border:1px solid #ddd;font-size:11pt;text-align:center">${qtd}&nbsp;un</td>
        <td style="padding:8px 12px;border:1px solid #ddd;font-size:11pt">${pagInfo}</td>
      </tr>`;
    }).join('');

    const pb = pageIdx > 0
      ? '<br style="mso-special-character:line-break;page-break-before:always">'
      : '';

    return `${pb}
<div style="font-family:Arial,sans-serif;padding:24px">
  <div style="border-bottom:3px solid #C65D3C;padding-bottom:12px;margin-bottom:20px">
    <div style="font-size:9pt;color:#aaa;letter-spacing:1px">${label}</div>
    <div style="font-size:22pt;font-weight:bold;color:#1C1917;margin-top:4px">${kitNome}</div>
    <div style="font-size:10pt;color:#C65D3C;font-weight:bold;margin-top:4px">
      ${kitOrders.length} pedido${kitOrders.length!==1?'s':''} &nbsp;&middot;&nbsp; ${totalQtd} unidade${totalQtd!==1?'s':''}
    </div>
  </div>

  <table style="border-collapse:collapse;width:100%">
    <thead>
      <tr>
        <th style="padding:8px 12px;border:1px solid #ccc;background:#F5EBE5;font-size:9pt;text-align:left">Pedido&nbsp;&middot;&nbsp;Data</th>
        <th style="padding:8px 12px;border:1px solid #ccc;background:#F5EBE5;font-size:9pt;text-align:left">Cliente</th>
        <th style="padding:8px 12px;border:1px solid #ccc;background:#F5EBE5;font-size:9pt;text-align:center;width:70px">Qtd</th>
        <th style="padding:8px 12px;border:1px solid #ccc;background:#F5EBE5;font-size:9pt;text-align:left">Pagamento</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>
</div>`;
  }).join('');

  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
                      xmlns:w="urn:schemas-microsoft-com:office:word">
<head><meta charset="UTF-8">
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
<style>
  @page { margin: 2cm; size: A4 portrait; }
  body  { font-family: Arial, sans-serif; }
</style>
</head>
<body>
<div style="font-family:Arial;font-size:9pt;color:#aaa;border-bottom:1px solid #eee;padding-bottom:6px;margin-bottom:0">
  SONHO DOS PAIN&Eacute;IS &nbsp;&middot;&nbsp; ${label} &nbsp;&middot;&nbsp;
  Gerado em: ${DT(new Date().toISOString())}
</div>
${pages}
</body></html>`;

  const blob = new Blob(['﻿' + html], {type: 'application/msword;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `Pedidos_${label.replace(/[\s/]/g,'_')}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}


// ─────────────────────────────────────────────
// MICRO COMPONENTS
// ─────────────────────────────────────────────
function Fld({l,children}){return <div><label className="block text-[11px] uppercase tracking-wider text-stone-500 mb-1.5">{l}</label>{children}</div>;}
function Sel({value,onChange,children}){return <div className="relative"><select value={value} onChange={onChange} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400 appearance-none">{children}</select><ChevronDown size={16} className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-stone-400"/></div>;}
function NIn({v,f}){return <input type="number" inputMode="decimal" step="0.01" value={v} onChange={e=>f(parseFloat(e.target.value)||0)} className="w-full bg-stone-50 border border-stone-200 rounded-xl px-3 py-2.5 text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums"/>;}
function RInput({v,f}){return <div className="relative"><span className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 text-sm">R$</span><input type="number" inputMode="decimal" step="0.01" value={v} onChange={e=>f(parseFloat(e.target.value)||0)} className="w-full bg-stone-50 border border-stone-200 rounded-xl pl-9 pr-4 py-3 text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums"/></div>;}
function ErrBox({msg}){return <div className="flex items-center gap-2 text-sm p-3 rounded-xl" style={{background:'#FCE8E6',color:'#B5302B'}}><AlertCircle size={14}/>{msg}</div>;}
function MiniCard({l,v,num,cor,sm}){return <div className="bg-white rounded-2xl p-3 border border-stone-200/80"><div className="text-[10px] uppercase tracking-wider text-stone-500">{l}</div><div className={`fd mt-0.5 ${sm?'text-lg':'text-xl'} font-semibold`} style={{color:cor||'#1C1917'}}>{v}</div></div>;}
function KCard({l,v,acc}){return <div className={`rounded-xl p-2.5 ${acc?'':'bg-stone-50'}`} style={acc?{background:'#F5EBE5'}:{}}><div className="text-[10px] uppercase tracking-wider text-stone-500">{l}</div><div className={`fd text-base font-semibold ${acc?'':'text-stone-900'}`} style={acc?{color:'#A04A2E'}:{}}>{v}</div></div>;}
function InCard({t,d,v1,v2,un,warn}){return <div className="bg-white rounded-2xl p-5 border border-stone-200/80"><div className="flex items-start justify-between mb-3"><div><div className="fd text-base text-stone-900">{t}</div><div className="text-xs text-stone-500 mt-0.5">{d}</div></div><div className="text-right"><div className="text-[10px] uppercase tracking-wider text-stone-500">Custo</div><div className="fd text-sm font-semibold text-stone-900 tabular-nums">{un}</div></div></div><div className="grid grid-cols-2 gap-3"><Fld l={v1.l}><NIn v={v1.v} f={v1.f}/></Fld><Fld l={v2.l}><NIn v={v2.v} f={v2.f}/></Fld></div>{warn&&<div className="mt-3 flex gap-2 p-3 rounded-xl text-xs" style={{background:'#FFF8EB',color:'#8B6B1A'}}><AlertCircle size={13} className="flex-shrink-0 mt-0.5"/><span>{warn}</span></div>}</div>;}
function SimpleInCard({t,d,v,f,un}){return <div className="bg-white rounded-2xl p-5 border border-stone-200/80 flex items-center justify-between gap-4"><div className="flex-1 min-w-0"><div className="fd text-base text-stone-900">{t}</div><div className="text-xs text-stone-500 mt-0.5">{d}</div></div><div className="flex items-center gap-2"><div className="relative w-32"><span className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 text-sm">R$</span><input type="number" inputMode="decimal" step="0.01" value={v} onChange={e=>f(parseFloat(e.target.value)||0)} className="w-full bg-stone-50 border border-stone-200 rounded-xl pl-9 pr-3 py-2.5 text-stone-900 focus:outline-none focus:border-stone-400 tabular-nums text-right"/></div>{un&&<span className="text-xs text-stone-500 whitespace-nowrap">{un}</span>}</div></div>;}
