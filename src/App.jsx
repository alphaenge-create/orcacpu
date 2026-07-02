import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import { doc, getDoc, setDoc } from "firebase/firestore";
import {
  Plus, Trash2, Pencil, X, Search, Upload, Download,
  ChevronDown, ChevronRight, Database, Calculator, Copy, Save, Percent, TrendingUp, RefreshCw,
  Tags, AlertTriangle, Check, FolderKanban, HardHat, User, LogIn
} from "lucide-react";

const uid = () => Math.random().toString(36).slice(2, 9);
const norm = (s) =>
  String(s ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
const fmt = (n) =>
  (Number(n) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const num = (v) => (v === "" || v === null || v === undefined ? 0 : Number(v));
const sanitize = (v) => {
  if (v === undefined) return null;
  if (v === null) return null;
  if (Array.isArray(v)) return v.map(sanitize);
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    const out = {};
    for (const k of Object.keys(v)) {
      const s = sanitize(v[k]);
      if (s !== undefined) out[k] = s;
    }
    return out;
  }
  return v;
};
const TIPOS = [
  { v: "MO", label: "Mão de obra" },
  { v: "MAT", label: "Material" },
  { v: "EQUIP", label: "Equipamento" },
  { v: "OUTROS", label: "Outros" },
];

const FONTES_PADRAO = ["SUDECAP", "DER-MG", "SEINFRA", "SINAPI", "Própria"];

const cpuValorUnit = (insumos) =>
  (insumos || []).reduce((s, i) => s + num(i.coeficiente) * num(i.valorUnitario), 0);

const precoKey = (descricao) => norm(descricao);

// O catálogo descobre os insumos existentes nas CPUs da Base Geral + Insumos já lançados nos projetos
function buildCatalog(cpus, projetos, projetoAtivoId, precos) {
  const map = new Map();

  // 1) Varre as CPUs da Base Geral para mapear insumos disponíveis
  (cpus || []).forEach((cpu) => {
    (cpu.insumos || []).forEach((i) => {
      const key = precoKey(i.descricao);
      if (!key) return;
      if (!map.has(key)) {
        map.set(key, {
          key,
          id: key,
          tipo: i.tipo,
          descricao: i.descricao,
          unidade: i.unidade,
          ocorrencias: 0,
          valoresEncontrados: new Set(),
          valorUnitario: "",
        });
      }
    });
  });

  // 2) Varre especificamente as ocorrências reais de custos do projeto selecionado
  const pAtivo = (projetos || []).find((p) => p.id === projetoAtivoId);
  if (pAtivo && pAtivo.etapas) {
    pAtivo.etapas.forEach((e) => {
      (e.itens || []).forEach((it) => {
        (it.insumos || []).forEach((i) => {
          const key = precoKey(i.descricao);
          if (!key) return;
          if (!map.has(key)) {
            map.set(key, {
              key,
              id: key,
              tipo: i.tipo,
              descricao: i.descricao,
              unidade: i.unidade,
              ocorrencias: 0,
              valoresEncontrados: new Set(),
              valorUnitario: "",
            });
          }
          const entry = map.get(key);
          entry.ocorrencias += 1;
          const v = i.valorUnitario;
          if (v !== "" && v !== null && v !== undefined && !Number.isNaN(Number(v))) {
            entry.valoresEncontrados.add(Number(v));
          }
        });
      });
    });
  }

  // 3) Carrega a tabela de referência customizada do Banco de Preços
  (precos || []).forEach((p) => {
    const key = precoKey(p.descricao);
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, {
        key,
        id: p.id || key,
        tipo: p.tipo,
        descricao: p.descricao,
        unidade: p.unidade,
        ocorrencias: 0,
        valoresEncontrados: new Set(),
        valorUnitario: "",
      });
    }
    const entry = map.get(key);
    entry.id = p.id || entry.id;
    entry.tipo = p.tipo || entry.tipo;
    entry.descricao = p.descricao || entry.descricao;
    entry.unidade = p.unidade || entry.unidade;
    entry.valorUnitario = p.valorUnitario;
  });

  return Array.from(map.values())
    .map((e) => ({
      ...e,
      divergente:
        e.valoresEncontrados.size > 1 ||
        (e.valoresEncontrados.size === 1 &&
          e.valorUnitario !== "" &&
          e.valorUnitario !== null &&
          e.valorUnitario !== undefined &&
          !e.valoresEncontrados.has(Number(e.valorUnitario))),
    }))
    .sort((a, b) => a.descricao.localeCompare(b.descricao, "pt-BR"));
}

const applyCatalogToInsumos = (insumos, catalogMap) =>
  (insumos || []).map((i) => {
    const entry = catalogMap.get(precoKey(i.descricao));
    if (entry && entry.valorUnitario !== "" && entry.valorUnitario !== null && entry.valorUnitario !== undefined) {
      return { ...i, valorUnitario: entry.valorUnitario };
    }
    return i;
  });

const calcBdi = (b, custoInicialOverride) => {
  const custoInicial = num(custoInicialOverride !== undefined ? custoInicialOverride : b.custoInicial);
  const ac = num(b.admCentral);
  const ct = num(b.contabilidade);
  const co = num(b.contingenciamento);
  const cf = num(b.custoFinanceiro);
  const lucro = num(b.lucro);
  const das = num(b.dasAnexoIV);
  const art = num(b.art);
  const pv = das + art;
  const numerador = (1 + ac) * (1 + ct) * (1 + co) * (1 + cf) * (1 + lucro);
  const denominador = 1 - pv;
  const FatorBdi = denominador <= 0 ? 1 : numerador / denominador;
  const bdiRate = FatorBdi - 1;
  const valorVenda = custoInicial * FatorBdi;
  return { bdiRate, FatorBdi, valorVenda };
};

const seedCpus = () => [
  {
    id: uid(),
    codigo: "SUDECAP 04.20.0010",
    fonte: "SUDECAP",
    descricao: "Alvenaria de vedação em bloco cerâmico 9x19x19cm, assentado com argamassa",
    unidade: "m²",
    insumos: [
      { id: uid(), tipo: "MO", descricao: "Pedreiro", unidade: "h", coeficiente: 0.7, valorUnitario: "" },
      { id: uid(), tipo: "MO", descricao: "Servente", unidade: "h", coeficiente: 0.45, valorUnitario: "" },
      { id: uid(), tipo: "MAT", descricao: "Bloco cerâmico 9x19x19", unidade: "un", coeficiente: 25, valorUnitario: "" },
      { id: uid(), tipo: "MAT", descricao: "Argamassa de assentamento", unidade: "m³", coeficiente: 0.012, valorUnitario: "" },
    ],
  },
];

export default function App() {
  const [tab, setTab] = useState("projetos");
  const [cpus, setCpus] = useState([]);
  const [projetos, setProjetos] = useState([]);
  const [projetoAtivoId, setProjetoAtivoId] = useState("");
  const [precos, setPrecos] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [status, setStatus] = useState("");
  const saveTimer = useRef(null);
  const fileInputRef = useRef(null);
  // Novos estados para controle de recolhimento/expansão das camadas
  const [etapasExpandidas, setEtapasExpandidas] = useState({});
  const [cpusExpandidas, setCpusExpandidas] = useState({});

  // Inicialização e Leitura
  useEffect(() => {
    (async () => {
      try {
        const ref = doc(db, "orcacpu", "data_v2");
        const snap = await getDoc(ref);
        if (!snap.exists()) throw new Error("sem dados");
        const data = snap.data();
        setCpus(data.cpus || []);
        setProjetos(data.projetos || []);
        setPrecos(data.precos || []);
        setProjetoAtivoId(data.projetoAtivoId || "");
      } catch {
        setCpus(seedCpus());
        const pId = uid();
        const defaultProj = {
          id: pId,
          nome: "Orçamento Padrão Inicial",
          cliente: "Cliente Geral",
          etapas: [{ id: uid(), nome: "Etapa Inicial", itens: [] }],
          bdi: {
            custoInicial: 0,
            admCentral: 0.04,
            contabilidade: 0.01,
            contingenciamento: 0.02,
            custoFinanceiro: 0.03,
            dasAnexoIV: 0.13,
            art: 0,
            lucro: 0.42,
          }
        };
        setProjetos([defaultProj]);
        setProjetoAtivoId(pId);
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  // Salvamento Automático (Corrigido com Sanitize)
  useEffect(() => {
    if (!loaded) return;
    setStatus("Salvando...");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        // Limpa os dados de valores 'undefined' antes de enviar ao Firestore
        const payload = sanitize({ cpus, projetos, precos, projetoAtivoId });
        await setDoc(doc(db, "orcacpu", "data_v2"), payload);
        setStatus("Salvo");
      } catch (e) {
        console.error("Erro real ao salvar no Firestore:", e);
        setStatus("Falha ao salvar: " + (e?.message || e));
      }
      setTimeout(() => setStatus(""), 4000);
    }, 600);
  }, [cpus, projetos, precos, projetoAtivoId, loaded]);
  
  // Projeto Corrente Detectado
  const projetoAtivo = useMemo(() => {
    return projetos.find((p) => p.id === projetoAtivoId) || projetos[0] || null;
  }, [projetos, projetoAtivoId]);

  const etapas = useMemo(() => projetoAtivo?.etapas || [], [projetoAtivo]);
  const bdi = useMemo(() => projetoAtivo?.bdi || {
    custoInicial: 0, admCentral: 0.04, contabilidade: 0.01, contingenciamento: 0.02, custoFinanceiro: 0.03, dasAnexoIV: 0.13, art: 0, lucro: 0.42
  }, [projetoAtivo]);

  const setEtapas = (novasEtapas) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => (p.id === projetoAtivoId ? { ...p, etapas: typeof novasEtapas === "function" ? novasEtapas(p.etapas) : novasEtapas } : p))
    );
  };

  const setBdi = (novoBdi) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => (p.id === projetoAtivoId ? { ...p, bdi: typeof novoBdi === "function" ? novoBdi(p.bdi) : novoBdi } : p))
    );
  };

  const catalog = useMemo(() => buildCatalog(cpus, projetos, projetoAtivoId, precos), [cpus, projetos, projetoAtivoId, precos]);
  const catalogMap = useMemo(() => new Map(catalog.map((c) => [c.key, c])), [catalog]);

  const upsertPreco = (descricao, tipo, unidade, valorUnitario) => {
    const key = precoKey(descricao);
    if (!key) return;
    setPrecos((prev) => {
      const idx = prev.findIndex((p) => precoKey(p.descricao) === key);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = { ...next[idx], descricao, tipo, unidade, valorUnitario };
        return next;
      }
      return [...prev, { id: uid(), descricao, tipo, unidade, valorUnitario }];
    });
  };

  const removePreco = (descricao) => {
    const key = precoKey(descricao);
    setPrecos((prev) => prev.filter((p) => precoKey(p.descricao) !== key));
  };

  // Melhoria Crítica solicitada: Altera APENAS os insumos associados à aba CUSTOS do projeto ativo
  const aplicarPrecoNoOrcamentoAtivo = (descricao, valorUnitario) => {
    const key = precoKey(descricao);
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projetoAtivoId) return p;
        return {
          ...p,
          etapas: p.etapas.map((e) => ({
            ...e,
            itens: e.itens.map((it) => ({
              ...it,
              insumos: it.insumos.map((i) => (precoKey(i.descricao) === key ? { ...i, valorUnitario } : i))
            }))
          }))
        };
      })
    );
  };

  const aplicarTodosPrecosNoOrcamentoAtivo = () => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projetoAtivoId) return p;
        return {
          ...p,
          etapas: p.etapas.map((e) => ({
            ...e,
            itens: e.itens.map((it) => ({
              ...it,
              insumos: applyCatalogToInsumos(it.insumos, catalogMap)
            }))
          }))
        };
      })
    );
  };

// NOVO: Função para varrer e consolidar o quantitativo de materiais
  const processarMateriais = useMemo(() => {
    const resumoMAT = {};
    etapas.forEach((etapa) => {
      (etapa.itens || []).forEach((item) => {
        const qtdItem = num(item.quantidade);
        (item.insumos || []).forEach((insumo) => {
          // Filtra o que for do tipo "MAT" ou o que NÃO for Mão de Obra (MO) ou Equipamento (EQUIP)
          if (insumo.tipo === "MAT" || (insumo.tipo !== "MO" && insumo.tipo !== "EQUIP" && insumo.unidade?.toLowerCase() !== "h")) {
            const nomeMat = (insumo.descricao || "").toUpperCase().trim();
            if (!nomeMat) return;

            // Busca o valor unitário atualizado diretamente do catálogo/banco de preços de referência
            const entry = catalogMap.get(precoKey(insumo.descricao));
            const precoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(insumo.valorUnitario);

            const qtdTotal = num(insumo.coeficiente) * qtdItem;
            const custoTotal = qtdTotal * precoUnit;

            if (!resumoMAT[nomeMat]) {
              resumoMAT[nomeMat] = {
                material: insumo.descricao,
                unidade: insumo.unidade || "un",
                quantidade: 0,
                valorUnitario: precoUnit,
                valorTotal: 0,
              };
            }
            resumoMAT[nomeMat].quantidade += qtdTotal;
            resumoMAT[nomeMat].valorTotal += custoTotal;
          }
        });
      });
    });
    return Object.values(resumoMAT).sort((a, b) => b.valorTotal - a.valorTotal); // Ordena do mais caro para o mais barato
  }, [etapas, catalogMap]);

  const grandTotal = useMemo(() => {
    return etapas.reduce(
      (s, e) => s + e.itens.reduce((s2, it) => s2 + num(it.quantidade) * cpuValorUnit(it.insumos), 0),
      0
    );
  }, [etapas]);

  const bdiCalc = useMemo(() => {
    const calcularFatorBdiQualquer = (t) => {
      const ac = num(t.admCentral);
      const c = num(t.contabilidade);
      const co = num(t.contingenciamento);
      const cf = num(t.custoFinanceiro);
      const l = num(t.lucro);
      const das = num(t.dasAnexoIV);
      const art = num(t.art);

      const pv = das + art;
      const numerador = (1 + ac) * (1 + c) * (1 + co) * (1 + cf) * (1 + l);
      const denominador = 1 - pv;
      return denominador <= 0 ? 1 : numerador / denominador;
    };

    const FatorBdiGeral = calcularFatorBdiQualquer(bdi);
    
    const faturamentoDireto = !!bdi.faturamentoDireto;
    const FatorBdiMateriais = (faturamentoDireto && bdi.materiais) 
      ? calcularFatorBdiQualquer(bdi.materiais) 
      : FatorBdiGeral;

    // Calcular o preço de venda de forma ponderada analisando insumo por insumo
    let totalCustoDireto = 0;
    let totalPrecoVenda = 0;

    (etapas || []).forEach(e => {
      (e.itens || []).forEach(it => {
        const qtdCpu = num(it.quantidade);
        (it.insumos || []).forEach(ins => {
          const tipo = String(ins.tipo || "").toUpperCase().trim();
          const custoInsumoTotal = num(ins.coeficiente) * qtdCpu * num(ins.valorUnitario);
          totalCustoDireto += custoInsumoTotal;

          // Se for Material (MAT ou MATERIAL) usa o BDI de materiais, senão usa o Geral
          if (faturamentoDireto && (tipo === "MAT" || tipo === "MATERIAL" || (!tipo.includes("MO") && !tipo.includes("MÃO") && !tipo.includes("MAO") && !tipo.includes("EQUIP")))) {
            totalPrecoVenda += custoInsumoTotal * FatorBdiMateriais;
          } else {
            totalPrecoVenda += custoInsumoTotal * FatorBdiGeral;
          }
        });
      });
    });

    const totalDiValor = Math.max(0, totalPrecoVenda - totalCustoDireto);
    const totalDiRate = totalCustoDireto > 0 ? (totalDiValor / totalCustoDireto) : 0;

    return {
      bdiRate: FatorBdiGeral - 1,
      bdiRateMateriais: FatorBdiMateriais - 1,
      FatorBdi: FatorBdiGeral,
      FatorBdiMateriais: FatorBdiMateriais,
      faturamentoDireto: faturamentoDireto,
      totalDiValor: totalDiValor,
      totalDiRate: totalDiRate,
      valorVenda: totalPrecoVenda
    };
  }, [bdi, etapas]);

  // Abas disponíveis apenas dentro de um projeto ativo
  const abasProjeto = ["custo", "planilha", "bdi", "precovenda", "maoobra", "materiais", "precos"];
  const tabEhDeProjeto = abasProjeto.includes(tab);

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900">

      <div className="max-w-6xl mx-auto px-4 py-6">
        {/* ── HEADER ── */}
        <header className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-2xl font-semibold tracking-tight">Orçamentador por CPU</h1>
            <p className="text-sm text-stone-500">
              {projetoAtivo
                ? `Orçamento: ${projetoAtivo.nome} · ${projetoAtivo.cliente || "Geral"}`
                : "Crie ou selecione um orçamento para começar"}
            </p>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs text-stone-400">{status}</span>
          </div>
        </header>

        {/* ── NAV GLOBAL ── */}
        <nav className="flex gap-1 mb-1 flex-wrap">
          <TabBtn active={tab === "projetos"} onClick={() => setTab("projetos")} icon={<FolderKanban size={15} />}>
            Orçamentos ({projetos.length})
          </TabBtn>
          <TabBtn active={tab === "cpus"} onClick={() => setTab("cpus")} icon={<Database size={15} />}>
            Base de CPUs ({cpus.length})
          </TabBtn>
        </nav>

        {/* ── NAV DO PROJETO ATIVO (só aparece quando há projeto) ── */}
        {projetoAtivo && (
          <nav className="flex gap-1 mb-6 border-b border-stone-200 flex-wrap pt-1">
            <span className="self-center text-[10px] font-semibold text-stone-400 uppercase pr-2 pl-1">
              {projetoAtivo.nome}:
            </span>
            <TabBtn active={tab === "custo"} onClick={() => setTab("custo")} icon={<Calculator size={15} />}>
  Lançamento CPU
</TabBtn>
<TabBtn active={tab === "planilha"} onClick={() => setTab("planilha")} icon={<FolderKanban size={15} />}>
  Planilha de custo
</TabBtn>
            <TabBtn active={tab === "bdi"} onClick={() => setTab("bdi")} icon={<Percent size={15} />}>
              BDI — {fmt(bdiCalc.bdiRate * 100)}%
            </TabBtn>
            <TabBtn active={tab === "precovenda"} onClick={() => setTab("precovenda")} icon={<TrendingUp size={15} />}>
              Venda — R$ {fmt(bdiCalc.valorVenda)}
            </TabBtn>
            <TabBtn active={tab === "maoobra"} onClick={() => setTab("maoobra")} icon={<HardHat size={15} />}>
              Mão de Obra
            </TabBtn>
            <TabBtn active={tab === "materiais"} onClick={() => setTab("materiais")} icon={<Database size={15} />}>
              Materiais
            </TabBtn>
            <TabBtn active={tab === "precos"} onClick={() => setTab("precos")} icon={<Tags size={15} />}>
              Banco de Preços ({catalog.length})
              {catalog.some((c) => c.divergente) && (
                <span className="ml-1 w-1.5 h-1.5 rounded-full bg-amber-500 inline-block" />
              )}
            </TabBtn>
          </nav>
        )}

        {/* divisor para abas globais (sem projeto ativo) */}
        {!projetoAtivo && <div className="border-b border-stone-200 mb-6" />}

        <datalist id="insumos-catalogo">
          {catalog.map((c) => <option key={c.key} value={c.descricao} />)}
        </datalist>

        {/* ── CONTEÚDO DAS ABAS ── */}
        {tab === "projetos" && (
          <div className="space-y-4">
            <div className="flex justify-between items-center bg-white border border-stone-200 rounded-lg p-4 shadow-xs">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Seus Orçamentos</h2>
                <p className="text-xs text-stone-500">Gerencie, selecione ou crie novas pastas de projetos e fechamentos comerciais.</p>
              </div>
              <button
                onClick={() => {
                  const pId = uid();
                  setProjetos((prev) => [
                    ...prev,
                    {
                      id: pId,
                      nome: `Novo Orçamento — ${prev.length + 1}`,
                      cliente: "Cliente Geral",
                      etapas: [{ id: uid(), nome: "Etapa Inicial", itens: [] }],
                      bdi: {
                        custoInicial: 0,
                        admCentral: 0,
                        contabilidade: 0,
                        contingenciamento: 0,
                        custoFinanceiro: 0,
                        dasAnexoIV: 0,
                        art: 0,
                        lucro: 0,
                        faturamentoDireto: false,
                        materiais: { admCentral: 0, contabilidade: 0, contingenciamento: 0, custoFinanceiro: 0, lucro: 0, dasAnexoIV: 0, art: 0 }
                      }
                    }
                  ]);
                  setProjetoAtivoId(pId);
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-stone-900 text-white rounded-lg text-xs font-medium hover:bg-stone-800"
              >
                <Plus size={14} /> Novo Orçamento
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {projetos.map((p) => {
                const isActive = p.id === projetoAtivoId;
                
                // Calcula o custo direto acumulado deste projeto específico
                const cDiretoTotal = (p.etapas || []).reduce(
                  (s, e) => s + (e.itens || []).reduce((s2, it) => s2 + num(it.quantidade) * cpuValorUnit(it.insumos), 0),
                  0
                );

                // Calcula o preço final ponderado analisando insumo por insumo deste projeto específico
                const valorVendaCalculado = (() => {
                  const calcFator = (t) => {
                    const ac = num(t.admCentral || t.adminCentral);
                    const c = num(t.contabilidade);
                    const co = num(t.contingenciamento);
                    const cf = num(t.custoFinanceiro);
                    const l = num(t.lucro);
                    const das = num(t.dasAnexoIV || 0);
                    const art = num(t.art);
                    const pv = das + art;
                    const numr = (1 + ac) * (1 + c) * (1 + co) * (1 + cf) * (1 + l);
                    const den = 1 - pv;
                    return den <= 0 ? 1 : numr / den;
                  };

                  const fGeral = calcFator(p.bdi || {});
                  const fatDireto = !!p.bdi?.faturamentoDireto;
                  const fMats = (fatDireto && p.bdi?.materiais) ? calcFator(p.bdi.materiais) : fGeral;

                  let totalVenda = 0;
                  (p.etapas || []).forEach(e => {
                    (e.itens || []).forEach(it => {
                      const qCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tIn = String(ins.tipo || "").toUpperCase().trim();
                        const cIn = num(ins.coeficiente) * qCpu * num(ins.valorUnitario);
                        if (fatDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")))) {
                          totalVenda += cIn * fMats;
                        } else {
                          totalVenda += cIn * fGeral;
                        }
                      });
                    });
                  });
                  return totalVenda;
                })();

                return (
                  <div
                    key={p.id}
                    className={`bg-white border rounded-xl p-4 shadow-xs space-y-3 cursor-pointer transition-all ${
                      isActive ? "border-stone-900 ring-1 ring-stone-900 bg-stone-50/20" : "border-stone-200 hover:border-stone-400"
                    }`}
                    onClick={() => { setProjetoAtivoId(p.id); setTab("custo"); }}
                  >
                    <div className="flex justify-between items-start">
                      <div>
                        <h3 className="font-semibold text-stone-800 text-sm flex items-center gap-1.5 uppercase">
                          📁 {p.nome}
                        </h3>
                        <p className="text-xs text-stone-400">Cliente: {p.cliente || "Geral"}</p>
                      </div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          if (projetos.length <= 1) return alert("Não é possível apagar todos os orçamentos.");
                          if (confirm(`Tem certeza que deseja apagar o orçamento "${p.nome}"?`)) {
                            setProjetos((prev) => prev.filter((item) => item.id !== p.id));
                            if (isActive) setProjetoAtivoId(projetos.find((item) => item.id !== p.id)?.id || "");
                          }
                        }}
                        className="text-stone-400 hover:text-red-600 p-1 rounded-md transition-colors"
                        title="Excluir Orçamento"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>

                    <div className="grid grid-cols-2 gap-2 pt-2 border-t border-stone-100 text-xs font-mono">
                      <div>
                        <span className="text-stone-400 block font-sans text-[10px] uppercase">Custo Direto:</span>
                        <span className="text-stone-600 font-medium">R$ {fmt(cDiretoTotal)}</span>
                      </div>
                      <div className="text-right">
                        <span className="text-stone-400 block font-sans text-[10px] uppercase">Preço Estimado (Venda):</span>
                        <span className="text-stone-900 font-bold text-sm">R$ {fmt(valorVendaCalculado)}</span>
                      </div>
                    </div>

                    <div className="flex justify-between items-center pt-1 text-[11px]">
                      <span className="text-stone-400">{(p.etapas || []).length} etapa(s) cadastrada(s)</span>
                      {isActive && (
                        <span className="text-stone-800 font-semibold bg-stone-200 px-2 py-0.5 rounded text-[10px] uppercase tracking-wider">
                          Selecionado ativo
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {tab === "cpus" && (
          <CpuLibrary cpus={cpus} setCpus={setCpus} fileInputRef={fileInputRef} catalogMap={catalogMap} />
        )}

        {/* Abas de projeto — só renderizam se houver projeto ativo */}
        {tabEhDeProjeto && !projetoAtivo && (
          <div className="text-center py-20 text-stone-400">
            <FolderKanban size={40} className="mx-auto mb-3 opacity-30" />
            <p className="text-sm font-medium">Nenhum orçamento selecionado.</p>
            <button onClick={() => setTab("projetos")} className="mt-3 text-xs underline">
              Criar ou selecionar um orçamento
            </button>
          </div>
        )}
        {tab === "custo" && projetoAtivo && (
          <Orcamento 
            etapas={etapas} 
            setEtapas={setEtapas} 
            cpus={cpus} 
            grandTotal={grandTotal} 
            catalogMap={catalogMap} 
          />
        )}
{tab === "bdi" && projetoAtivo && (
          <BdiTab bdi={bdi} setBdi={setBdi} bdiCalc={bdiCalc} grandTotal={grandTotal} />
        )}

        {tab === "planilha" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Planilha de Exploração de Custos Diretos</h2>
                <p className="text-xs text-stone-500">Visualização hierárquica completa: Etapa ➔ CPU ➔ Todos os Insumos Associados.</p>
              </div>
              <div className="flex gap-2">
                <button 
                  onClick={() => {
                  const objEtapas = {};
                  const objCpus = {};
                  etapas.forEach((etapa) => {
                    // Usa o ID real da etapa ou o índice fallback
                    const eId = etapa.id || `etapa-${etapas.indexOf(etapa)}`;
                    objEtapas[eId] = true;
                    (etapa.itens || []).forEach((item) => {
                      // Usa o ID real do item
                      objCpus[item.id] = true;
                    });
                  });
                  setEtapasExpandidas(objEtapas);
                  setCpusExpandidas(objCpus);
                }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600 flex items-center gap-1"
                >
                  Expandir Tudo
                </button>
                <button 
                  onClick={() => { setEtapasExpandidas({}); setCpusExpandidas({}); }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Recolher Tudo
                </button>

                {/* EXCEL DA PLANILHA DE CUSTO */}
                <button 
                  onClick={() => {
                    const data = [];
                    data.push(["ESTRUTURA", "DESCRIÇÃO", "UND", "QTD PROP.", "CUSTO UNIT", "CUSTO TOTAL"]);
                    etapas.forEach((etapa, idxE) => {
                      data.push([`${idxE + 1}`, etapa.nome, "", "", "", (etapa.itens || []).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos)), 0)]);
                      (etapa.itens || []).forEach((item, idxI) => {
                        const numCpu = `${idxE + 1}.${idxI + 1}`;
                        data.push([numCpu, item.servico || item.descricao, item.unidade, num(item.quantidade), cpuValorUnit(item.insumos), num(item.quantidade) * cpuValorUnit(item.insumos)]);
                        (item.insumos || []).forEach((ins, idxIn) => {
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const pUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          data.push([`${numCpu}.${idxIn + 1}`, `[${ins.tipo}] ${ins.descricao}`, ins.unidade || "un", num(ins.coeficiente) * num(item.quantidade), pUnit, (num(ins.coeficiente) * num(item.quantidade)) * pUnit]);
                        });
                      });
                    });
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Planilha de Custo");
                    XLSX.writeFile(wb, `${projetoAtivo.nome || "Orcamento"}_Planilha_Custo.xlsx`);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-emerald-200 text-emerald-700 bg-emerald-50/50 rounded hover:bg-emerald-50 flex items-center gap-1"
                >
                  <Download size={12} /> Excel (.xlsx)
                </button>

                {/* PDF LIMPO DA PLANILHA DE CUSTO */}
                <button 
                  onClick={() => {
                    const tituloOriginal = document.title;
                    document.title = `${projetoAtivo.nome || "Orcamento"}_Planilha_Custo`;
                    const estiloPrint = document.createElement("style");
                    estiloPrint.innerHTML = `
                      @media print {
                        body * { visibility: hidden; }
                        #area-planilha-custo, #area-planilha-custo * { visibility: visible; }
                        #area-planilha-custo { position: absolute; left: 0; top: 0; width: 100%; background: white !important; }
                      }
                    `;
                    document.head.appendChild(estiloPrint);
                    window.print();
                    document.head.removeChild(estiloPrint);
                    document.title = tituloOriginal;
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1"
                >
                  <Download size={12} /> PDF (.pdf)
                </button>
              </div>
            </div>

            <div id="area-planilha-custo" className="border border-stone-200 rounded-lg overflow-hidden bg-white">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Estrutura (Etapa / CPU / Insumo)</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Prop.</span>
                <span className="col-span-1.5 text-right">Custo Unit</span>
                <span className="col-span-2 text-right">Custo Total</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[600px] overflow-y-auto">
                {etapas.length === 0 ? (
                  <div className="p-8 text-center text-stone-400 italic text-xs">
                    Nenhuma etapa cadastrada neste orçamento.
                  </div>
                ) : (
                  etapas.map((etapa, idxEtapa) => {
                    const numEtapa = idxEtapa + 1;
                    const etapaId = etapa.id || `etapa-${idxEtapa}`;
                    const isEtapaAberta = !!etapasExpandidas[etapaId];

                    return (
                      <div key={etapaId} className="bg-stone-50/30">
                        <div 
                          className="grid grid-cols-12 gap-2 px-4 py-2 bg-stone-200/60 text-stone-800 text-xs font-bold items-center uppercase tracking-wide cursor-pointer hover:bg-stone-200 select-none"
                          onClick={() => setEtapasExpandidas(p => ({ ...p, [etapaId]: !isEtapaAberta }))}
                        >
                          <span className="col-span-10 flex items-center gap-1.5">
                            {isEtapaAberta ? <ChevronDown size={14} className="text-stone-500 shrink-0" /> : <ChevronRight size={14} className="text-stone-500 shrink-0" />}
                            <span className="truncate">📁 {numEtapa}. {etapa.nome}</span>
                          </span>
                          <span className="col-span-2 text-right font-mono">
                            R$ {fmt((etapa.itens || []).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos)), 0))}
                          </span>
                        </div>

                        {isEtapaAberta && (etapa.itens || []).map((item, idxItem) => {
                          const numCpu = `${numEtapa}.${idxItem + 1}`;
                          const itemId = item.id || `item-${numCpu}`;
                          const isCpuAberta = !!cpusExpandidas[itemId];
                          const qtdItem = num(item.quantidade);
                          const custoUnitCpu = cpuValorUnit(item.insumos);

                          return (
                            <div key={itemId} className="border-b border-stone-100">
                              <div 
                                onClick={() => setCpusExpandidas(p => ({ ...p, [itemId]: !isCpuAberta }))}
                                className="grid grid-cols-12 gap-2 px-4 py-2 bg-white text-xs items-center font-semibold text-stone-700 pl-8 cursor-pointer hover:bg-stone-50 select-none"
                              >
                                <span className="col-span-6 truncate text-stone-900 flex items-center gap-1">
                                  {isCpuAberta ? <ChevronDown size={13} className="text-stone-400 shrink-0" /> : <ChevronRight size={13} className="text-stone-400 shrink-0" />}
                                  🔹 {numCpu}. {item.codigo ? `[${item.codigo}] ` : ""}{item.servico || item.descricao}
                                </span>
                                <span className="col-span-1 text-center font-mono text-stone-400">{item.unidade}</span>
                                <span className="col-span-1.5 text-right font-mono">{qtdItem.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                                <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(custoUnitCpu)}</span>
                                <span className="col-span-2 text-right font-mono text-stone-800">
                                  R$ {fmt(qtdItem * custoUnitCpu)}
                                </span>
                              </div>

                              {isCpuAberta && (item.insumos || []).length > 0 && (
                                <div className="bg-stone-50/50 divide-y divide-stone-100/60 border-t border-b border-stone-100">
                                  {(item.insumos || []).map((insumo, idxInsumo) => {
                                    const numInsumo = `${numCpu}.${idxInsumo + 1}`;
                                    const entry = catalogMap.get(precoKey(insumo.descricao));
                                    const precoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(insumo.valorUnitario);
                                    const qtdCalculada = num(insumo.coeficiente) * qtdItem;
                                    const custoTotalInsumo = qtdCalculada * precoUnit;

                                    return (
                                      <div key={insumo.id || idxInsumo} className="grid grid-cols-12 gap-2 px-4 py-1.5 text-[11px] items-center text-stone-600 pl-14 hover:bg-stone-100/40">
                                        <span className="col-span-6 truncate uppercase font-sans text-stone-500">
                                          {numInsumo}. <span className="text-[9px] font-mono font-bold text-stone-400 border border-stone-200 px-1 py-0.5 rounded bg-white mr-1">{insumo.tipo}</span> {insumo.descricao}
                                        </span>
                                        <span className="col-span-1 text-center font-mono text-stone-400 uppercase text-[10px]">{insumo.unidade || "un"}</span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-600">
                                          {qtdCalculada.toLocaleString("pt-BR", { minimumFractionDigits: 3, maximumFractionDigits: 3 })}
                                        </span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(precoUnit)}</span>
                                        <span className="col-span-2 text-right font-mono font-medium text-stone-700">
                                          R$ {fmt(custoTotalInsumo)}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })
                )}

                {/* LINHA DE TOTAL GERAL DA PLANILHA DE CUSTO */}
                <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold uppercase tracking-wider">
                  <span className="col-span-6">CUSTO DIRETO TOTAL</span>
                  <span className="col-span-1"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-2 text-right font-mono text-amber-400">
                    R$ {fmt(grandTotal)}
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}

        {tab === "precovenda" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Planilha de Preço de Venda (Custo + BDI)</h2>
                <p className="text-xs text-stone-500">
                  Visualização hierárquica por Etapa ➔ CPU ➔ Insumos aplicando BDI Geral de {fmt(bdiCalc.bdiRate * 100)}% {bdiCalc.faturamentoDireto && `e BDI de Materiais de ${fmt(bdiCalc.bdiRateMateriais * 100)}%`}. Valor Comercial Fechado: <span className="font-bold text-stone-800 font-mono">R$ {fmt(bdiCalc.valorVenda)}</span>
                </p>
              </div>
              <div className="flex gap-2">
                <button 
                  onClick={() => {
                    const obj = {};
                    etapas.forEach((e, idxE) => { 
                      const etapaId = e.id || `etapa-${idxE}`;
                      obj[etapaId] = true; 
                      (e.itens || []).forEach((it, idxIt) => { 
                        const numCpu = `${idxE + 1}.${idxIt + 1}`;
                        const itemId = it.id || `item-${numCpu}`;
                        obj[itemId] = true; 
                      }); 
                    });
                    setEtapasExpandidas(obj); setCpusExpandidas(obj);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Expandir Tudo
                </button>
                <button 
                  onClick={() => { setEtapasExpandidas({}); setCpusExpandidas({}); }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Recolher Tudo
                </button>

                {/* EXCEL DA PLANILHA DE VENDA (CORRIGIDO PARA LER O BDI DE MATERIAIS) */}
                <button 
                  onClick={() => {
                    const data = [];
                    data.push(["ESTRUTURA", "DESCRIÇÃO", "UND", "QTD PROP.", "PREÇO UNIT VENDA", "TOTAL VENDA"]);
                    etapas.forEach((etapa, idxE) => {
                      let totalEtapaVenda = 0;
                      (etapa.itens || []).forEach(it => {
                        const qCpu = num(it.quantidade);
                        (it.insumos || []).forEach(ins => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const cIn = num(ins.coeficiente) * qCpu * num(ins.valorUnitario);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          totalEtapaVenda += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                        });
                      });

                      data.push([`${idxE + 1}`, etapa.nome, "", "", "", totalEtapaVenda]);
                      
                      (etapa.itens || []).forEach((item, idxI) => {
                        const numCpu = `${idxE + 1}.${idxI + 1}`;
                        let totalItemVenda = 0;
                        (item.insumos || []).forEach(ins => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const cIn = num(ins.coeficiente) * num(item.quantidade) * num(ins.valorUnitario);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          totalItemVenda += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                        });

                        data.push([numCpu, item.servico || item.descricao, item.unidade, num(item.quantidade), totalItemVenda / num(item.quantidade), totalItemVenda]);
                        
                        (item.insumos || []).forEach((ins, idxIn) => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const custoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          const fatBdi = isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi;
                          
                          data.push([`${numCpu}.${idxIn + 1}`, `[${ins.tipo}] ${ins.descricao}`, ins.unidade || "un", num(ins.coeficiente) * num(item.quantidade), custoUnit * fatBdi, (num(ins.coeficiente) * num(item.quantidade)) * custoUnit * fatBdi]);
                        });
                      });
                    });
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Preço de Venda");
                    XLSX.writeFile(wb, `${projetoAtivo.nome || "Orcamento"}_Preco_Venda.xlsx`);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-emerald-200 text-emerald-700 bg-emerald-50/50 rounded hover:bg-emerald-50 flex items-center gap-1"
                >
                  <Download size={12} /> Excel (.xlsx)
                </button>

                {/* PDF LIMPO DA PLANILHA DE VENDA */}
                <button 
                  onClick={() => {
                    const tituloOriginal = document.title;
                    document.title = `${projetoAtivo.nome || "Orcamento"}_Preco_Venda`;
                    const estiloPrint = document.createElement("style");
                    estiloPrint.innerHTML = `
                      @media print {
                        body * { visibility: hidden; }
                        #area-planilha-venda, #area-planilha-venda * { visibility: visible; }
                        #area-planilha-venda { position: absolute; left: 0; top: 0; width: 100%; background: white !important; }
                      }
                    `;
                    document.head.appendChild(estiloPrint);
                    window.print();
                    document.head.removeChild(estiloPrint);
                    document.title = tituloOriginal;
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1"
                >
                  <Download size={12} /> PDF (.pdf)
                </button>
              </div>
            </div>

            <div id="area-planilha-venda" className="border border-stone-200 rounded-lg overflow-hidden bg-white">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Estrutura (Etapa / CPU / Insumo)</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Prop.</span>
                <span className="col-span-1.5 text-right">Preço Unit Venda</span>
                <span className="col-span-2 text-right">Total Venda</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[600px] overflow-y-auto">
                {etapas.length === 0 ? (
                  <div className="p-8 text-center text-stone-400 italic text-xs">
                    Nenhuma etapa cadastrada neste orçamento.
                  </div>
                ) : (
                  etapas.map((etapa, idxEtapa) => {
                    const numEtapa = idxEtapa + 1;
                    const etapaId = etapa.id || `etapa-${idxEtapa}`;
                    const isEtapaAberta = !!etapasExpandidas[etapaId];

                    let totalEtapaComBdi = 0;
                    (etapa.itens || []).forEach(it => {
                      const qCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tIn = String(ins.tipo || "").toUpperCase().trim();
                        const cIn = num(ins.coeficiente) * qCpu * num(ins.valorUnitario);
                        const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                        totalEtapaComBdi += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                      });
                    });

                    return (
                      <div key={etapaId} className="bg-stone-50/30">
                        <div 
                          onClick={() => setEtapasExpandidas(p => ({ ...p, [etapaId]: !isEtapaAberta }))}
                          className="grid grid-cols-12 gap-2 px-4 py-2 bg-stone-200/60 text-stone-800 text-xs font-bold items-center uppercase tracking-wide cursor-pointer hover:bg-stone-200 select-none"
                        >
                          <span className="col-span-10 flex items-center gap-1.5">
                            {isEtapaAberta ? <ChevronDown size={14} className="text-stone-500" /> : <ChevronRight size={14} className="text-stone-500" />}
                            📁 {numEtapa}. {etapa.nome}
                          </span>
                          <span className="col-span-2 text-right font-mono text-stone-900">
                            R$ {fmt(totalEtapaComBdi)}
                          </span>
                        </div>

                        {isEtapaAberta && (etapa.itens || []).map((item, idxItem) => {
                          const numCpu = `${numEtapa}.${idxItem + 1}`;
                          const itemId = item.id || `item-${numCpu}`;
                          const isCpuAberta = !!cpusExpandidas[itemId];
                          const qtdItem = num(item.quantidade);

                          let totalCpuComBdi = 0;
                          (item.insumos || []).forEach(ins => {
                            const tIn = String(ins.tipo || "").toUpperCase().trim();
                            const cIn = num(ins.coeficiente) * num(ins.valorUnitario);
                            const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                            totalCpuComBdi += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                          });

                          return (
                            <div key={itemId} className="border-b border-stone-100">
                              <div 
                                onClick={() => setCpusExpandidas(p => ({ ...p, [itemId]: !isCpuAberta }))}
                                className="grid grid-cols-12 gap-2 px-4 py-2 bg-white text-xs items-center font-semibold text-stone-700 pl-8 cursor-pointer hover:bg-stone-50 select-none"
                              >
                                <span className="col-span-6 truncate text-stone-900 flex items-center gap-1">
                                  {isCpuAberta ? <ChevronDown size={13} className="text-stone-400 shrink-0" /> : <ChevronRight size={13} className="text-stone-400 shrink-0" />}
                                  🔹 {numCpu}. {item.codigo ? `[${item.codigo}] ` : ""}{item.servico || item.descricao}
                                </span>
                                <span className="col-span-1 text-center font-mono text-stone-400">{item.unidade}</span>
                                <span className="col-span-1.5 text-right font-mono">{qtdItem.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                                <span className="col-span-1.5 text-right font-mono text-stone-500">R$ {fmt(totalCpuComBdi)}</span>
                                <span className="col-span-2 text-right font-mono text-stone-900">
                                  R$ {fmt(qtdItem * totalCpuComBdi)}
                                </span>
                              </div>

                              {isCpuAberta && (item.insumos || []).length > 0 && (
                                <div className="bg-stone-50/50 divide-y divide-stone-100/60 border-t border-b border-stone-100">
                                  {(item.insumos || []).map((insumo, idxInsumo) => {
                                    const numInsumo = `${numCpu}.${idxInsumo + 1}`;
                                    const tIn = String(insumo.tipo || "").toUpperCase().trim();
                                    const entry = catalogMap.get(precoKey(insumo.descricao));
                                    const custoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(insumo.valorUnitario);
                                    
                                    const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                                    const precoVendaInsumo = custoUnit * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                                    
                                    const qtdCalculada = num(insumo.coeficiente) * qtdItem;
                                    const vendaTotalInsumo = qtdCalculada * precoVendaInsumo;

                                    return (
                                      <div key={insumo.id || idxInsumo} className="grid grid-cols-12 gap-2 px-4 py-1.5 text-[11px] items-center text-stone-600 pl-14 hover:bg-stone-100/40">
                                        <span className="col-span-6 truncate uppercase font-sans text-stone-500">
                                          {numInsumo}. <span className="text-[9px] font-mono font-bold text-stone-400 border border-stone-200 px-1 py-0.5 rounded bg-white mr-1">{insumo.tipo}</span> {insumo.descricao}
                                        </span>
                                        <span className="col-span-1 text-center font-mono text-stone-400 uppercase text-[10px]">{insumo.unidade || "un"}</span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-600">
                                          {qtdCalculada.toLocaleString("pt-BR", { minimumFractionDigits: 3, maximumFractionDigits: 3 })}
                                        </span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(precoVendaInsumo)}</span>
                                        <span className="col-span-2 text-right font-mono font-medium text-blue-700">
                                          R$ {fmt(vendaTotalInsumo)}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })
                )}
                
                {/* LINHA DE TOTAIS GERAIS DA PLANILHA DE VENDA */}
                <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold uppercase tracking-wider">
                  <span className="col-span-6">VALOR FINAL DE VENDA COM BDI</span>
                  <span className="col-span-1"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-2 text-right font-mono text-amber-400">
                    R$ {fmt(bdiCalc.valorVenda)}
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}

        {tab === "maoobra" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div>
              <h2 className="text-base font-semibold text-stone-800">Consolidado Qualitativo de Mão de Obra</h2>
              <p className="text-xs text-stone-500">Visualização agrupada de todas as horas e custos de mão de obra alocados no orçamento.</p>
            </div>

            <div className="border border-stone-200 rounded-lg overflow-hidden">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Descrição do Profissional</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Horas Totais</span>
                <span className="col-span-1.5 text-right">Valor Unit.</span>
                <span className="col-span-2 text-right">Subtotal Direto</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[500px] overflow-y-auto">
                {(() => {
                  const mos = new Map();
                  (etapas || []).forEach(e => {
                    (e.itens || []).forEach(it => {
                      const qtdCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tipo = String(ins.tipo || "").toUpperCase().trim();
                        if (tipo === "MO" || tipo.includes("MÃO") || tipo.includes("MAO")) {
                          if (!String(ins.descricao || "").trim()) return;
                          
                          const chave = ins.descricao.trim().toLowerCase();
                          const qtdCalc = num(ins.coeficiente) * qtdCpu;
                          
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const vUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          
                          if (mos.has(chave)) {
                            const existente = mos.get(chave);
                            existente.qtd += qtdCalc;
                            existente.total += qtdCalc * vUnit;
                          } else {
                            mos.set(chave, {
                              descricao: ins.descricao,
                              unidade: ins.unidade || "h",
                              qtd: qtdCalc,
                              valorUnit: vUnit,
                              total: qtdCalc * vUnit
                            });
                          }
                        }
                      });
                    });
                  });

                  const listaMo = Array.from(mos.values()).sort((a, b) => b.total - a.total);
                  if (listaMo.length === 0) {
                    return <div className="p-8 text-center text-stone-400 italic text-xs">Nenhuma mão de obra localizada no orçamento.</div>;
                  }

                  return (
                    <>
                      {listaMo.map((r, idx) => (
                        <div key={idx} className="grid grid-cols-12 gap-2 px-4 py-2 text-xs items-center hover:bg-stone-50/60 uppercase">
                          <span className="col-span-6 font-medium text-stone-800 truncate">{r.descricao}</span>
                          <span className="col-span-1 text-center font-mono text-stone-400">{r.unidade}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-900">{r.qtd.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(r.valorUnit)}</span>
                          <span className="col-span-2 text-right font-mono font-semibold text-stone-700">R$ {fmt(r.total)}</span>
                        </div>
                      ))}
                      
                      <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold">
                        <span className="col-span-6">TOTAL GERAL EM MÃO DE OBRA</span>
                        <span className="col-span-1"></span>
                        <span className="col-span-1.5 text-right font-mono text-stone-300">
                          {listaMo.reduce((acc, curr) => acc + curr.qtd, 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-2 text-right font-mono text-amber-400">
                          R$ {fmt(listaMo.reduce((acc, curr) => acc + curr.total, 0))}
                        </span>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>
          </div>
        )}

        {tab === "materiais" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div>
              <h2 className="text-base font-semibold text-stone-800">Quantitativo de Materiais</h2>
              <p className="text-xs text-stone-500">Consolidação de todos os materiais físicos consumidos nas CPUs do orçamento ativo.</p>
            </div>

            <div className="border border-stone-200 rounded-lg overflow-hidden">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Material</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Total</span>
                <span className="col-span-1.5 text-right">Preço Unit.</span>
                <span className="col-span-2 text-right">Total Bruto</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[500px] overflow-y-auto">
                {(() => {
                  const mats = new Map();
                  (etapas || []).forEach(e => {
                    (e.itens || []).forEach(it => {
                      const qtdCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tipo = String(ins.tipo || "").toUpperCase().trim();
                        if (tipo === "MAT" || tipo === "MATERIAL" || (!tipo.includes("MO") && !tipo.includes("MÃO") && !tipo.includes("MAO") && !tipo.includes("EQUIP"))) {
                          if (!String(ins.descricao || "").trim()) return; 
                          
                          const chave = ins.descricao.trim().toLowerCase();
                          const qtdCalc = num(ins.coeficiente) * qtdCpu;
                          
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const vUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          
                          if (mats.has(chave)) {
                            const existente = mats.get(chave);
                            existente.qtd += qtdCalc;
                            existente.total += qtdCalc * vUnit;
                          } else {
                            mats.set(chave, {
                              descricao: ins.descricao,
                              unidade: ins.unidade || "un",
                              qtd: qtdCalc,
                              valorUnit: vUnit,
                              total: qtdCalc * vUnit
                            });
                          }
                        }
                      });
                    });
                  });

                  const listaMats = Array.from(mats.values()).sort((a, b) => b.total - a.total);
                  if (listaMats.length === 0) {
                    return <div className="p-8 text-center text-stone-400 italic text-xs">Nenhum material localizado neste orçamento.</div>;
                  }

                  return (
                    <>
                      {listaMats.map((r, idx) => (
                        <div key={idx} className="grid grid-cols-12 gap-2 px-4 py-2 text-xs items-center hover:bg-stone-50/60 uppercase">
                          <span className="col-span-6 font-medium text-stone-800 truncate">{r.descricao}</span>
                          <span className="col-span-1 text-center font-mono text-stone-400">{r.unidade}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-900">{r.qtd.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(r.valorUnit)}</span>
                          <span className="col-span-2 text-right font-mono font-semibold text-emerald-700">R$ {fmt(r.total)}</span>
                        </div>
                      ))}
                      
                      <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold">
                        <span className="col-span-6">TOTAL GERAL EM MATERIAIS</span>
                        <span className="col-span-1"></span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-2 text-right font-mono text-amber-400">
                          R$ {fmt(listaMats.reduce((acc, curr) => acc + curr.total, 0))}
                        </span>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>
          </div>
        )}
        {tab === "precos" && projetoAtivo && (
          <PrecosTab
            catalog={catalog}
            onUpsert={upsertPreco}
            onRemove={removePreco}
            onApplyToCpus={aplicarPrecoNoOrcamentoAtivo}
            onApplyAllToCpus={aplicarTodosPrecosNoOrcamentoAtivo}
          />
        )}
      </div>
    </div>
  );
}

function TabBtn({ active, onClick, icon, children, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
        disabled ? "opacity-40 cursor-not-allowed" : ""
      } ${active ? "border-stone-900 text-stone-900" : "border-transparent text-stone-400 hover:text-stone-600"}`}
    >
      {icon}
      {children}
    </button>
  );
}

function CpuLibrary({ cpus, setCpus, fileInputRef, catalogMap }) {
  const [query, setQuery] = useState("");
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [editing, setEditing] = useState(null);
  const [expanded, setExpanded] = useState({});
  const [importMsg, setImportMsg] = useState("");
  
  // NOVO: Controla qual linha do resultado filtrado está focada pelo teclado
  const [activeIndex, setActiveIndex] = useState(-1);

  const fontes = useMemo(
    () => ["Todas", ...Array.from(new Set(cpus.map((c) => c.fonte).filter(Boolean)))],
    [cpus]
  );

  const queryTokens = useMemo(() => {
    const tokens = [];
    const re = /"([^"]+)"|(\S+)/g;
    let m;
    while ((m = re.exec(query)) !== null) {
      const t = norm((m[1] || m[2] || "").trim());
      if (t) tokens.push(t);
    }
    return tokens;
  }, [query]);

  const filtered = cpus.filter((c) => {
    const matchesFonte = fonteFiltro === "Todas" || c.fonte === fonteFiltro;
    if (!matchesFonte) return false;
    if (queryTokens.length === 0) return true;
    const haystack = norm(
      c.codigo + " " + c.descricao + " " + c.fonte + " " + c.insumos.map((i) => i.descricao).join(" ")
    );
    return queryTokens.every((t) => haystack.includes(t));
  });

  const [confirmingDelete, setConfirmingDelete] = useState(null);

  const removeCpu = (id) => {
    setCpus(cpus.filter((c) => c.id !== id));
    setConfirmingDelete(null);
  };

  const duplicateCpu = (c) => {
    setCpus([...cpus, { ...c, id: uid(), codigo: c.codigo + " (cópia)", insumos: c.insumos.map((i) => ({ ...i, id: uid() })) }]);
  };

  const saveCpu = (cpu) => {
    if (cpus.find((c) => c.id === cpu.id)) {
      setCpus(cpus.map((c) => (c.id === cpu.id ? cpu : c)));
    } else {
      setCpus([...cpus, cpu]);
    }
    setEditing(null);
  };

  // NOVO: Gerencia a navegação por setas e Enter na listagem
  const handleKeyDown = (evt) => {
    if (filtered.length === 0) return;

    if (evt.key === "ArrowDown") {
      evt.preventDefault();
      setActiveIndex((prev) => (prev + 1) % filtered.length);
    } else if (evt.key === "ArrowUp") {
      evt.preventDefault();
      setActiveIndex((prev) => (prev - 1 + filtered.length) % filtered.length);
    } else if (evt.key === "Enter") {
      if (activeIndex >= 0 && activeIndex < filtered.length) {
        evt.preventDefault();
        const targetCpu = filtered[activeIndex];
        setExpanded((prev) => ({ ...prev, [targetCpu.id]: !prev[targetCpu.id] }));
      }
    } else if (evt.key === "Escape") {
      setQuery("");
      setActiveIndex(-1);
    }
  };

  const handleImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportMsg("Lendo planilha...");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const getField = (row, names) => {
        for (const key of Object.keys(row)) {
          if (names.includes(norm(key))) return row[key];
        }
        return "";
      };

      const headers = rows.length ? Object.keys(rows[0]).map(norm) : [];
      const hasInsumoColumn = headers.some((h) => ["insumo", "item", "insumo_descricao"].includes(h));

      const inferTipo = (desc) => {
        const d = norm(desc);
        if (/^chp\/|^chi\/|caminhao|trator|escavadeira|pa carregadeira|guindaste|compactador|motoniveladora|retroescavadeira/.test(d)) return "EQUIP";
        if (/servente|pedreiro|oficial|ajudante|encarregado|mestre de obras|carpinteiro|armador|eletricista|pintor/.test(d)) return "MO";
        return "MAT";
      };

      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerText = rawRows.slice(0, 5).flat().join(" ").toLowerCase();
      const isRelatorioSudecap = headerText.includes("relatório de composiç") || headerText.includes("relatorio de composic");

      const novas = [];

      if (isRelatorioSudecap) {
        const fonteNome = file.name.toLowerCase().includes("der") ? "DER-MG" : "SUDECAP";
        let atual = null;
        rawRows.slice(3).forEach((row) => {
          const c0 = String(row[0] ?? "").trim();
          const c1 = String(row[1] ?? "").trim();
          const c2 = String(row[2] ?? "").trim();
          const und = String(row[7] ?? "").trim();
          const consumoRaw = row[9];
          if (c0) {
            if (und) {
              atual = { id: uid(), codigo: c0, fonte: fonteNome, descricao: c1, unidade: und, insumos: [] };
              novas.push(atual);
            } else {
              atual = null;
            }
          } else if (c1 && atual) {
            atual.insumos.push({
              id: uid(), tipo: inferTipo(c2), descricao: c2, unidade: und || "un", coeficiente: consumoRaw ? num(consumoRaw) : 0, valorUnitario: ""
            });
          }
        });
      } else if (hasInsumoColumn) {
        const grouped = {};
        rows.forEach((row) => {
          const codigo = String(getField(row, ["codigo", "código", "code"])).trim();
          if (!codigo) return;
          if (!grouped[codigo]) {
            grouped[codigo] = {
              id: uid(), codigo,
              fonte: String(getField(row, ["fonte", "tabela", "origem"])) || "Própria",
              descricao: String(getField(row, ["descricao", "descrição", "servico", "serviço"])),
              unidade: String(getField(row, ["unidade", "un", "unid"])) || "un",
              insumos: []
            };
          }
          const insumoDesc = String(getField(row, ["insumo", "item", "insumo_descricao"]));
          if (insumoDesc) {
            const rawValor = getField(row, ["valor_unitario", "valor unitário", "valor", "preco", "preço"]);
            grouped[codigo].insumos.push({
              id: uid(),
              tipo: (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("MAT") && "MAT") ||
                    (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("EQUIP") && "EQUIP") ||
                    (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("MO") && "MO") || inferTipo(insumoDesc),
              descricao: insumoDesc,
              unidade: String(getField(row, ["unidade_insumo", "un_insumo", "unidade insumo"])) || "un",
              coeficiente: num(getField(row, ["coeficiente", "coef", "indice", "índice", "produtividade"])),
              valorUnitario: rawValor === "" ? "" : num(rawValor),
            });
          }
        });
        novas.push(...Object.values(grouped));
      } else {
        let atual = null;
        rows.forEach((row) => {
          const codigo = String(getField(row, ["codigo", "código", "code"])).trim();
          const descricao = String(getField(row, ["descricao", "descrição", "servico", "serviço", "item"])).trim();
          const unidade = String(getField(row, ["unidade", "un", "unid"])).trim();
          const coefRaw = getField(row, ["coeficiente", "coef", "indice", "índice", "produtividade"]);
          if (!codigo && !descricao) return;
          if (coefRaw === "" || coefRaw === undefined || coefRaw === null) {
            atual = { id: uid(), codigo: codigo || "(sem código)", fonte: "Própria", descricao, unidade: unidade || "un", insumos: [] };
            novas.push(atual);
          } else if (atual) {
            atual.insumos.push({ id: uid(), tipo: inferTipo(descricao), descricao, unidade: unidade || "un", coeficiente: num(coefRaw), valorUnitario: "" });
          }
        });
      }

      if (novas.length === 0) {
        setImportMsg("Nenhuma composição reconhecida.");
      } else {
        setCpus([...cpus, ...novas]);
        setImportMsg(`${novas.length} CPU(s) importada(s) para a Base Geral.`);
      }
    } catch (err) {
      setImportMsg("Erro ao ler: " + err.message);
    }
    e.target.value = "";
    setTimeout(() => setImportMsg(""), 4000);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <div className="relative flex-1 min-w-[200px]">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
          <input
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setActiveIndex(-1); // Reseta a linha ativa ao digitar
            }}
            onKeyDown={handleKeyDown} // NOVO: Gatilho para monitorar as setas do teclado
            placeholder='Buscar na biblioteca... ex: "alvenaria" "bloco"'
            className="w-full pl-8 pr-3 py-2 text-sm border border-stone-300 rounded-lg bg-white focus:outline-none focus:ring-1 focus:ring-stone-500"
          />
        </div>
        <select value={fonteFiltro} onChange={(e) => { setFonteFiltro(e.target.value); setActiveIndex(-1); }} className="px-3 py-2 text-sm border border-stone-300 rounded-lg bg-white">
          {fontes.map((f) => <option key={f}>{f}</option>)}
        </select>
        <label className="flex items-center gap-1.5 px-3 py-2 text-sm border border-stone-300 rounded-lg bg-white cursor-pointer hover:bg-stone-100">
          <Upload size={15} /> Importar Planilha
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImport} />
        </label>
        <button onClick={() => setEditing("new")} className="flex items-center gap-1.5 px-3 py-2 text-sm bg-stone-900 text-white rounded-lg hover:bg-stone-700">
          <Plus size={15} /> Nova CPU Base
        </button>
      </div>

      {importMsg && <div className="mb-4 text-xs px-3 py-2 rounded-lg bg-amber-50 text-amber-800 border border-amber-200">{importMsg}</div>}

      <div className="space-y-2">
        {filtered.map((c, index) => (
          <div 
            key={c.id} 
            className={`border rounded-lg bg-white transition-all ${
              index === activeIndex ? "border-stone-500 ring-1 ring-stone-500 bg-stone-50/40" : "border-stone-200"
            }`}
          >
            <div className="flex items-center gap-3 px-4 py-3 cursor-pointer" onClick={() => { setExpanded({ ...expanded, [c.id]: !expanded[c.id] }); setActiveIndex(index); }}>
              {expanded[c.id] ? <ChevronDown size={16} className="text-stone-400 shrink-0" /> : <ChevronRight size={16} className="text-stone-400 shrink-0" />}
              <span className="text-[11px] font-mono px-1.5 py-0.5 bg-stone-100 rounded text-stone-500 shrink-0">{c.fonte}</span>
              <span className="text-xs font-mono text-stone-500 shrink-0">{c.codigo}</span>
              <span className={`text-sm flex-1 truncate ${index === activeIndex ? "font-medium text-stone-900" : "text-stone-800"}`}>{c.descricao}</span>
              <span className="text-xs text-stone-400 shrink-0">/{c.unidade}</span>
              <div className="flex gap-1 shrink-0" onClick={(e) => e.stopPropagation()}>
                <IconBtn onClick={() => duplicateCpu(c)} title="Duplicar"><Copy size={14} /></IconBtn>
                <IconBtn onClick={() => setEditing(c)} title="Editar"><Pencil size={14} /></IconBtn>
                {confirmingDelete === c.id ? (
                  <span className="flex items-center gap-1 text-xs">
                    <button onClick={() => removeCpu(c.id)} className="px-1.5 py-0.5 bg-red-600 text-white rounded">Sim</button>
                    <button onClick={() => setConfirmingDelete(null)} className="px-1.5 py-0.5 border border-stone-300 rounded">Não</button>
                  </span>
                ) : (
                  <IconBtn onClick={() => setConfirmingDelete(c.id)} title="Excluir"><Trash2 size={14} /></IconBtn>
                )}
              </div>
            </div>
            {expanded[c.id] && (
              <div className="px-4 pb-3 border-t border-stone-100 pt-2 bg-stone-50/50">
                <InsumoTable insumos={c.insumos} readOnly />
              </div>
            )}
          </div>
        ))}
      </div>

      {editing && <CpuEditor cpu={editing === "new" ? null : editing} onCancel={() => setEditing(null)} onSave={saveCpu} catalogMap={catalogMap} />}
    </div>
  );
}

function IconBtn({ onClick, title, children }) {
  return <button onClick={onClick} title={title} className="p-1.5 text-stone-400 hover:text-stone-700 hover:bg-stone-100 rounded">{children}</button>;
}

/* ---------------- TABELA DE INSUMOS PADRONIZADA ---------------- */
function InsumoTable({ insumos, readOnly, onChange, catalogMap }) {
  const setMany = (id, patch) => onChange(insumos.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  const set = (id, field, value) => setMany(id, { [field]: value });
  const remove = (id) => onChange(insumos.filter((i) => i.id !== id));

  const handleDescricaoBlur = (i) => {
    if (!catalogMap) return;
    const entry = catalogMap.get(norm(i.descricao));
    if (!entry) return;
    const semValor = i.valorUnitario === "" || i.valorUnitario === null || i.valorUnitario === undefined;
    if (semValor && entry.valorUnitario !== "" && entry.valorUnitario !== null) {
      setMany(i.id, { valorUnitario: entry.valorUnitario, tipo: i.tipo || entry.tipo, unidade: i.unidade || entry.unidade });
    }
  };

  return (
    <table className="w-full text-xs">
      <thead>
        <tr className="text-stone-400 text-left">
          <th className="font-normal py-1 pr-2 w-24">Tipo</th>
          <th className="font-normal py-1 pr-2">Insumo</th>
          <th className="font-normal py-1 pr-2 w-16">Un.</th>
          <th className="font-normal py-1 pr-2 w-24 text-right">Coeficiente</th>
          <th className="font-normal py-1 pr-2 w-28 text-right">Valor Unit. (R$)</th>
          <th className="font-normal py-1 pr-2 w-24 text-right">Subtotal</th>
          {!readOnly && <th className="w-7"></th>}
        </tr>
      </thead>
      <tbody>
        {insumos.map((i) => (
          <tr key={i.id} className="border-t border-stone-100">
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className="text-[10px] px-1 py-0.5 bg-stone-100 rounded text-stone-600 font-medium">{i.tipo}</span>
              ) : (
                <select value={i.tipo || "MAT"} onChange={(e) => set(i.id, "tipo", e.target.value)} className="w-full border border-stone-200 rounded p-0.5 bg-white">
                  {TIPOS.map((t) => <option key={t.v} value={t.v}>{t.label}</option>)}
                </select>
              )}
            </td>
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className="text-stone-700">{i.descricao}</span>
              ) : (
                <input
                  value={i.descricao || ""}
                  onChange={(e) => set(i.id, "descricao", e.target.value)}
                  onBlur={() => handleDescricaoBlur(i)}
                  list="insumos-catalogo"
                  className="w-full border border-stone-200 rounded px-1 py-0.5"
                />
              )}
            </td>
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className="text-stone-500">{i.unidade}</span>
              ) : (
                <input value={i.unidade || ""} onChange={(e) => set(i.id, "unidade", e.target.value)} className="w-full border border-stone-200 rounded px-1 py-0.5" />
              )}
            </td>
            <td className="py-1 pr-2 text-right">
              {readOnly ? (
                <span className="font-mono">{i.coeficiente}</span>
              ) : (
                <input type="number" step="any" value={i.coeficiente ?? ""} onChange={(e) => set(i.id, "coeficiente", e.target.value)} className="w-20 border border-stone-200 rounded px-1 py-0.5 text-right font-mono" />
              )}
            </td>
            <td className="py-1 pr-2 text-right">
              {readOnly ? (
                <span className="font-mono text-stone-600">{i.valorUnitario !== "" ? `R$ ${fmt(i.valorUnitario)}` : "—"}</span>
              ) : (
                <input type="number" step="any" value={i.valorUnitario ?? ""} onChange={(e) => set(i.id, "valorUnitario", e.target.value)} placeholder="0,00" className="w-24 border border-stone-200 rounded px-1 py-0.5 text-right font-mono" />
              )}
            </td>
            <td className="py-1 pr-2 text-right font-mono text-stone-600">
              R$ {fmt(num(i.coeficiente) * num(i.valorUnitario))}
            </td>
            {!readOnly && (
              <td className="py-1 text-center">
                <button onClick={() => remove(i.id)} className="text-stone-300 hover:text-red-500"><X size={13} /></button>
              </td>
            )}
          </tr>
        ))}
        {!readOnly && (
          <tr>
            <td colSpan="7" className="py-2">
              <button type="button" onClick={() => onChange([...insumos, { id: uid(), tipo: "MAT", descricao: "", unidade: "un", coeficiente: 1, valorUnitario: "" }])} className="text-stone-500 hover:text-stone-900 font-medium flex items-center gap-1">
                <Plus size={12} /> Adicionar Insumo
              </button>
            </td>
          </tr>
        )}
      </tbody>
    </table>
  );
}

/* ---------------- ABA BANCO DE PREÇOS (PRECOS) ---------------- */
function PrecosTab({ catalog, onUpsert, onRemove, onApplyToCpus, onApplyAllToCpus }) {
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const exportarXls = () => {
    const wb = XLSX.utils.book_new();
    const rows = [
      ["Banco de Preços — Catálogo de Insumos"],
      [],
      ["Tipo", "Descrição", "Unidade", "Valor Unitário (R$)", "Ocorrências na Planilha"],
      ...catalog.map((c) => [c.tipo, c.descricao, c.unidade, c.valorUnitario !== "" ? c.valorUnitario : "", c.ocorrencias]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 45 }, { wch: 10 }, { wch: 20 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, ws, "Banco de Preços");
    XLSX.writeFile(wb, "banco_de_precos.xlsx");
  };

  const filtered = catalog.filter((c) => {
    // Divide o texto digitado por espaços e remove itens vazios
    const searchTerms = norm(query).split(/\s+/).filter(Boolean);
    const targetText = norm(c.descricao);
    
    // Verifica se TODAS as palavras buscadas estão presentes na descrição do insumo
    return searchTerms.every((term) => targetText.includes(term));
  });

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
        <div className="relative w-72">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Filtrar catálogo de preços..." className="w-full pl-8 pr-3 py-2 text-sm border border-stone-300 rounded-lg" />
        </div>
        <div className="flex gap-2">
          <button onClick={onApplyAllToCpus} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-stone-50 hover:bg-stone-100 text-stone-700">
            <RefreshCw size={13} /> Sincronizar Tudo na Planilha de Custos
          </button>
          <button onClick={exportarXls} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700">
            <Download size={13} /> Exportar .xlsx
          </button>
          <button onClick={() => setEditing({ id: null, descricao: "", tipo: "MAT", unidade: "un", valorUnitario: "" })} className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg font-medium hover:bg-stone-700">
            <Plus size={13} /> Novo Insumo Manual
          </button>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs border-collapse">
          <thead>
            <tr className="border-b border-stone-200 text-stone-400 font-normal">
              <th className="py-2 pr-3 w-28">Tipo</th>
              <th className="py-2 pr-3">Descrição Única do Insumo</th>
              <th className="py-2 pr-3 w-20">Un.</th>
              <th className="py-2 pr-3 w-32 text-right">Preço Padrão (R$)</th>
              <th className="py-2 pr-3 w-28 text-center">Na Planilha Ativa</th>
              <th className="py-2 w-24 text-center">Ações</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((c) => (
              <tr key={c.key} className="border-b border-stone-100 hover:bg-stone-50/50">
                <td className="py-2 pr-3">
                  <span className="px-1.5 py-0.5 bg-stone-100 text-stone-600 rounded font-medium text-[10px]">{c.tipo}</span>
                </td>
                <td className="py-2 pr-3 font-medium text-stone-800">
                  {c.descricao}
                  {c.divergente && (
                    <span className="ml-2 text-[10px] bg-amber-50 text-amber-700 border border-amber-200 px-1.5 py-0.2 rounded inline-flex items-center gap-1">
                      <AlertTriangle size={10} /> Preço Divergente
                    </span>
                  )}
                </td>
                <td className="py-2 pr-3 text-stone-500 font-mono">{c.unidade}</td>
                <td className="py-2 pr-3 text-right font-mono font-medium text-stone-900">
                  {c.valorUnitario !== "" ? `R$ ${fmt(c.valorUnitario)}` : <span className="text-stone-300">Não definido</span>}
                </td>
                <td className="py-2 pr-3 text-center text-stone-500">{c.ocorrencias} item(ns)</td>
                <td className="py-2 text-center flex justify-center gap-1">
                  <button onClick={() => setEditing(c)} className="p-1 border border-stone-200 rounded text-stone-600 hover:bg-stone-100" title="Editar Preço">
                    <Pencil size={12} />
                  </button>
                  {c.valorUnitario !== "" && c.ocorrencias > 0 && (
                    <button onClick={() => onApplyToCpus(c.descricao, c.valorUnitario)} className="p-1 border border-stone-200 rounded bg-stone-50 text-stone-700 hover:bg-stone-100" title="Forçar este valor nos Custos desta obra">
                      <Check size={12} />
                    </button>
                  )}
                  <button onClick={() => onRemove(c.descricao)} className="p-1 border border-stone-200 rounded text-stone-400 hover:text-red-600" title="Remover Referência">
                    <X size={12} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 bg-stone-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white border border-stone-200 rounded-xl max-w-md w-full p-5 shadow-lg">
            <h3 className="font-semibold text-sm text-stone-900 mb-4">{editing.key ? "Editar Insumo do Catálogo" : "Novo Insumo no Banco"}</h3>
            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-stone-500 mb-1">Descrição</label>
                <input disabled={!!editing.key} value={editing.descricao} onChange={(e) => setEditing({ ...editing, descricao: e.target.value })} className="w-full border border-stone-300 rounded-lg px-3 py-2 disabled:bg-stone-50" />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-stone-500 mb-1">Tipo</label>
                  <select value={editing.tipo} onChange={(e) => setEditing({ ...editing, tipo: e.target.value })} className="w-full border border-stone-300 rounded-lg px-2 py-2 bg-white">
                    {TIPOS.map((t) => <option key={t.v} value={t.v}>{t.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-stone-500 mb-1">Unidade</label>
                  <input value={editing.unidade} onChange={(e) => setEditing({ ...editing, unidade: e.target.value })} className="w-full border border-stone-300 rounded-lg px-3 py-2" />
                </div>
              </div>
              <div>
                <label className="block text-stone-500 mb-1">Valor Unitário Homologado (R$)</label>
                <input type="number" step="any" value={editing.valorUnitario} onChange={(e) => setEditing({ ...editing, valorUnitario: e.target.value })} placeholder="0,00" className="w-full border border-stone-300 rounded-lg px-3 py-2 font-mono text-sm" />
              </div>
            </div>
            <div className="mt-5 flex justify-end gap-2 text-xs">
              <button onClick={() => setEditing(null)} className="px-3 py-2 border border-stone-300 rounded-lg">Cancelar</button>
              <button
                onClick={() => {
                  onUpsert(editing.descricao, editing.tipo, editing.unidade, editing.valorUnitario === "" ? "" : num(editing.valorUnitario));
                  if (editing.key && editing.valorUnitario !== "") {
                    onApplyToCpus(editing.descricao, num(editing.valorUnitario));
                  }
                  setEditing(null);
                }}
                className="px-3 py-2 bg-stone-900 text-white rounded-lg hover:bg-stone-700"
              >
                Salvar e Replicar nos Custos
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- EDITOR DE CPUS INDIVIDUAIS ---------------- */
function CpuEditor({ cpu, onCancel, onSave, catalogMap }) {
  const [codigo, setCodigo] = useState(cpu?.codigo || "");
  const [fonte, setFonte] = useState(cpu?.fonte || "Própria");
  const [descricao, setDescricao] = useState(cpu?.descricao || "");
  const [unidade, setUnidade] = useState(cpu?.unidade || "m²");
  const [insumos, setInsumos] = useState(cpu?.insumos ? JSON.parse(JSON.stringify(cpu.insumos)) : []);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!descricao.trim()) return;
    onSave({ id: cpu?.id || uid(), codigo, fonte, descricao, unidade, insumos });
  };

  return (
    <div className="fixed inset-0 bg-stone-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-40 overflow-y-auto">
      <form onSubmit={handleSubmit} className="bg-white border border-stone-200 rounded-xl max-w-2xl w-full p-5 shadow-lg my-8">
        <h3 className="font-semibold text-sm mb-4">{cpu ? "Editar Composição da Base" : "Nova Composição Técnica"}</h3>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs mb-4">
          <div>
            <label className="block text-stone-500 mb-1">Tabela / Fonte</label>
            <input value={fonte} onChange={(e) => setFonte(e.target.value)} placeholder="Ex: SINAPI, SUDECAP" className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
          <div>
            <label className="block text-stone-500 mb-1">Código Identificador</label>
            <input value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ex: 12.34.56" className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
          <div>
            <label className="block text-stone-500 mb-1">Unidade Principal</label>
            <input value={unidade} onChange={(e) => setUnidade(e.target.value)} className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
        </div>
        <div className="text-xs mb-4">
          <label className="block text-stone-500 mb-1">Descrição Técnica da Composição</label>
          <input value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex: Concreto armado fck=25mpa..." className="w-full border border-stone-300 rounded-lg px-3 py-2" />
        </div>

        <div className="border-t border-stone-200 pt-3">
          <h4 className="text-xs font-semibold text-stone-700 mb-2">Estrutura de Insumos da CPU</h4>
          <InsumoTable insumos={insumos} onChange={setInsumos} catalogMap={catalogMap} />
        </div>

        <div className="mt-6 pt-3 border-t border-stone-200 flex justify-end gap-2 text-xs">
          <button type="button" onClick={onCancel} className="px-4 py-2 border border-stone-300 rounded-lg">Cancelar</button>
          <button type="submit" className="px-4 py-2 bg-stone-900 text-white rounded-lg hover:bg-stone-700">Salvar na Biblioteca</button>
        </div>
      </form>
    </div>
  );
}

/* ---------------- PLANILHA DE ORÇAMENTO / CUSTO ---------------- */
function Orcamento({ etapas, setEtapas, cpus, grandTotal, catalogMap }) {
  const [buscasPorEtapa, setBuscasPorEtapa] = useState({}); // Controla a busca de cada etapa individualmente
  const [editingEtapaId, setEditingEtapaId] = useState(null);
  const [editingEtapaNome, setEditingEtapaNome] = useState("");
  
  // Controla o índice do item selecionado via teclado para cada etapa
  const [activeIndices, setActiveIndices] = useState({}); 

  // NOVO: Controla quais itens da etapa estão expandidos (mostrando insumos)
  const [itensExpandidos, setItensExpandidos] = useState({});

  const adicionarEtapa = () => {
    setEtapas([...etapas, { id: uid(), nome: `Nova Etapa ${etapas.length + 1}`, itens: [] }]);
  };

  const removerEtapa = (id) => {
    if (etapas.length <= 1) return;
    setEtapas(etapas.filter((e) => e.id !== id));
  };

  const salvarNomeEtapa = (id) => {
    setEtapas(etapas.map((e) => (e.id === id ? { ...e, nome: editingEtapaNome } : e)));
    setEditingEtapaId(null);
  };

  const lancarCpuNaEtapa = (etapaId, cpu) => {
    const insumosAjustados = applyCatalogToInsumos(cpu.insumos, catalogMap);
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: [
            ...e.itens,
            {
              id: uid(),
              cpuId: cpu.id,
              codigo: cpu.codigo,
              servico: cpu.descricao,
              unidade: cpu.unidade,
              quantidade: 1,
              insumos: insumosAjustados,
            },
          ],
        };
      })
    );
  };

  const mudarQuantidadeItem = (etapaId, itemId, Qtd) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: e.itens.map((it) => (it.id === itemId ? { ...it, quantidade: Qtd } : it)),
        };
      })
    );
  };

  const mudarInsumosDoItem = (etapaId, itemId, novosInsumos) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: e.itens.map((it) => (it.id === itemId ? { ...it, insumos: novosInsumos } : it)),
        };
      })
    );
  };

  const removerItemDaEtapa = (etapaId, itemId) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return { ...e, itens: e.itens.filter((it) => it.id !== itemId) };
      })
    );
  };

  const toggleExpandirItem = (itemId) => {
    setItensExpandidos((prev) => ({ ...prev, [itemId]: !prev[itemId] }));
  };

  const obterCpusFiltradas = (textoBusca) => {
    if (!textoBusca || !textoBusca.trim()) return [];
    const searchTerms = norm(textoBusca).split(/\s+/).filter(Boolean);
    return cpus
      .filter((c) => {
        const targetText = norm(c.codigo + " " + c.descricao);
        return searchTerms.every((term) => targetText.includes(term));
      })
      .slice(0, 10);
  };

  const handleKeyDown = (evt, etapaId, listaCpus) => {
    if (listaCpus.length === 0) return;
    
    const currentIndex = activeIndices[etapaId] !== undefined ? activeIndices[etapaId] : -1;

    if (evt.key === "ArrowDown") {
      evt.preventDefault();
      const nextIndex = (currentIndex + 1) % listaCpus.length;
      setActiveIndices({ ...activeIndices, [etapaId]: nextIndex });
    } else if (evt.key === "ArrowUp") {
      evt.preventDefault();
      const prevIndex = (currentIndex - 1 + listaCpus.length) % listaCpus.length;
      setActiveIndices({ ...activeIndices, [etapaId]: prevIndex });
    } else if (evt.key === "Enter") {
      if (currentIndex >= 0 && currentIndex < listaCpus.length) {
        evt.preventDefault();
        lancarCpuNaEtapa(etapaId, listaCpus[currentIndex]);
        setBuscasPorEtapa({ ...buscasPorEtapa, [etapaId]: "" });
        setActiveIndices({ ...activeIndices, [etapaId]: -1 });
      }
    } else if (evt.key === "Escape") {
      setBuscasPorEtapa({ ...buscasPorEtapa, [etapaId]: "" });
      setActiveIndices({ ...activeIndices, [etapaId]: -1 });
    }
  };
  
  return (
    <div className="space-y-4">
      {/* Topo da aba: apenas o botão de Adicionar Etapa */}
      <div className="flex justify-end items-center bg-white border border-stone-200 rounded-lg p-3">
        <button onClick={adicionarEtapa} className="flex items-center gap-1 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg hover:bg-stone-700">
          <Plus size={14} /> Adicionar Nova Etapa
        </button>
      </div>

      {/* Listagem das Etapas */}
      <div className="space-y-4">
        {etapas.map((e) => {
          const termoBuscaEtapa = buscasPorEtapa[e.id] || "";
          const filtradasParaEstaEtapa = obterCpusFiltradas(termoBuscaEtapa);
          const activeIndex = activeIndices[e.id] !== undefined ? activeIndices[e.id] : -1;

          return (
            <div key={e.id} className="bg-white border border-stone-200 rounded-lg overflow-visible">
              <div className="bg-stone-50/70 px-4 py-2.5 flex justify-between items-center border-b border-stone-200">
                {editingEtapaId === e.id ? (
                  <div className="flex items-center gap-2">
                    <input value={editingEtapaNome} onChange={(e) => setEditingEtapaNome(e.target.value)} className="border border-stone-300 text-xs rounded px-2 py-1 bg-white" />
                    <button onClick={() => salvarNomeEtapa(e.id)} className="text-stone-800 font-bold text-xs">Salvar</button>
                  </div>
                ) : (
                  <h3 className="font-medium text-sm text-stone-800 flex items-center gap-2">
                    {e.nome}
                    <button onClick={() => { setEditingEtapaId(e.id); setEditingEtapaNome(e.nome); }} className="text-stone-400 hover:text-stone-700"><Pencil size={12} /></button>
                  </h3>
                )}
                {etapas.length > 1 && (
                  <button onClick={() => removerEtapa(e.id)} className="text-stone-400 hover:text-red-500"><Trash2 size={14} /></button>
                )}
              </div>

              {/* Contêiner expande dinamicamente ao digitar na busca */}
              <div className={`p-4 space-y-3 transition-all ${termoBuscaEtapa.trim() ? "min-h-[400px]" : "min-h-0"}`}>
                {/* Campo de busca exclusivo DESTA ETAPA - LARGURA TOTAL */}
                <div className="relative w-full mb-3">
                  <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
                  <input 
                    value={termoBuscaEtapa} 
                    onChange={(evt) => {
                      setBuscasPorEtapa({ ...buscasPorEtapa, [e.id]: evt.target.value });
                      setActiveIndices({ ...activeIndices, [e.id]: -1 });
                    }}
                    onKeyDown={(evt) => handleKeyDown(evt, e.id, filtradasParaEstaEtapa)}
                    placeholder="Pesquisar CPU para lançar NESTA etapa..." 
                    className="w-full pl-8 pr-3 py-1.5 text-xs border border-stone-300 rounded-lg bg-stone-50/40 focus:bg-white" 
                  />
                  
                  {termoBuscaEtapa.trim() && (
                    <div className="absolute left-0 right-0 top-full bg-white border border-stone-200 rounded-b-lg shadow-xl mt-1 z-50 max-h-[350px] overflow-y-auto text-xs">
                      {filtradasParaEstaEtapa.length === 0 && <p className="p-3 text-stone-400">Nenhuma composição encontrada.</p>}
                      {filtradasParaEstaEtapa.map((c, index) => (
                        <div 
                          key={c.id} 
                          className={`p-2 border-b border-stone-100 last:border-0 cursor-pointer flex justify-between items-center transition-colors ${
                            index === activeIndex ? "bg-stone-100 font-medium" : "hover:bg-stone-50"
                          }`} 
                          onClick={() => {
                            lancarCpuNaEtapa(e.id, c);
                            setBuscasPorEtapa({ ...buscasPorEtapa, [e.id]: "" });
                            setActiveIndices({ ...activeIndices, [e.id]: -1 });
                          }}
                        >
                          <div className="flex-1 min-w-0 pr-2">
                            <span className="font-mono text-[10px] text-stone-400 block">{c.codigo}</span>
                            <p className="truncate text-stone-800">{c.descricao}</p>
                          </div>
                          <span className={`text-[10px] px-2 py-0.5 rounded shrink-0 transition-colors ${
                            index === activeIndex ? "bg-stone-900 text-white" : "bg-stone-200 text-stone-700"
                          }`}>
                            Lançar
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Exibição dos itens da Etapa */}
                {e.itens.length === 0 && !termoBuscaEtapa.trim() && (
                  <p className="text-xs text-stone-400 italic pt-1">Nenhuma CPU lançada nesta etapa.</p>
                )}
                {e.itens.map((it) => {
                  const estaExpandido = !!itensExpandidos[it.id]; // Por padrão, undefined avalia como falso (recolhido)

                  return (
                    <div key={it.id} className="border border-stone-100 rounded-lg p-3 bg-stone-50/30">
                      {/* Cabeçalho do item - Clicável para expandir/recolher */}
                      <div className="flex flex-wrap items-center justify-between gap-3 mb-1 pb-1">
                        <div 
                          className="flex-1 min-w-0 flex items-center gap-2 cursor-pointer select-none"
                          onClick={() => toggleExpandirItem(it.id)}
                          title="Clique para alternar entre nome principal e composição completa"
                        >
                          {estaExpandido ? (
                            <ChevronDown size={14} className="text-stone-400 shrink-0" />
                          ) : (
                            <ChevronRight size={14} className="text-stone-400 shrink-0" />
                          )}
                          <div className="min-w-0">
                            <span className="font-mono text-[10px] text-stone-400">{it.codigo}</span>
                            <h4 className="text-xs font-semibold text-stone-800 truncate">{it.servico}</h4>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-3 text-xs">
                          <div className="flex items-center gap-1.5">
                            <span className="text-stone-400">Qtd:</span>
                            <input type="number" step="any" value={it.quantidade} onChange={(evt) => mudarQuantidadeItem(e.id, it.id, evt.target.value)} className="w-16 border border-stone-200 rounded px-1.5 py-0.5 text-right font-mono bg-white" />
                            <span className="text-stone-500 font-medium">/{it.unidade}</span>
                          </div>
                          <div className="text-right">
                            <span className="text-[10px] block text-stone-400 font-mono">Unit: R$ {fmt(cpuValorUnit(it.insumos))}</span>
                            <span className="font-semibold text-stone-900 font-mono">Total: R$ {fmt(num(it.quantidade) * cpuValorUnit(it.insumos))}</span>
                          </div>
                          <button onClick={() => removerItemDaEtapa(e.id, it.id)} className="text-stone-300 hover:text-red-500 ml-2"><Trash2 size={14} /></button>
                        </div>
                      </div>

                      {/* Exibe a tabela de insumos apenas se o usuário expandir o item */}
                      {estaExpandido && (
                        <div className="mt-2 pt-2 border-t border-stone-100 transition-all">
                          <InsumoTable insumos={it.insumos} onChange={(novos) => mudarInsumosDoItem(e.id, it.id, novos)} catalogMap={catalogMap} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------- ABA PLANILHA DE BDI ---------------- */
function BdiTab({ bdi, setBdi, bdiCalc, grandTotal }) {
  const faturamentoDireto = !!bdi.faturamentoDireto;

  // Inicializa taxas de materiais se não existirem
  const bdiMats = bdi.materiais || {
    admCentral: 0,
    contabilidade: 0,
    contingenciamento: 0,
    custoFinanceiro: 0,
    lucro: 0,
    dasAnexoIV: 0,
    art: 0
  };

  const handleGeralChange = (campo, valor) => {
    setBdi(prev => ({ ...prev, [campo]: valor }));
  };

  const handleMatChange = (campo, valor) => {
    setBdi(prev => ({
      ...prev,
      materiais: { ...bdiMats, [campo]: valor }
    }));
  };

  // Função auxiliar para calcular taxas somadas ou BDI para o painel resumo
  const calcularFatorQualquer = (t) => {
    const ac = num(t.admCentral);
    const c = num(t.contabilidade);
    const co = num(t.contingenciamento);
    const cf = num(t.custoFinanceiro);
    const l = num(t.lucro);
    const das = num(t.dasAnexoIV);
    const art = num(t.art);

    const pv = das + art;
    const numerador = (1 + ac) * (1 + c) * (1 + co) * (1 + cf) * (1 + l);
    const denominador = 1 - pv;
    return denominador <= 0 ? 1 : numerador / denominador;
  };

  const bdiGeralRate = calcularFatorQualquer(bdi) - 1;
  const bdiMatRate = faturamentoDireto ? (calcularFatorQualquer(bdiMats) - 1) : bdiGeralRate;

  return (
    <div className="space-y-6">
      {/* Barra de controle superior */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 flex justify-between items-center flex-wrap gap-3">
        <div>
          <h3 className="font-semibold text-sm text-stone-800">Opções do Regime de Faturamento</h3>
          <p className="text-xs text-stone-400">Ative o BDI diferenciado se houver materiais faturados direto pelo fornecedor.</p>
        </div>
        <label className="flex items-center gap-2 bg-stone-50 border border-stone-200 px-3 py-1.5 rounded-md cursor-pointer select-none hover:bg-stone-100 transition-colors text-xs font-semibold text-stone-700">
          <input
            type="checkbox"
            checked={faturamentoDireto}
            onChange={(e) => setBdi(prev => ({ ...prev, faturamentoDireto: e.target.checked }))}
            className="w-4 h-4 accent-stone-900 rounded"
          />
          Habilitar Faturamento Direto (BDI Diferenciado para Materiais)
        </label>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className={`bg-white border border-stone-200 rounded-lg p-4 text-xs space-y-4 ${faturamentoDireto ? "lg:col-span-2" : "lg:col-span-2"}`}>
          <h3 className="font-semibold text-sm text-stone-800 border-b border-stone-100 pb-2">Composição Analítica do BDI</h3>
          
          <div className={`grid grid-cols-1 gap-6 ${faturamentoDireto ? "sm:grid-cols-2" : "sm:grid-cols-2"}`}>
            {/* GRUPO 1: BDI GERAL */}
            <div className="space-y-4">
              <h4 className="font-bold text-stone-700 uppercase text-[10px] bg-stone-100 px-2 py-1 rounded tracking-wide">
                {faturamentoDireto ? "1. Taxas Gerais (Serviços e MO)" : "Taxas Gerais / Padrão"}
              </h4>
              
              <div className="space-y-3">
                <h5 className="font-medium text-stone-400 uppercase text-[9px]">Administração e Riscos</h5>
                <BdiInput label="Administração Central" value={bdi.admCentral} onChange={(v) => handleGeralChange("admCentral", v)} />
                <BdiInput label="Contabilidade / Seguros" value={bdi.contabilidade} onChange={(v) => handleGeralChange("contabilidade", v)} />
                <BdiInput label="Contingenciamento" value={bdi.contingenciamento} onChange={(v) => handleGeralChange("contingenciamento", v)} />
                <BdiInput label="Custo Financeiro" value={bdi.custoFinanceiro} onChange={(v) => handleGeralChange("custoFinanceiro", v)} />
                
                <h5 className="font-medium text-stone-400 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                <BdiInput label="Lucro Real de Venda" value={bdi.lucro} onChange={(v) => handleGeralChange("lucro", v)} />
                <BdiInput label="DAS / Tributos (Anexo IV)" value={bdi.dasAnexoIV} onChange={(v) => handleGeralChange("dasAnexoIV", v)} />
                <BdiInput label="ART / Encargos Contrato" value={bdi.art} onChange={(v) => handleGeralChange("art", v)} />
              </div>

              <div className="pt-2 border-t border-stone-100 flex justify-between items-center text-[11px] font-bold text-stone-700">
                <span>Taxa BDI Geral:</span>
                <span className="font-mono bg-stone-100 text-stone-800 px-1.5 py-0.5 rounded">{fmt(bdiGeralRate * 100)}%</span>
              </div>
            </div>

            {/* GRUPO 2: BDI MATERIAIS (Só renderiza se a caixa estiver marcada) */}
            {faturamentoDireto && (
              <div className="space-y-4 border-l border-stone-100 pl-4 sm:pl-6">
                <h4 className="font-bold text-emerald-800 uppercase text-[10px] bg-emerald-50 px-2 py-1 rounded tracking-wide">
                  2. Taxas Exclusivas para Materiais
                </h4>
                
                <div className="space-y-3">
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px]">Administração e Riscos</h5>
                  <BdiInput label="Administração Central" value={bdiMats.admCentral} onChange={(v) => handleMatChange("admCentral", v)} />
                  <BdiInput label="Contabilidade / Seguros" value={bdiMats.contabilidade} onChange={(v) => handleMatChange("contabilidade", v)} />
                  <BdiInput label="Contingenciamento" value={bdiMats.contingenciamento} onChange={(v) => handleMatChange("contingenciamento", v)} />
                  <BdiInput label="Custo Financeiro" value={bdiMats.custoFinanceiro} onChange={(v) => handleMatChange("custoFinanceiro", v)} />
                  
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                  <BdiInput label="Lucro Real de Venda" value={bdiMats.lucro} onChange={(v) => handleMatChange("lucro", v)} />
                  <BdiInput label="DAS / Tributos (Anexo IV)" value={bdiMats.dasAnexoIV} onChange={(v) => handleMatChange("dasAnexoIV", v)} />
                  <BdiInput label="ART / Encargos Contrato" value={bdiMats.art} onChange={(v) => handleMatChange("art", v)} />
                </div>

                <div className="pt-2 border-t border-stone-100 flex justify-between items-center text-[11px] font-bold text-emerald-800">
                  <span>Taxa BDI Materiais:</span>
                  <span className="font-mono bg-emerald-50 text-emerald-800 px-1.5 py-0.5 rounded">{fmt(bdiMatRate * 100)}%</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* PAINEL DA DIREITA: RESUMO TOTALIZADOR */}
        <div className="bg-stone-900 text-stone-100 rounded-lg p-5 flex flex-col justify-between h-full min-h-[320px]">
          <div>
            <h3 className="font-semibold text-xs uppercase tracking-wider text-stone-400 mb-4">Resumo Geral de Fechamento</h3>
            <div className="space-y-3 text-xs">
              <div className="flex justify-between"><span className="text-stone-400">Custo Direto Base:</span><span className="font-mono">R$ {fmt(grandTotal)}</span></div>
              <div className="flex justify-between">
                <span className="text-stone-400">BDI Geral Aplicado:</span>
                <span className="font-mono text-stone-300">{fmt(bdiCalc.bdiRate * 100)}%</span>
              </div>
              {faturamentoDireto && (
                <div className="flex justify-between">
                  <span className="text-emerald-400">BDI Materiais Aplicado:</span>
                  <span className="font-mono text-emerald-300">{fmt(bdiCalc.bdiRateMateriais * 100)}%</span>
                </div>
              )}
              <div className="flex justify-between border-t border-stone-800 pt-2">
                <span className="text-stone-400">Total BDI (Rateio):</span>
                <span className="font-mono">R$ {fmt(bdiCalc.totalDiValor)}</span>
              </div>
            </div>
          </div>
          <div className="mt-6 pt-4 border-t border-stone-800 text-right">
            <span className="text-[10px] text-stone-400 block uppercase font-medium">Preço Final de Venda</span>
            <span className="text-2xl font-bold font-mono text-white">R$ {fmt(bdiCalc.valorVenda)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function BdiInput({ label, value, onChange }) {
  return (
    <div className="flex items-center justify-between gap-3">
      <span className="text-stone-600">{label}</span>
      <input type="number" step="any" value={value === 0 ? "" : num(value) * 100} onChange={(e) => onChange(e.target.value === "" ? 0 : num(e.target.value) / 100)} className="w-20 border border-stone-300 rounded px-2 py-1 text-right font-mono" placeholder="0.00" />
    </div>
  );
}

/* ---------------- ABA FECHAMENTO: PREÇO DE VENDA ---------------- */
function PrecoVenda({ etapas, FatorBdi, grandTotal, nomeProjeto }) {
  const exportarXls = () => {
    const wb = XLSX.utils.book_new();
    const rows = [
      [`Planilha de Preço de Venda — ${nomeProjeto || "Orçamento"}`],
      [`Fator BDI aplicado: ${FatorBdi.toLocaleString("pt-BR", { minimumFractionDigits: 4, maximumFractionDigits: 4 })}`],
      [],
      ["Etapa", "Serviço", "Qtd.", "Un.", "Custo Unit. (R$)", "Preço Venda Unit. (R$)", "Total Venda (R$)"],
    ];
    (etapas || []).forEach((e) => {
      (e.itens || []).forEach((it) => {
        const uCusto = cpuValorUnit(it.insumos);
        rows.push([e.nome, it.servico, num(it.quantidade), it.unidade, uCusto, uCusto * FatorBdi, num(it.quantidade) * uCusto * FatorBdi]);
      });
    });
    rows.push([]);
    rows.push(["", "", "", "", "", "TOTAL GERAL", grandTotal * FatorBdi]);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 20 }, { wch: 45 }, { wch: 8 }, { wch: 6 }, { wch: 20 }, { wch: 22 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws, "Preço de Venda");
    XLSX.writeFile(wb, "preco_de_venda.xlsx");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 space-y-4">
      <div className="border-b border-stone-100 pb-2 flex justify-between items-center">
        <h3 className="font-semibold text-sm text-stone-800">Planilha Sintética de Fechamento (Preço de Venda)</h3>
        <button onClick={exportarXls} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700">
          <Download size={13} /> Exportar .xlsx
        </button>
      </div>
      <div className="space-y-3">
        {etapas.map((e) => {
          const custoEtapa = e.itens.reduce((s, it) => s + num(it.quantidade) * cpuValorUnit(it.insumos), 0);
          return (
            <div key={e.id} className="border border-stone-100 rounded-lg overflow-hidden">
              <div className="bg-stone-50/50 px-4 py-2 flex justify-between text-xs font-semibold text-stone-700">
                <span>{e.nome}</span>
                <span className="font-mono">R$ {fmt(custoEtapa * FatorBdi)}</span>
              </div>
              <div className="divide-y divide-stone-50">
                {e.itens.map((it) => {
                  const uCusto = cpuValorUnit(it.insumos);
                  const totalVendaItem = num(it.quantidade) * (uCusto * FatorBdi);
                  return (
                    <div key={it.id} className="flex items-center justify-between gap-4 px-4 py-2 text-xs">
                      <span className="text-stone-700 truncate flex-1">{it.servico}</span>
                      <span className="text-stone-400 font-mono w-24 text-right">{it.quantidade} {it.unidade}</span>
                      <span className="text-stone-500 font-mono w-28 text-right">R$ {fmt(uCusto * FatorBdi)}/un.</span>
                      <span className="font-medium font-mono text-stone-900 w-28 text-right">R$ {fmt(totalVendaItem)}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
      <div className="pt-4 border-t border-stone-200 flex justify-end">
        <div className="text-right p-2">
          <span className="text-xs text-stone-400 block font-medium">Valor Total do Fechamento Comercial</span>
          <span className="text-xl font-bold font-mono text-stone-900">R$ {fmt(grandTotal * FatorBdi)}</span>
        </div>
      </div>
    </div>
  );
}
