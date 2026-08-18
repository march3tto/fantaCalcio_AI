import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Search, Plus, Trash2, Check, X, Wallet, ListChecks,
  LayoutGrid, Settings, RotateCcw, AlertTriangle, Undo2, Upload, Download, ChevronDown, FileSpreadsheet,
  Star, Users, BookOpen, Heart, Target, TrendingUp, TrendingDown, ShieldAlert, ThumbsUp, ThumbsDown, Mic, LayoutTemplate, Gavel, Trophy, Cpu
} from "lucide-react";
import guideDatabaseDefault from "./guideDatabase.json";

// guideDatabase parte dal file in bundle (src/guideDatabase.json) ma può essere
// aggiornato a runtime dalla tab Guida (upload di un JSON più recente): il binding
// è mutabile e l'eventuale aggiornamento viene salvato in localStorage così resta
// attivo anche dopo un refresh. Vedi GUIDA_OVERRIDE_KEY / useGuidaDatabase più sotto.
const GUIDA_OVERRIDE_KEY = "fantacalcio-guida-override";

function caricaGuidaOverride() {
  try {
    const raw = localStorage.getItem(GUIDA_OVERRIDE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (parsed && parsed.dati && typeof parsed.dati === "object") return parsed;
  } catch (e) {
    // localStorage non disponibile o dato corrotto: si ignora e si riparte dal bundle
  }
  return null;
}

const guidaOverrideIniziale = caricaGuidaOverride();
let guideDatabase = guidaOverrideIniziale ? guidaOverrideIniziale.dati : guideDatabaseDefault;

// ---------- Costanti dominio Fantacalcio Mantra ----------
const RUOLI = ["Por", "Dd", "Ds", "Dc", "B", "E", "M", "C", "W", "T", "A", "Pc"];

const RUOLO_LABEL = {
  Por: "Portiere", Dd: "Difensore Destro", Ds: "Difensore Sinistro", Dc: "Difensore Centrale",
  B: "Braccetto", E: "Esterno", M: "Mediano", C: "Centrocampista", W: "Ala",
  T: "Trequartista", A: "Attaccante", Pc: "Prima Punta",
};

const GRUPPO = {
  Por: "POR", Dd: "DIF", Ds: "DIF", Dc: "DIF", B: "DIF",
  E: "CEN", M: "CEN", C: "CEN", W: "CEN", T: "CEN",
  A: "ATT", Pc: "ATT",
};

const GRUPPO_LABEL = { POR: "Portieri", DIF: "Difensori", CEN: "Centrocampisti", ATT: "Attaccanti" };
const GRUPPO_ACCENT = { POR: "text-amber-600", DIF: "text-sky-600", CEN: "text-emerald-600", ATT: "text-rose-600" };
const GRUPPO_BORDER = { POR: "border-amber-500/40", DIF: "border-sky-500/40", CEN: "border-emerald-500/40", ATT: "border-rose-500/40" };
const GRUPPO_BG = { POR: "bg-amber-500/8", DIF: "bg-sky-500/8", CEN: "bg-emerald-500/8", ATT: "bg-rose-500/8" };

const MODULI = {
  "3-4-3": ["Dc", "Dc", "Dc", "E", "M", "C", "E", "W", "A", "Pc"],
  "3-4-1-2": ["Dc", "Dc", "Dc", "E", "M", "C", "E", "T", "Pc", "A"],
  "3-5-2": ["Dc", "Dc", "Dc", "E", "M", "C", "T", "E", "A", "Pc"],
  "4-3-3": ["Dd", "Dc", "Dc", "Ds", "M", "C", "T", "W", "Pc", "W"],
  "4-3-1-2": ["Dd", "Dc", "Dc", "Ds", "M", "C", "C", "T", "Pc", "A"],
  "4-4-2": ["Dd", "Dc", "Dc", "Ds", "E", "M", "C", "E", "A", "Pc"],
  "4-4-1-1": ["Dd", "Dc", "Dc", "Ds", "E", "M", "C", "E", "T", "Pc"],
  "4-5-1": ["Dd", "Dc", "Dc", "Ds", "E", "M", "C", "T", "E", "Pc"],
  "4-2-3-1": ["Dd", "Dc", "Dc", "Ds", "M", "C", "T", "W", "W", "Pc"],
  "4-1-4-1": ["Dd", "Dc", "Dc", "Ds", "M", "E", "C", "C", "E", "Pc"],
  "4-1-3-2": ["Dd", "Dc", "Dc", "Ds", "M", "C", "T", "C", "Pc", "A"],
  "5-3-2": ["Dd", "Dc", "Dc", "Dc", "Ds", "M", "C", "T", "A", "Pc"],
  "5-4-1": ["Dd", "Dc", "Dc", "Dc", "Ds", "E", "M", "C", "E", "Pc"],
  "5-2-3": ["Dd", "Dc", "Dc", "Dc", "Ds", "M", "C", "W", "Pc", "W"],
};

const DEFAULT_SPLIT = { POR: 5, DIF: 20, CEN: 30, ATT: 45 };

const DEFAULT_ALGORITMI = {
  // Valore Reale (FVM)
  defaultTitolare: 0.75,
  baseTitolarita: 0.70,
  coefTitolarita: 0.30,
  bonusMultiruolo: 0.05,
  bonusRigorista: 0.15,
  bonusPunizioni: 0.05,
  bonusAngoli: 0.05,
  misterValorizzato: 1.10,
  misterPenalizzato: 0.90,

  // Decisione Rilancio CPU
  urgenzaMinimo: 1.6,
  febbreAstaCoef: 0.12,
  febbreAstaCap: 0.5,
  bonusFasciaCoef: 0.06,
  probColpoDiTestaBase: 0.22,
  probColpoDiTestaFascia: 0.05,
  moltiplicatoreColpoDiTestaMin: 1.4,
  moltiplicatoreColpoDiTestaMax: 1.9,
};

// Metadati per la Tab "Algoritmi": raggruppano DEFAULT_ALGORITMI in sezioni
// leggibili, con range/step adatti a ciascun parametro per gli slider di modifica.
const ALGORITMI_GRUPPI = [
  {
    titolo: "Valore Reale (FVM)",
    descrizione: "Come calcoliamo il valore reale stimato di un giocatore a partire da quotazione/FVM, titolarità, ruoli multipli e giudizio del mister in Guida.",
    campi: [
      { chiave: "defaultTitolare", label: "Titolarità di default", desc: "Probabilità di titolarità usata quando il giocatore non è presente in Guida.", min: 0, max: 1, step: 0.05, decimali: 2 },
      { chiave: "baseTitolarita", label: "Base titolarità", desc: "Quota fissa del fattore titolarità, indipendente dalla probabilità.", min: 0, max: 1, step: 0.01, decimali: 2 },
      { chiave: "coefTitolarita", label: "Coefficiente titolarità", desc: "Peso della probabilità di titolarità sul fattore finale.", min: 0, max: 1, step: 0.01, decimali: 2 },
      { chiave: "bonusMultiruolo", label: "Bonus multiruolo", desc: "Bonus percentuale per ogni ruolo aggiuntivo oltre al primo.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "bonusRigorista", label: "Bonus rigorista", desc: "Bonus percentuale se il giocatore è il rigorista designato.", min: 0, max: 0.5, step: 0.01, decimali: 2 },
      { chiave: "bonusPunizioni", label: "Bonus punizioni", desc: "Bonus percentuale se il giocatore batte le punizioni.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "bonusAngoli", label: "Bonus angoli", desc: "Bonus percentuale se il giocatore batte i corner.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "misterValorizzato", label: "Moltiplicatore valorizzato", desc: "Moltiplicatore applicato ai giocatori segnalati come valorizzati dal mister in Guida.", min: 1, max: 1.5, step: 0.01, decimali: 2 },
      { chiave: "misterPenalizzato", label: "Moltiplicatore penalizzato", desc: "Moltiplicatore applicato ai giocatori segnalati come penalizzati dal mister in Guida.", min: 0.5, max: 1, step: 0.01, decimali: 2 },
    ],
  },
  {
    titolo: "Decisione Rilancio CPU",
    descrizione: "Parametri che guidano il comportamento delle squadre CPU (e dell'agente \"Tu\") nella Simulazione asta: quando rilanciano e con quale aggressività.",
    campi: [
      { chiave: "urgenzaMinimo", label: "Urgenza sotto minimo", desc: "Moltiplicatore di urgenza quando la squadra è sotto il minimo di giocatori richiesto per il reparto.", min: 1, max: 3, step: 0.1, decimali: 1 },
      { chiave: "febbreAstaCoef", label: "Coefficiente febbre d'asta", desc: "Quanto cresce l'aggressività per ogni concorrente attivo oltre i primi due.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "febbreAstaCap", label: "Tetto febbre d'asta", desc: "Aumento massimo di aggressività dovuto alla febbre d'asta.", min: 0, max: 1, step: 0.05, decimali: 2 },
      { chiave: "bonusFasciaCoef", label: "Bonus fascia/stelle", desc: "Bonus di valore percepito per ogni stella oltre la terza.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "probColpoDiTestaBase", label: "Prob. colpo di testa (base)", desc: "Probabilità base che una CPU faccia un rilancio a sorpresa ('colpo di testa').", min: 0, max: 1, step: 0.01, decimali: 2 },
      { chiave: "probColpoDiTestaFascia", label: "Prob. colpo di testa (per stella)", desc: "Incremento di probabilità di colpo di testa per ogni stella del giocatore.", min: 0, max: 0.3, step: 0.01, decimali: 2 },
      { chiave: "moltiplicatoreColpoDiTestaMin", label: "Colpo di testa · moltiplicatore min", desc: "Moltiplicatore minimo applicato al valore percepito durante un colpo di testa.", min: 1, max: 2, step: 0.05, decimali: 2 },
      { chiave: "moltiplicatoreColpoDiTestaMax", label: "Colpo di testa · moltiplicatore max", desc: "Moltiplicatore massimo applicato al valore percepito durante un colpo di testa.", min: 1, max: 3, step: 0.05, decimali: 2 },
    ],
  },
];

// Indicatori mostrati nella tab Squadre: non sono parametri modificabili ma numeri
// derivati dal Valore Reale calcolato coi parametri qui sopra — li spieghiamo nella
// tab Algoritmi per tenere insieme "come li calcoliamo" e "cosa vuol dire il numero".
const INDICATORI_FORZA_ROSA = [
  { label: "Indice forza rosa", desc: "Somma del Valore Reale stimato (vedi sezione qui sotto) di tutti i giocatori in rosa. È il numero grezzo che misura quanto vale la squadra secondo il modello." },
  { label: "#N forza", desc: "Etichetta accanto al nome della squadra: la sua posizione in classifica tra tutte le squadre ordinate per Indice forza rosa decrescente." },
  { label: "Efficienza spesa", desc: "Indice forza rosa diviso i crediti spesi, in percentuale. Sopra 100% vuol dire aver ottenuto più valore reale di quanto pagato; sotto 100% il contrario." },
  { label: "Miglior affare", desc: "Etichetta assegnata alla squadra con l'Efficienza spesa più alta: chi, a parità di crediti spesi, ha portato a casa il valore reale stimato maggiore." },
];

const STORAGE_KEY = "asta-mantra-2026-27";

// Limiti di rosa (validi per ogni squadra, compresa la propria)
const CAP_POR = 4;
const CAP_ALTRI = 40;

const uid = () => Math.random().toString(36).slice(2, 10);

// Avvia il download lato browser di un file testuale (CSV/JSON) generato al volo,
// tramite un link temporaneo con attributo "download": non richiede backend.
function scaricaFile(nomeFile, contenuto, mime = "text/csv;charset=utf-8;") {
  const blob = new Blob([contenuto], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeFile;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function dataFileOggi() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

function buildSlotsFromModulo(nomeModulo) {
  const outfield = MODULI[nomeModulo] || MODULI["3-4-3"];
  return [
    { id: uid(), ruolo: "Por", giocatoreId: null, prezzo: null },
    ...outfield.map((r) => ({ id: uid(), ruolo: r, giocatoreId: null, prezzo: null })),
  ];
}

// Crea le squadre partecipanti: la prima è sempre "Io" (isMia: true).
// rosa: [{ giocatoreId, ruolo, prezzo }] — la rosa reale acquistata, senza limite di slot tattici,
// vincolata solo dai limiti CAP_POR / CAP_ALTRI.
function buildSquadre(numPartecipanti, budgetTotale) {
  const n = Math.max(1, numPartecipanti || 1);
  const squadre = [{ id: uid(), nome: "Io", isMia: true, budgetTotale, rosa: [] }];
  for (let i = 2; i <= n; i++) {
    squadre.push({ id: uid(), nome: `Squadra ${i}`, isMia: false, budgetTotale, rosa: [] });
  }
  return squadre;
}

function contaRosa(rosa) {
  const por = rosa.filter((r) => r.ruolo === "Por").length;
  const altri = rosa.length - por;
  return { por, altri };
}

function puoAssegnare(squadra, ruolo) {
  if (!squadra) return false;
  const { por, altri } = contaRosa(squadra.rosa);
  return ruolo === "Por" ? por < CAP_POR : altri < CAP_ALTRI;
}

// ---------- Analisi ruoli/moduli di una rosa ----------
// Conta, per ciascun ruolo Mantra, quanti giocatori della rosa sono eleggibili
// (un giocatore multi-ruolo, es. Dc/Ds, viene conteggiato in entrambi i ruoli:
// è quello che conta per capire quali moduli puoi effettivamente schierare).
function contaRuoliEleggibili(rosa, gById) {
  const conteggio = Object.fromEntries(RUOLI.map((r) => [r, 0]));
  rosa.forEach((r) => {
    const g = gById[r.giocatoreId];
    const ruoli = g?.ruoli?.length ? g.ruoli : [r.ruolo];
    ruoli.forEach((ru) => { if (conteggio[ru] !== undefined) conteggio[ru] += 1; });
  });
  return conteggio;
}

// Massimo abbinamento bipartito (algoritmo di Kuhn) tra gli slot di un modulo
// e i giocatori disponibili, per verificare se un modulo è davvero copribile
// tenendo conto che ogni giocatore può occupare un solo slot.
function copriturSlot(slotRuoli, giocatoriRuoli) {
  const matchGiocatore = new Array(giocatoriRuoli.length).fill(-1);
  function provaAssegna(slotIdx, visitati) {
    for (let gi = 0; gi < giocatoriRuoli.length; gi++) {
      if (visitati[gi]) continue;
      if (!giocatoriRuoli[gi].includes(slotRuoli[slotIdx])) continue;
      visitati[gi] = true;
      if (matchGiocatore[gi] === -1 || provaAssegna(matchGiocatore[gi], visitati)) {
        matchGiocatore[gi] = slotIdx;
        return true;
      }
    }
    return false;
  }
  let assegnati = 0;
  for (let s = 0; s < slotRuoli.length; s++) {
    const visitati = new Array(giocatoriRuoli.length).fill(false);
    if (provaAssegna(s, visitati)) assegnati++;
  }
  return assegnati;
}

// Per ogni modulo standard verifica se la rosa lo copre interamente:
// serve almeno un portiere + un abbinamento completo dei 10 slot di movimento.
function moduliCopribili(rosa, gById) {
  const haPortiere = rosa.some((r) => {
    const g = gById[r.giocatoreId];
    return (g?.ruoli?.length ? g.ruoli : [r.ruolo]).includes("Por");
  });
  const giocatoriMovimento = rosa
    .map((r) => {
      const g = gById[r.giocatoreId];
      const ruoli = (g?.ruoli?.length ? g.ruoli : [r.ruolo]).filter((ru) => ru !== "Por");
      return ruoli.length ? ruoli : null;
    })
    .filter(Boolean);

  const risultati = {};
  Object.entries(MODULI).forEach(([nome, slotRuoli]) => {
    const assegnati = copriturSlot(slotRuoli, giocatoriMovimento);
    risultati[nome] = { coperto: haPortiere && assegnati === slotRuoli.length, assegnati, totale: slotRuoli.length };
  });
  return { haPortiere, risultati };
}

// ---------- Motore di valutazione: collega i giocatori della lista d'asta ----------
// alle informazioni della Guida (guideDatabase.json) per stimare un "valore reale"
// più affidabile della sola quotazione, tenendo conto di titolarità, ruolo di
// specialista (rigori/punizioni/angoli) e giudizio del mister (valorizzato/penalizzato).

function normalizza(str) {
  return String(str || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // rimuove accenti
    .replace(/[^a-z0-9\s]/g, "")
    .trim();
}

// Le chiavi di guideDatabase sono già slug delle squadre (es. "atalanta", "milan").
// Gestiamo qualche alias comune presente nei file di quotazione ufficiali.
const TEAM_ID_ALIASES = {
  "hellas verona": "verona", "verona": "verona",
  "inter": "inter", "internazionale": "inter",
  "ac milan": "milan", "milan": "milan",
  "as roma": "roma", "roma": "roma",
};

function teamIdDaNomeSquadra(nomeSquadra) {
  const n = normalizza(nomeSquadra);
  if (!n) return null;
  if (TEAM_ID_ALIASES[n]) return TEAM_ID_ALIASES[n];
  return n.replace(/\s+/g, "");
}

function trovaVoceGuida(nomeSquadra) {
  const teamId = teamIdDaNomeSquadra(nomeSquadra);
  if (!teamId) return null;
  const voce = guideDatabase[teamId];
  return voce && voce.modulo ? voce : null; // scarta stub vuoti (modulo:null)
}

// Cerca un giocatore per nome/cognome all'interno di tutte le sezioni della Guida
// per una squadra e restituisce le informazioni utili alla valutazione.
function trovaInfoGiocatoreInGuida(giocatore, guida) {
  if (!guida) return null;
  const target = normalizza(giocatore.nome);
  if (!target) return null;
  const somiglia = (nome) => {
    const n = normalizza(nome);
    if (!n) return false;
    return n === target || n.includes(target) || target.includes(n);
  };

  let probTitolare = null;
  let rivale = null; // avversario diretto nel ballottaggio, se presente
  let motivi = [];
  const info = {
    isRigorista: false, isPunizioni: false, isAngoli: false,
    isValorizzato: false, isPenalizzato: false, isGiovane: false, isScommessa: false,
  };

  if ((guida.titolari || []).some((t) => somiglia(t.nome))) probTitolare = 1;

  (guida.ballottaggi || []).forEach((b) => {
    if (somiglia(b.giocatoreA?.nome)) { probTitolare = (b.giocatoreA.percentuale ?? 50) / 100; rivale = b.giocatoreB || null; }
    if (somiglia(b.giocatoreB?.nome)) { probTitolare = (b.giocatoreB.percentuale ?? 50) / 100; rivale = b.giocatoreA || null; }
  });

  if ((guida.specialisti?.rigoristi || []).some((p) => somiglia(p.nome))) info.isRigorista = true;
  if ((guida.specialisti?.punizioni || []).some((p) => somiglia(p.nome))) info.isPunizioni = true;
  if ((guida.specialisti?.angoli || []).some((p) => somiglia(p.nome))) info.isAngoli = true;

  const valorizzato = (guida.valorizzati || []).find((p) => somiglia(p.nome));
  if (valorizzato) { info.isValorizzato = true; motivi = motivi.concat(valorizzato.motivi || []); }
  const penalizzato = (guida.penalizzati || []).find((p) => somiglia(p.nome));
  if (penalizzato) { info.isPenalizzato = true; motivi = motivi.concat(penalizzato.motivi || []); }

  if (guida.giovaneDaSeguire && somiglia(guida.giovaneDaSeguire.nome)) info.isGiovane = true;
  if (guida.scommessa && somiglia(guida.scommessa.nome)) info.isScommessa = true;

  if (probTitolare === null) return null; // non trovato in Guida: nessuna informazione aggiuntiva
  return { probTitolare, rivale, motivi, ...info };
}

// Valore reale stimato = quotazione/FVM corretta per titolarità, ruoli multipli,
// bonus da specialista sui piazzati e giudizio del mister sulla squadra.
function calcolaValoreReale(giocatore, guidaSquadra, algoritmi = DEFAULT_ALGORITMI) {
  const base = (giocatore.fvm && giocatore.fvm > 0 ? giocatore.fvm : giocatore.quotazione) || 1;
  const info = trovaInfoGiocatoreInGuida(giocatore, guidaSquadra);
  const probTitolare = info ? info.probTitolare : (algoritmi.defaultTitolare ?? 0.75); // default neutro se non presente in Guida
  const fattoreTitolarita = (algoritmi.baseTitolarita ?? 0.70) + (algoritmi.coefTitolarita ?? 0.30) * probTitolare;
  const bonusRuoli = 1 + (algoritmi.bonusMultiruolo ?? 0.05) * Math.max(0, (giocatore.ruoli || []).length - 1);
  const bonusSpecialista = 1 + (info?.isRigorista ? (algoritmi.bonusRigorista ?? 0.15) : 0) 
                             + (info?.isPunizioni ? (algoritmi.bonusPunizioni ?? 0.05) : 0) 
                             + (info?.isAngoli ? (algoritmi.bonusAngoli ?? 0.05) : 0);
  const fattoreMister = info?.isValorizzato ? (algoritmi.misterValorizzato ?? 1.10) 
                      : info?.isPenalizzato ? (algoritmi.misterPenalizzato ?? 0.90) : 1;
  const valore = base * fattoreTitolarita * bonusRuoli * bonusSpecialista * fattoreMister;
  return { valore: Math.round(valore * 10) / 10, info };
}

function valoreGiocatore(giocatore, algoritmi = DEFAULT_ALGORITMI) {
  const guida = trovaVoceGuida(giocatore.squadra);
  return calcolaValoreReale(giocatore, guida, algoritmi);
}

// Scheda Guida completa per un singolo giocatore (usata in Asta Live): oltre al
// valore reale, recupera la voce squadra e le info testuali (ballottaggio con
// avversario, motivi valorizzato/penalizzato, ecc.) da mostrare durante la chiamata.
function trovaGuidaGiocatore(giocatore) {
  const guida = trovaVoceGuida(giocatore.squadra);
  const info = trovaInfoGiocatoreInGuida(giocatore, guida);
  return { guida, info };
}

// ---------- Dettatura vocale: parsing numeri italiani ----------
// L'app non può usare un microfono continuo in ascolto (non è supportato da
// iOS Safari/WebKit): la soluzione è lasciare che l'utente usi il pulsante
// microfono nativo della tastiera iOS per dettare nel campo di ricerca,
// es. "Scamacca quaranta", e qui interpretiamo il testo dettato per
// selezionare automaticamente il giocatore e precompilare il prezzo.

const IT_UNITA = ["zero", "uno", "due", "tre", "quattro", "cinque", "sei", "sette", "otto", "nove"];
const IT_DECINE_TEENS = ["dieci", "undici", "dodici", "tredici", "quattordici", "quindici", "sedici", "diciassette", "diciotto", "diciannove"];
const IT_DECINE = { 2: "venti", 3: "trenta", 4: "quaranta", 5: "cinquanta", 6: "sessanta", 7: "settanta", 8: "ottanta", 9: "novanta" };

function numeroInParoleIt(n) {
  if (n === 0) return "zero";
  if (n === 1000) return "mille";
  if (n < 10) return IT_UNITA[n];
  if (n < 20) return IT_DECINE_TEENS[n - 10];
  if (n < 100) {
    const decina = IT_DECINE[Math.floor(n / 10)];
    const unita = n % 10;
    if (unita === 0) return decina;
    // elisione della vocale finale davanti a "uno" e "otto" (es. ventuno, ventotto)
    if (unita === 1 || unita === 8) return decina.slice(0, -1) + IT_UNITA[unita];
    return decina + IT_UNITA[unita];
  }
  // 100-999
  const centinaia = Math.floor(n / 100);
  const resto = n % 100;
  const parolaCentinaia = centinaia === 1 ? "cento" : IT_UNITA[centinaia] + "cento";
  return resto === 0 ? parolaCentinaia : parolaCentinaia + numeroInParoleIt(resto);
}

// Mappa parola-italiana -> numero, generata una sola volta (0-1000).
const IT_PAROLA_A_NUMERO = (() => {
  const mappa = new Map();
  for (let n = 0; n <= 1000; n++) mappa.set(numeroInParoleIt(n), n);
  return mappa;
})();

// Estrae dal testo dettato l'eventuale prezzo finale (in cifre o in parole) e
// restituisce separatamente la parte che dovrebbe essere il nome del giocatore.
// Esempi riconosciuti: "Scamacca quaranta" -> {nome:"Scamacca", prezzo:40}
//                       "Retegui 55"        -> {nome:"Retegui", prezzo:55}
// La dettatura iOS a volte duplica una parola quando c'è una breve pausa
// (es. "Scamacca quaranta" -> "ScamaccaScamacca quaranta" oppure "Scamacca Scamacca quaranta").
// Questa funzione ripulisce il testo dettato prima di interpretarlo:
// 1) rimuove la punteggiatura residua
// 2) collassa una parola raddoppiata attaccata (metà1 === metà2, es. "ScamaccaScamacca" -> "Scamacca")
// 3) rimuove token identici consecutivi (es. "Scamacca Scamacca" -> "Scamacca")
function ripulisciDettatura(testoGrezzo) {
  const senzaPunteggiatura = String(testoGrezzo || "").replace(/[.,!?]/g, "").trim();
  if (!senzaPunteggiatura) return "";
  // fase 1: collassa una singola parola raddoppiata e attaccata (es. "ScamaccaScamacca")
  let token = senzaPunteggiatura.split(/\s+/).map((tok) => {
    const len = tok.length;
    if (len >= 4 && len % 2 === 0) {
      const meta = len / 2;
      const a = tok.slice(0, meta).toLowerCase();
      const b = tok.slice(meta).toLowerCase();
      if (a === b) return tok.slice(0, meta);
    }
    return tok;
  });
  // fase 2: individua la più lunga sequenza di parole consecutive ripetuta due volte
  // (copre sia "Scamacca Scamacca" sia nomi doppi come "De Ketelaere De Ketelaere")
  for (let n = Math.floor(token.length / 2); n >= 1; n--) {
    for (let start = 0; start + 2 * n <= token.length; start++) {
      const a = token.slice(start, start + n).join(" ").toLowerCase();
      const b = token.slice(start + n, start + 2 * n).join(" ").toLowerCase();
      if (a === b) {
        token = [...token.slice(0, start), ...token.slice(start, start + n), ...token.slice(start + 2 * n)];
        return token.join(" ");
      }
    }
  }
  return token.join(" ");
}

// Interpreta un prezzo dettato a voce durante un rilancio: la tastiera può
// restituire sia cifre ("45") sia la parola ("quarantacinque") a seconda di
// come iOS interpreta la dettatura in un campo numerico.
function interpretaPrezzoDettato(testo) {
  const pulito = ripulisciDettatura(testo).toLowerCase();
  if (!pulito) return null;
  if (/^\d+$/.test(pulito)) return parseInt(pulito, 10);
  if (IT_PAROLA_A_NUMERO.has(pulito)) return IT_PAROLA_A_NUMERO.get(pulito);
  const senzaSpazi = pulito.replace(/\s+/g, "");
  if (IT_PAROLA_A_NUMERO.has(senzaSpazi)) return IT_PAROLA_A_NUMERO.get(senzaSpazi);
  return null;
}

function estraiPrezzoDaTesto(testoGrezzo) {
  const testo = ripulisciDettatura(testoGrezzo);
  if (!testo) return { nome: "", prezzo: null };
  const token = testo.split(/\s+/);
  const ultimo = (token[token.length - 1] || "").toLowerCase();

  if (/^\d+$/.test(ultimo)) {
    return { nome: token.slice(0, -1).join(" "), prezzo: parseInt(ultimo, 10) };
  }
  if (IT_PAROLA_A_NUMERO.has(ultimo)) {
    return { nome: token.slice(0, -1).join(" "), prezzo: IT_PAROLA_A_NUMERO.get(ultimo) };
  }
  // prova a combinare le ultime due parole (es. dettatura con spazio anomalo)
  if (token.length >= 2) {
    const combinato = (token[token.length - 2] + token[token.length - 1]).toLowerCase();
    if (IT_PAROLA_A_NUMERO.has(combinato)) {
      return { nome: token.slice(0, -2).join(" "), prezzo: IT_PAROLA_A_NUMERO.get(combinato) };
    }
  }
  return { nome: testo, prezzo: null };
}

const initialState = {
  setup: { budgetTotale: 1000, numPartecipanti: 10, modulo: "3-4-3", split: { ...DEFAULT_SPLIT } },
  slots: buildSlotsFromModulo("3-4-3"),
  giocatori: [], // {id, ruoli, nome, squadra, quotazione, note, stato: 'disponibile'|'preso_altri'|'mio', presoDa, preferito}
  squadre: buildSquadre(10, 1000),
  algoritmi: { ...DEFAULT_ALGORITMI },
};

// ---------- Persistenza ----------
function useAstaStorage() {
  const [state, setState] = useState(initialState);
  const [loaded, setLoaded] = useState(false);
  const saveTimer = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setState((prev) => {
            const merged = { ...prev, ...parsed };
            merged.algoritmi = { ...prev.algoritmi, ...parsed.algoritmi };
            return merged;
          });
        }
      } catch (e) {
        // nessun dato salvato ancora, si parte da zero
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await window.storage.set(STORAGE_KEY, JSON.stringify(state), false);
      } catch (e) {
        console.error("Errore salvataggio", e);
      }
    }, 400);
    return () => clearTimeout(saveTimer.current);
  }, [state, loaded]);

  return [state, setState, loaded];
}

// Gestisce l'aggiornamento a runtime di guideDatabase (upload dalla tab Guida) e la
// sua persistenza in localStorage. Le altre tab che leggono guideDatabase (Asta,
// Squadre, Simula) sono montate/smontate col cambio scheda in App, quindi leggono
// sempre dati freschi al mount successivo; solo GuidaTab resta montata mentre la si
// aggiorna, e usa "aggiornataIl" come dipendenza per ricalcolare la lista squadre
// senza perdere il proprio stato locale (messaggi, pannello aperto).
function useGuidaDatabase() {
  const [aggiornataIl, setAggiornataIl] = useState(guidaOverrideIniziale?.aggiornataIl || null);

  function aggiornaGuida(datiCaricati, modalita) {
    const nuovo = modalita === "sostituisci" ? datiCaricati : { ...guideDatabase, ...datiCaricati };
    guideDatabase = nuovo;
    const timestamp = new Date().toLocaleString("it-IT");
    try {
      localStorage.setItem(GUIDA_OVERRIDE_KEY, JSON.stringify({ dati: nuovo, aggiornataIl: timestamp }));
    } catch (e) {
      // storage pieno o non disponibile: l'aggiornamento resta comunque attivo per questa sessione
    }
    setAggiornataIl(timestamp);
  }

  function ripristinaGuida() {
    guideDatabase = guideDatabaseDefault;
    try { localStorage.removeItem(GUIDA_OVERRIDE_KEY); } catch (e) { /* ignora */ }
    setAggiornataIl(null);
  }

  return { aggiornataIl, aggiornaGuida, ripristinaGuida };
}

// ---------- App ----------
export default function App() {
  const [state, setState, loaded] = useAstaStorage();
  const { aggiornataIl: guidaAggiornataIl, aggiornaGuida, ripristinaGuida } = useGuidaDatabase();
  const [tab, setTab] = useState("asta");

  // Stato del giocatore "chiamato" nell'asta live: sollevato qui (invece che dentro
  // AstaTab) così sopravvive quando l'utente naviga verso altre schermate e poi torna
  // su "Asta" — altrimenti, essendo AstaTab montato/smontato in base a tab, si perdeva.
  const [astaQuery, setAstaQuery] = useState("");
  const [astaSelezionato, setAstaSelezionato] = useState(null);
  const [astaPrezzoAttuale, setAstaPrezzoAttuale] = useState("");
  const [astaModalitaAssegna, setAstaModalitaAssegna] = useState(null); // null | 'io' | 'altri'
  const [astaRuoloScelto, setAstaRuoloScelto] = useState("");
  const [astaSquadraScelta, setAstaSquadraScelta] = useState("");
  const [astaPrezzoFinale, setAstaPrezzoFinale] = useState("");
  const [astaFiltroRuolo, setAstaFiltroRuolo] = useState("TUTTI");
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmApplicaSim, setConfirmApplicaSim] = useState(false);
  const [simApplicaMsg, setSimApplicaMsg] = useState("");

  // Stato della Simulazione asta (agente CPU) — sollevato qui per lo stesso motivo
  // dello stato di Asta Live: sopravvive al cambio di scheda. È completamente
  // separato da `state`: non viene mai persistito né scritto sui dati veri dell'asta.
  // Anche "Tu" è un agente automatico: nessun rilancio manuale, si guarda la simulazione
  // e si valuta il risultato — le sue scelte seguono i tuoi preferiti, il modulo
  // personalizzato (tab Moduli) e le ripartizioni di budget (tab Setup).
  const [simStarted, setSimStarted] = useState(false);
  const [simSquadre, setSimSquadre] = useState([]);
  const [simPool, setSimPool] = useState([]);
  const [simAuction, setSimAuction] = useState(null);
  const [simLog, setSimLog] = useState([]);
  const [simCallerIdx, setSimCallerIdx] = useState(0);
  const [simRosaTarget, setSimRosaTarget] = useState(25);
  const [simPorMinimo, setSimPorMinimo] = useState(3);
  const [simInPausa, setSimInPausa] = useState(false);
  const simTimeoutRef = useRef(null);
  const SIM_UTENTE_ID = "tu";
  // Minimi per ruolo classico (Por/Dif/Cen/Att) derivati da obiettivo rosa + minimo
  // portieri: si ricalcolano solo quando cambia uno dei due parametri.
  const simMinimi = useMemo(() => simMinimiClassici(simRosaTarget, simPorMinimo), [simRosaTarget, simPorMinimo]);

  function simPushLog(testo) {
    setSimLog((l) => [...l.slice(-80), testo]);
  }

  // Usa le squadre REALI configurate in Setup: stesso numero, stessi nomi (comprese le
  // rinomine), stesso budget di ciascuna. Alle CPU (tutte tranne "Io") assegniamo a
  // rotazione un profilo di personalità; "Io" diventa l'agente "Tu".
  function avviaSimulazione() {
    const disponibili = giocatori.filter((g) => g.stato === "disponibile");
    const pool = simCalcolaFasce(disponibili.map((g) => ({ id: g.id, nome: g.nome, ruoli: g.ruoli, squadra: g.squadra, quotazione: g.quotazione, fvm: g.fvm })), algoritmi);

    // L'agente "Tu": personalità neutra sui reparti perché il peso lo danno già le
    // ripartizioni reali (setup.split); slotObiettivo viene dal modulo personalizzato
    // in tab Moduli (state.slots), preferitiIds dai giocatori segnati con la stella.
    const slotObiettivo = {};
    RUOLI.forEach((r) => { slotObiettivo[r] = 0; });
    slots.forEach((s) => { slotObiettivo[s.ruolo] = (slotObiettivo[s.ruolo] || 0) + 1; });
    const preferitiIds = new Set(giocatori.filter((g) => g.preferito && g.stato === "disponibile").map((g) => g.id));

    let indiceProfilo = 0;
    const squadreSimulate = squadre.map((s) => {
      if (s.isMia) {
        return {
          id: SIM_UTENTE_ID, nome: s.nome || "Tu", isUtente: true, accent: "text-emerald-400",
          personalita: { POR: 1, DIF: 1, CEN: 1, ATT: 1 }, aggressivita: 0.5,
          budgetTotale: s.budgetTotale, rosa: [], slotObiettivo, preferitiIds,
        };
      }
      const profilo = SIM_BOT_PROFILES[indiceProfilo % SIM_BOT_PROFILES.length];
      indiceProfilo++;
      return {
        id: s.id, nome: s.nome, isUtente: false, accent: profilo.accent,
        personalita: profilo.personalita, aggressivita: profilo.aggressivita,
        budgetTotale: s.budgetTotale, rosa: [],
      };
    });

    setSimSquadre(squadreSimulate);
    setSimPool(pool);
    setSimCallerIdx(0);
    setSimAuction(null);
    setSimInPausa(false);
    setSimApplicaMsg("");
    const notaPref = preferitiIds.size > 0 ? `${preferitiIds.size} preferiti` : "nessun preferito segnato";
    setSimLog([`🏟️ Simulazione avviata: ${squadreSimulate.length} squadre (budget ${setup.budgetTotale} cr. ciascuna), modulo ${setup.modulo}, ${notaPref}. Ordine: ${squadreSimulate.map((s) => s.nome).join(" → ")}`]);
    setSimStarted(true);
  }

  function resetSimulazione() {
    clearTimeout(simTimeoutRef.current);
    setSimStarted(false);
    setSimSquadre([]);
    setSimPool([]);
    setSimAuction(null);
    setSimLog([]);
    setSimCallerIdx(0);
    setSimInPausa(false);
    setSimApplicaMsg("");
  }

  function simAvanzaChiamata(squadreCorrenti, poolCorrente) {
    for (let step = 1; step <= squadreCorrenti.length; step++) {
      const idx = (simCallerIdx + step) % squadreCorrenti.length;
      if (!simRosaCompleta(squadreCorrenti[idx], simRosaTarget, poolCorrente, simMinimi)) {
        setSimCallerIdx(idx);
        return;
      }
    }
    const sottoObiettivo = squadreCorrenti.filter((s) => s.rosa.length < simRosaTarget);
    if (sottoObiettivo.length > 0) {
      simPushLog(`🏁 Simulazione conclusa: il mercato è finito prima che ${sottoObiettivo.map((s) => s.nome).join(", ")} raggiungesse l'obiettivo di ${simRosaTarget} giocatori.`);
    } else {
      simPushLog("🏁 Tutte le rose hanno raggiunto l'obiettivo (minimo portieri incluso). Simulazione conclusa!");
    }
  }

  function avviaAstaSimulata(giocatore, ruolo, chiamanteId, scommessa) {
    const eleggibili = simSquadre.filter((s) => simPuoComprare(s, ruolo, simRosaTarget)).map((s) => s.id);
    const participants = simSquadre.map((s) => s.id).filter((id) => eleggibili.includes(id) || id === chiamanteId);
    const chiamanteNome = simSquadre.find((s) => s.id === chiamanteId)?.nome;
    const tagScommessa = scommessa ? "🎲 scommessa low-cost: " : "📣 ";
    simPushLog(`${tagScommessa}${chiamanteNome} chiama ${giocatore.nome} (${ruolo}) — base 1 credito`);

    const cursorChiamante = participants.indexOf(chiamanteId);
    const startCursor = simProssimoAttore(participants, cursorChiamante, chiamanteId, new Set());

    if (startCursor === -1) {
      simFinalizzaVendita(chiamanteId, giocatore, ruolo, 1);
      return;
    }
    setSimAuction({ giocatore, ruolo, participants, leaderId: chiamanteId, prezzoAttuale: 1, passati: new Set(), cursor: startCursor });
  }

  function simRisolviAzione(attoreId, rilancia) {
    setSimAuction((prev) => {
      if (!prev) return prev;
      const attore = simSquadre.find((s) => s.id === attoreId);
      const passati = new Set(prev.passati);
      let leaderId = prev.leaderId;
      let prezzoAttuale = prev.prezzoAttuale;
      if (rilancia) {
        prezzoAttuale = prev.prezzoAttuale + 1;
        leaderId = attoreId;
        simPushLog(`💰 ${attore.nome} rilancia: ${prezzoAttuale} crediti`);
      } else {
        passati.add(attoreId);
        simPushLog(`🚫 ${attore.nome} passa`);
      }
      const nuovoCursor = simProssimoAttore(prev.participants, prev.cursor, leaderId, passati);
      if (nuovoCursor === -1) {
        simFinalizzaVendita(leaderId, prev.giocatore, prev.ruolo, prezzoAttuale);
        return null;
      }
      return { ...prev, leaderId, prezzoAttuale, passati, cursor: nuovoCursor };
    });
  }

  function simFinalizzaVendita(vincitoreId, giocatore, ruolo, prezzo) {
    setSimSquadre((prev) => {
      const aggiornate = prev.map((s) => s.id === vincitoreId ? { ...s, rosa: [...s.rosa, { giocatoreId: giocatore.id, ruolo, prezzo }] } : s);
      const vincitore = aggiornate.find((s) => s.id === vincitoreId);
      const marcatore = vincitore.isUtente && vincitore.preferitiIds?.has(giocatore.id) ? " ⭐ (era un tuo preferito)" : "";
      simPushLog(`🏆 ${vincitore.nome} si aggiudica ${giocatore.nome} per ${prezzo} crediti!${marcatore}`);
      setSimPool((prevPool) => {
        const nuovoPool = prevPool.filter((p) => p.id !== giocatore.id);
        simAvanzaChiamata(aggiornate, nuovoPool);
        return nuovoPool;
      });
      return aggiornate;
    });
  }

  // Motore: TUTTE le squadre agiscono da sole, "Tu" compresa — nessun rilancio manuale.
  // "Tu" usa le stesse funzioni delle CPU (simSceltaChiamata/simDecisioneRilancio), ma
  // la sua squadra porta con sé slotObiettivo (dal modulo in tab Moduli) e preferitiIds
  // (dai giocatori con la stella), che pesano le sue scelte — vedi simFattoreRuolo e
  // simValorePercepito. Un breve ritardo tra un'azione e l'altra rende leggibile la
  // cronaca; "Pausa" ferma il motore, "Salta alla fine" lo risolve all'istante.
  useEffect(() => {
    if (!simStarted || simInPausa) return;
    clearTimeout(simTimeoutRef.current);

    if (simAuction) {
      const attoreId = simAuction.participants[simAuction.cursor];
      const attore = simSquadre.find((s) => s.id === attoreId);
      const numAttivi = simAuction.participants.length - simAuction.passati.size;
      simTimeoutRef.current = setTimeout(() => {
        const rilancia = simDecisioneRilancio(attore, simAuction.giocatore, simAuction.ruolo, simAuction.prezzoAttuale, simRosaTarget, simMinimi, numAttivi, simPool, algoritmi);
        simRisolviAzione(attoreId, rilancia);
      }, 550);
      return;
    }

    const chiamante = simSquadre[simCallerIdx];
    if (!chiamante) return;
    if (simRosaCompleta(chiamante, simRosaTarget, simPool, simMinimi)) {
      simAvanzaChiamata(simSquadre, simPool);
      return;
    }
    simTimeoutRef.current = setTimeout(() => {
      const scelta = simSceltaChiamata(chiamante, simPool, setup, simRosaTarget, simMinimi, algoritmi);
      if (!scelta) { simAvanzaChiamata(simSquadre, simPool); return; }
      avviaAstaSimulata(scelta.giocatore, scelta.ruolo, chiamante.id, scelta.scommessa);
    }, 700);
    return () => clearTimeout(simTimeoutRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [simStarted, simInPausa, simAuction, simCallerIdx, simSquadre, simPool]);

  // Risolve l'intera simulazione rimasta in un colpo solo, in modo sincrono, senza
  // passare dai timeout: utile per saltare subito al risultato finale.
  function simSaltaAllaFine() {
    clearTimeout(simTimeoutRef.current);
    let squadre = simSquadre.map((s) => ({ ...s, rosa: [...s.rosa] }));
    let pool = [...simPool];
    let callerIdx = simCallerIdx;
    const logExtra = [];
    let auction = simAuction ? { ...simAuction, passati: new Set(simAuction.passati) } : null;
    let guardia = 0;

    function risolviAsta(a) {
      let corrente = a;
      while (corrente) {
        const attoreId = corrente.participants[corrente.cursor];
        const attore = squadre.find((s) => s.id === attoreId);
        const numAttivi = corrente.participants.length - corrente.passati.size;
        const rilancia = simDecisioneRilancio(attore, corrente.giocatore, corrente.ruolo, corrente.prezzoAttuale, simRosaTarget, simMinimi, numAttivi, pool, algoritmi);
        const passati = new Set(corrente.passati);
        let leaderId = corrente.leaderId, prezzoAttuale = corrente.prezzoAttuale;
        if (rilancia) { prezzoAttuale += 1; leaderId = attoreId; } else { passati.add(attoreId); }
        const nuovoCursor = simProssimoAttore(corrente.participants, corrente.cursor, leaderId, passati);
        if (nuovoCursor === -1) {
          squadre = squadre.map((s) => s.id === leaderId ? { ...s, rosa: [...s.rosa, { giocatoreId: corrente.giocatore.id, ruolo: corrente.ruolo, prezzo: prezzoAttuale }] } : s);
          const vincitore = squadre.find((s) => s.id === leaderId);
          logExtra.push(`🏆 ${vincitore.nome} si aggiudica ${corrente.giocatore.nome} per ${prezzoAttuale} crediti!`);
          pool = pool.filter((p) => p.id !== corrente.giocatore.id);
          corrente = null;
        } else {
          corrente = { ...corrente, leaderId, prezzoAttuale, passati, cursor: nuovoCursor };
        }
      }
    }

    if (auction) risolviAsta(auction);

    while (guardia++ < 5000) {
      if (squadre.every((s) => simRosaCompleta(s, simRosaTarget, pool, simMinimi))) break;
      let idx = -1;
      for (let step = 0; step < squadre.length; step++) {
        const cand = (callerIdx + step) % squadre.length;
        if (!simRosaCompleta(squadre[cand], simRosaTarget, pool, simMinimi)) { idx = cand; break; }
      }
      if (idx === -1) break;
      callerIdx = idx;
      const chiamante = squadre[callerIdx];
      const scelta = simSceltaChiamata(chiamante, pool, setup, simRosaTarget, simMinimi, algoritmi);
      if (!scelta) { callerIdx = (callerIdx + 1) % squadre.length; continue; }
      const eleggibili = squadre.filter((s) => simPuoComprare(s, scelta.ruolo, simRosaTarget)).map((s) => s.id);
      const participants = squadre.map((s) => s.id).filter((id) => eleggibili.includes(id) || id === chiamante.id);
      const cursorChiamante = participants.indexOf(chiamante.id);
      const startCursor = simProssimoAttore(participants, cursorChiamante, chiamante.id, new Set());
      if (startCursor === -1) {
        squadre = squadre.map((s) => s.id === chiamante.id ? { ...s, rosa: [...s.rosa, { giocatoreId: scelta.giocatore.id, ruolo: scelta.ruolo, prezzo: 1 }] } : s);
        logExtra.push(`🏆 ${chiamante.nome} si aggiudica ${scelta.giocatore.nome} per 1 credito!`);
        pool = pool.filter((p) => p.id !== scelta.giocatore.id);
      } else {
        risolviAsta({ giocatore: scelta.giocatore, ruolo: scelta.ruolo, participants, leaderId: chiamante.id, prezzoAttuale: 1, passati: new Set(), cursor: startCursor });
      }
    }

    const sottoObiettivo = squadre.filter((s) => s.rosa.length < simRosaTarget);
    logExtra.push(
      sottoObiettivo.length > 0
        ? `🏁 Simulazione risolta: il mercato è finito prima che ${sottoObiettivo.map((s) => s.nome).join(", ")} raggiungesse l'obiettivo di ${simRosaTarget} giocatori.`
        : "🏁 Simulazione risolta fino alla fine."
    );
    setSimSquadre(squadre);
    setSimPool(pool);
    setSimCallerIdx(callerIdx);
    setSimAuction(null);
    setSimLog((l) => [...l.slice(-80), ...logExtra]);
  }

  // Applica il risultato della simulazione ai dati VERI dell'asta: scrive le rose
  // simulate dentro state.squadre/state.giocatori (stessa forma prodotta da
  // assegnaGiocatore), così il risultato compare in Squadre e — per "Tu" — in Rosa.
  // Azione esplicita e confermata, perché modifica i dati reali: salta i giocatori
  // già assegnati nel frattempo o oltre i limiti di rosa, senza fermarsi.
  function applicaRisultatoSimulazione() {
    const ioId = state.squadre.find((s) => s.isMia)?.id;
    let giocatoriNuovi = state.giocatori;
    let squadreNuove = state.squadre;
    let applicati = 0, saltati = 0;

    simSquadre.forEach((simSquadra) => {
      const squadraRealeId = simSquadra.isUtente ? ioId : simSquadra.id;
      simSquadra.rosa.forEach((r) => {
        const giocatoreReale = giocatoriNuovi.find((g) => g.id === r.giocatoreId);
        const squadraReale = squadreNuove.find((s) => s.id === squadraRealeId);
        if (!giocatoreReale || giocatoreReale.stato !== "disponibile" || !squadraReale || !puoAssegnare(squadraReale, r.ruolo)) {
          saltati++;
          return;
        }
        giocatoriNuovi = giocatoriNuovi.map((g) => g.id === r.giocatoreId ? { ...g, stato: squadraReale.isMia ? "mio" : "preso_altri", presoDa: squadraRealeId } : g);
        squadreNuove = squadreNuove.map((s) => s.id === squadraRealeId ? { ...s, rosa: [...s.rosa, { giocatoreId: r.giocatoreId, ruolo: r.ruolo, prezzo: r.prezzo }] } : s);
        applicati++;
      });
    });

    setState((prev) => ({ ...prev, giocatori: giocatoriNuovi, squadre: squadreNuove }));
    setConfirmApplicaSim(false);
    setSimApplicaMsg(
      saltati > 0
        ? `Applicati ${applicati} giocatori. ${saltati} saltati (già assegnati nella tua asta reale o limite di rosa raggiunto).`
        : `Applicati ${applicati} giocatori alla tua asta reale.`
    );
  }

  const { setup, slots, giocatori, squadre, algoritmi = DEFAULT_ALGORITMI } = state;

  function updateAlgoritmi(patch) {
    setState((prev) => ({
      ...prev,
      algoritmi: { ...prev.algoritmi, ...patch },
    }));
  }

  const ioSquadra = useMemo(() => squadre.find((s) => s.isMia) || squadre[0], [squadre]);
  const rosaIo = ioSquadra.rosa;

  const budgetSpeso = useMemo(
    () => rosaIo.reduce((acc, r) => acc + (r.prezzo || 0), 0),
    [rosaIo]
  );
  const budgetResiduo = ioSquadra.budgetTotale - budgetSpeso;

  const spesoPerGruppo = useMemo(() => {
    const acc = { POR: 0, DIF: 0, CEN: 0, ATT: 0 };
    rosaIo.forEach((r) => { acc[GRUPPO[r.ruolo]] += (r.prezzo || 0); });
    return acc;
  }, [rosaIo]);

  // Non essendoci più slot tattici fissi, i "posti liberi" per reparto derivano
  // dai limiti di rosa: 4 portieri totali, 40 condivisi tra Dif/Cen/Att.
  const { por: porContatiIo, altri: altriContatiIo } = contaRosa(rosaIo);
  const slotLiberiPerGruppo = useMemo(() => ({
    POR: Math.max(0, CAP_POR - porContatiIo),
    DIF: Math.max(0, CAP_ALTRI - altriContatiIo),
    CEN: Math.max(0, CAP_ALTRI - altriContatiIo),
    ATT: Math.max(0, CAP_ALTRI - altriContatiIo),
  }), [porContatiIo, altriContatiIo]);

  const slotLiberiTotali = Math.max(0, CAP_POR - porContatiIo) + Math.max(0, CAP_ALTRI - altriContatiIo);

  function updateSetup(patch) {
    setState((prev) => {
      let squadre = prev.squadre;
      if (patch.budgetTotale !== undefined) {
        squadre = squadre.map((s) => ({ ...s, budgetTotale: patch.budgetTotale }));
      }
      if (patch.numPartecipanti !== undefined) {
        const n = Math.max(1, patch.numPartecipanti);
        if (n > squadre.length) {
          const extra = [];
          for (let i = squadre.length + 1; i <= n; i++) {
            extra.push({ id: uid(), nome: `Squadra ${i}`, isMia: false, budgetTotale: patch.budgetTotale ?? prev.setup.budgetTotale, rosa: [] });
          }
          squadre = [...squadre, ...extra];
        } else if (n < squadre.length) {
          const keep = [...squadre];
          while (keep.length > n) {
            const last = keep[keep.length - 1];
            if (!last.isMia && last.rosa.length === 0) keep.pop();
            else break; // non rimuovo squadre con giocatori già assegnati
          }
          squadre = keep;
        }
      }
      return { ...prev, setup: { ...prev.setup, ...patch }, squadre };
    });
  }

  function rinominaSquadra(squadraId, nome) {
    setState((prev) => ({
      ...prev,
      squadre: prev.squadre.map((s) => (s.id === squadraId ? { ...s, nome } : s)),
    }));
  }

  function togglePreferito(giocatoreId) {
    setState((prev) => ({
      ...prev,
      giocatori: prev.giocatori.map((g) => (g.id === giocatoreId ? { ...g, preferito: !g.preferito } : g)),
    }));
  }

  // Assegna un giocatore (mio o di un altro partecipante) alla rosa di una squadra,
  // rispettando i limiti 4 Por / 40 altri ruoli. Ritorna false se il limite è già raggiunto.
  function assegnaGiocatore(squadraId, giocatoreId, ruolo, prezzo) {
    let ok = true;
    setState((prev) => {
      const squadra = prev.squadre.find((s) => s.id === squadraId);
      if (!squadra || !puoAssegnare(squadra, ruolo)) { ok = false; return prev; }
      return {
        ...prev,
        giocatori: prev.giocatori.map((g) =>
          g.id === giocatoreId ? { ...g, stato: squadra.isMia ? "mio" : "preso_altri", presoDa: squadraId } : g
        ),
        squadre: prev.squadre.map((s) =>
          s.id === squadraId ? { ...s, rosa: [...s.rosa, { giocatoreId, ruolo, prezzo }] } : s
        ),
      };
    });
    return ok;
  }

  function annullaAssegnazione(squadraId, giocatoreId) {
    setState((prev) => ({
      ...prev,
      giocatori: prev.giocatori.map((g) => (g.id === giocatoreId ? { ...g, stato: "disponibile", presoDa: null } : g)),
      squadre: prev.squadre.map((s) =>
        s.id === squadraId ? { ...s, rosa: s.rosa.filter((r) => r.giocatoreId !== giocatoreId) } : s
      ),
    }));
  }

  function cambiaModulo(nomeModulo) {
    setState((prev) => ({
      ...prev,
      setup: { ...prev.setup, modulo: nomeModulo },
      slots: buildSlotsFromModulo(nomeModulo),
    }));
  }

  function cambiaRuoloSlot(slotId, nuovoRuolo) {
    setState((prev) => ({
      ...prev,
      slots: prev.slots.map((s) => (s.id === slotId ? { ...s, ruolo: nuovoRuolo } : s)),
    }));
  }

  function aggiungiGiocatore(g) {
    setState((prev) => ({ ...prev, giocatori: [...prev.giocatori, { ...g, id: uid(), stato: "disponibile", preferito: false }] }));
  }

  function rimuoviGiocatore(id) {
    setState((prev) => ({ ...prev, giocatori: prev.giocatori.filter((g) => g.id !== id) }));
  }

  // Formato riga testuale: Ruoli(separati da ; o ,);Nome;Squadra;Quotazione;Fvm(opzionale);Preferito(opzionale)
  // La colonna Preferito è lo stesso formato prodotto dall'esportazione CSV della
  // lista (bottone "Esporta CSV"): 1/true/si per i preferiti, così il file scaricato
  // si può reimportare qui e ritrovare le stelline già segnate.
  function importaCsv(testo, sostituisci) {
    const righe = testo.split("\n").map((r) => r.trim()).filter(Boolean);
    const nuovi = [];
    righe.forEach((riga) => {
      const parti = riga.split(";").map((p) => p.trim());
      if (parti.length < 4) return;
      const [ruoliRaw, nome, squadra, quotazione, fvm, preferitoRaw] = parti;
      const ruoli = ruoliRaw.split(/[,\/]/).map((r) => r.trim()).filter((r) => RUOLI.includes(r));
      if (ruoli.length === 0 || !nome) return;
      nuovi.push({
        id: uid(), ruoli, nome, squadra,
        quotazione: parseFloat((quotazione || "0").replace(",", ".")) || 0,
        fvm: fvm ? parseFloat(fvm.replace(",", ".")) || null : null,
        note: "", stato: "disponibile",
        preferito: ["1", "true", "si", "sì", "x"].includes((preferitoRaw || "").trim().toLowerCase()),
      });
    });
    setState((prev) => ({
      ...prev,
      giocatori: sostituisci ? nuovi : mergeGiocatori(prev.giocatori, nuovi),
    }));
    return nuovi.length;
  }

  // Importa da file Excel/CSV ufficiale (es. export di fantacalcio.it, foglio "Tutti").
  // Cerca la riga di intestazione vera (alcuni export hanno un titolo nella prima riga)
  // e riconosce le colonne standard: Id, R, RM, Nome, Squadra, Qt.A, Qt.A M, FVM, FVM M.
  function importaFile(rows, sostituisci) {
    if (!rows || rows.length < 2) return 0;

    let headerIdx = 0;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const cells = (rows[i] || []).map((c) => String(c || "").toLowerCase().trim());
      if (cells.some((c) => c === "nome" || c === "cognome" || c === "calciatore")) { headerIdx = i; break; }
    }
    const header = (rows[headerIdx] || []).map((h) => String(h || "").toLowerCase().trim());
    const findCol = (candidates) => header.findIndex((h) => candidates.some((c) => h === c || h.includes(c)));

    const idxRuolo = findCol(["rm", "ruolo m", "ruolo mantra", "r.m."]);
    const idxRuoloFallback = findCol(["r", "ruolo"]);
    const idxNome = findCol(["nome", "cognome", "calciatore"]);
    const idxSquadra = findCol(["squadra"]);
    const idxQuot = findCol(["qt.a m", "qta m", "quotazione m", "quotazione mantra"]);
    const idxQuotFallback = findCol(["qt.a", "qta", "quotazione"]);
    const idxFvm = findCol(["fvm m", "fantavalore m", "fvm mantra"]);
    const idxFvmFallback = findCol(["fvm", "fantavalore"]);

    const colRuolo = idxRuolo >= 0 ? idxRuolo : idxRuoloFallback;
    const colQuot = idxQuot >= 0 ? idxQuot : idxQuotFallback;
    const colFvm = idxFvm >= 0 ? idxFvm : idxFvmFallback;

    if (colRuolo < 0 || idxNome < 0) return 0;

    const nuovi = [];
    rows.slice(headerIdx + 1).forEach((r) => {
      const ruoliRaw = String(r[colRuolo] || "").trim();
      const nome = String(r[idxNome] || "").trim();
      if (!ruoliRaw || !nome) return;
      const ruoli = ruoliRaw.split(/[;,\/]/).map((x) => x.trim()).filter((x) => RUOLI.includes(x));
      if (ruoli.length === 0) return;
      nuovi.push({
        id: uid(),
        ruoli,
        nome,
        squadra: idxSquadra >= 0 ? String(r[idxSquadra] || "").trim() : "",
        quotazione: colQuot >= 0 ? parseFloat(r[colQuot]) || 0 : 0,
        fvm: colFvm >= 0 && r[colFvm] !== undefined && r[colFvm] !== "" ? parseFloat(r[colFvm]) || null : null,
        note: "", stato: "disponibile", preferito: false,
      });
    });
    setState((prev) => ({
      ...prev,
      giocatori: sostituisci ? nuovi : mergeGiocatori(prev.giocatori, nuovi),
    }));
    return nuovi.length;
  }

  function mergeGiocatori(esistenti, nuovi) {
    const chiave = (g) => (g.nome + "|" + g.squadra).toLowerCase();
    const mappa = new Map(esistenti.map((g) => [chiave(g), g]));
    nuovi.forEach((g) => mappa.set(chiave(g), { ...mappa.get(chiave(g)), ...g, id: mappa.get(chiave(g))?.id || g.id }));
    return Array.from(mappa.values());
  }

  async function resetTutto() {
    setState(initialState);
    try { await window.storage.delete(STORAGE_KEY, false); } catch (e) { }
    setConfirmReset(false);
    setTab("setup");
  }

  if (!loaded) {
    return (
      <div className="min-h-screen bg-inkbg flex items-center justify-center text-inkdim font-mono">
        Caricamento asta...
      </div>
    );
  }

  return (
    <div className="h-screen w-full overflow-hidden bg-inkbg text-ink flex flex-col font-sans" style={{ height: "100dvh" }}>
      {/* Header scoreboard */}
      <header className="shrink-0 z-20 bg-inkbg/95 backdrop-blur border-b border-line px-4 pt-4 pb-3">
        <div className="flex items-center justify-between">
          <div className="text-[10px] tracking-[0.2em] text-inkdim uppercase font-semibold">Asta Mantra · Serie A 26/27</div>
          <button
            onClick={() => setConfirmReset(true)}
            className="p-2 rounded-lg border border-line text-inkdim active:scale-95 transition"
            aria-label="Azzera asta"
          >
            <RotateCcw size={22} />
          </button>
        </div>
        {/* barra reparti (solo in Rosa) */}
        {tab === "rosa" && (
          <div className="flex gap-1.5 mt-3">
            {(["POR", "DIF", "CEN", "ATT"]).map((g) => {
              const alloc = Math.round((setup.budgetTotale * setup.split[g]) / 100);
              const residuo = alloc - spesoPerGruppo[g];
              return (
                <div key={g} className={`flex-1 rounded-md border ${GRUPPO_BORDER[g]} ${GRUPPO_BG[g]} px-2 py-1`}>
                  <div className={`text-[9px] uppercase tracking-wide font-bold ${GRUPPO_ACCENT[g]}`}>{g}</div>
                  <div className="font-mono text-sm font-bold tabular-nums">{residuo}</div>
                </div>
              );
            })}
          </div>
        )}
      </header>

      {/* Contenuto */}
      <main className="flex-1 overflow-y-auto px-4 pt-4 pb-4">
        {tab === "setup" && (
          <SetupTab
            setup={setup} updateSetup={updateSetup}
            squadre={squadre} rinominaSquadra={rinominaSquadra}
          />
        )}
        {tab === "moduli" && (
          <ModuliTab setup={setup} cambiaModulo={cambiaModulo} slots={slots} cambiaRuoloSlot={cambiaRuoloSlot} />
        )}
        {tab === "giocatori" && (
          <GiocatoriTab
            giocatori={giocatori} squadre={squadre} aggiungiGiocatore={aggiungiGiocatore} rimuoviGiocatore={rimuoviGiocatore}
            importaCsv={importaCsv} importaFile={importaFile} togglePreferito={togglePreferito}
          />
        )}
        {tab === "asta" && (
          <AstaTab
            giocatori={giocatori} squadre={squadre} ioSquadraId={ioSquadra.id} setup={setup}
            spesoPerGruppo={spesoPerGruppo} slotLiberiPerGruppo={slotLiberiPerGruppo}
            budgetResiduo={budgetResiduo} slotLiberiTotali={slotLiberiTotali}
            assegnaGiocatore={assegnaGiocatore}
            query={astaQuery} setQuery={setAstaQuery}
            selezionato={astaSelezionato} setSelezionato={setAstaSelezionato}
            prezzoAttuale={astaPrezzoAttuale} setPrezzoAttuale={setAstaPrezzoAttuale}
            modalitaAssegna={astaModalitaAssegna} setModalitaAssegna={setAstaModalitaAssegna}
            ruoloScelto={astaRuoloScelto} setRuoloScelto={setAstaRuoloScelto}
            squadraScelta={astaSquadraScelta} setSquadraScelta={setAstaSquadraScelta}
            prezzoFinale={astaPrezzoFinale} setPrezzoFinale={setAstaPrezzoFinale}
            filtroRuoloAsta={astaFiltroRuolo} setFiltroRuoloAsta={setAstaFiltroRuolo}
          />
        )}
        {tab === "rosa" && (
          <RosaTab giocatori={giocatori} ioSquadra={ioSquadra} budgetSpeso={budgetSpeso} budgetResiduo={budgetResiduo} annullaAssegnazione={annullaAssegnazione} setup={setup} updateSetup={updateSetup} />
        )}
        {tab === "squadre" && (
          <SquadreTab squadre={squadre} giocatori={giocatori} />
        )}
        {tab === "guida" && (
          <GuidaTab
            giocatori={giocatori} squadre={squadre}
            aggiornataIl={guidaAggiornataIl} aggiornaGuida={aggiornaGuida} ripristinaGuida={ripristinaGuida}
          />
        )}
        {tab === "simulazione" && (
          <SimulazioneTab
            giocatori={giocatori} setup={setup} squadreReali={squadre}
            simStarted={simStarted} simSquadre={simSquadre} simPool={simPool} simAuction={simAuction}
            simLog={simLog} simCallerIdx={simCallerIdx}
            simRosaTarget={simRosaTarget} setSimRosaTarget={setSimRosaTarget}
            simPorMinimo={simPorMinimo} setSimPorMinimo={setSimPorMinimo} simMinimi={simMinimi}
            simInPausa={simInPausa} setSimInPausa={setSimInPausa}
            avviaSimulazione={avviaSimulazione} resetSimulazione={resetSimulazione} simSaltaAllaFine={simSaltaAllaFine}
            richiediApplicaSimulazione={() => setConfirmApplicaSim(true)}
            simApplicaMsg={simApplicaMsg}
            vaiATab={setTab}
          />
        )}
        {tab === "algoritmi" && (
          <AlgoritmiTab algoritmi={algoritmi} updateAlgoritmi={updateAlgoritmi} />
        )}
      </main>

      {/* Nav inferiore, bloccata in fondo: fa parte del layout flex, non "fixed",
          così su iOS Safari non si sposta quando la barra degli indirizzi appare/scompare */}
      <nav
        className="shrink-0 bg-inkbg/95 backdrop-blur border-t border-line flex z-20 overflow-x-auto"
        style={{ paddingBottom: "env(safe-area-inset-bottom)" }}
      >
        {[
          { id: "setup", label: "Setup", icon: Settings },
          { id: "moduli", label: "Moduli", icon: LayoutTemplate },
          { id: "giocatori", label: "Giocatori", icon: ListChecks },
          { id: "asta", label: "Asta Live", icon: Search },
          { id: "rosa", label: "Rosa", icon: LayoutGrid },
          { id: "squadre", label: "Squadre", icon: Users },
          { id: "guida", label: "Guida", icon: BookOpen },
          { id: "simulazione", label: "Simula", icon: Gavel },
          { id: "algoritmi", label: "Algoritmi", icon: Cpu },
        ].map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            className={`flex-1 flex flex-col items-center gap-1 py-2.5 transition ${tab === id ? "text-emerald-400" : "text-inkdim"}`}
          >
            <Icon size={24} />
            <span className="text-[10px] font-semibold tracking-wide">{label}</span>
          </button>
        ))}
      </nav>

      {confirmReset && (
        <ConfirmModal
          title="Azzerare tutta l'asta?"
          text="Verranno cancellati budget, schema, lista giocatori, squadre partecipanti e rose acquistate. Operazione irreversibile."
          onCancel={() => setConfirmReset(false)}
          onConfirm={resetTutto}
        />
      )}
      {confirmApplicaSim && (
        <ConfirmModal
          title="Applicare il risultato della simulazione?"
          text="I giocatori acquistati nella simulazione verranno assegnati alle squadre reali (Squadre) e alla tua rosa (Rosa), come se l'asta fosse avvenuta davvero. I giocatori già presi nella tua asta reale non vengono toccati."
          onCancel={() => setConfirmApplicaSim(false)}
          onConfirm={applicaRisultatoSimulazione}
        />
      )}
    </div>
  );
}

// ---------- Setup Tab ----------
function SetupTab({ setup, updateSetup, squadre, rinominaSquadra }) {
  return (
    <div className="space-y-6">
      <Section title="Parametri asta">
        <div className="space-y-3">
          <Field label="Budget totale (cr)">
            <input
              type="number" value={setup.budgetTotale}
              onChange={(e) => updateSetup({ budgetTotale: parseInt(e.target.value) || 0 })}
              className="input-dark"
            />
          </Field>
          <Field label="Partecipanti">
            <input
              type="number" value={setup.numPartecipanti}
              onChange={(e) => updateSetup({ numPartecipanti: parseInt(e.target.value) || 0 })}
              className="input-dark"
            />
          </Field>
        </div>
        <p className="text-xs text-inkdim mt-2">
          Ogni rosa, compresa la tua, può avere al massimo <span className="text-ink font-semibold">{CAP_POR} portieri</span> e{" "}
          <span className="text-ink font-semibold">{CAP_ALTRI} giocatori</span> negli altri ruoli.
        </p>
      </Section>

      <Section title={`Squadre partecipanti (${squadre.length})`}>
        <div className="space-y-1.5">
          {squadre.map((s) => (
            <div key={s.id} className="flex items-center gap-2">
              {s.isMia && <span className="text-[10px] font-bold text-emerald-400 shrink-0 w-6">TU</span>}
              <input
                value={s.nome}
                onChange={(e) => rinominaSquadra(s.id, e.target.value)}
                className={`input-dark flex-1 ${s.isMia ? "border-emerald-400/40" : ""}`}
              />
            </div>
          ))}
        </div>
        <p className="text-xs text-inkdim mt-2">Rinomina le squadre con i nomi reali dei partecipanti: le userai per assegnare i giocatori presi da altri durante l'asta.</p>
      </Section>
    </div>
  );
}

// ---------- Moduli Tab (schema tattico di riferimento) ----------
function ModuliTab({ setup, cambiaModulo, slots, cambiaRuoloSlot }) {
  return (
    <div className="space-y-6">
      <Section title="Modulo di riferimento">
        <div className="flex flex-wrap gap-2">
          {Object.keys(MODULI).map((m) => (
            <button
              key={m}
              onClick={() => cambiaModulo(m)}
              className={`px-3 py-1.5 rounded-lg text-sm font-mono font-bold border transition ${setup.modulo === m ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
                }`}
            >
              {m}
            </button>
          ))}
        </div>
        <p className="text-xs text-inkdim mt-2">Schema puramente indicativo per la disposizione in campo: acquisti e limiti di rosa non dipendono da questi slot.</p>
      </Section>

      <Section title="Personalizza ruoli per slot (riferimento tattico)">
        <div className="grid grid-cols-2 gap-2">
          {slots.map((s, i) => (
            <div key={s.id} className="flex items-center justify-between bg-panel border border-line rounded-lg px-2 py-1.5">
              <span className="text-[11px] text-inkdim">Slot {i + 1}</span>
              <select
                value={s.ruolo}
                onChange={(e) => cambiaRuoloSlot(s.id, e.target.value)}
                className="bg-transparent text-sm font-mono font-bold text-ink outline-none"
              >
                {RUOLI.map((r) => <option key={r} value={r} className="bg-panel">{r}</option>)}
              </select>
            </div>
          ))}
        </div>
      </Section>
    </div>
  );
}

// ---------- Giocatori Tab ----------
function GiocatoriTab({ giocatori, squadre, aggiungiGiocatore, rimuoviGiocatore, importaCsv, importaFile, togglePreferito }) {
  const [ruoliSel, setRuoliSel] = useState(["Por"]);
  const [nome, setNome] = useState("");
  const [squadra, setSquadra] = useState("");
  const [quotazione, setQuotazione] = useState("");
  const [filtro, setFiltro] = useState("");
  const [filtroRuolo, setFiltroRuolo] = useState("TUTTI");
  const [soloPreferiti, setSoloPreferiti] = useState(false);
  const [csvText, setCsvText] = useState("");
  const [showImport, setShowImport] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const fileInputRef = useRef(null);

  const squadreById = Object.fromEntries((squadre || []).map((s) => [s.id, s]));

  function toggleRuoloSel(r) {
    setRuoliSel((prev) => prev.includes(r) ? prev.filter((x) => x !== r) : [...prev, r]);
  }

  function handleAdd() {
    if (!nome.trim() || ruoliSel.length === 0) return;
    aggiungiGiocatore({ ruoli: ruoliSel, nome: nome.trim(), squadra: squadra.trim(), quotazione: parseFloat(quotazione) || 0, fvm: null, note: "" });
    setNome(""); setSquadra(""); setQuotazione("");
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const nomeFoglio = wb.SheetNames.find((n) => n.toLowerCase().trim() === "tutti") || wb.SheetNames[0];
        const sheet = wb.Sheets[nomeFoglio];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const n = importaFile(rows, false);
        setImportMsg(n > 0 ? `Importati/aggiornati ${n} giocatori dal file` : "Non ho riconosciuto le colonne del file. Prova con l'incolla testo qui sotto.");
      } catch (err) {
        setImportMsg("Errore nella lettura del file.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }

  const filtrati = giocatori.filter((g) => {
    const matchTesto = (g.nome + g.squadra).toLowerCase().includes(filtro.toLowerCase());
    const matchRuolo = filtroRuolo === "TUTTI" || (g.ruoli || []).includes(filtroRuolo);
    const matchPreferiti = !soloPreferiti || g.preferito;
    return matchTesto && matchRuolo && matchPreferiti;
  });

  return (
    <div className="space-y-5">
      <Section title="Importa la tua lista">
        <p className="text-xs text-inkdim mb-2">
          Scarica il listone da fantacalcio.it (pagina Quotazioni → Esporta) o da un altro sito ufficiale, poi carica qui il file: riconosco automaticamente ruolo Mantra, quotazione e Fvm.
        </p>
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} className="hidden" />
        <button onClick={() => fileInputRef.current?.click()} className="btn-primary inline-flex items-center gap-1.5 text-xs px-3 py-1.5">
          <FileSpreadsheet size={17} /> Carica file Excel / CSV
        </button>
        {importMsg && <p className="text-xs text-emerald-400 mt-2">{importMsg}</p>}

        <button onClick={() => setShowImport((v) => !v)} className="mt-2 flex items-center gap-1.5 text-xs text-inkdim py-1.5">
          <Upload size={16} /> oppure incolla lista via testo
        </button>
        {showImport && (
          <div className="mt-2 space-y-2">
            <textarea
              value={csvText} onChange={(e) => setCsvText(e.target.value)}
              placeholder={"Ruoli;Nome;Squadra;Quotazione;Fvm\nPor;Maignan;Milan;15;50\nM;C;Modric;Milan;13;60"}
              rows={4} className="input-dark w-full font-mono text-xs"
            />
            <div className="flex gap-2">
              <button onClick={() => { const n = importaCsv(csvText, false); setImportMsg(`Aggiunti/aggiornati ${n} giocatori`); setCsvText(""); }} className="btn-secondary flex-1">Aggiungi</button>
              <button onClick={() => { const n = importaCsv(csvText, true); setImportMsg(`Lista sostituita: ${n} giocatori`); setCsvText(""); }} className="btn-secondary flex-1">Sostituisci tutto</button>
            </div>
          </div>
        )}
      </Section>

      <Section title="Aggiungi giocatore manualmente">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {RUOLI.map((r) => (
            <button
              key={r} onClick={() => toggleRuoloSel(r)}
              className={`px-2 py-1 rounded-md text-xs font-mono font-bold border transition ${ruoliSel.includes(r) ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
                }`}
            >{r}</button>
          ))}
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={quotazione} onChange={(e) => setQuotazione(e.target.value)} placeholder="Quotazione" type="number" className="input-dark" />
          <input value={squadra} onChange={(e) => setSquadra(e.target.value)} placeholder="Squadra" className="input-dark" />
          <input value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Nome giocatore" className="input-dark col-span-2" />
        </div>
        <button onClick={handleAdd} className="btn-primary inline-flex items-center gap-1.5 text-xs px-3 py-1.5">
          <Plus size={17} /> Aggiungi alla lista
        </button>
      </Section>

      <Section title={`Lista giocatori (${giocatori.length}) · ${giocatori.filter((g) => g.preferito).length} preferiti`}>
        <button
          onClick={() => {
            const righe = giocatori.map((g) => [
              (g.ruoli || []).join(","), g.nome, g.squadra || "",
              g.quotazione ?? 0, g.fvm ?? "", g.preferito ? "1" : "0",
            ].join(";"));
            scaricaFile(`fantacalcio-giocatori-${dataFileOggi()}.csv`, ["Ruoli;Nome;Squadra;Quotazione;Fvm;Preferito", ...righe].join("\n"));
          }}
          disabled={giocatori.length === 0}
          className="btn-secondary w-full mb-2 inline-flex items-center justify-center gap-1.5 text-xs px-3 py-1.5 disabled:opacity-40"
        >
          <Download size={16} /> Esporta CSV (con preferiti)
        </button>
        <div className="flex gap-2 mb-2">
          <input value={filtro} onChange={(e) => setFiltro(e.target.value)} placeholder="Cerca..." className="input-dark flex-1" />
          <button
            onClick={() => setSoloPreferiti((v) => !v)}
            aria-label="Solo preferiti"
            className={`px-2.5 rounded-lg border flex items-center justify-center shrink-0 ${soloPreferiti ? "bg-amber-400/15 border-amber-400 text-amber-300" : "border-line text-inkdim"}`}
          >
            <Star size={19} fill={soloPreferiti ? "currentColor" : "none"} />
          </button>
        </div>
        <div className="flex flex-wrap gap-1.5 mb-3">
          <button
            onClick={() => setFiltroRuolo("TUTTI")}
            className={`px-2.5 py-1.5 rounded-md text-xs font-mono font-bold border transition ${filtroRuolo === "TUTTI" ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
              }`}
          >Tutti</button>
          {RUOLI.map((r) => (
            <button
              key={r} onClick={() => setFiltroRuolo(r)}
              className={`px-2.5 py-1.5 rounded-md text-xs font-mono font-bold border transition ${filtroRuolo === r ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
                }`}
            >{r}</button>
          ))}
        </div>
        <div className="space-y-1.5 max-h-[50vh] overflow-y-auto">
          {filtrati.length === 0 && <p className="text-sm text-inkdim text-center py-6">Nessun giocatore. Carica un file o aggiungine uno a mano.</p>}
          {filtrati.map((g) => (
            <div key={g.id} className="flex items-center justify-between bg-panel border border-line rounded-lg px-3 py-2">
              <div className="flex items-center gap-2 min-w-0">
                <div className="flex gap-1 shrink-0">
                  {(g.ruoli || []).map((r) => (
                    <span key={r} className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${GRUPPO_BG[GRUPPO[r]]} ${GRUPPO_ACCENT[GRUPPO[r]]}`}>{r}</span>
                  ))}
                </div>
                <div className="min-w-0">
                  <div className="text-sm font-semibold truncate">{g.nome}</div>
                  <div className="text-[11px] text-inkdim truncate">
                    {g.squadra} {g.stato !== "disponibile" && `· ${g.stato === "mio" ? "in rosa" : `preso da ${squadreById[g.presoDa]?.nome || "altri"}`}`}
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <span className="font-mono text-sm">{g.quotazione}</span>
                <button onClick={() => togglePreferito(g.id)} className={g.preferito ? "text-amber-400 p-1" : "text-inkdim p-1"} aria-label="Preferito">
                  <Star size={17} fill={g.preferito ? "currentColor" : "none"} />
                </button>
                <button onClick={() => rimuoviGiocatore(g.id)} className="text-inkdim p-1"><Trash2 size={17} /></button>
              </div>
            </div>
          ))}
        </div>
      </Section>
    </div>
  );
}

// ---------- Asta Live Tab ----------
function QuickBidButtons({ value, onChange }) {
  const incrementa = (delta) => {
    const attuale = parseInt(value, 10) || 0;
    onChange(String(Math.max(0, attuale + delta)));
  };
  return (
    <div className="flex gap-1.5">
      {[1, 5, 10].map((d) => (
        <button
          key={d}
          type="button"
          onClick={() => incrementa(d)}
          className="flex-1 py-1.5 rounded-md border border-line bg-inkbg text-sm font-bold text-emerald-400 active:bg-panelhover"
        >
          +{d}
        </button>
      ))}
    </div>
  );
}

// Nessun giocatore dovrebbe assorbire più di questa quota del budget totale della lega,
// indipendentemente da quanto sembri conveniente: è il tetto assoluto del Max Consigliato.
const TETTO_PERCENTUALE_BUDGET = 0.25;

// Quanti slot di un dato ruolo Mantra prevede il modulo tattico target (usato solo
// come riferimento strategico in asta: NON limita gli acquisti, che restano vincolati
// solo da CAP_POR/CAP_ALTRI — vedi contaRosa/puoAssegnare).
function ruoloSlotCountInModulo(modulo, ruolo) {
  if (ruolo === "Por") return 1;
  const outfield = MODULI[modulo] || MODULI["3-4-3"];
  return outfield.filter((r) => r === ruolo).length;
}

// Classifica il prezzo attuale del rilancio rispetto al budget max consigliato,
// per l'alert di convenienza a colori (decisione in meno di 2 secondi).
function classificaPrezzo(prezzo, maxConsigliato) {
  if (!prezzo || prezzo <= 0 || !maxConsigliato) return null;
  if (prezzo <= maxConsigliato) return "verde";
  if (prezzo <= Math.round(maxConsigliato * 1.4)) return "giallo";
  return "rosso";
}

const LIVELLO_CONFIG = {
  verde: { label: "AFFARE", bg: "bg-emerald-400/15", border: "border-emerald-400", text: "text-emerald-300", pulse: "" },
  giallo: { label: "IN TARGET", bg: "bg-amber-400/15", border: "border-amber-400", text: "text-amber-300", pulse: "" },
  rosso: { label: "OVERBIDDING", bg: "bg-rose-400/15", border: "border-rose-400", text: "text-rose-300", pulse: "animate-pulse" },
};

function AstaTab({
  giocatori, squadre, ioSquadraId, setup, spesoPerGruppo, slotLiberiPerGruppo, budgetResiduo, slotLiberiTotali, assegnaGiocatore,
  query, setQuery, selezionato, setSelezionato, prezzoAttuale, setPrezzoAttuale,
  modalitaAssegna, setModalitaAssegna, ruoloScelto, setRuoloScelto,
  squadraScelta, setSquadraScelta, prezzoFinale, setPrezzoFinale,
  filtroRuoloAsta, setFiltroRuoloAsta,
}) {
  const inputRef = useRef(null);

  const ioSquadra = squadre.find((s) => s.id === ioSquadraId);
  const altreSquadre = squadre.filter((s) => s.id !== ioSquadraId);

  // Se nel testo (digitato o dettato) c'è un numero finale, lo separiamo dal nome
  // così la ricerca funziona anche mentre l'utente sta ancora dettando il prezzo.
  const { nome: queryNomeParte, prezzo: queryPrezzoParte } = useMemo(() => estraiPrezzoDaTesto(query), [query]);

  const risultati = useMemo(() => {
    const base = (queryNomeParte || query).trim();
    const q = normalizza(base);
    return giocatori
      .filter((g) => g.stato === "disponibile")
      .filter((g) => filtroRuoloAsta === "TUTTI" || (g.ruoli || []).includes(filtroRuoloAsta))
      .filter((g) => !q || normalizza(g.nome).includes(q) || q.includes(normalizza(g.nome)))
      .slice(0, 8);
  }, [query, queryNomeParte, giocatori, filtroRuoloAsta]);

  function selezionaGiocatore(g, prezzoIniziale) {
    setSelezionato(g);
    setQuery(g.nome);
    setPrezzoAttuale(prezzoIniziale != null ? String(prezzoIniziale) : "");
    setModalitaAssegna(null);
    setRuoloScelto(g.ruoli[0]);
    setSquadraScelta(altreSquadre[0]?.id || "");
  }

  // Dettatura vocale: quando la tastiera (dettatura nativa iOS) produce un testo
  // tipo "Scamacca quaranta", riconosciamo automaticamente il giocatore e
  // precompiliamo il prezzo del rilancio attuale, senza bisogno di altri tocchi.
  // Il match è bidirezionale (nome-in-testo o testo-in-nome) e ignora gli accenti,
  // per assorbire eventuali imperfezioni residue della trascrizione vocale.
  useEffect(() => {
    if (selezionato) return;
    if (queryPrezzoParte === null || !queryNomeParte.trim()) return;
    const q = normalizza(queryNomeParte);
    if (!q) return;
    const match = giocatori.find((g) => {
      if (g.stato !== "disponibile") return false;
      const n = normalizza(g.nome);
      return n.includes(q) || q.includes(n);
    });
    if (match) selezionaGiocatore(match, queryPrezzoParte);
  }, [queryNomeParte, queryPrezzoParte, selezionato, giocatori]);

  const analisi = useMemo(() => {
    if (!selezionato) return null;
    const gruppo = GRUPPO[ruoloScelto] || GRUPPO[selezionato.ruoli[0]];
    const alloc = Math.round((setup.budgetTotale * setup.split[gruppo]) / 100);
    const residuoGruppo = alloc - spesoPerGruppo[gruppo];
    const slotLiberi = Math.max(1, slotLiberiPerGruppo[gruppo]);
    const budgetMedioSlot = residuoGruppo / slotLiberi;

    const stessGruppo = giocatori.filter((g) => g.ruoli.some((r) => GRUPPO[r] === gruppo) && g.stato === "disponibile");
    const quotazioneMedia = stessGruppo.length
      ? stessGruppo.reduce((a, g) => a + g.quotazione, 0) / stessGruppo.length
      : selezionato.quotazione || 1;
    const peso = quotazioneMedia > 0 ? selezionato.quotazione / quotazioneMedia : 1;

    let maxConsigliato = Math.round(budgetMedioSlot * peso);

    // Priorità di ruolo: quanto è urgente completare proprio QUESTO ruolo Mantra per il
    // modulo target, in base a quanti slot restano scoperti in rosa e a quanti giocatori
    // di quel ruolo sono ancora disponibili sul mercato (scarsità).
    const slotModuloRuolo = ruoloSlotCountInModulo(setup.modulo, ruoloScelto);
    const posseduti = ioSquadra
      ? ioSquadra.rosa.filter((r) => {
        const g = giocatori.find((gg) => gg.id === r.giocatoreId);
        const ruoli = g?.ruoli?.length ? g.ruoli : [r.ruolo];
        return ruoli.includes(ruoloScelto);
      }).length
      : 0;
    const slotRuoloLiberi = Math.max(0, slotModuloRuolo - posseduti);
    const disponibiliRuolo = giocatori.filter((g) => g.stato === "disponibile" && g.ruoli.includes(ruoloScelto)).length;

    let fattoreRuolo = 1;
    if (slotModuloRuolo === 0) fattoreRuolo = 0.6; // ruolo fuori dal modulo target: non prioritario
    else if (slotRuoloLiberi === 0) fattoreRuolo = 0.7; // ruolo già coperto: sarebbe solo un backup
    else if (disponibiliRuolo <= slotRuoloLiberi * 2) fattoreRuolo = 1.15; // slot scoperti e pochi disponibili: urgenza
    maxConsigliato = Math.round(maxConsigliato * fattoreRuolo);

    // Riserva dinamica per gli altri slot ancora da riempire: più severa a inizio asta
    // (rosa quasi vuota, serve prudenza), più leggera quando restano pochi slot totali.
    const slotTotali = CAP_POR + CAP_ALTRI;
    const percentualeRosaCompleta = 1 - slotLiberiTotali / slotTotali;
    const altriSlotLiberi = Math.max(0, slotLiberiTotali - 1);
    const riserva = Math.round(altriSlotLiberi * (1 + (1 - percentualeRosaCompleta) * 2));

    // Tetto assoluto: nessun giocatore dovrebbe assorbire più del 25% del budget totale.
    const tettoAssoluto = Math.round(setup.budgetTotale * TETTO_PERCENTUALE_BUDGET);

    maxConsigliato = Math.min(maxConsigliato, tettoAssoluto, budgetResiduo - riserva);
    maxConsigliato = Math.max(1, maxConsigliato);

    let affidabilita = 100;
    const prezzo = parseFloat(prezzoAttuale);
    if (prezzo > 0) affidabilita = Math.max(0, Math.min(100, Math.round((maxConsigliato / prezzo) * 100)));

    return {
      gruppo, residuoGruppo, slotLiberi: slotLiberiPerGruppo[gruppo], maxConsigliato, affidabilita, haPrezzo: prezzo > 0,
      fattoreRuolo, slotRuoloLiberi, tettoAssoluto,
    };
  }, [selezionato, ruoloScelto, setup, spesoPerGruppo, slotLiberiPerGruppo, giocatori, budgetResiduo, slotLiberiTotali, prezzoAttuale, ioSquadra]);

  const valoreSelezionato = useMemo(() => (selezionato ? valoreGiocatore(selezionato) : null), [selezionato]);

  const gById = useMemo(() => Object.fromEntries(giocatori.map((g) => [g.id, g])), [giocatori]);

  // Impatto sulla rosa Mantra: quanti slot del ruolo scelto prevede il modulo target,
  // quanti sono già coperti da giocatori in rosa, e chi sarebbe il "titolare" attuale
  // di cui questo giocatore diventerebbe backup/sostituto diretto.
  const impattoRosa = useMemo(() => {
    if (!selezionato || !ruoloScelto || !ioSquadra) return null;
    const slotModulo = ruoloSlotCountInModulo(setup.modulo, ruoloScelto);
    const posseduti = ioSquadra.rosa.filter((r) => {
      const g = gById[r.giocatoreId];
      const ruoli = g?.ruoli?.length ? g.ruoli : [r.ruolo];
      return ruoli.includes(ruoloScelto);
    });
    const nomiPosseduti = posseduti.map((r) => gById[r.giocatoreId]?.nome).filter(Boolean);
    const liberi = Math.max(0, slotModulo - posseduti.length);
    return { slotModulo, posseduti: posseduti.length, nomiPosseduti, liberi };
  }, [selezionato, ruoloScelto, ioSquadra, gById, setup.modulo]);

  const prezzoNum = parseFloat(prezzoAttuale) || 0;
  const livelloPrezzo = analisi ? classificaPrezzo(prezzoNum, analisi.maxConsigliato) : null;

  function confermaAssegnaIo() {
    if (!selezionato || !ruoloScelto || !ioSquadra) return;
    if (!puoAssegnare(ioSquadra, ruoloScelto)) return;
    const prezzo = parseInt(prezzoFinale) || 0;
    assegnaGiocatore(ioSquadra.id, selezionato.id, ruoloScelto, prezzo);
    reset();
  }

  function confermaAssegnaAltri() {
    if (!selezionato || !squadraScelta || !ruoloScelto) return;
    const squadra = squadre.find((s) => s.id === squadraScelta);
    if (!puoAssegnare(squadra, ruoloScelto)) return;
    const prezzo = parseInt(prezzoFinale) || 0;
    assegnaGiocatore(squadraScelta, selezionato.id, ruoloScelto, prezzo);
    reset();
  }

  function reset() {
    setSelezionato(null); setQuery(""); setPrezzoAttuale(""); setPrezzoFinale("");
    setModalitaAssegna(null); setRuoloScelto(""); setSquadraScelta("");
  }

  return (
    <div className="space-y-4">
      <div className="flex gap-2">
        <div className="relative flex-1">
          <Search size={19} className="absolute left-3 top-1/2 -translate-y-1/2 text-inkdim" />
          <input
            ref={inputRef}
            value={query}
            onChange={(e) => { setQuery(e.target.value); setSelezionato(null); }}
            placeholder="Nome giocatore chiamato all'asta..."
            className="input-dark w-full pl-9 pr-10"
            autoFocus
          />
          {query && !selezionato ? (
            <button onClick={reset} className="absolute right-3 top-1/2 -translate-y-1/2 text-inkdim"><X size={19} /></button>
          ) : !selezionato ? (
            <button
              onClick={() => inputRef.current?.focus()}
              aria-label="Detta col microfono della tastiera"
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-inkdim p-1"
            >
              <Mic size={19} />
            </button>
          ) : null}
        </div>
      </div>

      {selezionato && (
        <div>
          <label className="text-xs text-inkdim mb-1 block">Prezzo attuale del rilancio</label>
          <div className="relative mb-2">
            <input
              type="text" inputMode="numeric" value={prezzoAttuale}
              onChange={(e) => {
                const parlato = interpretaPrezzoDettato(e.target.value);
                setPrezzoAttuale(parlato !== null ? String(parlato) : e.target.value);
              }}
              placeholder="Inserisci man mano che l'asta sale, o detta col microfono" className="input-dark w-full pr-9"
            />
            <Mic size={18} className="absolute right-3 top-1/2 -translate-y-1/2 text-inkdim pointer-events-none" />
          </div>
          <QuickBidButtons value={prezzoAttuale} onChange={setPrezzoAttuale} />
        </div>
      )}

      {!selezionato && (
        <div className="flex flex-wrap gap-1.5 -mt-2">
          <button
            onClick={() => setFiltroRuoloAsta("TUTTI")}
            className={`px-2.5 py-1.5 rounded-md text-xs font-mono font-bold border transition ${filtroRuoloAsta === "TUTTI" ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
              }`}
          >Tutti</button>
          {RUOLI.map((r) => (
            <button
              key={r} onClick={() => setFiltroRuoloAsta(r)}
              className={`px-2.5 py-1.5 rounded-md text-xs font-mono font-bold border transition ${filtroRuoloAsta === r ? "bg-emerald-400/15 border-emerald-400 text-emerald-300" : "border-line text-inkdim"
                }`}
            >{r}</button>
          ))}
        </div>
      )}
      {!selezionato && (
        <p className="text-[11px] text-inkdim flex items-center gap-1.5">
          <Mic size={14} className="shrink-0" /> Tocca il campo e usa il microfono della tastiera: detta "nome prezzo", es. "Scamacca quaranta"
        </p>
      )}

      {!selezionato && risultati.length > 0 && (
        <div className="bg-panel border border-line rounded-lg divide-y divide-[#22352A] overflow-hidden">
          {risultati.map((g) => (
            <button key={g.id} onClick={() => selezionaGiocatore(g)} className={`w-full flex items-center justify-between px-3 py-2.5 active:bg-panelhover ${g.preferito ? "bg-amber-400/5" : ""}`}>
              <div className="flex items-center gap-2 min-w-0">
                {g.preferito && <Star size={16} className="text-amber-400 shrink-0" fill="currentColor" />}
                <div className="flex gap-1 shrink-0">
                  {g.ruoli.map((r) => (
                    <span key={r} className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${GRUPPO_BG[GRUPPO[r]]} ${GRUPPO_ACCENT[GRUPPO[r]]}`}>{r}</span>
                  ))}
                </div>
                <span className="text-sm font-semibold truncate">{g.nome}</span>
                <span className="text-[11px] text-inkdim truncate">{g.squadra}</span>
              </div>
              <span className="font-mono text-sm text-inkdim shrink-0">{g.quotazione}</span>
            </button>
          ))}
        </div>
      )}

      {!selezionato && (query || filtroRuoloAsta !== "TUTTI") && risultati.length === 0 && (
        <p className="text-sm text-inkdim text-center py-4">Nessun giocatore disponibile trovato nella tua lista.</p>
      )}

      {selezionato && analisi && (
        <div className="bg-panel border border-line rounded-xl p-4 space-y-4">
          {selezionato.preferito && (
            <div className="flex items-center gap-1.5 text-amber-300 bg-amber-400/10 border border-amber-400/40 rounded-lg px-2.5 py-1.5 text-xs font-bold">
              <Star size={16} fill="currentColor" /> È uno dei tuoi preferiti!
            </div>
          )}
          <div className="flex items-center justify-between">
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                {selezionato.ruoli.map((r) => (
                  <span key={r} className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${GRUPPO_BG[GRUPPO[r]]} ${GRUPPO_ACCENT[GRUPPO[r]]}`}>{r}</span>
                ))}
                <span className="font-bold text-lg truncate">{selezionato.nome}</span>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap text-xs text-inkdim mt-0.5">
                <span>{selezionato.squadra} · base {selezionato.quotazione}{selezionato.fvm ? ` · Fvm ${selezionato.fvm}` : ""}</span>
                {valoreSelezionato?.info ? (
                  <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${valoreSelezionato.info.probTitolare >= 1 ? "bg-emerald-400/15 text-emerald-300" : "bg-amber-400/15 text-amber-300"}`}>
                    {valoreSelezionato.info.probTitolare >= 1 ? "TITOLARE" : `BALLOTTAGGIO ${Math.round(valoreSelezionato.info.probTitolare * 100)}%`}
                  </span>
                ) : (
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-inkbg text-inkdim border border-line">TITOLARITÀ N.D.</span>
                )}
              </div>
            </div>
            <button onClick={reset} className="text-inkdim p-1 shrink-0"><X size={22} /></button>
          </div>

          {selezionato.ruoli.length > 1 && (
            <div className="flex items-center gap-1.5 flex-wrap -mt-2">
              <span className="text-[10px] text-inkdim uppercase shrink-0">Valuta per ruolo</span>
              {selezionato.ruoli.map((r) => (
                <button
                  key={r} type="button" onClick={() => setRuoloScelto(r)}
                  className={`text-[11px] font-bold px-2 py-1 rounded-md border transition ${ruoloScelto === r ? `${GRUPPO_BG[GRUPPO[r]]} ${GRUPPO_ACCENT[GRUPPO[r]]} border-current` : "border-line text-inkdim"
                    }`}
                >{r}</button>
              ))}
            </div>
          )}

          {livelloPrezzo ? (
            <div className={`flex items-center justify-between rounded-lg border px-3 py-2 ${LIVELLO_CONFIG[livelloPrezzo].bg} ${LIVELLO_CONFIG[livelloPrezzo].border} ${LIVELLO_CONFIG[livelloPrezzo].pulse}`}>
              <span className={`text-sm font-black tracking-wide ${LIVELLO_CONFIG[livelloPrezzo].text}`}>{LIVELLO_CONFIG[livelloPrezzo].label}</span>
              <span className={`text-[11px] font-mono ${LIVELLO_CONFIG[livelloPrezzo].text}`}>{analisi.affidabilita}% affidabilità</span>
            </div>
          ) : (
            <div className="rounded-lg border border-line px-3 py-2 text-center text-[11px] text-inkdim">
              Inserisci il prezzo attuale per l'alert di convenienza
            </div>
          )}

          <div className="grid grid-cols-3 gap-1.5">
            <div className="bg-inkbg rounded-lg p-2 border border-line">
              <div className="text-[9px] text-inkdim uppercase leading-tight">Max consigl.</div>
              <div className="font-mono text-lg font-black text-emerald-400">{analisi.maxConsigliato}</div>
              {analisi.maxConsigliato >= analisi.tettoAssoluto ? (
                <div className="text-[8px] text-amber-400 leading-tight mt-0.5">tetto 25% budget</div>
              ) : analisi.fattoreRuolo > 1 ? (
                <div className="text-[8px] text-emerald-400 leading-tight mt-0.5">ruolo scarso: priorità alta</div>
              ) : analisi.fattoreRuolo < 1 && analisi.slotRuoloLiberi === 0 ? (
                <div className="text-[8px] text-inkdim leading-tight mt-0.5">ruolo coperto: solo backup</div>
              ) : analisi.fattoreRuolo < 1 ? (
                <div className="text-[8px] text-inkdim leading-tight mt-0.5">fuori dal modulo</div>
              ) : null}
            </div>
            <div className={`rounded-lg p-2 border ${livelloPrezzo === "rosso" ? "bg-rose-400/10 border-rose-400" : livelloPrezzo === "giallo" ? "bg-amber-400/10 border-amber-400" : "bg-inkbg border-line"
              }`}>
              <div className="text-[9px] text-inkdim uppercase leading-tight">Prezzo attuale</div>
              <div className={`font-mono text-lg font-black ${livelloPrezzo === "rosso" ? "text-rose-400" : livelloPrezzo === "giallo" ? "text-amber-400" : livelloPrezzo === "verde" ? "text-emerald-400" : "text-ink"
                }`}>{prezzoNum > 0 ? prezzoNum : "–"}</div>
            </div>
            <div className="bg-inkbg rounded-lg p-2 border border-line">
              <div className="text-[9px] text-inkdim uppercase leading-tight">Residuo ({analisi.slotLiberi} lib.)</div>
              <div className="font-mono text-lg font-black">{analisi.residuoGruppo}</div>
            </div>
          </div>

          {valoreSelezionato && (
            <div className="flex items-center justify-between bg-inkbg border border-line rounded-lg px-3 py-2">
              <div>
                <div className="text-[10px] text-inkdim uppercase">
                  {valoreSelezionato.info ? "Valore reale stimato (da Guida)" : "Stima base (quotazione/FVM)"}
                </div>
                <div className={`font-mono text-lg font-bold ${valoreSelezionato.info ? "text-sky-400" : "text-inkdim"}`}>{valoreSelezionato.valore}</div>
              </div>
              <div className="flex flex-wrap gap-1 justify-end max-w-[55%]">
                {valoreSelezionato.info?.isValorizzato && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-emerald-400/15 text-emerald-300">VALORIZZATO</span>}
                {valoreSelezionato.info?.isPenalizzato && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-rose-400/15 text-rose-300">PENALIZZATO</span>}
                {valoreSelezionato.info?.isRigorista && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-amber-400/15 text-amber-300">RIGORISTA</span>}
                {valoreSelezionato.info === null && <span className="text-[9px] text-inkdim">non presente in Guida</span>}
              </div>
            </div>
          )}

          <GuidaGiocatoreCard giocatore={selezionato} />

          <PannelloInteresseAvversari
            altreSquadre={altreSquadre} gById={gById} ruolo={ruoloScelto}
            gruppo={analisi.gruppo} setup={setup} fattoreRuolo={analisi.fattoreRuolo}
          />

          {impattoRosa && (
            <div className="bg-inkbg border border-line rounded-lg p-2.5 space-y-1.5">
              <div className="flex items-center justify-between">
                <span className="flex items-center gap-1.5 text-[10px] text-inkdim uppercase"><Users size={14} /> Impatto rosa · {setup.modulo} · {ruoloScelto}</span>
                <span className={`font-mono text-sm font-bold ${impattoRosa.liberi > 0 ? "text-emerald-400" : "text-amber-400"}`}>
                  {impattoRosa.posseduti}/{impattoRosa.slotModulo}
                </span>
              </div>
              {impattoRosa.slotModulo === 0 ? (
                <p className="text-[11px] text-inkdim">Ruolo {ruoloScelto} non previsto nel modulo {setup.modulo}.</p>
              ) : impattoRosa.liberi > 0 ? (
                <p className="text-[11px] text-emerald-300">Slot {ruoloScelto} libero: completa il reparto.</p>
              ) : impattoRosa.nomiPosseduti.length > 0 ? (
                <p className="text-[11px] text-amber-300">Ruolo già coperto — sarebbe backup di {impattoRosa.nomiPosseduti.join(", ")}.</p>
              ) : (
                <p className="text-[11px] text-inkdim">Nessuno slot {ruoloScelto} nel modulo scelto.</p>
              )}
            </div>
          )}

          {!modalitaAssegna ? (
            <div className="flex flex-col gap-2 pt-1">
              <button onClick={() => { setModalitaAssegna("io"); setPrezzoFinale(prezzoAttuale); }} className="btn-primary w-full py-3.5 text-base flex items-center justify-center gap-2">
                <Check size={23} /> AGGIUDICATO A ME
              </button>
              <button onClick={() => { setModalitaAssegna("altri"); setPrezzoFinale(prezzoAttuale); }} className="btn-secondary w-full">
                Preso da un altro
              </button>
            </div>
          ) : modalitaAssegna === "io" ? (
            <div className="space-y-2 pt-1 border-t border-line">
              <label className="text-xs text-inkdim block pt-2">Ruolo con cui lo prendi</label>
              <select value={ruoloScelto} onChange={(e) => setRuoloScelto(e.target.value)} className="input-dark w-full">
                {selezionato.ruoli.map((r) => (
                  <option key={r} value={r} className="bg-panel">{r} · {RUOLO_LABEL[r]}</option>
                ))}
              </select>
              {!puoAssegnare(ioSquadra, ruoloScelto) && (
                <p className="text-xs text-rose-400 flex items-center gap-1.5">
                  <AlertTriangle size={16} /> Limite di rosa raggiunto ({ruoloScelto === "Por" ? `${CAP_POR} portieri` : `${CAP_ALTRI} giocatori negli altri ruoli`}).
                </p>
              )}
              <label className="text-xs text-inkdim block">Prezzo finale pagato</label>
              <div className="relative">
                <input
                  type="text" inputMode="numeric" value={prezzoFinale}
                  onChange={(e) => {
                    const parlato = interpretaPrezzoDettato(e.target.value);
                    setPrezzoFinale(parlato !== null ? String(parlato) : e.target.value);
                  }}
                  className="input-dark w-full pr-9" autoFocus
                />
                <Mic size={18} className="absolute right-3 top-1/2 -translate-y-1/2 text-inkdim pointer-events-none" />
              </div>
              <QuickBidButtons value={prezzoFinale} onChange={setPrezzoFinale} />
              <div className="flex gap-2 pt-1">
                <button onClick={confermaAssegnaIo} disabled={!puoAssegnare(ioSquadra, ruoloScelto)} className="btn-primary flex-1 disabled:opacity-40">Conferma acquisto</button>
                <button onClick={() => setModalitaAssegna(null)} className="btn-secondary px-4">Annulla</button>
              </div>
            </div>
          ) : (
            <div className="space-y-2 pt-1 border-t border-line">
              <label className="text-xs text-inkdim block pt-2">Squadra che se lo aggiudica</label>
              <select value={squadraScelta} onChange={(e) => setSquadraScelta(e.target.value)} className="input-dark w-full">
                {altreSquadre.map((s) => (
                  <option key={s.id} value={s.id} className="bg-panel">{s.nome}</option>
                ))}
              </select>
              <label className="text-xs text-inkdim block">Ruolo assegnato</label>
              <select value={ruoloScelto} onChange={(e) => setRuoloScelto(e.target.value)} className="input-dark w-full">
                {selezionato.ruoli.map((r) => (
                  <option key={r} value={r} className="bg-panel">{r} · {RUOLO_LABEL[r]}</option>
                ))}
              </select>
              {!puoAssegnare(squadre.find((s) => s.id === squadraScelta), ruoloScelto) && (
                <p className="text-xs text-rose-400 flex items-center gap-1.5">
                  <AlertTriangle size={16} /> Quella squadra ha già raggiunto il limite per questo ruolo.
                </p>
              )}
              <label className="text-xs text-inkdim block">Prezzo pagato</label>
              <div className="relative">
                <input
                  type="text" inputMode="numeric" value={prezzoFinale}
                  onChange={(e) => {
                    const parlato = interpretaPrezzoDettato(e.target.value);
                    setPrezzoFinale(parlato !== null ? String(parlato) : e.target.value);
                  }}
                  className="input-dark w-full pr-9" autoFocus
                />
                <Mic size={18} className="absolute right-3 top-1/2 -translate-y-1/2 text-inkdim pointer-events-none" />
              </div>
              <QuickBidButtons value={prezzoFinale} onChange={setPrezzoFinale} />
              <div className="flex gap-2 pt-1">
                <button onClick={confermaAssegnaAltri} disabled={!puoAssegnare(squadre.find((s) => s.id === squadraScelta), ruoloScelto)} className="btn-primary flex-1 disabled:opacity-40">Conferma</button>
                <button onClick={() => setModalitaAssegna(null)} className="btn-secondary px-4">Annulla</button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Simulazione Asta (agente CPU) ----------
// Motore di un'asta finta, giocabile in pratica, che usa la lista giocatori reale
// (schermata "Giocatori") e il motore ValoreReale già esistente (calcolaValoreReale)
// per far ragionare le CPU. Stato ENTIRELY separato da state.giocatori/state.squadre:
// la simulazione non scrive mai sui dati veri dell'asta, solo su una copia locale.

const SIM_BOT_PROFILES = [
  { nome: "CPU Bomber", accent: "text-rose-400", personalita: { POR: 0.8, DIF: 0.85, CEN: 1.0, ATT: 1.35 }, aggressivita: 0.64 },
  { nome: "CPU Diesse", accent: "text-sky-400", personalita: { POR: 1.0, DIF: 1.3, CEN: 1.05, ATT: 0.8 }, aggressivita: 0.6 },
  { nome: "CPU Equilibrata", accent: "text-emerald-400", personalita: { POR: 1, DIF: 1, CEN: 1, ATT: 1 }, aggressivita: 0.56 },
  { nome: "CPU Risparmio", accent: "text-violet-400", personalita: { POR: 0.7, DIF: 0.9, CEN: 0.9, ATT: 0.9 }, aggressivita: 0.4 },
  { nome: "CPU Azzardo", accent: "text-amber-400", personalita: { POR: 1.1, DIF: 1.1, CEN: 1.2, ATT: 1.3 }, aggressivita: 0.76 },
];

// Fasce di mercato in stile listone "vero": calcolate una volta sola all'avvio della
// simulazione, per percentile di ValoreReale DENTRO il proprio reparto (un portiere non
// compete con un attaccante per la fascia TOP). Restano fisse per tutta la simulazione,
// come un listone reale non cambia fascia a metà asta.
const SIM_FASCE = [
  { soglia: 0.08, nome: "TOP", stelle: 5, colore: "text-amber-300 bg-amber-400/15 border-amber-400/40" },
  { soglia: 0.25, nome: "SEMI-TOP", stelle: 4, colore: "text-violet-300 bg-violet-400/15 border-violet-400/40" },
  { soglia: 0.60, nome: "BUONA ROTAZIONE", stelle: 3, colore: "text-sky-300 bg-sky-400/15 border-sky-400/40" },
  { soglia: 0.85, nome: "RISERVA", stelle: 2, colore: "text-inkdim2 bg-white/5 border-line" },
  { soglia: 1.01, nome: "SCOMMESSA", stelle: 1, colore: "text-rose-300 bg-rose-400/15 border-rose-400/40" },
];

// Arricchisce il pool con fascia/stelle/prezzo medio/prezzo consigliato per ciascun
// giocatore. "Prezzo medio" = quotazione/FVM di listino (il prezzo "grezzo"); "prezzo
// consigliato" = ValoreReale già corretto da titolarità, ruoli multipli, specialisti e
// giudizio Guida — può essere più alto o più basso del medio, come nell'app di
// riferimento (screenshot: Pr. Medio 15 vs Prezzo Consigliato 18).
function simCalcolaFasce(pool, algoritmi = DEFAULT_ALGORITMI) {
  const perGruppo = { POR: [], DIF: [], CEN: [], ATT: [] };
  pool.forEach((p) => perGruppo[GRUPPO[p.ruoli[0]]].push(p));

  const risultato = new Map();
  Object.values(perGruppo).forEach((lista) => {
    const conValore = lista
      .map((p) => {
        const { valore } = valoreGiocatore(p, algoritmi);
        const base = (p.fvm && p.fvm > 0 ? p.fvm : p.quotazione) || 1;
        return { p, valore, base };
      })
      .sort((a, b) => b.valore - a.valore);
    const n = conValore.length;
    conValore.forEach((item, i) => {
      const percentile = n <= 1 ? 0 : i / n;
      const fasciaInfo = SIM_FASCE.find((f) => percentile < f.soglia) || SIM_FASCE[SIM_FASCE.length - 1];
      risultato.set(item.p.id, {
        ...item.p,
        fascia: fasciaInfo.nome,
        stelle: fasciaInfo.stelle,
        fasciaColore: fasciaInfo.colore,
        prezzoMedio: Math.max(1, Math.round(item.base)),
        prezzoConsigliato: Math.max(1, Math.round(item.valore)),
      });
    });
  });
  return pool.map((p) => risultato.get(p.id));
}

// Quanto è urgente, per QUESTA squadra, completare un certo ruolo Mantra — usato solo
// dall'agente "Tu": si basa sugli slot reali del modulo scelto (tab Moduli, con le tue
// personalizzazioni), non sul modulo di default. Le CPU non hanno slotObiettivo e quindi
// restano con fattore neutro (1): il loro comportamento non cambia.
function simFattoreRuolo(squadra, ruolo) {
  if (!squadra.slotObiettivo) return 1;
  const obiettivo = squadra.slotObiettivo[ruolo] || 0;
  const posseduti = squadra.rosa.filter((r) => r.ruolo === ruolo).length;
  if (obiettivo === 0) return 0.65; // ruolo fuori dal tuo modulo: non prioritario
  if (posseduti < obiettivo) return 1.25; // slot ancora scoperto nel tuo modulo
  return 0.75; // ruolo già coperto, sarebbe solo un backup
}

// Valore percepito di un giocatore da parte di questa squadra: ValoreReale, con un bonus
// se il giocatore è tra i tuoi preferiti (solo per l'agente "Tu", che porta preferitiIds).
function simValorePercepito(squadra, giocatore, algoritmi = DEFAULT_ALGORITMI) {
  const base = valoreGiocatore(giocatore, algoritmi).valore;
  const bonusPreferito = squadra.preferitiIds && squadra.preferitiIds.has(giocatore.id) ? 1.4 : 1;
  return base * bonusPreferito;
}

function simBudgetResiduo(squadra) {
  return squadra.budgetTotale - squadra.rosa.reduce((a, r) => a + (r.prezzo || 0), 0);
}
function simSlotsLiberi(squadra, rosaTarget) {
  const { por, altri } = contaRosa(squadra.rosa);
  const capResiduo = Math.max(0, CAP_POR - por) + Math.max(0, CAP_ALTRI - altri);
  const targetResiduo = Math.max(0, rosaTarget - squadra.rosa.length);
  return Math.min(capResiduo, targetResiduo);
}
function simPuoComprare(squadra, ruolo, rosaTarget) {
  const { por, altri } = contaRosa(squadra.rosa);
  if (ruolo === "Por" ? por >= CAP_POR : altri >= CAP_ALTRI) return false;
  return simSlotsLiberi(squadra, rosaTarget) > 0;
}
const SIM_RUOLO_RAPPRESENTATIVO = { POR: "Por", DIF: "Dc", CEN: "C", ATT: "A" };
// Quanti giocatori possiede già questa squadra in un dato reparto "classico".
function simPossedutiGruppo(squadra, gruppo) {
  if (gruppo === "POR") return contaRosa(squadra.rosa).por;
  return squadra.rosa.filter((r) => GRUPPO[r.ruolo] === gruppo).length;
}
// Minimi "da ruolo classico" (Portiere/Difensore/Centrocampista/Attaccante), come in
// un'asta reale: il minimo portieri lo decidi tu (simPorMinimo), gli altri tre reparti
// vengono derivati dalla proporzione classica 3-8-8-6 su 25, riscalata sull'obiettivo
// di rosa che hai impostato — con arrotondamento a somma esatta (metodo dei resti).
const SIM_RATIO_CLASSICA = { DIF: 8 / 22, CEN: 8 / 22, ATT: 6 / 22 };
function simMinimiClassici(rosaTarget, porMinimo) {
  // L'outfield usato per derivare i minimi non supera mai CAP_ALTRI: se l'obiettivo di
  // rosa digitato è molto alto, i minimi restano comunque un traguardo raggiungibile
  // (oltre CAP_ALTRI il motore satura comunque, vedi simSlotsLiberi).
  const outfield = Math.max(0, Math.min(rosaTarget - porMinimo, CAP_ALTRI));
  const grezzi = { DIF: outfield * SIM_RATIO_CLASSICA.DIF, CEN: outfield * SIM_RATIO_CLASSICA.CEN, ATT: outfield * SIM_RATIO_CLASSICA.ATT };
  const base = { DIF: Math.floor(grezzi.DIF), CEN: Math.floor(grezzi.CEN), ATT: Math.floor(grezzi.ATT) };
  let resto = outfield - (base.DIF + base.CEN + base.ATT);
  const ordineResti = ["DIF", "CEN", "ATT"].sort((a, b) => (grezzi[b] - base[b]) - (grezzi[a] - base[a]));
  for (let i = 0; resto > 0 && i < ordineResti.length; i++, resto--) base[ordineResti[i]] += 1;
  return { POR: porMinimo, DIF: Math.max(1, base.DIF), CEN: Math.max(1, base.CEN), ATT: Math.max(1, base.ATT) };
}
// Comprare (o rilanciare su) un giocatore di un certo reparto rischierebbe di lasciare
// questa squadra senza più spazio per i minimi ancora scoperti negli ALTRI reparti?
// Va controllato anche sui rilanci per giocatori chiamati DA ALTRE squadre, non solo
// sulle proprie chiamate: altrimenti una squadra può accumulare troppo in un reparto,
// esaurire gli slot e restare bloccata per sempre sotto ai minimi di un altro.
function simRischiaSaltoMinimo(squadra, pool, rosaTarget, minimi, gruppoInteressato) {
  const riservaAltri = ["POR", "DIF", "CEN", "ATT"]
    .filter((g) => g !== gruppoInteressato)
    .reduce((tot, g) => {
      const mancanti = Math.max(0, (minimi[g] || 0) - simPossedutiGruppo(squadra, g));
      if (mancanti === 0) return tot;
      // se il mercato ha già esaurito quel reparto, non ha senso continuare a riservargli spazio
      if (!pool.some((p) => p.ruoli.some((r) => GRUPPO[r] === g))) return tot;
      return tot + mancanti;
    }, 0);
  if (riservaAltri === 0) return false;
  return simSlotsLiberi(squadra, rosaTarget) - 1 < riservaAltri;
}
// C'è ancora, nel mercato residuo, almeno un giocatore che questa squadra potrebbe
// comprare (reparto non ancora al limite, non a rischio salto-minimo, e presente nel pool)?
function simMercatoDisponibilePer(squadra, pool, rosaTarget, minimi) {
  return ["POR", "DIF", "CEN", "ATT"].some((g) => {
    if (!simPuoComprare(squadra, SIM_RUOLO_RAPPRESENTATIVO[g], rosaTarget)) return false;
    if (simRischiaSaltoMinimo(squadra, pool, rosaTarget, minimi, g)) return false;
    return pool.some((p) => p.ruoli.some((r) => GRUPPO[r] === g));
  });
}
// Una rosa si considera completa quando raggiunge l'obiettivo (con i minimi per ruolo
// classico soddisfatti), OPPURE quando il mercato non ha più nulla che questa squadra
// possa ancora comprare — altrimenti, se rosaTarget non è raggiungibile con i giocatori
// realmente disponibili, il motore resterebbe bloccato a girare senza mai concludere.
function simRosaCompleta(squadra, rosaTarget, pool, minimi) {
  const minimiSoddisfatti = ["POR", "DIF", "CEN", "ATT"].every((g) => {
    if (simPossedutiGruppo(squadra, g) >= minimi[g]) return true;
    return !pool.some((p) => p.ruoli.some((r) => GRUPPO[r] === g)); // mercato esaurito per quel reparto
  });
  if (!minimiSoddisfatti) return !simMercatoDisponibilePer(squadra, pool, rosaTarget, minimi);
  const slotsFiniti = simSlotsLiberi(squadra, rosaTarget) <= 0;
  if (slotsFiniti) return true;
  return !simMercatoDisponibilePer(squadra, pool, rosaTarget, minimi);
}

// Una CPU sceglie un reparto: se uno dei quattro ruoli classici (Portiere, Difensore,
// Centrocampista, Attaccante) è ancora sotto il proprio minimo, la chiamata viene
// FORZATA lì (priorità al più scoperto, coi portieri sempre in testa per il loro cap
// ridotto) — regola dura, non pesata, perché in un'asta reale nessuno resta troppo a
// lungo sotto ai minimi di un reparto. Solo a minimi soddisfatti si sceglie il reparto
// pesando ripartizione budget reale + personalità. Quando la squadra è già avanti con
// gli acquisti "titolari" (o quel ruolo è già coperto per "Tu"), c'è una probabilità di
// pescare una "scommessa": un giocatore economico e non di primo piano, magari
// segnalato tale nella Guida, invece del migliore disponibile.
function simSceltaChiamata(squadra, pool, setup, rosaTarget, minimi, algoritmi = DEFAULT_ALGORITMI) {
  const gruppiSottoMinimo = ["POR", "DIF", "CEN", "ATT"].filter((g) => {
    if (simPossedutiGruppo(squadra, g) >= minimi[g]) return false;
    if (!simPuoComprare(squadra, SIM_RUOLO_RAPPRESENTATIVO[g], rosaTarget)) return false;
    return pool.some((p) => p.ruoli.some((r) => GRUPPO[r] === g));
  });

  let gruppo;
  const forzato = gruppiSottoMinimo.length > 0;
  if (forzato) {
    gruppo = gruppiSottoMinimo.sort((a, b) => {
      if (a === "POR" && b !== "POR") return -1;
      if (b === "POR" && a !== "POR") return 1;
      const mancA = minimi[a] - simPossedutiGruppo(squadra, a);
      const mancB = minimi[b] - simPossedutiGruppo(squadra, b);
      return mancB - mancA;
    })[0];
  } else {
    const gruppiValidi = ["POR", "DIF", "CEN", "ATT"].filter((g) => {
      if (!simPuoComprare(squadra, SIM_RUOLO_RAPPRESENTATIVO[g], rosaTarget)) return false;
      if (simRischiaSaltoMinimo(squadra, pool, rosaTarget, minimi, g)) return false;
      return pool.some((p) => p.ruoli.some((r) => GRUPPO[r] === g));
    });
    if (gruppiValidi.length === 0) return null;
    const fattoreGruppo = (g) => {
      if (!squadra.slotObiettivo) return 1;
      const ruoliDelGruppo = RUOLI.filter((r) => GRUPPO[r] === g);
      return ruoliDelGruppo.reduce((max, r) => Math.max(max, simFattoreRuolo(squadra, r)), 0.5);
    };
    const pesi = gruppiValidi.map((g) => (setup.split[g] || 1) * (squadra.personalita[g] || 1) * fattoreGruppo(g));
    const totale = pesi.reduce((a, b) => a + b, 0) || 1;
    let pick = Math.random() * totale;
    gruppo = gruppiValidi[0];
    for (let i = 0; i < gruppiValidi.length; i++) { pick -= pesi[i]; if (pick <= 0) { gruppo = gruppiValidi[i]; break; } }
  }

  const candidati = pool
    .filter((p) => p.ruoli.some((r) => GRUPPO[r] === gruppo))
    .map((p) => {
      const ruolo = p.ruoli.find((r) => GRUPPO[r] === gruppo) || p.ruoli[0];
      const { info } = valoreGiocatore(p, algoritmi);
      return { p, ruolo, valore: simValorePercepito(squadra, p, algoritmi) * simFattoreRuolo(squadra, ruolo), scommessaGuida: !!(info?.isScommessa || info?.isGiovane) };
    })
    .sort((a, b) => b.valore - a.valore);
  if (candidati.length === 0) return null;

  const residuo = simBudgetResiduo(squadra);
  const riserva = Math.max(0, simSlotsLiberi(squadra, rosaTarget) - 1);
  const maxSpendibile = residuo - riserva;

  // Squadra già "avanti" con gli acquisti, o ruolo già coperto per "Tu": possibile
  // acquisto di profondità a basso costo invece che del migliore sul mercato — ma solo
  // qualche volta, per non far restare troppo budget inutilizzato a fine simulazione.
  const giaAvanti = squadra.rosa.length >= Math.round(rosaTarget * 0.6);
  const ruoloGiaCoperto = squadra.slotObiettivo ? simFattoreRuolo(squadra, candidati[0].ruolo) <= 0.75 : false;
  const provaScommessa = !forzato && (giaAvanti || ruoloGiaCoperto) && Math.random() < 0.3;

  let fetta;
  let scommessa = false;
  if (provaScommessa) {
    const segnalati = candidati.filter((c) => c.scommessaGuida);
    const fasciaEconomica = candidati.slice(-Math.max(3, Math.ceil(candidati.length * 0.22)));
    fetta = segnalati.length > 0 && Math.random() < 0.5 ? segnalati : fasciaEconomica;
    scommessa = true;
  } else {
    let papabili = candidati.filter((c) => c.valore <= Math.max(1, maxSpendibile * 1.3));
    if (papabili.length === 0) papabili = candidati.slice(-6);
    fetta = papabili.slice(0, Math.max(3, Math.ceil(papabili.length * 0.35)));
  }
  const scelto = fetta[Math.floor(Math.random() * fetta.length)];
  if (!scelto) return null;
  return { giocatore: scelto.p, ruolo: scelto.ruolo, scommessa };
}

// Decisione di rilancio: quanto vale per questa squadra il giocatore rispetto a
// ValoreReale (più eventuale bonus preferiti/urgenza modulo per "Tu" e priorità
// assoluta se manca ancora il minimo di quel ruolo classico), budget e slot residui.
// Include la "febbre d'asta": più bidder sono ancora in corsa, più cresce la voglia di
// superare il valore stimato, e c'è una probabilità di "colpo di testa" che spinge il
// rilancio anche oltre il +40% del valore reale — cosa che nelle aste vere capita
// spesso, e ancora di più sui giocatori di fascia TOP/SEMI-TOP (vedi simCalcolaFasce).
// Rifiuta di rilanciare su un reparto se così facendo rischierebbe di restare senza
// spazio per i minimi ancora scoperti negli altri reparti — anche quando il giocatore è
// stato chiamato da un'altra squadra, non solo sulle proprie chiamate.
function simDecisioneRilancio(squadra, giocatore, ruolo, prezzoAttuale, rosaTarget, minimi, numAttivi, pool, algoritmi = DEFAULT_ALGORITMI) {
  if (!simPuoComprare(squadra, ruolo, rosaTarget)) return false;
  const gruppo = GRUPPO[ruolo];
  if (simRischiaSaltoMinimo(squadra, pool, rosaTarget, minimi, gruppo)) return false;
  const nextBid = prezzoAttuale + 1;
  const riserva = Math.max(0, simSlotsLiberi(squadra, rosaTarget) - 1);
  const maxSpendibile = simBudgetResiduo(squadra) - riserva;
  if (nextBid > maxSpendibile) return false;

  const valoreBase = simValorePercepito(squadra, giocatore, algoritmi) * simFattoreRuolo(squadra, ruolo);
  const sottoMinimo = simPossedutiGruppo(squadra, gruppo) < minimi[gruppo];
  const urgenzaMinimo = sottoMinimo ? (algoritmi.urgenzaMinimo ?? 1.6) : 1;
  const febbreAsta = 1 + Math.min((algoritmi.febbreAstaCap ?? 0.5), Math.max(0, numAttivi - 2) * (algoritmi.febbreAstaCoef ?? 0.12));
  const bonusFascia = giocatore.stelle ? 1 + Math.max(0, giocatore.stelle - 3) * (algoritmi.bonusFasciaCoef ?? 0.06) : 1;

  let desiderioMax = Math.round(valoreBase * (squadra.personalita[gruppo] || 1) * urgenzaMinimo * febbreAsta * bonusFascia * (0.95 + Math.random() * 0.65));
  // Colpo di testa: a volte si va comunque oltre il valore stimato di oltre il 40%,
  // più spesso sui giocatori TOP/SEMI-TOP che scatenano vere e proprie aste al rialzo.
  const probColpoDiTestaBase = (algoritmi.probColpoDiTestaBase ?? 0.22);
  const probColpoDiTestaFascia = (algoritmi.probColpoDiTestaFascia ?? 0.05);
  const probColpoDiTesta = probColpoDiTestaBase + (giocatore.stelle ? Math.max(0, giocatore.stelle - 3) * probColpoDiTestaFascia : 0);
  if (Math.random() < probColpoDiTesta) {
    const moltiplicatoreMin = (algoritmi.moltiplicatoreColpoDiTestaMin ?? 1.4);
    const moltiplicatoreMax = (algoritmi.moltiplicatoreColpoDiTestaMax ?? 1.9);
    desiderioMax = Math.max(desiderioMax, Math.round(valoreBase * (moltiplicatoreMin + Math.random() * (moltiplicatoreMax - moltiplicatoreMin))));
  }
  const effettivo = Math.min(maxSpendibile, Math.max(desiderioMax, 1));
  if (nextBid > effettivo) return false;
  const aggressivitaEffettiva = sottoMinimo ? Math.max(squadra.aggressivita, 0.65) : squadra.aggressivita;
  const prob = aggressivitaEffettiva + ((effettivo - nextBid) / (effettivo + 1)) * 0.3;
  return Math.random() < Math.min(prob, 0.95);
}

function simProssimoAttore(partecipanti, daIdx, leaderId, passati) {
  const n = partecipanti.length;
  for (let step = 1; step <= n; step++) {
    const idx = (daIdx + step) % n;
    const pid = partecipanti[idx];
    if (pid !== leaderId && !passati.has(pid)) return idx;
  }
  return -1;
}

function SimulazioneTab({
  giocatori, setup, squadreReali,
  simStarted, simSquadre, simPool, simAuction, simLog, simCallerIdx, simRosaTarget, setSimRosaTarget,
  simPorMinimo, setSimPorMinimo, simMinimi,
  simInPausa, setSimInPausa,
  avviaSimulazione, resetSimulazione, simSaltaAllaFine,
  richiediApplicaSimulazione, simApplicaMsg, vaiATab,
}) {
  const logRef = useRef(null);
  const [rosaTargetTxt, setRosaTargetTxt] = useState(String(simRosaTarget));
  const [porMinimoTxt, setPorMinimoTxt] = useState(String(simPorMinimo));

  useEffect(() => { logRef.current?.scrollIntoView({ behavior: "smooth" }); }, [simLog]);

  const disponibiliReali = useMemo(() => giocatori.filter((g) => g.stato === "disponibile"), [giocatori]);
  const preferitiCount = useMemo(() => giocatori.filter((g) => g.preferito && g.stato === "disponibile").length, [giocatori]);

  // Campi liberi: mentre scrivi non c'è alcun clamp automatico (altrimenti digitare
  // "25" passa per "2" che veniva subito corretto al minimo). Il valore viene validato
  // e applicato solo quando esci dal campo (onBlur), non ad ogni tasto premuto.
  function commitRosaTarget() {
    const n = parseInt(rosaTargetTxt, 10);
    const valido = Number.isFinite(n) ? Math.max(5, Math.min(999, n)) : simRosaTarget;
    setRosaTargetTxt(String(valido));
    setSimRosaTarget(valido);
  }
  function commitPorMinimo() {
    const n = parseInt(porMinimoTxt, 10);
    const valido = Number.isFinite(n) ? Math.max(1, Math.min(CAP_POR, n)) : simPorMinimo;
    setPorMinimoTxt(String(valido));
    setSimPorMinimo(valido);
  }

  if (!simStarted) {
    return (
      <div className="flex flex-col gap-5">
        <Section title="Simulazione asta (guarda il risultato, senza rilanciare tu)">
          <p className="text-xs text-inkdim leading-relaxed mb-3">
            Un'asta completa in automatico, anche dal tuo lato: "Tu" chiama e rilancia da solo, seguendo
            i giocatori che hai segnato preferiti, il modulo personalizzato in tab Moduli e le ripartizioni
            di budget di Setup. Le CPU restano avversari con strategie proprie. Include acquisti di profondità
            a basso costo ("scommesse") e rilanci che a volte superano il valore stimato, come in un'asta vera.
            Nulla tocca la tua asta reale finché non scegli di applicare il risultato.
          </p>
          {disponibiliReali.length < 20 ? (
            <div className="flex items-center gap-2 text-amber-300 text-xs bg-amber-400/10 border border-amber-400/30 rounded-lg p-2.5">
              <AlertTriangle size={17} className="shrink-0" />
              Importa prima la lista giocatori nella scheda "Giocatori": servono almeno una ventina di giocatori disponibili.
            </div>
          ) : (
            <>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Giocatori per rosa (obiettivo)">
                  <input
                    type="text" inputMode="numeric" pattern="[0-9]*"
                    value={rosaTargetTxt}
                    onChange={(e) => setRosaTargetTxt(e.target.value.replace(/[^0-9]/g, ""))}
                    onBlur={commitRosaTarget}
                    className="input-dark w-full"
                  />
                </Field>
                <Field label="Portieri minimi per squadra">
                  <input
                    type="text" inputMode="numeric" pattern="[0-9]*"
                    value={porMinimoTxt}
                    onChange={(e) => setPorMinimoTxt(e.target.value.replace(/[^0-9]/g, ""))}
                    onBlur={commitPorMinimo}
                    className="input-dark w-full"
                  />
                </Field>
              </div>
              <div className="text-[11px] text-inkdim mt-2">
                Il tetto reale per squadra è {CAP_POR + CAP_ALTRI} ({CAP_POR} portieri + {CAP_ALTRI} altri, dalle regole rosa): oltre quel numero la simulazione si ferma comunque lì.
              </div>
              <div className="text-[11px] text-inkdim mt-2">
                Minimi per ruolo classico (derivati dall'obiettivo di rosa): <span className="text-ink font-semibold">{simMinimi.POR} Por</span>,{" "}
                <span className="text-ink font-semibold">{simMinimi.DIF} Dif</span>,{" "}
                <span className="text-ink font-semibold">{simMinimi.CEN} Cen</span>,{" "}
                <span className="text-ink font-semibold">{simMinimi.ATT} Att</span>
              </div>
              <div className="mt-3">
                <div className="text-[11px] text-inkdim uppercase tracking-wide mb-1.5">
                  Squadre partecipanti ({squadreReali.length}) · budget {setup.budgetTotale} cr. ciascuna
                </div>
                <div className="flex flex-wrap gap-1.5">
                  {squadreReali.map((s) => (
                    <span
                      key={s.id}
                      className={`text-[11px] px-2 py-1 rounded-md border ${s.isMia ? "border-emerald-400/50 text-emerald-400 bg-emerald-400/5" : "border-line text-inkdim"}`}
                    >
                      {s.nome}{s.isMia ? " (Tu)" : ""}
                    </span>
                  ))}
                </div>
              </div>
              <div className="text-[11px] text-inkdim mt-3">
                Modulo di riferimento: <span className="text-ink font-semibold">{setup.modulo}</span> ·{" "}
                {preferitiCount > 0 ? (
                  <>preferiti segnati: <span className="text-amber-400 font-semibold">{preferitiCount}</span></>
                ) : (
                  <span className="text-amber-300">nessun preferito segnato — la simulazione userà comunque modulo e ripartizioni.</span>
                )}
              </div>
            </>
          )}
          <button
            disabled={disponibiliReali.length < 20}
            onClick={() => avviaSimulazione()}
            className="btn-primary w-full mt-4 disabled:opacity-40"
          >
            Avvia simulazione
          </button>
        </Section>
      </div>
    );
  }

  const tu = simSquadre[0];
  const callingTeam = simSquadre[simCallerIdx];
  const tuttiCompleti = simSquadre.every((s) => simRosaCompleta(s, simRosaTarget, simPool, simMinimi));
  const gById = Object.fromEntries(giocatori.map((g) => [g.id, g]));

  return (
    <div className="flex flex-col gap-4">
      {/* Tabellone */}
      <div className="rounded-xl border border-line bg-panel2 p-5 flex flex-col items-center text-center gap-1.5">
        {simAuction ? (
          <>
            <span className={`text-[10px] font-bold uppercase tracking-wide ${GRUPPO_ACCENT[GRUPPO[simAuction.ruolo]]}`}>
              {RUOLO_LABEL[simAuction.ruolo]}
            </span>
            <div className="text-2xl font-black text-ink flex items-center gap-1.5">
              {simAuction.giocatore.nome}
              {tu.preferitiIds?.has(simAuction.giocatore.id) && <Star size={19} className="text-amber-400" fill="currentColor" />}
            </div>
            <div className="text-[11px] text-inkdim font-mono">{simAuction.giocatore.squadra || "—"}</div>
            <BadgeFascia giocatore={simAuction.giocatore} />
            <BadgeGuida giocatore={simAuction.giocatore} />
            <div className="text-4xl font-black text-emerald-400 font-mono mt-1">{simAuction.prezzoAttuale}</div>
            <div className="flex gap-3 text-[11px] text-inkdim font-mono">
              <span>Pr. Medio <b className="text-ink">{simAuction.giocatore.prezzoMedio}</b></span>
              <span>Consigliato <b className="text-ink">{simAuction.giocatore.prezzoConsigliato}</b></span>
            </div>
            <div className="text-[11px] text-inkdim">
              in testa: <span className="font-semibold text-ink">{simSquadre.find((s) => s.id === simAuction.leaderId)?.nome}</span>
            </div>
          </>
        ) : tuttiCompleti ? (
          <div className="text-lg font-bold text-ink">Simulazione conclusa 🏁</div>
        ) : (
          <div className="text-sm text-inkdim">Turno di chiamata: <span className="text-ink font-semibold">{callingTeam?.nome}</span></div>
        )}
      </div>

      {/* Squadre in gara */}
      <div className="flex gap-2 overflow-x-auto pb-1">
        {simSquadre.map((s) => {
          const { por, altri } = contaRosa(s.rosa);
          return (
            <div key={s.id} className={`shrink-0 rounded-lg border border-line bg-panel p-2.5 min-w-[120px] ${s.isUtente ? "border-emerald-400/50" : ""}`}>
              <div className={`text-[11px] font-bold truncate ${s.isUtente ? "text-emerald-400" : s.accent}`}>{s.nome}{s.isUtente ? " ★" : ""}</div>
              <div className="font-mono text-sm font-bold text-ink">{simBudgetResiduo(s)} cr</div>
              <div className="text-[9px] text-inkdim font-mono">P {por}/{CAP_POR} · Altri {altri}</div>
            </div>
          );
        })}
      </div>

      {/* Controlli riproduzione */}
      {!tuttiCompleti && (
        <div className="flex gap-2">
          <button onClick={() => setSimInPausa((p) => !p)} className="btn-secondary flex-1">
            {simInPausa ? "Riprendi" : "Pausa"}
          </button>
          <button onClick={simSaltaAllaFine} className="btn-secondary flex-1">Salta alla fine</button>
        </div>
      )}

      {/* Stato azione corrente (sola lettura) */}
      <div className="rounded-xl border border-line bg-panel p-4">
        {tuttiCompleti ? (
          <div className="flex flex-col items-center gap-3 py-2">
            <Trophy className="text-amber-400" size={34} />
            <p className="text-sm text-inkdim text-center">
              Tutte le rose hanno raggiunto l'obiettivo di {simRosaTarget} giocatori
              (minimi: {simMinimi.POR} Por · {simMinimi.DIF} Dif · {simMinimi.CEN} Cen · {simMinimi.ATT} Att).
            </p>
            <RiepilogoSquadre squadre={simSquadre} gById={gById} />
            {simApplicaMsg ? (
              <div className="w-full flex flex-col gap-2">
                <div className="text-xs text-emerald-300 bg-emerald-400/10 border border-emerald-400/30 rounded-lg p-2.5 text-center">{simApplicaMsg}</div>
                <div className="flex gap-2">
                  <button onClick={() => vaiATab("rosa")} className="btn-secondary flex-1">Vai a Rosa</button>
                  <button onClick={() => vaiATab("squadre")} className="btn-secondary flex-1">Vai a Squadre</button>
                </div>
              </div>
            ) : (
              <button onClick={richiediApplicaSimulazione} className="btn-primary w-full">Applica risultato alla mia asta</button>
            )}
            <button onClick={resetSimulazione} className="btn-secondary w-full">Nuova simulazione</button>
          </div>
        ) : simInPausa ? (
          <div className="flex flex-col items-center gap-2 py-6 text-inkdim">
            <span className="text-sm">In pausa — premi "Riprendi" per continuare.</span>
          </div>
        ) : simAuction ? (
          <div className="flex flex-col items-center gap-2 py-6 text-inkdim">
            <span className="text-sm">{simSquadre.find((s) => s.id === simAuction.participants[simAuction.cursor])?.nome} sta decidendo...</span>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-2 py-6 text-inkdim">
            <Users size={22} />
            <span className="text-sm">{callingTeam?.nome} sta scegliendo chi chiamare...</span>
          </div>
        )}
      </div>

      {/* Cronaca */}
      <div className="rounded-xl border border-line bg-panel p-3 max-h-52 overflow-y-auto flex flex-col gap-1">
        {simLog.map((riga, i) => (
          <div key={i} className="text-[11px] text-inkdim leading-snug">{riga}</div>
        ))}
        <div ref={logRef} />
      </div>

      {!tuttiCompleti && (
        <button onClick={resetSimulazione} className="text-xs text-inkdim underline self-center">Annulla simulazione</button>
      )}
    </div>
  );
}

// Badge Guida: mostra titolarità/ballottaggio e valorizzato/penalizzato per il
// giocatore attualmente in asta, leggendo dal motore ValoreReale già esistente
// (guideDatabase → trovaVoceGuida/trovaInfoGiocatoreInGuida).
// Badge di fascia: TOP/SEMI-TOP/BUONA ROTAZIONE/RISERVA/SCOMMESSA con stelline, calcolati
// da simCalcolaFasce sul percentile di ValoreReale dentro il reparto del giocatore.
function BadgeFascia({ giocatore }) {
  if (!giocatore.fascia) return null;
  return (
    <div className={`flex items-center gap-1.5 text-[10px] font-bold px-2 py-1 rounded-md border ${giocatore.fasciaColore}`}>
      <span>{giocatore.fascia}</span>
      <span className="tracking-tighter">{"★".repeat(giocatore.stelle)}{"☆".repeat(5 - giocatore.stelle)}</span>
    </div>
  );
}

function BadgeGuida({ giocatore }) {
  const { info } = valoreGiocatore(giocatore);
  if (!info) return null;
  return (
    <div className="flex gap-1 flex-wrap justify-center">
      {info.probTitolare >= 1 ? (
        <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-emerald-400/15 text-emerald-300">TITOLARE</span>
      ) : (
        <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-amber-400/15 text-amber-300">BALLOTTAGGIO {Math.round(info.probTitolare * 100)}%</span>
      )}
      {info.isValorizzato && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-emerald-400/15 text-emerald-300">VALORIZZATO</span>}
      {info.isPenalizzato && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-rose-400/15 text-rose-300">PENALIZZATO</span>}
      {info.isRigorista && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-amber-400/15 text-amber-300">RIGORISTA</span>}
    </div>
  );
}

// Scheda Guida per il giocatore chiamato in Asta Live: ballottaggio con l'avversario
// diretto (se presente), specialità sui piazzati, giudizio del mister con motivi, e
// il contesto tattico della squadra. Chiusa di default per non affollare lo schermo
// durante l'asta; si apre a tocco.
function GuidaGiocatoreCard({ giocatore }) {
  const [aperto, setAperto] = useState(false);
  const { guida, info } = useMemo(() => trovaGuidaGiocatore(giocatore), [giocatore]);

  if (!guida) {
    return (
      <div className="bg-inkbg border border-line rounded-lg px-3 py-2 text-[11px] text-inkdim flex items-center gap-1.5">
        <BookOpen size={14} className="shrink-0" /> Nessun dato Guida per {giocatore.squadra || "questa squadra"}.
      </div>
    );
  }

  const percentualeMia = info?.rivale ? Math.round(info.probTitolare * 100) : null;

  const badges = [];
  if (info?.isRigorista) badges.push({ label: "RIGORISTA", cls: "bg-amber-400/15 text-amber-300" });
  if (info?.isPunizioni) badges.push({ label: "PUNIZIONI", cls: "bg-amber-400/15 text-amber-300" });
  if (info?.isAngoli) badges.push({ label: "ANGOLI", cls: "bg-amber-400/15 text-amber-300" });
  if (info?.isGiovane) badges.push({ label: "GIOVANE DA SEGUIRE", cls: "bg-violet-400/15 text-violet-300" });
  if (info?.isScommessa) badges.push({ label: "SCOMMESSA", cls: "bg-sky-400/15 text-sky-300" });

  return (
    <div className="bg-inkbg border border-line rounded-lg overflow-hidden">
      <button type="button" onClick={() => setAperto((v) => !v)} className="w-full flex items-center justify-between px-3 py-2">
        <span className="flex items-center gap-1.5 text-[11px] text-inkdim uppercase font-bold">
          <BookOpen size={14} /> Guida · {guida.teamName}
        </span>
        <ChevronDown size={17} className={`shrink-0 text-inkdim transition-transform ${aperto ? "rotate-180" : ""}`} />
      </button>
      {aperto && (
        <div className="px-3 pb-3 space-y-2.5 border-t border-line pt-2.5">
          <div className="flex items-center justify-between text-[11px] text-inkdim">
            <span>{guida.modulo} · {guida.allenatore}</span>
            <span>Att {"★".repeat(guida.rating?.attacco || 0)}{"☆".repeat(5 - (guida.rating?.attacco || 0))} · Dif {"★".repeat(guida.rating?.difesa || 0)}{"☆".repeat(5 - (guida.rating?.difesa || 0))}</span>
          </div>

          {info?.rivale ? (
            <div>
              <div className="flex items-center justify-between text-xs mb-1">
                <span className="font-mono font-bold text-emerald-400">{percentualeMia}%</span>
                <span className="text-ink font-semibold truncate px-2 text-center">{giocatore.nome} · {info.rivale.nome}</span>
                <span className="font-mono font-bold text-inkdim">{info.rivale.percentuale}%</span>
              </div>
              <div className="h-1.5 rounded-full bg-panel overflow-hidden flex">
                <div className="h-full bg-emerald-400" style={{ width: `${percentualeMia}%` }} />
                <div className="h-full bg-line" style={{ width: `${info.rivale.percentuale}%` }} />
              </div>
            </div>
          ) : info?.probTitolare >= 1 ? (
            <p className="text-[11px] text-emerald-300">Titolare fisso secondo la Guida.</p>
          ) : (
            <p className="text-[11px] text-inkdim">{giocatore.nome} non è indicato tra i titolari dalla Guida.</p>
          )}

          {badges.length > 0 && (
            <div className="flex flex-wrap gap-1">
              {badges.map((b) => (
                <span key={b.label} className={`text-[9px] font-bold px-1.5 py-0.5 rounded ${b.cls}`}>{b.label}</span>
              ))}
            </div>
          )}

          {(info?.isValorizzato || info?.isPenalizzato) && info.motivi.length > 0 && (
            <div className={`rounded-md px-2.5 py-1.5 border ${info.isValorizzato ? "border-emerald-400/30 bg-emerald-400/5" : "border-rose-400/30 bg-rose-400/5"}`}>
              <div className={`flex items-center gap-1.5 text-[10px] font-bold mb-1 ${info.isValorizzato ? "text-emerald-300" : "text-rose-300"}`}>
                {info.isValorizzato ? <ThumbsUp size={13} /> : <ThumbsDown size={13} />}
                {info.isValorizzato ? "Valorizzato dal mister" : "Penalizzato dal mister"}
              </div>
              {info.motivi.map((m, i) => <div key={i} className="text-[11px] text-inkdim">· {m}</div>)}
            </div>
          )}

          {guida.ruoliChiave?.descrizione && (
            <p className="text-[11px] text-inkdim">{guida.ruoliChiave.descrizione}</p>
          )}
        </div>
      )}
    </div>
  );
}

// Pannello "chi altro potrebbe volerlo" in Asta Live: stima, per ciascun avversario,
// quanto è interessato al ruolo del giocatore chiamato (analizzaSquadra +
// interesseSquadraPerRuolo) e apre con un consiglio rapido di strategia che confronta
// il tuo bisogno (fattoreRuolo, lo stesso usato per il tetto di spesa consigliato) con
// quello stimato degli avversari: aspettare, spendere, o rilanciare solo per disturbo.
function PannelloInteresseAvversari({ altreSquadre, gById, ruolo, gruppo, setup, fattoreRuolo }) {
  const [aperto, setAperto] = useState(false);

  const righe = useMemo(() => {
    return altreSquadre
      .map((s) => ({ squadra: s, ...interesseSquadraPerRuolo(analizzaSquadra(s, gById), ruolo, gruppo, setup) }))
      .sort((a, b) => LIVELLO_INTERESSE_ORDINE[b.livello] - LIVELLO_INTERESSE_ORDINE[a.livello] || b.residuoGruppo - a.residuoGruppo);
  }, [altreSquadre, gById, ruolo, gruppo, setup]);

  if (righe.length === 0) return null;

  const numInteressati = righe.filter((r) => r.livello === "alto" || r.livello === "medio").length;
  const verdetto = testoStrategiaInteresse(fattoreRuolo, numInteressati, righe.length);

  return (
    <div className="bg-inkbg border border-line rounded-lg overflow-hidden">
      <div className="px-3 pt-2.5 pb-2">
        <div className="flex items-center gap-1.5 text-[11px] text-inkdim uppercase font-bold mb-1">
          <ShieldAlert size={14} /> Interesse avversari
        </div>
        <p className="text-xs font-semibold leading-snug">{verdetto}</p>
      </div>
      <button
        type="button" onClick={() => setAperto((v) => !v)}
        className="w-full flex items-center justify-between px-3 py-1.5 border-t border-line text-[11px] text-inkdim"
      >
        <span>{numInteressati}/{righe.length} squadre potenzialmente interessate</span>
        <ChevronDown size={17} className={`shrink-0 transition-transform ${aperto ? "rotate-180" : ""}`} />
      </button>
      {aperto && (
        <div className="px-3 pb-3 space-y-1.5 border-t border-line pt-2.5">
          {righe.map(({ squadra: s, livello, label, residuoGruppo, creditoMedioSlot, giaPresi }) => (
            <div key={s.id} className="flex items-center justify-between gap-2 text-[11px]">
              <span className="truncate flex-1 font-semibold">{s.nome}</span>
              <span className="text-inkdim font-mono shrink-0 text-right">
                {giaPresi} {ruolo} · {livello === "saturo" ? "rosa completa" : `${residuoGruppo} cr${creditoMedioSlot ? ` (~${creditoMedioSlot}/slot)` : ""}`}
              </span>
              <span className={`shrink-0 text-[9px] font-bold px-1.5 py-0.5 rounded ${LIVELLO_INTERESSE_CLS[livello]}`}>{label}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Riepilogo di fine simulazione: la rosa composta da OGNI squadra partecipante
// (compresa "Tu", evidenziata), per reparto, con lo speso totale. Per "Tu" mostra
// anche quanti dei tuoi preferiti reali sei riuscito ad aggiudicarti in questa prova.
function RiepilogoSquadre({ squadre, gById }) {
  return (
    <div className="w-full flex flex-col gap-3 text-left">
      {squadre.map((s) => {
        const perGruppo = { POR: [], DIF: [], CEN: [], ATT: [] };
        s.rosa.forEach((r) => perGruppo[GRUPPO[r.ruolo]].push(r));
        const speso = s.rosa.reduce((a, r) => a + (r.prezzo || 0), 0);
        const preferitiTotali = s.preferitiIds ? s.preferitiIds.size : 0;
        const preferitiPresi = s.rosa.filter((r) => s.preferitiIds?.has(r.giocatoreId)).length;
        return (
          <div key={s.id} className={`rounded-lg border p-3 ${s.isUtente ? "border-emerald-400/50 bg-emerald-400/5" : "border-line bg-panel2"}`}>
            <div className="flex items-center justify-between mb-2">
              <span className={`text-xs font-bold ${s.isUtente ? "text-emerald-400" : s.accent}`}>{s.nome}{s.isUtente ? " (Tu)" : ""}</span>
              <span className="font-mono text-[11px] text-inkdim">
                speso {speso} · residuo {simBudgetResiduo(s)}
                {s.isUtente && preferitiTotali > 0 && <> · preferiti {preferitiPresi}/{preferitiTotali}</>}
              </span>
            </div>
            <div className="grid grid-cols-2 gap-2 text-[11px]">
              {["POR", "DIF", "CEN", "ATT"].map((g) => (
                <div key={g}>
                  <span className={`font-mono ${GRUPPO_ACCENT[g]}`}>{GRUPPO_LABEL[g]}: </span>
                  <span className="text-inkdim">
                    {perGruppo[g].map((r) => {
                      const nome = gById[r.giocatoreId]?.nome;
                      const preferito = s.isUtente && s.preferitiIds?.has(r.giocatoreId);
                      return nome ? `${nome}${preferito ? " ★" : ""}` : null;
                    }).filter(Boolean).join(", ") || "—"}
                  </span>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ---------- Rosa Tab ----------
function RosaTab({ giocatori, ioSquadra, budgetSpeso, budgetResiduo, annullaAssegnazione, setup, updateSetup }) {
  const gById = Object.fromEntries(giocatori.map((g) => [g.id, g]));
  const rosa = ioSquadra.rosa;
  const perGruppo = { POR: [], DIF: [], CEN: [], ATT: [] };
  rosa.forEach((r) => perGruppo[GRUPPO[r.ruolo]].push(r));
  const { por: porCount, altri: altriCount } = contaRosa(rosa);
  const totalSplit = Object.values(setup.split).reduce((a, b) => a + b, 0);

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-3 gap-2">
        <StatBox label="Budget totale" value={ioSquadra.budgetTotale} />
        <StatBox label="Speso" value={budgetSpeso} accent="text-amber-400" />
        <StatBox label="Residuo" value={budgetResiduo} accent={budgetResiduo < 0 ? "text-rose-400" : "text-emerald-400"} />
      </div>

      <div className="grid grid-cols-2 gap-2">
        <div className="bg-panel border border-line rounded-lg p-2.5">
          <div className="flex items-center justify-between text-xs text-inkdim mb-1">
            <span>Portieri</span><span className="font-mono">{porCount}/{CAP_POR}</span>
          </div>
          <div className="h-1.5 rounded-full bg-inkbg overflow-hidden">
            <div className={`h-full ${porCount >= CAP_POR ? "bg-rose-400" : "bg-amber-400"}`} style={{ width: `${Math.min(100, (porCount / CAP_POR) * 100)}%` }} />
          </div>
        </div>
        <div className="bg-panel border border-line rounded-lg p-2.5">
          <div className="flex items-center justify-between text-xs text-inkdim mb-1">
            <span>Altri ruoli</span><span className="font-mono">{altriCount}/{CAP_ALTRI}</span>
          </div>
          <div className="h-1.5 rounded-full bg-inkbg overflow-hidden">
            <div className={`h-full ${altriCount >= CAP_ALTRI ? "bg-rose-400" : "bg-emerald-400"}`} style={{ width: `${Math.min(100, (altriCount / CAP_ALTRI) * 100)}%` }} />
          </div>
        </div>
      </div>

      <Section title="Ruoli mantra e moduli">
        <RuoliModuliPanel rosa={rosa} gById={gById} />
      </Section>

      <Section title={`Ripartizione budget per reparto (${totalSplit}%)`}>
        {totalSplit !== 100 && (
          <div className="flex items-center gap-1.5 text-amber-400 text-xs mb-2">
            <AlertTriangle size={16} /> La somma dovrebbe essere 100%
          </div>
        )}
        <div className="space-y-2">
          {(["POR", "DIF", "CEN", "ATT"]).map((g) => (
            <div key={g} className="flex items-center gap-3">
              <span className={`w-24 text-xs font-bold uppercase ${GRUPPO_ACCENT[g]}`}>{GRUPPO_LABEL[g]}</span>
              <input
                type="range" min={0} max={80} value={setup.split[g]}
                onChange={(e) => updateSetup({ split: { ...setup.split, [g]: parseInt(e.target.value) } })}
                className="flex-1 accent-emerald-400"
              />
              <span className="w-12 text-right font-mono text-sm">{setup.split[g]}%</span>
            </div>
          ))}
        </div>
      </Section>

      {(["POR", "DIF", "CEN", "ATT"]).map((g) => (
        <Section key={g} title={`${GRUPPO_LABEL[g]} (${perGruppo[g].length})`}>
          <div className="space-y-1.5">
            {perGruppo[g].length === 0 && (
              <p className="text-sm text-inkdim px-1">Nessun giocatore ancora in questo reparto.</p>
            )}
            {perGruppo[g].map((r) => {
              const giocatore = gById[r.giocatoreId];
              if (!giocatore) return null;
              return (
                <div key={r.giocatoreId} className="flex items-center justify-between rounded-lg px-3 py-2 border bg-panel border-line">
                  <div className="flex items-center gap-2 min-w-0">
                    <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${GRUPPO_BG[g]} ${GRUPPO_ACCENT[g]}`}>{r.ruolo}</span>
                    <div className="min-w-0">
                      <div className="text-sm font-semibold truncate">{giocatore.nome}</div>
                      <div className="text-[11px] text-inkdim truncate">{giocatore.squadra} · {giocatore.ruoli.join("/")}</div>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <span className="font-mono text-sm text-emerald-400">{r.prezzo} cr</span>
                    <button onClick={() => annullaAssegnazione(ioSquadra.id, r.giocatoreId)} className="text-inkdim p-1"><Undo2 size={17} /></button>
                  </div>
                </div>
              );
            })}
          </div>
        </Section>
      ))}
    </div>
  );
}

// ---------- Pannello riutilizzabile: conteggio per ruolo Mantra + moduli copribili ----------
function RuoliModuliPanel({ rosa, gById, compatto, nascondiRuoli }) {
  const conteggio = useMemo(() => contaRuoliEleggibili(rosa, gById), [rosa, gById]);
  const { haPortiere, risultati } = useMemo(() => moduliCopribili(rosa, gById), [rosa, gById]);
  const moduliOk = Object.entries(risultati).filter(([, v]) => v.coperto).map(([nome]) => nome);
  const moduliNo = Object.entries(risultati).filter(([, v]) => !v.coperto);

  return (
    <div className="space-y-2.5">
      {!nascondiRuoli && (
        <div>
          <div className="text-[10px] text-inkdim uppercase font-bold mb-1">Giocatori per ruolo mantra</div>
          <div className="grid grid-cols-6 gap-1">
            {RUOLI.map((r) => (
              <div key={r} className={`rounded-md border px-1 py-1 text-center ${GRUPPO_BORDER[GRUPPO[r]]} ${GRUPPO_BG[GRUPPO[r]]}`}>
                <div className={`text-[9px] font-bold ${GRUPPO_ACCENT[GRUPPO[r]]}`}>{r}</div>
                <div className="font-mono text-sm font-bold">{conteggio[r]}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div>
        <div className="text-[10px] text-inkdim uppercase font-bold mb-1">Moduli che puoi schierare</div>
        {!haPortiere && (
          <div className="flex items-center gap-1.5 text-[11px] text-amber-300 bg-amber-400/10 border border-amber-400/30 rounded-md px-2 py-1 mb-1.5">
            <AlertTriangle size={14} className="shrink-0" /> Manca un portiere: nessun modulo è schierabile.
          </div>
        )}
        {moduliOk.length > 0 ? (
          <div className="flex flex-wrap gap-1.5">
            {moduliOk.map((m) => (
              <span key={m} className="text-xs font-mono font-bold px-2 py-1 rounded-md bg-emerald-400/15 border border-emerald-400/40 text-emerald-300">{m}</span>
            ))}
          </div>
        ) : (
          <p className="text-xs text-inkdim">Ancora nessun modulo completamente copribile con la rosa attuale.</p>
        )}
        {!compatto && moduliNo.length > 0 && (
          <div className="flex flex-wrap gap-1.5 mt-1.5">
            {moduliNo.map(([m, v]) => (
              <span key={m} className="text-[11px] font-mono px-2 py-1 rounded-md bg-inkbg border border-line text-inkdim">
                {m} · {v.assegnati}/{v.totale}
              </span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
// Analisi generica di una squadra (usata dalla tab Squadre e dal pannello "Interesse
// avversari" in Asta Live): crediti spesi/residui, anche spalmati per reparto, conteggio
// giocatori per reparto/ruolo Mantra, slot di rosa ancora liberi (Por/altri, coi CAP
// dell'asta), indice di forza rosa ed efficienza di spesa in base al Valore Reale.
function analizzaSquadra(s, gById) {
  const speso = s.rosa.reduce((acc, r) => acc + (r.prezzo || 0), 0);
  const residuo = s.budgetTotale - speso;
  const conteggio = { POR: 0, DIF: 0, CEN: 0, ATT: 0 };
  const spesoPerGruppo = { POR: 0, DIF: 0, CEN: 0, ATT: 0 };
  let valoreRosaTotale = 0;
  s.rosa.forEach((r) => {
    const gruppo = GRUPPO[r.ruolo];
    conteggio[gruppo] += 1;
    spesoPerGruppo[gruppo] += (r.prezzo || 0);
    const giocatore = gById[r.giocatoreId];
    if (giocatore) valoreRosaTotale += valoreGiocatore(giocatore).valore;
  });
  const conteggioRuoli = contaRuoliEleggibili(s.rosa, gById);
  const { por: porCount, altri: altriCount } = contaRosa(s.rosa);
  const slotResidui = Math.max(0, CAP_POR - porCount) + Math.max(0, CAP_ALTRI - altriCount);
  const creditoMedioSlot = slotResidui > 0 ? residuo / slotResidui : 0;
  const efficienza = speso > 0 ? valoreRosaTotale / speso : null;
  const repartiScoperti = (["DIF", "CEN", "ATT"]).filter((g) => conteggio[g] === 0);
  return { squadra: s, speso, residuo, conteggio, spesoPerGruppo, conteggioRuoli, porCount, altriCount, slotResidui, creditoMedioSlot, valoreRosaTotale, efficienza, repartiScoperti };
}

// Stima quanto un reparto è ancora "appetibile" per una squadra avversaria durante
// l'Asta Live, incrociando: slot di rosa ancora liberi per quel ruolo, quota di budget
// residua nel reparto (secondo la stessa ripartizione target impostata in Setup) e
// quanti giocatori di quel ruolo Mantra ha già in rosa. Non sapendo il modulo scelto
// dagli avversari, il "bisogno" è valutato a livello di reparto (POR/DIF/CEN/ATT),
// non di singolo slot tattico come per la propria rosa.
const LIVELLO_INTERESSE_ORDINE = { alto: 3, medio: 2, basso: 1, saturo: 0 };
const LIVELLO_INTERESSE_LABEL = { alto: "INTERESSE ALTO", medio: "INTERESSE MEDIO", basso: "INTERESSE BASSO", saturo: "ROSA COMPLETA" };
const LIVELLO_INTERESSE_CLS = {
  alto: "bg-rose-400/15 text-rose-300", medio: "bg-amber-400/15 text-amber-300",
  basso: "bg-line text-inkdim", saturo: "bg-line text-inkdim",
};

function interesseSquadraPerRuolo(analisi, ruolo, gruppo, setup) {
  const { squadra, residuo, conteggio, spesoPerGruppo, conteggioRuoli, porCount, altriCount } = analisi;
  const slotLiberiRuolo = ruolo === "Por" ? Math.max(0, CAP_POR - porCount) : Math.max(0, CAP_ALTRI - altriCount);
  const giaPresi = conteggioRuoli[ruolo] || 0;

  if (slotLiberiRuolo === 0 || residuo <= 0) {
    return { livello: "saturo", label: LIVELLO_INTERESSE_LABEL.saturo, residuoGruppo: 0, creditoMedioSlot: 0, giaPresi };
  }

  const allocGruppo = Math.round((squadra.budgetTotale * (setup.split[gruppo] ?? 25)) / 100);
  const residuoGruppo = allocGruppo - (spesoPerGruppo[gruppo] || 0);
  const creditoMedioSlot = Math.round(residuoGruppo / slotLiberiRuolo);
  // Credito medio "normale" per slot, come metro di paragone: budget totale diviso
  // tutti gli slot di rosa possibili (stesso per tutte le squadre, stesso setup.budgetTotale).
  const baselineSlot = (setup.budgetTotale || squadra.budgetTotale) / (CAP_POR + CAP_ALTRI);

  let livello = "medio";
  if (residuoGruppo <= 0 || giaPresi >= 3) livello = "basso";
  else if (conteggio[gruppo] === 0) livello = "alto"; // reparto ancora del tutto scoperto: urgenza
  else if (creditoMedioSlot >= baselineSlot) livello = "alto";
  else if (creditoMedioSlot < baselineSlot * 0.4) livello = "basso";

  return { livello, label: LIVELLO_INTERESSE_LABEL[livello], residuoGruppo, creditoMedioSlot, giaPresi };
}

// Confronta il bisogno stimato per "Tu" (derivato dallo stesso fattoreRuolo già usato
// per il tetto di spesa consigliato in Asta Live) con quanti avversari sembrano
// interessati, per un consiglio rapido di strategia (aspettare, spendere, disturbare).
function testoStrategiaInteresse(fattoreRuoloMio, numInteressati, numAltre) {
  const mioLivello = fattoreRuoloMio >= 1.1 ? "alto" : fattoreRuoloMio <= 0.8 ? "basso" : "medio";
  const nessunRivale = numInteressati === 0;
  if (mioLivello === "alto") {
    return nessunRivale
      ? "Ti serve e nessun avversario sembra pronto a spingere: puoi trattare con calma."
      : `Ti serve, e anche ${numInteressati}/${numAltre} avversari sono interessati: preparati a spendere.`;
  }
  if (mioLivello === "basso") {
    return nessunRivale
      ? "Non ti serve e nessun avversario sembra interessato: lascialo andare, il prezzo dovrebbe restare basso."
      : `Non ti serve, ma ${numInteressati}/${numAltre} avversari lo cercano: un rilancio di disturbo potrebbe fargli spendere di più.`;
  }
  return nessunRivale
    ? "Nessun avversario sembra interessato: puoi negoziare con margine."
    : `${numInteressati}/${numAltre} avversari potrebbero essere interessati: valuta il rilancio con un margine di sicurezza.`;
}

// Oltre a crediti/reparti mostra un indice di "forza rosa" (somma del valore reale
// stimato dei giocatori presi, corretto con i dati della Guida) e un indicatore
// di efficienza di spesa (valore reale ottenuto per credito speso), utile per capire
// durante l'asta chi ha comprato bene e chi rischia di dover rilanciare forte
// sui reparti ancora scoperti.
function SquadreTab({ squadre, giocatori }) {
  const gById = Object.fromEntries(giocatori.map((g) => [g.id, g]));
  const [espansa, setEspansa] = useState(null); // id squadra con pannello ruoli/moduli aperto

  const analisiSquadre = useMemo(() => squadre.map((s) => analizzaSquadra(s, gById)), [squadre, giocatori]);

  const ranking = useMemo(
    () => [...analisiSquadre].sort((a, b) => b.valoreRosaTotale - a.valoreRosaTotale).map((a) => a.squadra.id),
    [analisiSquadre]
  );
  const efficienzaValide = analisiSquadre.filter((a) => a.efficienza !== null);
  const miglioreEfficienzaId = efficienzaValide.length
    ? efficienzaValide.reduce((best, a) => (a.efficienza > best.efficienza ? a : best)).squadra.id
    : null;

  // Una riga per ogni giocatore in ogni rosa (le squadre senza acquisti fanno comunque
  // una riga, con le colonne giocatore vuote), con sia i totali di squadra sia i dettagli
  // del singolo acquisto: pensato come export completo per archiviare/condividere l'esito
  // dell'asta, non per essere reimportato.
  function esportaRoseCsv() {
    const header = "Squadra;BudgetTotale;Speso;Residuo;IndiceForzaRosa;EfficienzaSpesaPct;Giocatore;RuoloAssegnato;RuoliMantra;SquadraReale;Prezzo;Quotazione;Fvm;ValoreReale";
    const righe = [];
    analisiSquadre.forEach((a) => {
      const s = a.squadra;
      const effPct = a.efficienza !== null ? Math.round(a.efficienza * 100) : "";
      const base = [s.nome, s.budgetTotale, a.speso, a.residuo, a.valoreRosaTotale.toFixed(1), effPct];
      if (s.rosa.length === 0) {
        righe.push([...base, "", "", "", "", "", "", "", ""].join(";"));
        return;
      }
      s.rosa.forEach((r) => {
        const g = gById[r.giocatoreId];
        const valore = g ? valoreGiocatore(g).valore : "";
        righe.push([
          ...base,
          g?.nome || "", r.ruolo, (g?.ruoli || []).join(","), g?.squadra || "",
          r.prezzo ?? "", g?.quotazione ?? "", g?.fvm ?? "", valore,
        ].join(";"));
      });
    });
    scaricaFile(`fantacalcio-rose-${dataFileOggi()}.csv`, [header, ...righe].join("\n"));
  }

  return (
    <div className="space-y-3">
      <p className="text-xs text-inkdim">
        Forza rosa = valore reale stimato dai dati della Guida (titolarità, specialisti, giudizio del mister). Efficienza = valore reale ottenuto per ogni credito speso: sopra 100% vuol dire aver comprato bene.
      </p>
      <button
        onClick={esportaRoseCsv}
        disabled={squadre.every((s) => s.rosa.length === 0)}
        className="btn-secondary w-full inline-flex items-center justify-center gap-1.5 text-xs px-3 py-1.5 disabled:opacity-40"
      >
        <Download size={16} /> Scarica rose in CSV
      </button>
      {analisiSquadre.map((a) => {
        const s = a.squadra;
        const posizione = ranking.indexOf(s.id) + 1;
        const effPct = a.efficienza !== null ? Math.round(a.efficienza * 100) : null;
        return (
          <div key={s.id} className={`rounded-xl border p-3 ${s.isMia ? "border-emerald-400/50 bg-emerald-400/5" : "border-line bg-panel"}`}>
            <div className="flex items-center justify-between mb-2">
              <span className="font-bold text-sm flex items-center gap-1.5 truncate">
                {s.isMia && <span className="text-[10px] font-bold text-emerald-400 shrink-0">TU</span>}
                <span className="truncate">{s.nome}</span>
                {s.rosa.length > 0 && (
                  <span className="text-[10px] font-mono text-inkdim shrink-0">#{posizione} forza</span>
                )}
                {miglioreEfficienzaId === s.id && (
                  <span className="text-[9px] font-bold px-1.5 py-0.5 rounded bg-emerald-400/15 text-emerald-300 shrink-0">MIGLIOR AFFARE</span>
                )}
              </span>
              <span className={`font-mono text-sm font-bold shrink-0 ${a.residuo < 0 ? "text-rose-400" : "text-emerald-400"}`}>{a.residuo} cr</span>
            </div>

            <div className="grid grid-cols-6 gap-1 mb-1.5">
              {RUOLI.map((r) => {
                const g = GRUPPO[r];
                return (
                  <div key={r} className={`rounded-md border ${GRUPPO_BORDER[g]} ${GRUPPO_BG[g]} px-1 py-1 text-center`}>
                    <div className={`text-[9px] font-bold uppercase ${GRUPPO_ACCENT[g]}`}>{r}</div>
                    <div className="font-mono text-sm font-bold">{a.conteggioRuoli[r]}</div>
                  </div>
                );
              })}
            </div>

            <div className="grid grid-cols-2 gap-1.5 mb-1.5">
              <div className="bg-inkbg rounded-md px-2 py-1.5 border border-line">
                <div className="text-[9px] text-inkdim uppercase">Indice forza rosa</div>
                <div className="font-mono text-sm font-bold text-ink">{a.valoreRosaTotale ? a.valoreRosaTotale.toFixed(1) : "—"}</div>
              </div>
              <div className="bg-inkbg rounded-md px-2 py-1.5 border border-line">
                <div className="text-[9px] text-inkdim uppercase">Efficienza spesa</div>
                <div className={`font-mono text-sm font-bold flex items-center gap-1 ${effPct === null ? "text-inkdim" : effPct >= 105 ? "text-emerald-400" : effPct >= 95 ? "text-amber-400" : "text-rose-400"}`}>
                  {effPct === null ? "—" : (
                    <>
                      {effPct >= 100 ? <TrendingUp size={14} /> : <TrendingDown size={14} />}
                      {effPct}%
                    </>
                  )}
                </div>
              </div>
            </div>

            <div className="text-[10px] text-inkdim font-mono mb-1.5">
              {a.porCount}/{CAP_POR} portieri · {a.altriCount}/{CAP_ALTRI} altri ruoli · ~{Math.round(a.creditoMedioSlot)} cr/slot residuo
            </div>

            {a.repartiScoperti.length > 0 && (
              <div className="flex items-center gap-1.5 text-[11px] text-amber-300 bg-amber-400/10 border border-amber-400/30 rounded-md px-2 py-1 mb-1.5">
                <ShieldAlert size={16} className="shrink-0" />
                <span>Reparto scoperto: {a.repartiScoperti.map((g) => GRUPPO_LABEL[g]).join(", ")} — probabile rilancio in arrivo</span>
              </div>
            )}

            <button
              onClick={() => setEspansa((cur) => (cur === s.id ? null : s.id))}
              className="w-full flex items-center justify-between text-xs text-inkdim py-1.5 border-t border-line mt-0.5"
            >
              <span className="flex items-center gap-1.5"><LayoutTemplate size={16} /> Ruoli mantra e moduli</span>
              <ChevronDown size={17} className={`transition-transform ${espansa === s.id ? "rotate-180" : ""}`} />
            </button>
            {espansa === s.id && (
              <div className="pt-2">
                <RuoliModuliPanel rosa={s.rosa} gById={gById} compatto nascondiRuoli />
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ---------- Guida Tab ----------
// Mostra i dati di guideDatabase.json (formazioni, ballottaggi, specialisti,
// valorizzati/penalizzati, giovane da seguire, scommessa) e sovrappone un badge
// di stato incrociando il nome con la lista giocatori/rose ("in lista", "in rosa tua",
// "preso da altri"), così durante l'asta vedi subito chi è ancora disponibile.
const SENTIMENT_DOT = { positive: "bg-emerald-400", neutral: "bg-inkdim", negative: "bg-rose-400" };

function statoGiocatoreDaGuida(nomeGuida, giocatori, squadre) {
  const target = normalizza(nomeGuida);
  const giocatore = giocatori.find((g) => {
    const n = normalizza(g.nome);
    return n === target || n.includes(target) || target.includes(n);
  });
  if (!giocatore) return { label: "non in lista", cls: "bg-line text-inkdim" };
  if (giocatore.stato === "mio") return { label: "in rosa tua", cls: "bg-emerald-400/15 text-emerald-300 border border-emerald-400/40" };
  if (giocatore.stato === "preso_altri") {
    const sq = squadre.find((s) => s.id === giocatore.presoDa);
    return { label: `preso · ${sq ? sq.nome : "altri"}`, cls: "bg-rose-400/15 text-rose-300 border border-rose-400/40" };
  }
  return { label: "in lista", cls: "bg-sky-400/15 text-sky-300 border border-sky-400/40" };
}

function StatoBadge({ nomeGuida, giocatori, squadre }) {
  const stato = statoGiocatoreDaGuida(nomeGuida, giocatori, squadre);
  return <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded shrink-0 ${stato.cls}`}>{stato.label}</span>;
}

function GuidaTab({ giocatori, squadre, aggiornataIl, aggiornaGuida, ripristinaGuida }) {
  // "aggiornataIl" cambia valore a ogni aggiornaGuida/ripristinaGuida: usarlo come
  // dipendenza ricalcola la lista squadre sui dati freschi senza smontare la tab (che
  // farebbe perdere il pannello aperto e il messaggio di conferma appena mostrato).
  const squadreOrdinate = useMemo(() => {
    return Object.values(guideDatabase)
      .filter((t) => t && t.teamId)
      .sort((a, b) => a.teamName.localeCompare(b.teamName));
  }, [aggiornataIl]);
  const [teamId, setTeamId] = useState(squadreOrdinate.find((t) => t.modulo)?.teamId || squadreOrdinate[0]?.teamId);
  // Se un aggiornamento rimpiazza del tutto il database con id squadra diversi, la
  // selezione corrente può restare orfana: si ripiega sulla prima squadra disponibile.
  useEffect(() => {
    if (squadreOrdinate.length > 0 && !squadreOrdinate.some((t) => t.teamId === teamId)) {
      setTeamId(squadreOrdinate.find((t) => t.modulo)?.teamId || squadreOrdinate[0]?.teamId);
    }
  }, [squadreOrdinate, teamId]);
  const team = guideDatabase[teamId];

  const [mostraAggiorna, setMostraAggiorna] = useState(false);
  const [mostraTesto, setMostraTesto] = useState(false);
  const [testoJson, setTestoJson] = useState("");
  const [msgAggiorna, setMsgAggiorna] = useState("");
  const fileInputRef = useRef(null);

  // Valida il JSON caricato/incollato: deve avere la stessa forma di guideDatabase.json,
  // cioè un oggetto { "sluqsquadra": { teamId, teamName, ... }, ... }. "unisci" aggiorna
  // solo le squadre presenti nel file (utile per caricare aggiornamenti parziali durante
  // la settimana), "sostituisci" rimpiazza l'intero database.
  function validaEApplica(testoOrigine, modalita) {
    let parsed;
    try {
      parsed = JSON.parse(testoOrigine);
    } catch (e) {
      setMsgAggiorna("JSON non valido: controlla il file o il testo incollato.");
      return;
    }
    const valido = parsed && typeof parsed === "object" && !Array.isArray(parsed)
      && Object.values(parsed).every((v) => v && typeof v === "object");
    if (!valido) {
      setMsgAggiorna('Formato non riconosciuto: serve un oggetto tipo { "milan": { "teamId": "milan", ... }, ... } come guideDatabase.json.');
      return;
    }
    const n = Object.keys(parsed).length;
    aggiornaGuida(parsed, modalita);
    setMsgAggiorna(modalita === "sostituisci" ? `Guida sostituita: ${n} squadre.` : `Guida aggiornata: ${n} squadre unite ai dati esistenti.`);
    setTestoJson("");
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => validaEApplica(String(evt.target.result || ""), "unisci");
    reader.readAsText(file);
    e.target.value = "";
  }

  if (!team) return <p className="text-sm text-inkdim">Nessun dato disponibile.</p>;

  const haDati = !!team.modulo;

  return (
    <div className="space-y-4">
      <div className="bg-panel border border-line rounded-xl p-3">
        <div className="flex items-center justify-between gap-2">
          <div className="min-w-0">
            <div className="text-xs font-bold uppercase tracking-wide text-inkdim">Dati Guida</div>
            <div className="text-[11px] text-inkdim mt-0.5 truncate">
              {aggiornataIl ? `Aggiornati il ${aggiornataIl}` : "Dati originali (inclusi nell'app)"}
            </div>
          </div>
          <button onClick={() => setMostraAggiorna((v) => !v)} className="btn-secondary text-xs px-3 py-1.5 shrink-0">
            {mostraAggiorna ? "Chiudi" : "Aggiorna"}
          </button>
        </div>

        {mostraAggiorna && (
          <div className="mt-3 pt-3 border-t border-line space-y-2">
            <p className="text-[11px] text-inkdim">
              Carica un guideDatabase.json aggiornato, anche parziale con solo le squadre da aggiornare: quelle presenti nel file sostituiscono le corrispondenti, le altre restano invariate.
            </p>
            <input ref={fileInputRef} type="file" accept=".json,application/json" onChange={handleFile} className="hidden" />
            <button onClick={() => fileInputRef.current?.click()} className="btn-primary inline-flex items-center gap-1.5 text-xs px-3 py-1.5">
              <Upload size={14} /> Carica file JSON
            </button>

            <button onClick={() => setMostraTesto((v) => !v)} className="mt-1 flex items-center gap-1.5 text-xs text-inkdim py-1.5">
              <Upload size={16} /> oppure incolla JSON come testo
            </button>
            {mostraTesto && (
              <div className="space-y-2">
                <textarea
                  value={testoJson} onChange={(e) => setTestoJson(e.target.value)}
                  placeholder='{ "milan": { "teamId": "milan", "teamName": "Milan", ... } }'
                  rows={5} className="input-dark w-full font-mono text-xs"
                />
                <div className="flex gap-2">
                  <button onClick={() => validaEApplica(testoJson, "unisci")} className="btn-secondary flex-1">Unisci</button>
                  <button onClick={() => validaEApplica(testoJson, "sostituisci")} className="btn-secondary flex-1">Sostituisci tutto</button>
                </div>
              </div>
            )}

            {msgAggiorna && <p className="text-xs text-emerald-400">{msgAggiorna}</p>}

            {aggiornataIl && (
              <button
                onClick={() => { ripristinaGuida(); setMsgAggiorna("Ripristinati i dati originali."); }}
                className="text-xs text-rose-400 flex items-center gap-1.5 pt-1"
              >
                <RotateCcw size={16} /> Ripristina dati originali
              </button>
            )}
          </div>
        )}
      </div>

      <div className="flex gap-2 overflow-x-auto pb-1 -mx-4 px-4">
        {squadreOrdinate.map((t) => (
          <button
            key={t.teamId}
            onClick={() => setTeamId(t.teamId)}
            className={`shrink-0 flex flex-col items-center gap-1 w-14 py-1.5 rounded-lg border transition ${teamId === t.teamId ? "border-emerald-400 bg-emerald-400/10" : "border-line"
              }`}
          >
            <span className={`w-8 h-8 rounded-full flex items-center justify-center text-[10px] font-bold ${t.modulo ? "bg-panelhover text-ink" : "bg-panel text-inkdim"}`}>
              {t.teamName.slice(0, 3).toUpperCase()}
            </span>
            <span className="text-[9px] text-inkdim truncate w-full text-center">{t.teamName}</span>
          </button>
        ))}
      </div>

      {!haDati && (
        <div className="text-sm text-inkdim text-center py-10 border border-dashed border-line rounded-xl">
          Dati non ancora disponibili per {team.teamName}.<br />Arrivano a breve.
        </div>
      )}

      {haDati && (
        <>
          <div className="bg-panel border border-line rounded-xl p-3">
            <div className="flex items-center justify-between mb-1">
              <div>
                <div className="font-bold text-lg">{team.teamName}</div>
                <div className="text-xs text-inkdim">{team.modulo} · {team.allenatore}</div>
              </div>
              <div className="text-right text-[10px] text-inkdim">
                <div>Attacco {"★".repeat(team.rating?.attacco || 0)}{"☆".repeat(5 - (team.rating?.attacco || 0))}</div>
                <div>Difesa {"★".repeat(team.rating?.difesa || 0)}{"☆".repeat(5 - (team.rating?.difesa || 0))}</div>
              </div>
            </div>
            {team.updatedAt && <div className="text-[10px] text-emerald-400 mt-1">Aggiornata al {team.updatedAt}</div>}
          </div>

          <Section title="Titolari">
            <div className="space-y-1.5">
              {(team.titolari || []).map((p) => (
                <div key={p.nome} className="flex items-center justify-between bg-panel border border-line rounded-lg px-3 py-2">
                  <div className="flex items-center gap-2 min-w-0">
                    <div className="flex gap-1 shrink-0">
                      {p.ruoli.map((r) => (
                        <span key={r} className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${GRUPPO_BG[GRUPPO[r]]} ${GRUPPO_ACCENT[GRUPPO[r]]}`}>{r}</span>
                      ))}
                    </div>
                    <span className="text-sm font-semibold truncate">{p.nome}</span>
                  </div>
                  <StatoBadge nomeGuida={p.nome} giocatori={giocatori} squadre={squadre} />
                </div>
              ))}
            </div>
          </Section>

          {(team.ruoliChiave?.descrizione || (team.puntiChiave || []).length > 0) && (
            <Section title="Ruoli chiave e punti chiave">
              {team.ruoliChiave?.descrizione && <p className="text-xs text-inkdim mb-2">{team.ruoliChiave.descrizione}</p>}
              <div className="space-y-1.5">
                {(team.puntiChiave || []).map((p, i) => (
                  <div key={i} className="flex items-center gap-2 text-sm">
                    <span className={`w-2 h-2 rounded-full shrink-0 ${SENTIMENT_DOT[p.sentiment] || SENTIMENT_DOT.neutral}`} />
                    <span>{p.testo}</span>
                  </div>
                ))}
              </div>
            </Section>
          )}

          {(team.ballottaggi || []).length > 0 && (
            <Section title="Ballottaggi">
              <div className="space-y-2.5">
                {team.ballottaggi.map((b, i) => (
                  <div key={i}>
                    <div className="flex items-center justify-between text-xs mb-1">
                      <span className="font-mono font-bold text-emerald-400">{b.giocatoreA.percentuale}%</span>
                      <span className="text-ink font-semibold truncate px-2 text-center">{b.giocatoreA.nome} · {b.giocatoreB.nome}</span>
                      <span className="font-mono font-bold text-inkdim">{b.giocatoreB.percentuale}%</span>
                    </div>
                    <div className="h-1.5 rounded-full bg-inkbg overflow-hidden flex">
                      <div className="h-full bg-emerald-400" style={{ width: `${b.giocatoreA.percentuale}%` }} />
                      <div className="h-full bg-line" style={{ width: `${b.giocatoreB.percentuale}%` }} />
                    </div>
                  </div>
                ))}
              </div>
            </Section>
          )}

          {(team.specialisti?.rigoristi?.length || team.specialisti?.punizioni?.length || team.specialisti?.angoli?.length) > 0 && (
            <Section title="Specialisti">
              <div className="grid grid-cols-3 gap-2">
                {(["rigoristi", "punizioni", "angoli"]).map((k) => (
                  <div key={k}>
                    <div className="text-[10px] uppercase font-bold text-inkdim mb-1">{k}</div>
                    <div className="space-y-1">
                      {(team.specialisti?.[k] || []).map((p) => (
                        <div key={p.nome} className="text-xs truncate">{p.nome}</div>
                      ))}
                      {(team.specialisti?.[k] || []).length === 0 && <div className="text-xs text-inkdim">—</div>}
                    </div>
                  </div>
                ))}
              </div>
            </Section>
          )}

          {(team.valorizzati || []).length > 0 && (
            <div className="rounded-xl border border-emerald-400/30 bg-emerald-400/5 p-3">
              <div className="flex items-center gap-1.5 font-bold text-sm text-emerald-300 mb-2"><ThumbsUp size={17} /> Valorizzati</div>
              <div className="space-y-2">
                {team.valorizzati.map((p) => (
                  <div key={p.nome}>
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-semibold">{p.nome}</span>
                      <StatoBadge nomeGuida={p.nome} giocatori={giocatori} squadre={squadre} />
                    </div>
                    {(p.motivi || []).map((m, i) => <div key={i} className="text-xs text-inkdim">· {m}</div>)}
                  </div>
                ))}
              </div>
            </div>
          )}

          {(team.penalizzati || []).length > 0 && (
            <div className="rounded-xl border border-rose-400/30 bg-rose-400/5 p-3">
              <div className="flex items-center gap-1.5 font-bold text-sm text-rose-300 mb-2"><ThumbsDown size={17} /> Penalizzati</div>
              <div className="space-y-2">
                {team.penalizzati.map((p) => (
                  <div key={p.nome}>
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-semibold">{p.nome}</span>
                      <StatoBadge nomeGuida={p.nome} giocatori={giocatori} squadre={squadre} />
                    </div>
                    {(p.motivi || []).map((m, i) => <div key={i} className="text-xs text-inkdim">· {m}</div>)}
                  </div>
                ))}
              </div>
            </div>
          )}

          <div className="grid grid-cols-2 gap-2">
            {team.giovaneDaSeguire && (
              <div className="rounded-xl border border-violet-400/30 bg-violet-400/5 p-3">
                <div className="flex items-center gap-1.5 font-bold text-xs text-violet-300 mb-1.5"><Heart size={16} /> Giovane</div>
                <div className="text-sm font-semibold">{team.giovaneDaSeguire.nome}</div>
                <div className="text-[11px] text-inkdim mt-0.5">{team.giovaneDaSeguire.motivo}</div>
              </div>
            )}
            {team.scommessa && (
              <div className="rounded-xl border border-amber-400/30 bg-amber-400/5 p-3">
                <div className="flex items-center gap-1.5 font-bold text-xs text-amber-300 mb-1.5"><Target size={16} /> Scommessa</div>
                <div className="text-sm font-semibold">{team.scommessa.nome}</div>
                <div className="text-[11px] text-inkdim mt-0.5">{team.scommessa.motivo}</div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ---------- Algoritmi Tab (parametri del motore di valutazione e della CPU) ----------
function AlgoritmiTab({ algoritmi, updateAlgoritmi }) {
  return (
    <div className="space-y-6">
      {ALGORITMI_GRUPPI.map((gruppo, i) => (
        <React.Fragment key={gruppo.titolo}>
          <Section title={gruppo.titolo}>
            <p className="text-xs text-inkdim mb-3">{gruppo.descrizione}</p>
            <div className="space-y-4">
              {gruppo.campi.map((campo) => {
                const valore = algoritmi[campo.chiave] ?? DEFAULT_ALGORITMI[campo.chiave];
                return (
                  <div key={campo.chiave}>
                    <div className="flex items-center justify-between gap-2 mb-0.5">
                      <span className="text-xs font-semibold">{campo.label}</span>
                      <span className="font-mono text-sm text-emerald-400 shrink-0">{valore.toFixed(campo.decimali)}</span>
                    </div>
                    <input
                      type="range" min={campo.min} max={campo.max} step={campo.step}
                      value={valore}
                      onChange={(e) => updateAlgoritmi({ [campo.chiave]: parseFloat(e.target.value) })}
                      className="w-full accent-emerald-400"
                    />
                    <p className="text-[11px] text-inkdim mt-0.5">{campo.desc}</p>
                  </div>
                );
              })}
            </div>
            <button
              onClick={() => updateAlgoritmi(Object.fromEntries(gruppo.campi.map((c) => [c.chiave, DEFAULT_ALGORITMI[c.chiave]])))}
              className="btn-secondary mt-3 text-xs px-3 py-1.5 inline-flex items-center gap-1.5"
            >
              <RotateCcw size={16} /> Ripristina predefiniti
            </button>
          </Section>

          {/* Dopo "Valore Reale (FVM)": spiega i 4 indicatori della tab Squadre che
              derivano da questi parametri, prima di passare ai parametri della CPU. */}
          {i === 0 && (
            <Section title="Indicatori mostrati in Squadre">
              <p className="text-xs text-inkdim mb-3">
                Non sono parametri modificabili: sono numeri calcolati automaticamente a partire dal Valore Reale qui sopra, e li trovi nella tab Squadre per confrontare le rose.
              </p>
              <div className="space-y-3">
                {INDICATORI_FORZA_ROSA.map((ind) => (
                  <div key={ind.label} className="bg-panel border border-line rounded-lg px-3 py-2">
                    <div className="text-xs font-bold text-emerald-400 mb-0.5">{ind.label}</div>
                    <p className="text-[11px] text-inkdim">{ind.desc}</p>
                  </div>
                ))}
              </div>
            </Section>
          )}
        </React.Fragment>
      ))}
    </div>
  );
}

// ---------- UI helpers ----------
function Section({ title, children }) {
  return (
    <section>
      <h2 className="text-xs font-bold uppercase tracking-wide text-inkdim mb-2">{title}</h2>
      {children}
    </section>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="text-[11px] text-inkdim block mb-1">{label}</span>
      {children}
    </label>
  );
}

function StatBox({ label, value, accent = "text-ink" }) {
  return (
    <div className="bg-panel border border-line rounded-lg p-2.5 text-center">
      <div className="text-[10px] text-inkdim uppercase">{label}</div>
      <div className={`font-mono text-xl font-black ${accent}`}>{value}</div>
    </div>
  );
}

function ConfirmModal({ title, text, onCancel, onConfirm }) {
  return (
    <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-30 px-6">
      <div className="bg-panel border border-line rounded-xl p-5 max-w-sm w-full">
        <h3 className="font-bold text-lg mb-1.5">{title}</h3>
        <p className="text-sm text-inkdim mb-4">{text}</p>
        <div className="flex gap-2">
          <button onClick={onCancel} className="btn-secondary flex-1">Annulla</button>
          <button onClick={onConfirm} className="flex-1 bg-rose-500/90 rounded-lg py-2.5 font-semibold text-sm active:scale-95 transition">Conferma</button>
        </div>
      </div>
    </div>
  );
}

// Tailwind non supporta @apply senza config custom qui: definiamo classi via componenti inline
const styleTag = document.createElement("style");
styleTag.innerHTML = `
  .input-dark { background:#000000; border:1px solid #374151; border-radius:0.5rem; padding:0.55rem 0.7rem; font-size:0.875rem; color:#FFFFFF; outline:none; }
  .input-dark:focus { border-color:#10B981; }
  .input-dark option { background:#000000; color:#FFFFFF; }
  .btn-primary { background:#10B981; color:#FFFFFF; font-weight:700; border-radius:0.5rem; padding:0.65rem; font-size:0.875rem; transition:transform .1s; }
  .btn-primary:active { transform:scale(0.97); }
  .btn-secondary { background:#F3F4F6; border:1px solid #D1D5DB; color:#111827; font-weight:600; border-radius:0.5rem; padding:0.65rem; font-size:0.875rem; transition:transform .1s; }
  .btn-secondary:active { transform:scale(0.97); }
`;
if (typeof document !== "undefined" && !document.getElementById("asta-mantra-style")) {
  styleTag.id = "asta-mantra-style";
  document.head.appendChild(styleTag);
}