import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "motion/react";
import { 
  BarChart3, 
  TrendingUp, 
  Upload, 
  FileSpreadsheet, 
  RefreshCw, 
  Sparkles, 
  CheckCircle, 
  AlertTriangle, 
  Clock, 
  ArrowRight, 
  Database,
  Check,
  X,
  FileText,
  Info,
  ChevronRight,
  ShieldCheck,
  Brain,
  HelpCircle
} from "lucide-react";
import { IADetectionRow, IADetectionStats, AiAnalysisReport } from "./types";
import { demoIADetection } from "./demoData";

export default function App() {
  // App states
  const [detectionData, setDetectionData] = useState<IADetectionRow[]>(demoIADetection);
  const [sheetUrl, setSheetUrl] = useState<string>("");
  const [currentFileName, setCurrentFileName] = useState<string>("Données de démo (Détection IA)");
  
  // Loading states
  const [isImporting, setIsImporting] = useState<boolean>(false);
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [analysisStep, setAnalysisStep] = useState<string>("");
  const [aiReport, setAiReport] = useState<AiAnalysisReport | null>(null);
  
  // Feedback messages
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  
  // Filters and Interactive details
  const [selectedUserFilter, setSelectedUserFilter] = useState<string>("all");

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Normalization Helpers
  const normalizeChoice = (val: string): "Réelle" | "IA" | string => {
    if (!val) return "IA";
    const s = val.toLowerCase().trim();
    if (s.includes("reel") || s.includes("réel")) return "Réelle";
    if (s.includes("ia") || s.includes("ai")) return "IA";
    return val;
  };

  const normalizeYesNo = (val: string): "oui" | "non" => {
    if (!val) return "non";
    const s = val.toLowerCase().trim();
    if (s.includes("oui") || s.includes("yes") || s === "o" || s === "y") return "oui";
    return "non";
  };

  // Raw file rows processor
  const processRawData = (rows: any[][], sourceName: string) => {
    if (rows.length < 2) {
      setErrorMsg("Le document importé contient trop peu de lignes de données.");
      return;
    }

    const rawHeaders = rows[0].map(h => String(h || "").toLowerCase().trim());
    
    // Look for essential columns to match detection test rows
    const isDetectionTest = rawHeaders.some(h => 
      h.includes("reel") || h.includes("réel") || 
      h.includes("avis") || h.includes("ia") || 
      h.includes("confiance") || h.includes("changement")
    );

    if (!isDetectionTest) {
      setErrorMsg("Format de fichier non reconnu. Assurez-vous que le document contient des colonnes liées au test de détection IA (Avis, Confiance, IA, etc.).");
      return;
    }

    const mappedRows: IADetectionRow[] = [];
    
    // Auto-detect column mappings
    const idx = {
      date: rawHeaders.findIndex(h => h.includes("date") || h.includes("horodateur") || h.includes("time")),
      email: rawHeaders.findIndex(h => h.includes("email") || h.includes("courriel") || h.includes("user") || h.includes("nom")),
      urlImage: rawHeaders.findIndex(h => h.includes("url") || h.includes("image") || h.includes("lien")),
      typeReel: rawHeaders.findIndex(h => h.includes("type réel") || h.includes("type reel") || h.includes("réel") || h.includes("reel") || h === "type"),
      premierAvis: rawHeaders.findIndex(h => h.includes("premier avis") || h.includes("1er avis") || h.includes("guess 1") || h.includes("premier")),
      temps: rawHeaders.findIndex(h => h.includes("temps") || h.includes("durée") || h.includes("duration") || h.includes("seconds")),
      raison: rawHeaders.findIndex(h => h.includes("raison") || h.includes("explication") || h.includes("pourquoi")),
      confianceInitiale: rawHeaders.findIndex(h => h.includes("confiance initiale") || h.includes("confiance 1") || h.includes("initiale") || h.includes("confiance")),
      avisIA: rawHeaders.findIndex(h => h.includes("avis ia") || h.includes("suggestion") || h.includes("conseil")),
      changementAvis: rawHeaders.findIndex(h => h.includes("changement") || h.includes("avis 2") || h.includes("mind")),
      confianceFinale: rawHeaders.findIndex(h => h.includes("confiance finale") || h.includes("confiance 2") || h.includes("finale"))
    };

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0 || !row[0]) continue;

      const dateVal = idx.date !== -1 ? String(row[idx.date] || "") : `Test #${i}`;
      const emailVal = idx.email !== -1 ? String(row[idx.email] || "") : "anonyme@test.fr";
      const urlVal = idx.urlImage !== -1 ? String(row[idx.urlImage] || "") : "";
      const typeReelVal = idx.typeReel !== -1 ? String(row[idx.typeReel] || "") : "IA";
      const premierAvisVal = idx.premierAvis !== -1 ? String(row[idx.premierAvis] || "") : "Réelle";
      const tempsVal = idx.temps !== -1 ? parseFloat(String(row[idx.temps]).replace(",", ".")) || 6.5 : 6.5;
      const raisonVal = idx.raison !== -1 ? String(row[idx.raison] || "") : "";
      const confInitVal = idx.confianceInitiale !== -1 ? parseInt(String(row[idx.confianceInitiale])) || 70 : 70;
      const avisIAVal = idx.avisIA !== -1 ? String(row[idx.avisIA] || "") : "IA";
      const changAvisVal = idx.changementAvis !== -1 ? String(row[idx.changementAvis] || "") : "non";
      const confFinVal = idx.confianceFinale !== -1 ? parseInt(String(row[idx.confianceFinale])) || 70 : 70;

      mappedRows.push({
        date: dateVal,
        email: emailVal,
        urlImage: urlVal,
        typeReel: normalizeChoice(typeReelVal),
        premierAvis: normalizeChoice(premierAvisVal),
        temps: tempsVal,
        raison: raisonVal,
        confianceInitiale: confInitVal,
        avisIA: normalizeChoice(avisIAVal),
        changementAvis: normalizeYesNo(changAvisVal),
        confianceFinale: confFinVal
      });
    }

    if (mappedRows.length > 0) {
      setDetectionData(mappedRows);
      setAiReport(null);
      setCurrentFileName(sourceName);
      setSuccessMsg(`Importation réussie ! ${mappedRows.length} lignes de tests de détection IA chargées.`);
      setErrorMsg(null);
    } else {
      setErrorMsg("Aucune donnée de test valide n'a pu être extraite du fichier.");
    }
  };

  // Google Sheets Fetch (via proxy backend to bypass CORS)
  const handleGoogleSheetImport = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!sheetUrl) {
      setErrorMsg("Veuillez entrer une URL de Google Sheet valide.");
      return;
    }

    setIsImporting(true);
    setErrorMsg(null);
    setSuccessMsg(null);

    try {
      const response = await fetch("/api/import-sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url: sheetUrl }),
      });

      const result = await response.json();
      
      if (!response.ok) {
        throw new Error(result.error || "Erreur de chargement du document.");
      }

      processRawData(result.data, "Google Sheet Synchrone");
    } catch (err: any) {
      setErrorMsg(err.message || "Impossible d'importer. Vérifiez que votre Google Sheet est partagé en mode 'Lecteur Public'.");
    } finally {
      setIsImporting(false);
    }
  };

  // Local File upload (.xlsx, .xls, .csv)
  const handleLocalFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setErrorMsg(null);
    setSuccessMsg(null);
    setIsImporting(true);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = event.target?.result;
        if (!data) throw new Error("Fichier illisible.");
        
        const workbook = XLSX.read(data, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        processRawData(rows, file.name);
      } catch (err: any) {
        setErrorMsg(`Erreur lors du traitement du document : ${err.message || err}`);
      } finally {
        setIsImporting(false);
      }
    };

    reader.onerror = () => {
      setErrorMsg("Une erreur technique s'est produite lors de la lecture locale.");
      setIsImporting(false);
    };

    reader.readAsBinaryString(file);
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  // Reset to default demo data
  const handleLoadDemo = () => {
    setDetectionData(demoIADetection);
    setAiReport(null);
    setCurrentFileName("Données de démo (Détection IA)");
    setSuccessMsg("Données de démonstration chargées.");
    setErrorMsg(null);
  };

  // Calculate stats based on filters
  const stats: IADetectionStats = useMemo(() => {
    const filtered = detectionData.filter(row => 
      selectedUserFilter === "all" ? true : row.email === selectedUserFilter
    );

    const total = filtered.length;
    if (total === 0) {
      return {
        totalAnswers: 0,
        correctAnswersCount: 0,
        correctRate: 0,
        avgConfidenceInitial: 0,
        avgConfidenceFinal: 0,
        opinionChangeRate: 0,
        avgTime: 0,
        beneficialChanges: 0,
        detrimentalChanges: 0,
        neutralChanges: 0,
        correctByRealness: { realTotal: 0, realCorrect: 0, realCorrectRate: 0, iaTotal: 0, iaCorrect: 0, iaCorrectRate: 0 },
        byUser: []
      };
    }

    let correctCount = 0;
    let sumConfInit = 0;
    let sumConfFin = 0;
    let sumTime = 0;
    let changeCount = 0;
    let beneficial = 0;
    let detrimental = 0;
    let neutral = 0;

    let realTotal = 0;
    let realCorrect = 0;
    let iaTotal = 0;
    let iaCorrect = 0;

    const userMap: Record<string, { total: number; correct: number; time: number; change: number }> = {};

    filtered.forEach(row => {
      const typeNorm = row.typeReel;
      const firstNorm = row.premierAvis;
      const isFirstCorrect = firstNorm === typeNorm;
      
      const finalOpinion = row.changementAvis === "oui" 
        ? (firstNorm === "Réelle" ? "IA" : "Réelle")
        : firstNorm;
      
      const isFinalCorrect = finalOpinion === typeNorm;

      if (isFinalCorrect) {
        correctCount++;
      }

      sumConfInit += row.confianceInitiale;
      sumConfFin += row.confianceFinale;
      sumTime += row.temps;

      if (row.changementAvis === "oui") {
        changeCount++;
        if (!isFirstCorrect && isFinalCorrect) {
          beneficial++;
        } else if (isFirstCorrect && !isFinalCorrect) {
          detrimental++;
        } else {
          neutral++;
        }
      }

      if (typeNorm === "Réelle") {
        realTotal++;
        if (isFinalCorrect) realCorrect++;
      } else if (typeNorm === "IA") {
        iaTotal++;
        if (isFinalCorrect) iaCorrect++;
      }

      if (!userMap[row.email]) {
        userMap[row.email] = { total: 0, correct: 0, time: 0, change: 0 };
      }
      userMap[row.email].total++;
      if (isFinalCorrect) userMap[row.email].correct++;
      userMap[row.email].time += row.temps;
      if (row.changementAvis === "oui") userMap[row.email].change++;
    });

    const byUser = Object.entries(userMap).map(([email, u]) => ({
      email,
      total: u.total,
      correctRate: Math.round((u.correct / u.total) * 100),
      avgTime: Math.round((u.time / u.total) * 10) / 10,
      changeRate: Math.round((u.change / u.total) * 100)
    })).sort((a, b) => b.correctRate - a.correctRate);

    return {
      totalAnswers: total,
      correctAnswersCount: correctCount,
      correctRate: Math.round((correctCount / total) * 100),
      avgConfidenceInitial: Math.round((sumConfInit / total) * 10) / 10,
      avgConfidenceFinal: Math.round((sumConfFin / total) * 10) / 10,
      opinionChangeRate: Math.round((changeCount / total) * 100),
      avgTime: Math.round((sumTime / total) * 10) / 10,
      beneficialChanges: beneficial,
      detrimentalChanges: detrimental,
      neutralChanges: neutral,
      correctByRealness: {
        realTotal,
        realCorrect,
        realCorrectRate: realTotal > 0 ? Math.round((realCorrect / realTotal) * 100) : 0,
        iaTotal,
        iaCorrect,
        iaCorrectRate: iaTotal > 0 ? Math.round((iaCorrect / iaTotal) * 100) : 0
      },
      byUser
    };
  }, [detectionData, selectedUserFilter]);

  // Emails extracted for selector
  const userEmailsList = useMemo(() => {
    const list = new Set(detectionData.map(r => r.email));
    return Array.from(list);
  }, [detectionData]);

  // Dynamic Interpretations helper based on stats
  const interpretation = useMemo(() => {
    if (stats.totalAnswers === 0) return { status: "Aucun", text: "Veuillez charger des données.", color: "text-slate-400", bg: "border-slate-200" };
    
    const correct = stats.correctRate;
    const change = stats.opinionChangeRate;
    const confChange = stats.avgConfidenceFinal - stats.avgConfidenceInitial;

    if (correct >= 75 && change >= 25 && confChange > 0) {
      return {
        status: "Excellente",
        text: "Excellente performance ! Les testeurs font preuve d'une grande vigilance et l'apport de l'IA a permis d'optimiser significativement les détections ainsi que leur confiance.",
        color: "text-emerald-600",
        stroke: "#10b981",
        strokeOffset: 56.6
      };
    } else if (correct < 65 && change < 15) {
      return {
        status: "Biais cognitive",
        text: "Performance modeste. On observe un biais d'ancrage fort : les utilisateurs refusent d'ajuster leur avis malgré le support de l'IA, conservant un taux de réussite perfectible.",
        color: "text-rose-600",
        stroke: "#f43f5e",
        strokeOffset: 170
      };
    } else if (change >= 40 && stats.detrimentalChanges > stats.beneficialChanges) {
      return {
        status: "Sur-Influence",
        text: "Influence excessive. Les testeurs font trop confiance à l'IA, ce qui les amène à modifier des réponses initialement correctes pour faire de mauvaises détections finales.",
        color: "text-amber-500",
        stroke: "#f59e0b",
        strokeOffset: 113
      };
    } else {
      return {
        status: "Fiable",
        text: "Analyse globalement robuste et stable. Les utilisateurs écoutent l'IA de manière modérée pour corriger leurs doutes sans sur-influence majeure.",
        color: "text-blue-600",
        stroke: "#3b82f6",
        strokeOffset: 85
      };
    }
  }, [stats]);

  // Request Gemini Analysis endpoint
  const handleLaunchAiAnalysis = async () => {
    setIsAnalyzing(true);
    setErrorMsg(null);
    setAnalysisStep("Extraction des métriques cognitives...");

    try {
      setTimeout(() => setAnalysisStep("Alignement des indices de confiance..."), 600);
      setTimeout(() => setAnalysisStep("Consultation de l'expert cognitif Gemini 3.5..."), 1300);

      const summaryPayload = {
        totalTests: stats.totalAnswers,
        correctRate: stats.correctRate,
        initialConfidence: stats.avgConfidenceInitial,
        finalConfidence: stats.avgConfidenceFinal,
        opinionChangeRate: stats.opinionChangeRate,
        avgResponseTime: stats.avgTime,
        beneficialChanges: stats.beneficialChanges,
        detrimentalChanges: stats.detrimentalChanges,
        successOnRealImages: stats.correctByRealness.realCorrectRate,
        successOnAIImages: stats.correctByRealness.iaCorrectRate,
        samplesCount: detectionData.length
      };

      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mode: "detection-ia", dataSummary: summaryPayload })
      });

      if (!response.ok) {
        throw new Error("Impossible d'obtenir l'analyse de l'IA. Vérifiez que la clé API est active.");
      }

      const report = await response.json();
      setAiReport(report);
      
      // Scroll to insights
      setTimeout(() => {
        document.getElementById("ai-insights-section")?.scrollIntoView({ behavior: "smooth" });
      }, 300);
    } catch (err: any) {
      setErrorMsg(err.message || "Une erreur est survenue lors du calcul par le modèle d'IA.");
    } finally {
      setIsAnalyzing(false);
      setAnalysisStep("");
    }
  };



  return (
    <div id="neuro-analyzer-app" className="flex flex-col min-h-screen bg-[#F8FAFC] font-sans text-slate-800 overflow-x-hidden">
      
      {/* Top Navigation / Header */}
      <header className="flex flex-col lg:flex-row items-center justify-between px-8 py-4 bg-white border-b border-slate-200 shadow-sm z-20">
        <div className="flex items-center space-x-3 w-full lg:w-auto">
          <div className="w-9 h-9 bg-blue-600 rounded-lg flex items-center justify-center shadow-md shadow-blue-500/20 text-white shrink-0">
            <Brain size={19} />
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight text-slate-900 leading-tight">NeuroMetric Analyzer</h1>
            <p className="text-[10px] uppercase font-bold text-blue-500 tracking-wider">Métrique Cognitive & IA</p>
          </div>
        </div>

        {/* Action Controls Panel */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center space-y-2 sm:space-y-0 sm:space-x-3 w-full lg:w-auto mt-4 lg:mt-0">
          <form onSubmit={handleGoogleSheetImport} className="flex flex-1 items-center max-w-md bg-slate-100 border border-slate-200 rounded-lg overflow-hidden">
            <input 
              type="text" 
              placeholder="Collez le lien Google Sheet public..." 
              value={sheetUrl}
              onChange={(e) => setSheetUrl(e.target.value)}
              className="w-full pl-3 pr-2 py-1.5 bg-transparent text-xs text-slate-800 outline-none placeholder:text-slate-400 focus:ring-0"
            />
            <button 
              type="submit" 
              disabled={isImporting}
              className="px-3 text-slate-500 hover:text-blue-600 transition-colors shrink-0 border-l border-slate-200 py-1"
              title="Importer depuis Google Sheets"
            >
              {isImporting ? <RefreshCw className="animate-spin text-blue-600" size={13} /> : <FileSpreadsheet size={13} />}
            </button>
          </form>

          <input 
            type="file" 
            ref={fileInputRef}
            onChange={handleLocalFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden" 
          />

          <button 
            onClick={triggerFileSelect}
            disabled={isImporting}
            className="flex items-center justify-center space-x-2 px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold hover:bg-slate-50 transition-colors shadow-sm text-slate-700 shrink-0"
          >
            <Upload size={13} className="text-slate-500" />
            <span>Charger .xlsx</span>
          </button>

          <button
            onClick={handleLoadDemo}
            className="flex items-center justify-center space-x-1 px-2.5 py-1.5 bg-slate-100 border border-slate-200 rounded-lg text-xs hover:bg-slate-200 transition-colors text-slate-600 shrink-0"
            title="Recharger les données de test"
          >
            <Database size={12} />
            <span>Démo</span>
          </button>

          <button 
            onClick={handleLaunchAiAnalysis}
            disabled={isAnalyzing || detectionData.length === 0}
            className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold hover:shadow-lg hover:shadow-blue-500/20 transition-all shadow-md shrink-0 flex items-center space-x-1 disabled:opacity-50"
          >
            <Sparkles size={12} className={isAnalyzing ? "animate-pulse" : ""} />
            <span>{isAnalyzing ? "Analyse..." : "Lancer l'analyse"}</span>
          </button>
        </div>
      </header>

      {/* Global Status Banner Messages */}
      <AnimatePresence>
        {errorMsg && (
          <motion.div 
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="bg-rose-50 border-b border-rose-200 text-rose-700 px-8 py-2.5 text-xs flex items-center justify-between"
          >
            <span className="flex items-center space-x-2 font-medium">
              <AlertTriangle size={14} className="shrink-0" />
              <span>{errorMsg}</span>
            </span>
            <button onClick={() => setErrorMsg(null)} className="text-rose-400 hover:text-rose-600">
              <X size={14} />
            </button>
          </motion.div>
        )}
        {successMsg && (
          <motion.div 
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="bg-emerald-50 border-b border-emerald-200 text-emerald-800 px-8 py-2.5 text-xs flex items-center justify-between"
          >
            <span className="flex items-center space-x-2 font-medium">
              <CheckCircle size={14} className="shrink-0" />
              <span>{successMsg}</span>
            </span>
            <button onClick={() => setSuccessMsg(null)} className="text-emerald-500 hover:text-emerald-700">
              <X size={14} />
            </button>
          </motion.div>
        )}
      </AnimatePresence>

      {/* AI Loader Overlay when analyzing */}
      {isAnalyzing && (
        <div className="bg-slate-900/60 backdrop-blur-xs fixed inset-0 z-50 flex flex-col items-center justify-center text-white p-4">
          <div className="bg-slate-950/90 border border-slate-800 rounded-3xl p-8 max-w-sm w-full shadow-2xl text-center flex flex-col items-center">
            <div className="relative w-16 h-16 mb-4 flex items-center justify-center">
              <span className="absolute inset-0 rounded-full border-4 border-blue-500/10 border-t-blue-500 animate-spin"></span>
              <Brain className="text-blue-400 animate-pulse" size={28} />
            </div>
            <h3 className="font-display font-bold text-base mb-1.5">Moteur d'Analyse Actif</h3>
            <p className="text-[11px] text-slate-400 animate-pulse">{analysisStep || "Calcul en cours..."}</p>
          </div>
        </div>
      )}

      {/* Dashboard Grid Container */}
      <main className="flex-grow max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 flex flex-col space-y-6">

        {/* Primary Row: 4 Metric Cards */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          
          {/* Card 1: Correct response rate */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm transition-all hover:shadow-md flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Taux Réponse Correcte</p>
                <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded-sm">Final</span>
              </div>
              <div className="flex items-baseline space-x-2 mt-1">
                <span className="text-2xl font-black text-slate-900">{stats.correctRate}%</span>
                <span className="text-xs font-semibold text-slate-400">({stats.correctAnswersCount}/{stats.totalAnswers})</span>
              </div>
            </div>
            <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
              <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${stats.correctRate}%` }}></div>
            </div>
          </div>

          {/* Card 2: Avg initial confidence */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm transition-all hover:shadow-md flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Confiance Initiale Moyenne</p>
                <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded-sm">Avant IA</span>
              </div>
              <div className="flex items-baseline space-x-2 mt-1">
                <span className="text-2xl font-black text-slate-900">{stats.avgConfidenceInitial}%</span>
              </div>
            </div>
            <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
              <div className="bg-slate-300 h-full rounded-full" style={{ width: `${stats.avgConfidenceInitial}%` }}></div>
            </div>
          </div>

          {/* Card 3: Opinion change rate */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm transition-all hover:shadow-md flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Taux Changement d'Avis</p>
                <span className="text-[10px] font-bold text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded-sm">Flexibilité</span>
              </div>
              <div className="flex items-baseline space-x-2 mt-1">
                <span className="text-2xl font-black text-slate-900">{stats.opinionChangeRate}%</span>
              </div>
            </div>
            <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
              <div className="bg-amber-400 h-full rounded-full" style={{ width: `${stats.opinionChangeRate}%` }}></div>
            </div>
          </div>

          {/* Card 4: Avg final confidence */}
          <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm transition-all hover:shadow-md flex flex-col justify-between">
            <div>
              <div className="flex justify-between items-start">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Confiance Finale Moyenne</p>
                <span className="text-[10px] font-bold text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded-sm">Après IA</span>
              </div>
              <div className="flex items-baseline space-x-2 mt-1">
                <span className="text-2xl font-black text-slate-900">{stats.avgConfidenceFinal}%</span>
              </div>
            </div>
            <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
              <div className="bg-blue-600 h-full rounded-full" style={{ width: `${stats.avgConfidenceFinal}%` }}></div>
            </div>
          </div>

        </div>

        {/* Main Section: Graph, Table and Indicators */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* Left Column: Line Chart & Table (col-span-9) */}
          <div className="lg:col-span-9 flex flex-col space-y-6">

            {/* Figures de synthèse des métriques */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              
              {/* Figure A : Comparaison de Confiance (Slope / Progression) */}
              <div className="bg-white border border-slate-100 rounded-3xl p-6 shadow-sm flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <h4 className="font-display font-bold text-slate-900 text-sm">Évolution de la Certitude Moyenne</h4>
                    <span className="text-[10px] font-mono font-bold bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full">
                      Impact IA
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-6">Comparatif de l'assurance des testeurs avant et après la suggestion d'IA</p>
                </div>

                <div className="relative h-44 flex items-center justify-center">
                  <svg viewBox="0 0 400 150" className="w-full h-full">
                    {/* Background Grid Lines */}
                    <line x1="60" y1="20" x2="340" y2="20" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="3 3" />
                    <line x1="60" y1="75" x2="340" y2="75" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="3 3" />
                    <line x1="60" y1="130" x2="340" y2="130" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="3 3" />

                    {/* Connecting slope line */}
                    {/* Y scale is inverted: 20 is 100%, 130 is 0. Formula: 130 - (val/100)*110 */}
                    {(() => {
                      const y1 = 130 - (stats.avgConfidenceInitial / 100) * 110;
                      const y2 = 130 - (stats.avgConfidenceFinal / 100) * 110;
                      const isUp = stats.avgConfidenceFinal >= stats.avgConfidenceInitial;
                      const strokeColor = isUp ? "#3b82f6" : "#f43f5e";
                      const gradientId = "slope-grad";
                      return (
                        <>
                          <defs>
                            <linearGradient id={gradientId} x1="0" y1="0" x2="1" y2="0">
                              <stop offset="0%" stopColor="#94a3b8" />
                              <stop offset="100%" stopColor={strokeColor} />
                            </linearGradient>
                          </defs>
                          {/* Slope line */}
                          <line x1="80" y1={y1} x2="320" y2={y2} stroke={`url(#${gradientId})`} strokeWidth="4" strokeLinecap="round" />
                          
                          {/* Pulsing indicator on Final value */}
                          <circle cx="320" cy={y2} r="9" fill={strokeColor} fillOpacity="0.15" className="animate-pulse" />
                          
                          {/* Initial value circle */}
                          <circle cx="80" cy={y1} r="6" fill="#64748b" stroke="#ffffff" strokeWidth="2" />
                          <text x="80" y={y1 - 12} fill="#64748b" fontSize="12" fontWeight="bold" textAnchor="middle" className="font-mono">
                            {stats.avgConfidenceInitial}%
                          </text>

                          {/* Final value circle */}
                          <circle cx="320" cy={y2} r="7" fill={strokeColor} stroke="#ffffff" strokeWidth="2" />
                          <text x="320" y={y2 - 12} fill={strokeColor} fontSize="13" fontWeight="bold" textAnchor="middle" className="font-mono">
                            {stats.avgConfidenceFinal}%
                          </text>
                        </>
                      );
                    })()}

                    {/* Vertical label lines */}
                    <line x1="80" y1="20" x2="80" y2="135" stroke="#cbd5e1" strokeWidth="1" strokeDasharray="2 2" />
                    <line x1="320" y1="20" x2="320" y2="135" stroke="#cbd5e1" strokeWidth="1" strokeDasharray="2 2" />

                    {/* Bottom axis labels */}
                    <text x="80" y="148" fill="#64748b" fontSize="10" fontWeight="bold" textAnchor="middle">
                      AVANT IA
                    </text>
                    <text x="320" y="148" fill="#1e293b" fontSize="10" fontWeight="bold" textAnchor="middle">
                      APRÈS IA
                    </text>
                  </svg>
                </div>

                <div className="mt-4 pt-3 border-t border-slate-50 flex items-center justify-between text-xs">
                  <span className="text-slate-500">Variation nette :</span>
                  {(() => {
                    const diff = stats.avgConfidenceFinal - stats.avgConfidenceInitial;
                    const isUp = diff >= 0;
                    return (
                      <span className={`font-bold font-mono px-2 py-0.5 rounded-md ${isUp ? "bg-emerald-50 text-emerald-600" : "bg-rose-50 text-rose-600"}`}>
                        {isUp ? "+" : ""}{diff.toFixed(1)}% d'assurance
                      </span>
                    );
                  })()}
                </div>
              </div>

              {/* Figure B : Double Jauges Circulaires (Taux de Réponse Correcte vs Taux de Changement) */}
              <div className="bg-white border border-slate-100 rounded-3xl p-6 shadow-sm flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <h4 className="font-display font-bold text-slate-900 text-sm">Visualisation des Taux Majeurs</h4>
                    <span className="text-[10px] font-mono font-bold bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-full">
                      Performance & Flexibilité
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-6">Comparaison visuelle de l'exactitude finale et de la réceptivité à l'IA</p>
                </div>

                <div className="grid grid-cols-2 gap-4 items-center justify-center py-2">
                  
                  {/* Gauge 1: Taux de réussite */}
                  <div className="flex flex-col items-center">
                    <div className="relative w-24 h-24 flex items-center justify-center">
                      <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                        {/* Background track */}
                        <circle cx="50" cy="50" r="40" stroke="#f1f5f9" strokeWidth="6" fill="transparent" />
                        {/* Progress track */}
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          stroke="#10b981" 
                          strokeWidth="7" 
                          fill="transparent" 
                          strokeDasharray="251.2" 
                          strokeDashoffset={251.2 - (stats.correctRate / 100) * 251.2} 
                          strokeLinecap="round"
                          className="transition-all duration-1000"
                        />
                      </svg>
                      <div className="absolute flex flex-col items-center justify-center">
                        <span className="text-lg font-black text-slate-900 font-mono">{stats.correctRate}%</span>
                        <span className="text-[8px] uppercase tracking-wider font-bold text-slate-400">Réussite</span>
                      </div>
                    </div>
                    <span className="text-[10px] text-slate-400 font-medium text-center mt-2">
                      {stats.correctAnswersCount} / {stats.totalAnswers} corrects
                    </span>
                  </div>

                  {/* Gauge 2: Taux de changement */}
                  <div className="flex flex-col items-center">
                    <div className="relative w-24 h-24 flex items-center justify-center">
                      <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                        {/* Background track */}
                        <circle cx="50" cy="50" r="40" stroke="#f1f5f9" strokeWidth="6" fill="transparent" />
                        {/* Progress track */}
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          stroke="#f59e0b" 
                          strokeWidth="7" 
                          fill="transparent" 
                          strokeDasharray="251.2" 
                          strokeDashoffset={251.2 - (stats.opinionChangeRate / 100) * 251.2} 
                          strokeLinecap="round"
                          className="transition-all duration-1000"
                        />
                      </svg>
                      <div className="absolute flex flex-col items-center justify-center">
                        <span className="text-lg font-black text-slate-900 font-mono">{stats.opinionChangeRate}%</span>
                        <span className="text-[8px] uppercase tracking-wider font-bold text-slate-400">Changement</span>
                      </div>
                    </div>
                    <span className="text-[10px] text-slate-400 font-medium text-center mt-2">
                      {stats.beneficialChanges} cas utiles
                    </span>
                  </div>

                </div>

                <div className="mt-4 pt-3 border-t border-slate-50 flex items-center justify-between text-[11px] text-slate-500">
                  <span className="flex items-center">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 mr-1.5 inline-block"></span> Correct final
                  </span>
                  <span className="flex items-center">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-500 mr-1.5 inline-block"></span> Réceptivité
                  </span>
                </div>
              </div>

            </div>

            {/* Respondents list / Table */}
            <div className="bg-white border border-slate-100 rounded-3xl p-6 shadow-sm">
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between mb-4 gap-2">
                <div>
                  <h4 className="font-display font-bold text-slate-950 text-base">Données et Réponses</h4>
                  <p className="text-xs text-slate-500">Filtrer ou inspecter les détails des testeurs</p>
                </div>
                <div className="flex items-center space-x-2">
                  <span className="text-xs text-slate-500 shrink-0">Filtrer :</span>
                  <select 
                    value={selectedUserFilter}
                    onChange={(e) => setSelectedUserFilter(e.target.value)}
                    className="text-xs bg-slate-100 border border-slate-200 rounded-lg py-1 px-2 text-slate-700 focus:outline-none"
                  >
                    <option value="all">Tous les utilisateurs ({userEmailsList.length})</option>
                    {userEmailsList.map((email, idx) => (
                      <option key={idx} value={email}>{email}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="overflow-x-auto w-full">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-100 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      <th className="py-3 px-4">Utilisateur / Image</th>
                      <th className="py-3 px-4">Réel / Premier avis</th>
                      <th className="py-3 px-4">Conf. Initiale</th>
                      <th className="py-3 px-4">Conseil IA</th>
                      <th className="py-3 px-4">Changement</th>
                      <th className="py-3 px-4 text-right">Conf. Finale</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50 text-xs">
                    {detectionData
                      .filter(r => selectedUserFilter === "all" ? true : r.email === selectedUserFilter)
                      .slice(0, 8)
                      .map((row, idx) => {
                        const finalChoice = row.changementAvis === "oui" 
                          ? (row.premierAvis === "Réelle" ? "IA" : "Réelle")
                          : row.premierAvis;
                        const correct = finalChoice === row.typeReel;

                        return (
                          <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                            <td className="py-3 px-4">
                              <div className="font-semibold text-slate-900">{row.email.split("@")[0]}</div>
                              <div className="text-[10px] text-slate-400 truncate max-w-[150px]" title={row.urlImage || "N/A"}>
                                {row.urlImage ? row.urlImage.split("/").pop() : "Image test"}
                              </div>
                            </td>
                            <td className="py-3 px-4">
                              <div className="flex items-center space-x-1.5">
                                <span className={`w-1.5 h-1.5 rounded-full ${row.typeReel === "IA" ? "bg-blue-500" : "bg-emerald-500"}`}></span>
                                <span className="font-medium text-slate-700">{row.typeReel}</span>
                                <ArrowRight size={10} className="text-slate-300" />
                                <span className="text-slate-500">{row.premierAvis}</span>
                              </div>
                            </td>
                            <td className="py-3 px-4 font-mono font-medium text-slate-600">
                              {row.confianceInitiale}%
                            </td>
                            <td className="py-3 px-4">
                              <span className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${row.avisIA === row.typeReel ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"}`}>
                                {row.avisIA}
                              </span>
                            </td>
                            <td className="py-3 px-4">
                              <span className={`font-medium ${row.changementAvis === "oui" ? "text-amber-600" : "text-slate-400"}`}>
                                {row.changementAvis === "oui" ? "Oui" : "Non"}
                              </span>
                            </td>
                            <td className="py-3 px-4 text-right font-mono font-bold">
                              <span className={correct ? "text-emerald-600" : "text-rose-500"}>
                                {row.confianceFinale}%
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                  </tbody>
                </table>
              </div>
              
              {detectionData.filter(r => selectedUserFilter === "all" ? true : r.email === selectedUserFilter).length > 8 && (
                <div className="text-center pt-4 border-t border-slate-50 text-[11px] text-slate-400">
                  Affichage des 8 premiers résultats sur {detectionData.filter(r => selectedUserFilter === "all" ? true : r.email === selectedUserFilter).length} au total.
                </div>
              )}
            </div>
          </div>

          {/* Right Column: Confidence Gauge & Indicators (col-span-3) */}
          <div className="lg:col-span-3 flex flex-col space-y-6">
            
            {/* Quick Interpreter Gauge */}
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex flex-col items-center">
              <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">Interprète Statistique</h4>
              
              <div className="relative w-36 h-36 flex items-center justify-center">
                <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                  {/* Empty track */}
                  <circle cx="50" cy="50" r="43" stroke="#f1f5f9" strokeWidth="7" fill="transparent" />
                  {/* Progress track */}
                  <circle 
                    cx="50" 
                    cy="50" 
                    r="43" 
                    stroke={interpretation.stroke || "#3b82f6"} 
                    strokeWidth="8" 
                    fill="transparent" 
                    strokeDasharray="270" 
                    strokeDashoffset={interpretation.strokeOffset || "85"} 
                    strokeLinecap="round"
                    className="transition-all duration-1000"
                  />
                </svg>
                <div className="absolute flex flex-col items-center">
                  <span className={`text-xl font-black ${interpretation.color}`}>{interpretation.status}</span>
                  <span className="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Fiabilité</span>
                </div>
              </div>

              <p className="mt-4 text-center text-xs text-slate-500 leading-relaxed font-medium">
                {interpretation.text}
              </p>
            </div>

            {/* Live Indicators Details List */}
            <div className="bg-slate-900 rounded-3xl p-6 text-white shadow-xl flex-grow flex flex-col justify-between">
              <div>
                <h4 className="text-xs font-bold text-blue-400 uppercase tracking-widest mb-5 flex items-center">
                  <span className="w-2 h-2 bg-blue-500 rounded-full mr-2 inline-block"></span>
                  Indicateurs Cognitifs
                </h4>
                
                <div className="space-y-4">
                  {/* Row 1: Consensus */}
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-300 mb-1">
                      <span>Consensus Initial Correct</span>
                      <span className="font-bold text-emerald-400">
                        {Math.round((detectionData.filter(d => d.premierAvis === d.typeReel).length / detectionData.length) * 100)}%
                      </span>
                    </div>
                    <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                      <div 
                        className="bg-emerald-400 h-full rounded-full" 
                        style={{ width: `${(detectionData.filter(d => d.premierAvis === d.typeReel).length / detectionData.length) * 100}%` }}
                      ></div>
                    </div>
                  </div>

                  {/* Row 2: Helpful impact */}
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-300 mb-1">
                      <span>Changements Utiles</span>
                      <span className="font-bold text-blue-400">
                        {stats.beneficialChanges} cas
                      </span>
                    </div>
                    <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                      <div 
                        className="bg-blue-400 h-full rounded-full" 
                        style={{ width: `${stats.totalAnswers > 0 ? (stats.beneficialChanges / stats.totalAnswers) * 100 : 0}%` }}
                      ></div>
                    </div>
                  </div>

                  {/* Row 3: Response latency */}
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-300 mb-1">
                      <span>Latence de Décision</span>
                      <span className="font-bold text-amber-400">{stats.avgTime}s</span>
                    </div>
                    <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                      <div 
                        className="bg-amber-400 h-full rounded-full" 
                        style={{ width: `${Math.min((stats.avgTime / 20) * 100, 100)}%` }}
                      ></div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="mt-8 pt-6 border-t border-slate-800">
                <button 
                  onClick={handleLaunchAiAnalysis}
                  disabled={isAnalyzing}
                  className="w-full py-2.5 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold transition-all border border-slate-700/80 flex items-center justify-center space-x-2"
                >
                  <Sparkles size={13} className="text-blue-400" />
                  <span>{isAnalyzing ? "Calcul de l'IA..." : "Synthèse Cognitive IA"}</span>
                </button>
              </div>
            </div>

          </div>

        </div>

        {/* Dynamic AI Insights Report Render */}
        <AnimatePresence>
          {aiReport && (
            <motion.div 
              id="ai-insights-section"
              initial={{ opacity: 0, y: 30 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 20 }}
              className="bg-slate-900 border border-slate-800 p-8 rounded-3xl shadow-2xl relative overflow-hidden text-slate-100"
            >
              {/* Abs decorative light logo */}
              <div className="absolute top-0 right-0 p-8 opacity-5 pointer-events-none">
                <Sparkles size={180} className="text-blue-500 animate-pulse" />
              </div>

              <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 border-b border-slate-800 pb-6 relative z-10 gap-4">
                <div className="flex items-center space-x-3">
                  <div className="bg-blue-600 p-2 rounded-xl shadow-lg shadow-blue-500/30">
                    <Sparkles className="text-white" size={20} />
                  </div>
                  <div>
                    <span className="text-[10px] font-bold text-blue-400 tracking-wider uppercase font-mono">Expertise cognitive</span>
                    <h3 className="font-display font-extrabold text-white text-xl">Rapport Cognitif & Diagnostic</h3>
                  </div>
                </div>
                <div className="flex items-center space-x-2 text-xs font-mono bg-slate-800/80 border border-slate-700 px-3 py-1.5 rounded-full text-blue-300">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-ping"></span>
                  <span>Moteur Gemini 3.5-Flash</span>
                </div>
              </div>

              <div className="space-y-6 relative z-10 text-slate-300">
                {/* 1. Synthèse */}
                <div>
                  <h4 className="text-[10px] font-bold font-mono uppercase tracking-widest text-blue-400 mb-2 flex items-center">
                    <span className="w-1.5 h-1.5 bg-blue-400 rounded-full mr-2 inline-block"></span>
                    🎯 Synthèse Générale
                  </h4>
                  <p className="text-base text-white leading-relaxed font-medium">
                    {aiReport.performanceSummary}
                  </p>
                </div>

                {/* 2. Strengths vs Weaknesses */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                  <div className="bg-slate-950/40 border border-slate-800/80 p-5 rounded-2xl">
                    <h5 className="flex items-center space-x-2 font-display font-bold text-emerald-400 mb-3 text-sm">
                      <ShieldCheck size={16} />
                      <span>Atouts comportementaux</span>
                    </h5>
                    <ul className="space-y-2.5">
                      {aiReport.strengths.map((str, idx) => (
                        <li key={idx} className="flex items-start space-x-2.5 text-xs text-slate-200">
                          <span className="bg-emerald-950/80 text-emerald-400 p-0.5 rounded-full mt-0.5">
                            <Check size={10} strokeWidth={3} />
                          </span>
                          <span>{str}</span>
                        </li>
                      ))}
                    </ul>
                  </div>

                  <div className="bg-slate-950/40 border border-slate-800/80 p-5 rounded-2xl">
                    <h5 className="flex items-center space-x-2 font-display font-bold text-amber-400 mb-3 text-sm">
                      <AlertTriangle size={16} />
                      <span>Vulnérabilités de détection</span>
                    </h5>
                    <ul className="space-y-2.5">
                      {aiReport.vulnerabilities.map((vul, idx) => (
                        <li key={idx} className="flex items-start space-x-2.5 text-xs text-slate-200">
                          <span className="bg-amber-950/80 text-amber-400 p-0.5 rounded-full mt-0.5">
                            <AlertTriangle size={10} strokeWidth={3} />
                          </span>
                          <span>{vul}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>

                {/* 3. Confidence assessment */}
                <div className="bg-slate-950/20 border border-slate-800/80 p-5 rounded-2xl">
                  <h4 className="text-[10px] font-bold font-mono uppercase tracking-widest text-blue-400 mb-2 flex items-center">
                    <span className="w-1.5 h-1.5 bg-blue-400 rounded-full mr-2 inline-block"></span>
                    🧠 Évolution de la Confiance & Certitude
                  </h4>
                  <p className="text-xs leading-relaxed text-slate-200">
                    {aiReport.confidenceAnalysis}
                  </p>
                </div>

                {/* 4. Actions recommendations */}
                <div>
                  <h4 className="text-[10px] font-bold font-mono uppercase tracking-widest text-blue-400 mb-3 flex items-center">
                    <span className="w-1.5 h-1.5 bg-blue-400 rounded-full mr-2 inline-block"></span>
                    💡 Plan d'Action & Entraînement
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    {aiReport.recommendations.map((rec, idx) => (
                      <div key={idx} className="bg-slate-950/40 border border-slate-800/80 p-4 rounded-xl flex items-start space-x-3">
                        <div className="bg-slate-800 text-blue-400 font-mono font-bold text-[10px] w-5 h-5 rounded flex items-center justify-center shrink-0 mt-0.5">
                          {idx + 1}
                        </div>
                        <span className="text-xs text-slate-200 leading-normal">{rec}</span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>
            </motion.div>
          )}
        </AnimatePresence>

      </main>

      {/* Footer / Connection bar status */}
      <footer className="px-8 py-3 bg-white border-t border-slate-200 text-[10px] text-slate-400 flex flex-col sm:flex-row justify-between items-center mt-auto gap-2">
        <div className="flex flex-wrap items-center justify-center sm:justify-start gap-2">
          <span>Source : <strong className="text-slate-600">{currentFileName}</strong></span>
          <span className="hidden sm:inline">|</span>
          <span className="flex items-center">
            <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full mr-1.5 animate-pulse"></span>
            Analyse cognitive active
          </span>
        </div>
        <div className="flex space-x-4">
          <span className="hover:text-blue-600 cursor-pointer transition-colors" onClick={() => setSuccessMsg("Documentation API de détection IA en ligne.")}>Doc API</span>
          <span className="hover:text-blue-600 cursor-pointer transition-colors" onClick={() => setSuccessMsg("Historique des analyses chargé.")}>Journaux</span>
        </div>
      </footer>
    </div>
  );
}
