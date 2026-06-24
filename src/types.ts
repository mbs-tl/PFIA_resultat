export interface SimpleMeasure {
  date: string;
  value: number;
}

export interface IADetectionRow {
  date: string;
  email: string;
  urlImage: string;
  typeReel: "Réelle" | "IA" | string;
  premierAvis: "Réelle" | "IA" | string;
  temps: number; // in seconds
  raison: string;
  confianceInitiale: number; // 0 to 100
  avisIA: "Réelle" | "IA" | string;
  changementAvis: "oui" | "non" | string;
  confianceFinale: number; // 0 to 100
}

export interface IADetectionStats {
  totalAnswers: number;
  correctAnswersCount: number;
  correctRate: number; // percentage
  avgConfidenceInitial: number;
  avgConfidenceFinal: number;
  opinionChangeRate: number; // percentage
  avgTime: number; // in seconds
  beneficialChanges: number; // wrong -> correct
  detrimentalChanges: number; // correct -> wrong
  neutralChanges: number; // stayed correct or stayed wrong after change
  correctByRealness: {
    realTotal: number;
    realCorrect: number;
    realCorrectRate: number;
    iaTotal: number;
    iaCorrect: number;
    iaCorrectRate: number;
  };
  byUser: Array<{
    email: string;
    total: number;
    correctRate: number;
    avgTime: number;
    changeRate: number;
  }>;
}

export interface AiAnalysisReport {
  performanceSummary: string;
  strengths: string[];
  vulnerabilities: string[];
  confidenceAnalysis: string;
  recommendations: string[];
}
