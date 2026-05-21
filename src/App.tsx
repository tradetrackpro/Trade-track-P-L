/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useMemo, useEffect, ChangeEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  TrendingUp, 
  TrendingDown, 
  Calculator as CalcIcon, 
  RefreshCcw, 
  ChevronDown, 
  ChevronUp,
  ChevronLeft,
  ChevronRight,
  Sun,
  Moon,
  Settings,
  History,
  Info,
  IndianRupee,
  Percent,
  User,
  Plus,
  Trash2,
  Download,
  Upload,
  FileText,
  Table as TableIcon,
  UserCircle,
  Heart,
  Calendar,
  Cloud,
  CloudUpload,
  CloudDownload,
  LogIn,
  LogOut,
  Loader2,
  Fingerprint,
  Lock,
  Unlock,
  Wallet,
  ArrowUpRight,
  ArrowDownLeft,
  BookOpen,
  Edit3,
  Palmtree,
  BarChart3,
  MessageSquare,
  Send,
  Sparkles,
  Bot,
  Globe
} from 'lucide-react';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  ReferenceLine,
  AreaChart,
  Area,
  ReferenceArea,
  BarChart,
  Bar,
  Cell
} from 'recharts';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { initAuth, googleSignIn, logout, getAccessToken } from './lib/auth';
import { createOrUpdateBackup, restoreBackup } from './services/driveService';
import { listRecentEmails, sendTradingReport } from './services/gmailService';
import type { GmailMessageSnippet } from './services/gmailService';
import { User as FirebaseUser } from 'firebase/auth';

// --- Types ---

interface TradeResult {
  profit: number;
  profitPercentage: number;
  totalValue: number;
  fees: number;
  isProfit: boolean;
  roe: number; // Return on Equity for leverage
}

interface SavedTrade {
  id: string;
  mode: 'equity' | 'options';
  type: 'long' | 'short';
  entry: number;
  exit: number;
  quantity: number;
  leverage: number;
  profit: number;
  timestamp: number;
  entryDate?: string;
  exitDate?: string;
  lotSize?: number;
  strategy?: string;
  notes?: string;
  isFavorite?: boolean;
  ticker?: string;
  tradeSetup?: string;
  entryExitRationale?: string;
  lessonsLearned?: string;
  isManual?: boolean;
}

interface FundTransaction {
  id: string;
  type: 'deposit' | 'withdrawal';
  amount: number;
  timestamp: number;
  notes?: string;
}

// --- Components ---

const MARKET_HOLIDAYS_PRESET: Record<string, string> = {
  // 2025
  '2025-01-26': 'Republic Day',
  '2025-03-14': 'Holi',
  '2025-03-31': 'Ramzan Id',
  '2025-04-10': 'Mahavir Jayanti',
  '2025-04-11': 'Good Friday',
  '2025-04-14': 'Ambedkar Jayanti',
  '2025-05-01': 'Maharashtra Day',
  '2025-08-15': 'Independence Day',
  '2025-10-02': 'Mahatma Gandhi Jayanti',
  '2025-10-22': 'Diwali Balipratipada',
  '2025-11-05': 'Guru Nanak Jayanti',
  '2025-12-25': 'Christmas',

// 2026
  '2026-01-26': 'Republic Day',
  '2026-03-05': 'Holi',
  '2026-03-20': 'Ramzan Id',
  '2026-04-02': 'Mahavir Jayanti',
  '2026-04-03': 'Good Friday',
  '2026-04-14': 'Ambedkar Jayanti',
  '2026-05-01': 'Maharashtra Day',
  '2026-05-27': 'Bakri Id',
  '2026-08-15': 'Independence Day',
  '2026-10-02': 'Mahatma Gandhi Jayanti',
  '2026-10-20': 'Dussehra',
  '2026-11-09': 'Diwali Balipratipada',
  '2026-11-25': 'Guru Nanak Jayanti',
  '2026-12-25': 'Christmas'
};

const APP_TRANSLATIONS: Record<string, Record<string, string>> = {
  English: {
    appSettings: "App Settings",
    configSecurity: "Configuration & Security",
    userProfile: "User Profile",
    verifiedTrader: "Verified Trader",
    askJournalAi: "Ask Journal AI",
    signInGoogle: "Sign in with Google",
    googleDriveBackup: "Google Drive Backup",
    cloudSyncOffline: "Cloud Sync Offline",
    holidayColorThemeAccent: "Holiday Color Theme Accent",
    applyToSatSun: "Apply to Sat & Sun",
    backTheApp: "Back the App",
    supportFounder: "Support Founder",
    saveToMemory: "Save to Memory",
    restoreBackup: "Restore Backup",
    netProfitLoss: "Net Profit & Loss",
    winRateVol: "Win Rate / Vol",
    calendarHolidayColors: "Calendar Holiday Colors",
    noSavedTradesYet: "No saved trades yet",
    clear: "Clear",
    close: "Close",
    save: "Save",
    journal: "Journal",
    analytics: "Analytics",
    calendar: "Calendar",
    funds: "Funds",
    settings: "Settings",
    welcomeBack: "Welcome Back",
    totalTrades: "Total Trades",
    winRate: "Win Rate",
    lossTrades: "Loss Trades",
    winTrades: "Win Trades",
    netProfit: "Net Profit",
    activePnL: "Active P&L",
    newTradeLogs: "New Trade Logs",
    responseLanguage: "App / AI Language",
    gmailIntegration: "Gmail Workspace Integration",
    scanBrokerTrades: "Scan Inbox for Trades",
    recentBrokerEmails: "Recent Broker Emails & Alerts",
    sendPerfReport: "Send Performance Email Report",
    recipientEmail: "Recipient Email Address",
    customReportNote: "Custom Message / Note",
    sendMimeReport: "Send Email Report",
    importToJournal: "Import to Journal"
  },
  Hindi: {
    appSettings: "ऐप सेटिंग्स",
    configSecurity: "कॉन्फ़िगरेशन और सुरक्षा",
    userProfile: "उपयोगकर्ता प्रोफ़ाइल",
    verifiedTrader: "सत्यापित ट्रेडर",
    askJournalAi: "जर्नल एआई से पूछें",
    signInGoogle: "गूगल से साइन इन करें",
    googleDriveBackup: "गूगल ड्राइव बैकअप",
    cloudSyncOffline: "क्लाउड सिंक ऑफ़लाइन",
    holidayColorThemeAccent: "छुट्टी का रंग थीम",
    applyToSatSun: "शनि और रवि पर लागू करें",
    backTheApp: "ऐप का समर्थन करें",
    supportFounder: "संस्थापक का समर्थन करें",
    saveToMemory: "मेमोरी में सहेजें",
    restoreBackup: "बैकअप पुनर्स्थापित करें",
    netProfitLoss: "कुल लाभ और हानि",
    winRateVol: "जीत दर / वॉल्यूम",
    calendarHolidayColors: "कैलेंडर छुट्टी के रंग",
    noSavedTradesYet: "अभी तक कोई लॉग नहीं है",
    clear: "साफ करें",
    close: "बंद करें",
    save: "सहेजें",
    journal: "जर्नल",
    analytics: "विश्लेषण",
    calendar: "कैलेंडर",
    funds: "फंड",
    settings: "सेटिंग्स",
    welcomeBack: "आपका स्वागत है",
    totalTrades: "कुल ट्रेड",
    winRate: "जीत दर",
    lossTrades: "नुकसान वाले ट्रेड",
    winTrades: "जीत वाले ट्रेड",
    netProfit: "शुद्ध लाभ",
    activePnL: "सक्रिय पीएंडएल",
    newTradeLogs: "नया ट्रेड लॉग",
    responseLanguage: "ऐप / एआई भाषा"
  },
 Tamil: {
    appSettings: "செயலி அமைப்புகள்",
    configSecurity: "கட்டமைப்பு & பாதுகாப்பு",
    userProfile: "பயனர் சுயவிவரம்",
    verifiedTrader: "சரிபார்க்கப்பட்ட வர்த்தகர்",
    askJournalAi: "ஜர்னல் AI-யிடம் கேளுங்கள்",
    signInGoogle: "கூகுள் மூலம் உள்நுழைக",
    googleDriveBackup: "கூகுள் டிரைவ் காப்புப்பிரதி",
    cloudSyncOffline: "கிளவுட் ஒத்திசைவு ஆஃப்லைன்",
    holidayColorThemeAccent: "விடுமுறை வண்ண தீம்",
    applyToSatSun: "சனி & ஞாயிறுக்குப் பயன்படுத்துக",
    backTheApp: "செயலியை ஆதரிக்கவும்",
    supportFounder: "நிறுவனரை ஆதரிக்கவும்",
    saveToMemory: "நினைவகத்தில் சேமி",
    restoreBackup: "மீட்டமைக்க",
    netProfitLoss: "நிகர லாபம் & நஷ்டம்",
    winRateVol: "வெற்றி விகிதம் / அளவு",
    calendarHolidayColors: "நாட்காட்டி விடுமுறை நிறங்கள்",
    noSavedTradesYet: "இன்னும் வர்த்தகங்கள் இல்லை",
    clear: "அழி",
    close: "மூடு",
    save: "சேமி",
    journal: "டைரி",
    analytics: "பகுப்பாய்வு",
    calendar: "நாட்காட்டி",
    funds: "நிதி",
    settings: "அமைப்புகள்",
    welcomeBack: "வரவேற்கிறோம்",
    totalTrades: "மொத்த வர்த்தகங்கள்",
    winRate: "வெற்றி விகிதம்",
    lossTrades: "நஷ்ட வர்த்தகங்கள்",
    winTrades: "வெற்றி வர்த்தகங்கள்",
    netProfit: "நிகர லாபம்",
    activePnL: "செயலில் உள்ள P&L",
    newTradeLogs: "புதிய வர்த்தக பதிவு",
    responseLanguage: "செயலி / AI மொழி"
  },
  Telugu: {
    appSettings: "యాప్ సెట్టింగ్‌లు",
    configSecurity: "కాన్ఫిగరేషన్ & భద్రత",
    userProfile: "యూజర్ ప్రొఫైల్",
    verifiedTrader: "ధృవీకరించబడిన ట్రేడర్",
    askJournalAi: "జర్నల్ AIని అడగండి",
    signInGoogle: "గూగుల్‌తో సైన్ ఇన్",
    googleDriveBackup: "గూగుల్ డ్రైవ్ బ్యాకప్",
    cloudSyncOffline: "క్లౌడ్ సమకాలీకరణ ఆఫ్‌లైన్",
    holidayColorThemeAccent: "సెలవు దినాల రంగు థీమ్",
    applyToSatSun: "శని & ఆదివారాలకు వర్తింపజేయి",
    backTheApp: "యాప్‌కు మద్దతు ఇవ్వండి",
    supportFounder: "వ్యవస్థాపకునికి మద్దతు ఇవ్వండి",
    saveToMemory: "మెమరీలో సేవ్ చేయి",
    restoreBackup: "బ్యాకప్ పునరుద్ధరించు",
    netProfitLoss: "నికర లాభం & నష్టం",
    winRateVol: "విజయ శాతం / పరిమాణం",
    calendarHolidayColors: "క్యాలెండర్ సెలవు రంగులు",
    noSavedTradesYet: "ఇంకా ట్రేడ్‌లు లేవు",
    clear: "క్లియర్",
    close: "మూసివేయి",
    save: "సేవ్ చేయి",
    journal: "జర్నల్",
    analytics: "విశ్లేషణ",
    calendar: "క్యాలெండర్",
    funds: "ధన నిధులు",
    settings: "సెట్టింగులు",
    welcomeBack: "స్వాగతం",
    totalTrades: "మొత్తం ట్రేడ్‌లు",
    winRate: "విజయ రేటు",
    lossTrades: "నష్ట ట్రేడ్‌లు",
    winTrades: "విజయవంతమైన ట్రేడ్‌లు",
    netProfit: "నికర లాభం",
    activePnL: "సక్రియ P&L",
    newTradeLogs: "కొత్త ట్రేడ్ లాగ్",
    responseLanguage: "యాప్ / AI భాష"
  },
  Malayalam: {
    appSettings: "ആപ്പ് ക്രമീകരണങ്ങൾ",
    configSecurity: "കോൺഫിഗറേഷനും സുരക്ഷയും",
    userProfile: "യൂസർ പ്രൊഫൈൽ",
    verifiedTrader: "വെരിഫൈഡ് ട്രേഡർ",
    askJournalAi: "ജേണൽ AIയോട് ചോദിക്കുക",
    signInGoogle: "ഗൂഗിൾ വഴി ലോഗിൻ ചെയ്യുക",
    googleDriveBackup: "ഗൂഗിൾ ഡ്രൈവ് ബാക്കപ്പ്",
    cloudSyncOffline: "ക്ലൗഡ് സമന്വയം ഓഫ്‌ലൈനാണ്",
    holidayColorThemeAccent: "അവധിദിന വർണ്ണ തീം",
    applyToSatSun: "ശനി, ഞായർ ദിവസങ്ങളിൽ ബാധകമാക്കുക",
    backTheApp: "ആപ്പിനെ പിന്തുണയ്ക്കുക",
    supportFounder: "സ്ഥാപകനെ പിന്തുണയ്ക്കുക",
    saveToMemory: "മെമ്മറിയിലേക്ക് സേവ് ചെയ്യുക",
    restoreBackup: "ബാക്കപ്പ് പുനഃസ്ഥാപിക്കുക",
    netProfitLoss: "ആകെ ലാഭവും നഷ്ടവും",
    winRateVol: "വിജയ നിരക്ക് / വോളിയം",
    calendarHolidayColors: "കലണ്ടർ അവധി നിറങ്ങൾ",
    noSavedTradesYet: "ട്രേഡുകളൊന്നും ഇതുവരെ രേഖപ്പെടുത്തിയിട്ടില്ല",
    clear: "വ്യത്തിയാക്കുക",
    close: "അടയ്ക്കുക",
    save: "സേവ് ചെയ്യുക",
    journal: "ജേണൽ",
    analytics: "വിശകലനം",
    calendar: "കലണ്ടർ",
    funds: "ഫണ്ട്",
    settings: "സെറ്റിങ്സ്",
    welcomeBack: "വീണ്ടും സ്വാഗതം",
    totalTrades: "ആകെ ട്രേഡുകൾ",
    winRate: "വിജയ ശതമാനം",
    lossTrades: "നഷ്ടം വന്ന ട്രേഡുകൾ",
    winTrades: "ലാഭം വന്ന ട്രേഡുകൾ",
    netProfit: "അറ്റാദായം",
    activePnL: "ലൈവ് P&L",
    newTradeLogs: "പുതിയ ട്രേഡ് എൻട്രി",
    responseLanguage: "ആപ്പ് / AI ഭാഷ"
  },
  Kannada: {
    appSettings: "ಅಪ್ಲಿಕೇಶನ್ ಸೆಟ್ಟಿಂಗ್‌ಗಳು",
    configSecurity: "ಸಂರಚನೆ ಮತ್ತು ಭದ್ರತೆ",
    userProfile: "ಬಳಕೆದಾರರ ಪ್ರೊಫೈಲ್",
    verifiedTrader: "ಪರಿಶೀಲಿಸಿದ ಟ್ರೇಡರ್",
    askJournalAi: "ಜರ್ನಲ್ AI ಅನ್ನು ಕೇಳಿ",
    signInGoogle: "ಗೂಗಲ್‌ನೊಂದಿಗೆ ಸೈನ್ ಇನ್ ಮಾಡಿ",
    googleDriveBackup: "ಗೂಗಲ್ ಡ್ರೈವ್ ಬ್ಯಾಕಪ್",
    cloudSyncOffline: "ಕ್ಲೌಡ್ ಸಿಂಕ್ ಆಫ್‌ಲೈನ್",
    holidayColorThemeAccent: "ರಜಾದಿನದ ಬಣ್ಣದ ಥೀಮ್",
    applyToSatSun: "ಶನಿವಾರ ಮತ್ತು ಭಾನುವಾರಕ್ಕೆ ಅನ್ವಯಿಸಿ",
    backTheApp: "ಅಪ್ಲಿಕೇಶನ್ ಬೆಂಬಲಿಸಿ",
    supportFounder: "ಸಂಸ್ಥಾಪಕರನ್ನು ಬೆಂಬಲಿಸಿ",
    saveToMemory: "ಮೆಮೊರಿಗೆ ಉಳಿಸಿ",
    restoreBackup: "ಬ್ಯಾಕಪ್ ಮರುಸ್ಥಾಪಿಸಿ",
    netProfitLoss: "ನಿವ್ವಳ ಲಾಭ ಮತ್ತು నಷ್ಟ",
    winRateVol: "ಗೆಲುವಿನ ದರ / ವ್ಯಾಪಾರ",
    calendarHolidayColors: "ಕ್ಯಾಲೆಂಡರ್ ರಜಾದಿನ ಬಣ್ಣಗಳು",
    noSavedTradesYet: "ಇನ್ನೂ ಯಾವುದೇ ವ್ಯಾಪಾರಗಳಿಲ್ಲ",
    clear: "ತೆರವುಗೊಳಿಸು",
    close: "ಮುಚ್ಚು",
    save: "ಉಳಿಸಿ",
    journal: "ಜರ್ನಲ್",
    analytics: "ವಿಶ್ಲೇಷಣೆ",
    calendar: "ಕ್ಯಾಲೆಂಡರ್",
    funds: "ನಿಧಿಗಳು",
    settings: "ಸೆಟ್ಟಿಂಗ್ಸ್",
    welcomeBack: "ಸ್ವಾಗತ",
    totalTrades: "ಒಟ್ಟು ಟ್ರೇಡ್‌ಗಳು",
    winRate: "ಗೆಲುವಿನ ದರ",
    lossTrades: "ನಷ್ಟದ ಟ್ರೇಡ್‌ಗಳು",
    winTrades: "ಗೆದ್ದ ಟ್ರೇಡ್‌ಗಳು",
    netProfit: "ನಿವ್ವಳ ಲಾಭ",
    activePnL: "ಸಕ್ರಿಯ P&L",
    newTradeLogs: "ಹೊಸ ಟ್ರೇಡ್ ಲಾಗ್",
    responseLanguage: "ಅಪ್ಲಿಕೇಶನ್ / AI ಭಾಷೆ"
  },
  Bengali: {
    appSettings: "অ্যাপ সেটিংস",
    configSecurity: "কনফিগারেশন ও নিরাপত্তা",
    userProfile: "ব্যবহারকারী প্রোফাইল",
    verifiedTrader: "যাচাইকৃত ট্রেডার",
    askJournalAi: "জার্নাল এআই-কে জিজ্ঞাসা করুন",
    signInGoogle: "গুগল দিয়ে লগইন করুন",
    googleDriveBackup: "গুগল ড্রাইভ ব্যাকআপ",
    cloudSyncOffline: "ক্লাউড সিঙ্ক অফলাইন",
    holidayColorThemeAccent: "ছুটির দিনের কালার থিম",
    applyToSatSun: "শনি ও রবিবারেও প্রয়োগ করুন",
    backTheApp: "অ্যাপ্লিকেশনটিকে সমর্থন করুন",
    supportFounder: "প্রতিষ্ঠাতাকে সমর্থন করুন",
    saveToMemory: "মেমোরিতে সেভ করুন",
    restoreBackup: "ব্যাকআপ রিস্টোর করুন",
    netProfitLoss: "মোট লাভ ও ক্ষতি",
    winRateVol: "জয়ের হার / ভলিউম",
    calendarHolidayColors: "ক্যালেন্ডার ছুটির রঙ",
    noSavedTradesYet: "এখনো কোনো ট্রেড সেভ করা নেই",
    clear: "মুছে ফেলুন",
    close: "বন্ধ করুন",
    save: "সেভ করুন",
    journal: "জার্নাল",
    analytics: "বিশ্লেষণ",
    calendar: "ক্যালেন্ডার",
    funds: "তহবিল",
    settings: "সেটিংস",
    welcomeBack: "স্বাগতম",
    totalTrades: "মোট ট্রেড",
    winRate: "জয়ের হার",
    lossTrades: "লোকসানি ট্রেড",
    winTrades: "জয়ী ট্রেড",
    netProfit: "নিট লাভ",
    activePnL: "সক্রিয় P&L",
    newTradeLogs: "নতুন ট্রেড এন্ট্রি",
    responseLanguage: "অ্যাপ / এআই ভাষা"
  },
  Gujarati: {
    appSettings: "એપ સેટિંગ્સ",
    configSecurity: "રૂપરેખાંકન અને સુરક્ષા",
    userProfile: "વપરાશકર્તા પ્રોફાઇલ",
    verifiedTrader: "વેરિફાઇડ ટ્રેડર",
    askJournalAi: "જર્નલ AI ને પૂછો",
    signInGoogle: "ગૂગલ થી સાઇન ઇન કરો",
    googleDriveBackup: "ગૂગલ ડ્રાઇવ બેકઅપ",
    cloudSyncOffline: "ક્લાઉડ સમન્વયન ઑફલાઇન",
    holidayColorThemeAccent: "રજાઓ નો કલર થીમ",
    applyToSatSun: "શનિ અને રવિ પર લાગુ કરો",
    backTheApp: "એપને સપોર્ટ કરો",
    supportFounder: "સ્થાપકને સપોર્ટ કરો",
    saveToMemory: "મેમરીમાં સેવ કરો",
    restoreBackup: "બેકઅપ પુનર્સ્થાપિત કરો",
    netProfitLoss: "ચોખ્ખો નફો અને નુકસાન",
    winRateVol: "જીતનો દર / વોલ્યુમ",
    calendarHolidayColors: "કેલેન્ડર રજા ના રંગો",
    noSavedTradesYet: "હજી સુધી કોઈ ટ્રેડ સેવ નથી લોગ કરાયો",
    clear: "સાફ કરો",
    close: "બંધ કરો",
    save: "સેવ કરો",
    journal: "જર્નल",
    analytics: "વિશ્લેષણ",
    calendar: "કેલેન્ડર",
    funds: "ફંડ",
    settings: "સેટિંગ્સ",
    welcomeBack: "આપનું સ્વાગત છે",
    totalTrades: "કુલ ટ્રેડ",
    winRate: "જીતનો ગુણોત્તર",
    lossTrades: "નુકસાની ટ્રેડ",
    winTrades: "નફાકારક ટ્રેડ",
    netProfit: "ચોખ્ખો નફો",
    activePnL: "લાઈવ P&L",
    newTradeLogs: "નવી ટ્રેડ લોગ એન્ટ્રી",
    responseLanguage: "એપ / AI ભાષા"
  },
  Marathi: {
    appSettings: "अॅप सेटिंग्ज",
    configSecurity: "कॉन्फिगरेशन आणि सुरक्षा",
    userProfile: "वापरकर्ता प्रोफाइल",
    verifiedTrader: "सत्यापित ट्रेडर",
    askJournalAi: "जर्नल एआयला विचारा",
    signInGoogle: "गुगलने साइन इन करा",
    googleDriveBackup: "गुगल ड्राइव्हे बॅकअप",
    cloudSyncOffline: "क्लाउड सिंक ऑफलाइन",
    holidayColorThemeAccent: "सुट्टीची रंग थीम",
    applyToSatSun: "शनि आणि रवि लागू करा",
    backTheApp: "अॅपला पाठिंबा द्या",
    supportFounder: "संस्थापकाला पाठिंबा द्या",
    saveToMemory: "मेमरीमध्ये सेव्ह करा",
    restoreBackup: "बॅकअप पुनर्संचयित करा",
    netProfitLoss: "निव्वळ नफा आणि तोटा",
    winRateVol: "जीत दर / वॉल्यूम",
    calendarHolidayColors: "कॅलेंडर सुट्टीचे रंग",
    noSavedTradesYet: "अद्याप कोणताही ट्रेड सेव्ह केलेला नाही",
    clear: "स्वच्छ करा",
    close: "बंद करा",
    save: "सेव्ह करा",
    journal: "जर्नल",
    analytics: "विश्लेषण",
    calendar: "कॅलेंडर",
    funds: "निधी",
    settings: "सेटिंग्ज",
    welcomeBack: "आपले स्वागत आहे",
    totalTrades: "एकूण ट्रेड्स",
    winRate: "जीत दर",
    lossTrades: "तोट्यातील ट्रेड्स",
    winTrades: "नफ्यातील ट्रेड्स",
    netProfit: "निव्वळ नफा",
    activePnL: "सक्रिय P&L",
    newTradeLogs: "नवीन ट्रेड नोंदणी",
    responseLanguage: "अॅप / एआय भाषा"
  },
  Punjabi: {
    appSettings: "ਐਪ ਸੈਟਿੰਗਾਂ",
    configSecurity: "ਸੰਰਚਨਾ ਅਤੇ ਸੁਰੱਖਿਆ",
    userProfile: "ਉਪਭੋਗਤਾ ਪ੍ਰੋਫਾਈਲ",
    verifiedTrader: "ਪ੍ਰਮਾਣਿਤ ਟ੍ਰੇਡਰ",
    askJournalAi: "ਜਰਨਲ AI ਨੂੰ ਪੁੱਛੋ",
    signInGoogle: "ਗੂਗਲ ਨਾਲ ਸਾਈਨ ਇਨ ਕਰੋ",
    googleDriveBackup: "ਗੂਗਲ ਡਰਾਈਵ ਬੈਕਅੱਪ",
    cloudSyncOffline: "ਕਲਾਊਡ ਸਿੰਕ ਔਫਲਾਈਨ",
    holidayColorThemeAccent: "ਛੁੱਟੀ ਦਾ ਰੰਗ ਥੀਮ",
    applyToSatSun: "ਸ਼ਨਿੱਛਰਵਾਰ ਅਤੇ ਐਤਵਾਰ 'ਤੇ ਲਾਗੂ ਕਰੋ",
    backTheApp: "ਐਪ ਦਾ ਸਮਰਥਨ ਕਰੋ",
    supportFounder: "ਸੰਸਥਾਪਕ ਦਾ ਸਮਰਥਨ ਕਰੋ",
    saveToMemory: "ਮੈਮੋਰੀ ਵਿੱਚ ਸੁਰੱਖਿਅਤ ਕਰੋ",
    restoreBackup: "ਬੈਕਅੱਪ ਰੀਸਟੋਰ ਕਰੋ",
    netProfitLoss: "ਕੁੱਲ ਮੁਨਾਫਾ ਅਤੇ ਨੁਕਸਾਨ",
    winRateVol: "ਜਿੱਤ ਦੀ ਦਰ / ਵੌਲਯੂਮ",
    calendarHolidayColors: "ਕੈਲੰਡਰ ਛੁੱਟੀਆਂ ਦੇ ਰੰਗ",
    noSavedTradesYet: "ਅਜੇ ਤੱਕ ਕੋਈ ਟ੍ਰੇਡ ਸੁਰੱਖਿਅਤ ਨਹੀਂ ਹੈ",
    clear: "ਸਾਫ਼ ਕਰੋ",
    close: "ਬੰਦ ਕਰੋ",
    save: "ਸੁਰੱਖਿਅਤ ਕਰੋ",
    journal: "ਜਰਨਲ",
    analytics: "ਵਿਸ਼ਲੇਸ਼ਣ",
    calendar: "ਕੈਲੰਡਰ",
    funds: "ਫੰਡ",
    settings: "ਸੈਟਿੰਗਾਂ",
    welcomeBack: "ਜੀ ਆਇਆਂ ਨੂੰ",
    totalTrades: "ਕੁੱਲ ਟ੍ਰੇਡ",
    winRate: "ਜਿੱਤ ਦੀ ਦਰ",
    lossTrades: "ਘਾਟੇ ਵਾਲੇ ਟ੍ਰੇਡ",
    winTrades: "ਮੁਨਾਫੇ ਵਾਲੇ ਟ੍ਰੇਡ",
    netProfit: "ਸ਼ੁੱਧ ਮੁਨਾਫਾ",
    activePnL: "ਸਰਗਰਮ P&L",
    newTradeLogs: "ਨਵਾਂ ਟ੍ਰੇਡ ਲੌਗ",
    responseLanguage: "ਐਪ / AI ਭਾਸ਼า"
  }
};

export default function App() {
  // State
  const [mode, setMode] = useState<'equity' | 'options'>('equity');
  const [type, setType] = useState<'long' | 'short'>('long');
  const [entryPrice, setEntryPrice] = useState<string>('');
  const [exitPrice, setExitPrice] = useState<string>('');
  const [quantity, setQuantity] = useState<string>(''); // Quantity for equity, Lots for options
  const [lotSize, setLotSize] = useState<string>('65'); // Default lot size
  const [leverage, setLeverage] = useState<number>(1);
  const [feePercentage, setFeePercentage] = useState<string>('0.05'); // Adjusted default for India
  const [entryDate, setEntryDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [exitDate, setExitDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [strategy, setStrategy] = useState<string>('Intraday');
  const [notes, setNotes] = useState<string>('');
  const [tradeSetup, setTradeSetup] = useState<string>('');
  const [entryExitRationale, setEntryExitRationale] = useState<string>('');
  const [lessonsLearned, setLessonsLearned] = useState<string>('');
  const [graphTimeframe, setGraphTimeframe] = useState<'Day' | 'Week' | 'Month' | 'Year'>('Month');
  const [selectedBenchmark, setSelectedBenchmark] = useState<'Nifty' | 'Sensex' | 'Both' | 'None'>('Nifty');
  const [history, setHistory] = useState<SavedTrade[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [showCalendar, setShowCalendar] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showFunds, setShowFunds] = useState(false);
  const [totalFunds, setTotalFunds] = useState<number>(0);
  const [openingBalance, setOpeningBalance] = useState<number>(0);
  const [fundInput, setFundInput] = useState<string>('');
  const [fundTransactions, setFundTransactions] = useState<FundTransaction[]>([]);
  const [editingFundTxId, setEditingFundTxId] = useState<string | null>(null);
  const [fundNotesInput, setFundNotesInput] = useState<string>('');
  const [fundTab, setFundTab] = useState<'all' | 'deposit' | 'withdrawal'>('all');
  const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
  const [showOnlyFavorites, setShowOnlyFavorites] = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [showAdvancedJournal, setShowAdvancedJournal] = useState(false);
  const [userName, setUserName] = useState<string>('Trader');
  const [theme, setTheme] = useState<'light' | 'dark'>('light');

  // Calendar interactive states
  const [selectedCalendarDate, setSelectedCalendarDate] = useState<string | null>(null);
  const [editingTradeId, setEditingTradeId] = useState<string | null>(null);
  const [quickPnLAmount, setQuickPnLAmount] = useState<string>('');
  const [quickPnLType, setQuickPnLType] = useState<'profit' | 'loss'>('profit');
  const [quickPnLTicker, setQuickPnLTicker] = useState<string>('');
  const [quickPnLNotes, setQuickPnLNotes] = useState<string>('');

  const handleSelectCalendarDate = (dateStr: string | null) => {
    setSelectedCalendarDate(dateStr);
    setEditingTradeId(null);
    setQuickPnLAmount('');
    setQuickPnLType('profit');
    setQuickPnLTicker('');
    setQuickPnLNotes('');
  };
  const [customHolidays, setCustomHolidays] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem('trader_custom_holidays');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return {};
      }
    }
    return {};
  });

  useEffect(() => {
    localStorage.setItem('trader_custom_holidays', JSON.stringify(customHolidays));
  }, [customHolidays]);

  const handleSaveCustomHoliday = (dStr: string, name: string) => {
    if (!dStr || !name.trim()) return;
    setCustomHolidays(prev => ({
      ...prev,
      [dStr]: name.trim()
    }));
  };

  const handleRemoveCustomHoliday = (dStr: string) => {
    if (!dStr) return;
    setCustomHolidays(prev => {
      const copy = { ...prev };
      delete copy[dStr];
      return copy;
    });
  };

  const getHolidayName = (dStr: string) => {
    if (customHolidays[dStr]) return customHolidays[dStr];
    if (MARKET_HOLIDAYS_PRESET[dStr]) return MARKET_HOLIDAYS_PRESET[dStr];
    const dayOfWeek = new Date(dStr + 'T12:00:00').getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) return 'Weekend';
    return null;
  };

  // Support Founder states
  const [supportStep, setSupportStep] = useState<'menu' | 'info' | 'pay' | 'success'>('menu');
  const [supportName, setSupportName] = useState<string>('');
  const [supportLocation, setSupportLocation] = useState<string>('');
  const [supportEmail, setSupportEmail] = useState<string>('');
  const [supportAmount, setSupportAmount] = useState<string>('250');
  const [supportCustomAmount, setSupportCustomAmount] = useState<string>('');
  const [isProcessingPayment, setIsProcessingPayment] = useState<boolean>(false);
  const [activeUpiUrl, setActiveUpiUrl] = useState<string>('');
  const [universalUpiUrl, setUniversalUpiUrl] = useState<string>('');
  const [activeAppName, setActiveAppName] = useState<string>('');
  const [copiedUpi, setCopiedUpi] = useState<boolean>(false);

  // Trade Journal state
  const [isManualJournalOpen, setIsManualJournalOpen] = useState(false);
  const [selectedJournalTrade, setSelectedJournalTrade] = useState<SavedTrade | null>(null);

  // States for PnL chart drag ranges & dynamic analysis filtering
  const [historyRefLeft, setHistoryRefLeft] = useState<number | null>(null);
  const [historyRefRight, setHistoryRefRight] = useState<number | null>(null);
  const [historyZoomRange, setHistoryZoomRange] = useState<[number, number] | null>(null);

  // Projection chart zoom/range selection states
  const [projectionRefLeft, setProjectionRefLeft] = useState<number | null>(null);
  const [projectionRefRight, setProjectionRefRight] = useState<number | null>(null);
  const [projectionZoomRange, setProjectionZoomRange] = useState<[number, number] | null>(null);

  // Export menu visual state
  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);

  // States for manual trade log form
  const [manualTicker, setManualTicker] = useState('');
  const [manualMode, setManualMode] = useState<'equity' | 'options'>('equity');
  const [manualType, setManualType] = useState<'long' | 'short'>('long');
  const [manualEntryPrice, setManualEntryPrice] = useState('');
  const [manualExitPrice, setManualExitPrice] = useState('');
  const [manualQuantity, setManualQuantity] = useState('');
  const [manualLotSize, setManualLotSize] = useState('65');
  const [manualStrategy, setManualStrategy] = useState('Intraday');
  const [manualProfitCustom, setManualProfitCustom] = useState(false);
  const [manualProfitAmount, setManualProfitAmount] = useState('');
  const [manualEntryDate, setManualEntryDate] = useState(new Date().toISOString().split('T')[0]);
  const [manualExitDate, setManualExitDate] = useState(new Date().toISOString().split('T')[0]);
  const [manualTradeSetup, setManualTradeSetup] = useState('');
  const [manualEntryExitRationale, setManualEntryExitRationale] = useState('');
  const [manualLessonsLearned, setManualLessonsLearned] = useState('');
  const [manualNotes, setManualNotes] = useState('');

  const parsedQty = quantity === '' ? 1 : (parseFloat(quantity) || 0);

  // Auth State
  const [authUser, setAuthUser] = useState<FirebaseUser | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [isBackingUp, setIsBackingUp] = useState(false);
  const [isRestoring, setIsRestoring] = useState(false);
  const [needsAuth, setNeedsAuth] = useState(true);

  // Google Drive Sync States
  const [driveSyncStatus, setDriveSyncStatus] = useState<'idle' | 'checking' | 'backing-up' | 'restoring' | 'success' | 'error'>('idle');
  const [driveLastBackup, setDriveLastBackup] = useState<string | null>(() => localStorage.getItem('trade_track_drive_last_backup'));
  const [autoBackupEnabled, setAutoBackupEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem('trade_track_auto_backup_enabled');
    return saved === null ? true : saved === 'true';
  });
  const [holidayColorTheme, setHolidayColorTheme] = useState<'slate' | 'rose' | 'amber' | 'indigo' | 'emerald' | 'violet'>('rose');
  const [includeWeekendsInHolidayTheme, setIncludeWeekendsInHolidayTheme] = useState<boolean>(true);
  const [hasCheckedCloudBackupOnLoad, setHasCheckedCloudBackupOnLoad] = useState<boolean>(false);
  const [cloudNotification, setCloudNotification] = useState<{
    show: boolean;
    type: 'success' | 'restore_success' | 'error' | 'info';
    message: string;
  } | null>(null);

  // Gmail Workspace Integration States
  const [gmailEmails, setGmailEmails] = useState<GmailMessageSnippet[]>([]);
  const [isScanningGmail, setIsScanningGmail] = useState(false);
  const [gmailError, setGmailError] = useState<string | null>(null);
  const [reportRecipient, setReportRecipient] = useState('');
  const [reportSubject, setReportSubject] = useState('My Live Trading Performance Summary');
  const [reportCustomMessage, setReportCustomMessage] = useState('');
  const [isSendingReport, setIsSendingReport] = useState(false);

  // UI Dialog/Toast states
  const [confirmDialog, setConfirmDialog] = useState<{
    show: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
  } | null>(null);

  const [toastMessage, setToastMessage] = useState<{
    show: boolean;
    message: string;
    type: 'success' | 'error' | 'info';
  } | null>(null);

  const showToastAlert = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToastMessage({
      show: true,
      message,
      type
    });
  };

  // AI Chat States
  const INDIAN_LANGUAGES = [
    { code: 'English', name: 'English', native: 'English' },
    { code: 'Hindi', name: 'Hindi', native: 'हिन्दी' },
    { code: 'Tamil', name: 'Tamil', native: 'தமிழ்' },
    { code: 'Telugu', name: 'Telugu', native: 'తెలుగు' },
    { code: 'Malayalam', name: 'Malayalam', native: 'മലയാളം' },
    { code: 'Kannada', name: 'Kannada', native: 'ಕನ್ನಡ' },
    { code: 'Bengali', name: 'Bengali', native: 'বাংলা' },
    { code: 'Gujarati', name: 'Gujarati', native: 'ગુજરાતી' },
    { code: 'Marathi', name: 'Marathi', native: 'मराठी' },
    { code: 'Punjabi', name: 'Punjabi', native: 'ਪੰਜਾਬੀ' }
  ];

  const [showAIChat, setShowAIChat] = useState(false);
  const [aiLanguage, setAiLanguage] = useState<string>(() => {
    return localStorage.getItem('trade_track_ai_language') || 'English';
  });

  const [aiMessages, setAiMessages] = useState<{ role: 'user' | 'assistant'; content: string }[]>([
    { role: 'assistant', content: 'Hi there! I am your Trading Journal AI Assistant. Ask me anything about your trade stats, strategies, risk rules, or trading psychology.' }
  ]);
  const [aiInput, setAiInput] = useState('');
  const [aiSending, setAiSending] = useState(false);

  // Sync AI language choice
  useEffect(() => {
    localStorage.setItem('trade_track_ai_language', aiLanguage);
  }, [aiLanguage]);

  // Translation helper
  const t = (key: string): string => {
    return APP_TRANSLATIONS[aiLanguage]?.[key] || APP_TRANSLATIONS['English']?.[key] || key;
  };

  const handleSendAIMessage = async (customPrompt?: string) => {
    const textToSend = customPrompt || aiInput;
    if (!textToSend.trim() || aiSending) return;

    const userMsg = { role: 'user' as const, content: textToSend };
    setAiMessages(prev => [...prev, userMsg]);
    if (!customPrompt) setAiInput('');
    setAiSending(true);

    try {
      // Build trading history summary context to keep the model realistic
      let tradeContext = "Here is my trading history context for analysis:\n";
      if (!history || history.length === 0) {
        tradeContext += "I have not logged any trades in my journal yet.";
      } else {
        const totalTrades = history.length;
        const totalProfitVal = history.reduce((sum, t) => sum + (t.profit || 0), 0);
        const winTrades = history.filter(t => (t.profit || 0) > 0);
        const lossTrades = history.filter(t => (t.profit || 0) <= 0);
        const winRate = ((winTrades.length / totalTrades) * 100).toFixed(1);
        
        tradeContext += `Summary Stats:
- Total Trades: ${totalTrades}
- Net Profit/Loss: ₹${totalProfitVal.toFixed(2)}
- Win Trades: ${winTrades.length}
- Loss Trades: ${lossTrades.length}
- Win Rate: ${winRate}%
- Total Funds: ₹${totalFunds}

Individual Trade Logs (recent trades context):
${history.slice(-10).map((t, idx) => `- Trade ${idx+1}: ${t.ticker || 'N/A'}, Profit: ₹${(t.profit || 0).toFixed(2)}, Strategy: ${t.strategy || 'N/A'}, Lots/Qty: ${t.quantity}, Notes: ${t.notes || "None"}`).join('\n')}
`;
      }

      // Add special instruction when an Indian language is selected
      const languageInstruction = aiLanguage !== 'English' 
        ? `\n\n[LANGUAGE MANDATE]: The user has chosen to communicate in the "${aiLanguage}" language. You MUST respond completely in the ${aiLanguage} language using its native script (or standard local writing script). Translate all explanations, calculations, tables, psychological tips, and performance insights into fluent, natural ${aiLanguage}. Avoid complex English jargon where translated terms exist, but you may keep standard ticker codes (like TCS, Reliance, Nifty) or numbers unmodified.`
        : `\n\n[LANGUAGE MANDATE]: Please respond in premium professional English.`;

      const systemMsgContext = {
        role: 'user' as const,
        content: `[CONTEXT] The following is my current journal performance data. Keep this in mind when advising or answering any analytical questions. Be brief and supportive.\n\n${tradeContext}${languageInstruction}\n\n[USER QUERY]: ${textToSend}`
      };

      const payloadMessages = [
        ...aiMessages.map(m => ({ role: m.role, content: m.content })),
        systemMsgContext
      ];

      const res = await fetch("/api/gemini/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: payloadMessages })
      });

      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error || "Failed to contact Gemini");
      }

      const data = await res.json();
      setAiMessages(prev => [...prev, { role: 'assistant', content: data.reply }]);
    } catch (err: any) {
      console.error(err);
      setAiMessages(prev => [
        ...prev, 
        { role: 'assistant', content: `Sorry, I encountered an error: ${err.message || 'Unknown network error'}. Please verify that GEMINI_API_KEY is configured in your Secrets panel.` }
      ]);
    } finally {
      setAiSending(false);
    }
  };

  // Handle mobile and hardware back button actions to close panels instead of exiting the web page
  useEffect(() => {
    const isAnyOverlayOpen = showCalendar || showHistory || showFunds || showSettings || isManualJournalOpen || !!selectedJournalTrade || showAIChat;
    
    const handlePopState = (event: PopStateEvent) => {
      // If back is pressed, prevent closing the web app and instead close any open panels cleanly
      setShowCalendar(false);
      setShowHistory(false);
      setShowFunds(false);
      setShowSettings(false);
      setShowAIChat(false);
      setSupportStep('menu');
      setIsManualJournalOpen(false);
      setSelectedJournalTrade(null);
    };

    if (isAnyOverlayOpen) {
      if (!window.history.state?.isPanelOpen) {
        window.history.pushState({ isPanelOpen: true }, '');
      }
      window.addEventListener('popstate', handlePopState);
    } else {
      if (window.history.state?.isPanelOpen) {
        window.history.back();
      }
    }

    return () => {
      window.removeEventListener('popstate', handlePopState);
    };
  }, [showCalendar, showHistory, showFunds, showSettings, isManualJournalOpen, selectedJournalTrade, showAIChat]);

  // Load history and profile from localStorage
  useEffect(() => {
    const unsubscribe = initAuth(
      (user) => {
        setAuthUser(user);
        setNeedsAuth(false);
        if (user?.email) {
          setReportRecipient(user.email);
        }
      },
      () => {
        setAuthUser(null);
        setNeedsAuth(true);
      }
    );

    const savedHistory = localStorage.getItem('trade_history');
    const savedName = localStorage.getItem('user_name');
    const savedFunds = localStorage.getItem('trade_track_funds');
    const savedOpeningBalance = localStorage.getItem('trade_track_opening_balance');
    const savedTheme = localStorage.getItem('trade_track_theme') as 'light' | 'dark' | null;
    const savedFundTx = localStorage.getItem('trade_fund_transactions');
    const savedHolidayTheme = localStorage.getItem('trade_track_holiday_color_theme') as 'slate' | 'rose' | 'amber' | 'indigo' | 'emerald' | 'violet' | null;
    const savedIncludeWeekends = localStorage.getItem('trade_track_include_weekends_in_theme');
    
    if (savedHistory) {
      try {
        setHistory(JSON.parse(savedHistory));
      } catch (e) {
        console.error('Failed to parse history', e);
      }
    }

    if (savedFundTx) {
      try {
        setFundTransactions(JSON.parse(savedFundTx));
      } catch (e) {
        console.error('Failed to parse fund transactions', e);
      }
    }
    
    if (savedName) setUserName(savedName);
    if (savedFunds) setTotalFunds(parseFloat(savedFunds));
    if (savedOpeningBalance) setOpeningBalance(parseFloat(savedOpeningBalance));
    if (savedTheme) setTheme(savedTheme);
    if (savedHolidayTheme) setHolidayColorTheme(savedHolidayTheme);
    if (savedIncludeWeekends) setIncludeWeekendsInHolidayTheme(savedIncludeWeekends === 'true');

    return () => unsubscribe();
  }, []);

  // Save holiday settings
  useEffect(() => {
    localStorage.setItem('trade_track_holiday_color_theme', holidayColorTheme);
  }, [holidayColorTheme]);

  useEffect(() => {
    localStorage.setItem('trade_track_include_weekends_in_theme', String(includeWeekendsInHolidayTheme));
  }, [includeWeekendsInHolidayTheme]);

  // Save theme
  useEffect(() => {
    localStorage.setItem('trade_track_theme', theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'light' ? 'dark' : 'light');
  };

  // Save history
  useEffect(() => {
    localStorage.setItem('trade_history', JSON.stringify(history));
  }, [history]);

  // Save name
  useEffect(() => {
    localStorage.setItem('user_name', userName);
  }, [userName]);

  // Save funds
  useEffect(() => {
    localStorage.setItem('trade_track_funds', totalFunds.toString());
  }, [totalFunds]);

  // Save fund transactions
  useEffect(() => {
    localStorage.setItem('trade_fund_transactions', JSON.stringify(fundTransactions));
  }, [fundTransactions]);

  // Save opening balance
  useEffect(() => {
    localStorage.setItem('trade_track_opening_balance', openingBalance.toString());
  }, [openingBalance]);

  // Save Drive States
  useEffect(() => {
    if (driveLastBackup) {
      localStorage.setItem('trade_track_drive_last_backup', driveLastBackup);
    } else {
      localStorage.removeItem('trade_track_drive_last_backup');
    }
  }, [driveLastBackup]);

  useEffect(() => {
    localStorage.setItem('trade_track_auto_backup_enabled', String(autoBackupEnabled));
  }, [autoBackupEnabled]);

  // Auto Dismiss Cloud Notifications
  useEffect(() => {
    if (cloudNotification?.show) {
      const timer = setTimeout(() => {
        setCloudNotification(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [cloudNotification]);

  // Auto Dismiss Toast Messages
  useEffect(() => {
    if (toastMessage?.show) {
      const timer = setTimeout(() => {
        setToastMessage(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [toastMessage]);

  // Load & Auto-sync Google Drive cloud backup on app open
  useEffect(() => {
    const syncGoogleDriveOnLoad = async () => {
      const token = getAccessToken();
      if (!token || !authUser || hasCheckedCloudBackupOnLoad || !autoBackupEnabled) {
        return;
      }

      setHasCheckedCloudBackupOnLoad(true);
      setDriveSyncStatus('checking');

      try {
        const cloudData = await restoreBackup(token);
        
        if (cloudData) {
          // Check if local history exists
          const currentLocalHistoryJson = localStorage.getItem('trade_history');
          const isLocalHistoryEmpty = !currentLocalHistoryJson || JSON.parse(currentLocalHistoryJson).length === 0;

          if (isLocalHistoryEmpty) {
            setDriveSyncStatus('restoring');
            
            const restoredHistory = cloudData.history || [];
            setHistory(restoredHistory);
            
            if (cloudData.userName) {
              setUserName(cloudData.userName);
              localStorage.setItem('user_name', cloudData.userName);
            }
            if (cloudData.totalFunds !== undefined) {
              setTotalFunds(cloudData.totalFunds);
              localStorage.setItem('trade_track_funds', String(cloudData.totalFunds));
            }
            if (cloudData.openingBalance !== undefined) {
              setOpeningBalance(cloudData.openingBalance);
              localStorage.setItem('trade_track_opening_balance', String(cloudData.openingBalance));
            }
            if (cloudData.fundTransactions) {
              setFundTransactions(cloudData.fundTransactions);
              localStorage.setItem('trade_fund_transactions', JSON.stringify(cloudData.fundTransactions));
            }
            
            const backupDateStr = new Date(cloudData.timestamp).toLocaleString('en-IN');
            setDriveLastBackup(backupDateStr);
            setDriveSyncStatus('success');
            
            setCloudNotification({
              show: true,
              type: 'restore_success',
              message: `Cloud backup auto-restored! ${restoredHistory.length} trades synced.`
            });
          } else {
            // Local history exists. Auto back up local structure to modernise the Google Drive profile state
            setDriveSyncStatus('backing-up');
            
            const currentName = localStorage.getItem('user_name') || userName;
            const currentFundsStr = localStorage.getItem('trade_track_funds') || String(totalFunds);
            const currentOpenBalStr = localStorage.getItem('trade_track_opening_balance') || String(openingBalance);
            const currentTxStr = localStorage.getItem('trade_fund_transactions') || JSON.stringify(fundTransactions);
            const currentHistoryStr = localStorage.getItem('trade_history') || JSON.stringify(history);

            const latestHistory = currentHistoryStr ? JSON.parse(currentHistoryStr) : [];
            const latestTxs = currentTxStr ? JSON.parse(currentTxStr) : [];

            await createOrUpdateBackup(token, {
              history: latestHistory,
              userName: currentName,
              timestamp: Date.now(),
              totalFunds: parseFloat(currentFundsStr) || 0,
              openingBalance: parseFloat(currentOpenBalStr) || 0,
              fundTransactions: latestTxs
            });
            
            const nowStr = new Date().toLocaleString('en-IN');
            setDriveLastBackup(nowStr);
            setDriveSyncStatus('success');
            
            setCloudNotification({
              show: true,
              type: 'success',
              message: 'Google Drive cloud backup updated successfully!'
            });
          }
        } else {
          // No cloud file has been written yet. Force initial sync of local space
          setDriveSyncStatus('backing-up');
          
          await createOrUpdateBackup(token, {
            history,
            userName,
            timestamp: Date.now(),
            totalFunds,
            openingBalance,
            fundTransactions
          });
          
          const nowStr = new Date().toLocaleString('en-IN');
          setDriveLastBackup(nowStr);
          setDriveSyncStatus('success');
          
          setCloudNotification({
            show: true,
            type: 'success',
            message: 'First Google Drive backup created successfully!'
          });
        }
      } catch (err) {
        console.error('Auto Google Drive sync failed:', err);
        setDriveSyncStatus('error');
        setCloudNotification({
          show: true,
          type: 'error',
          message: 'Failed to auto-sync backup with Google Drive.'
        });
      }
    };

    if (authUser && !needsAuth && !hasCheckedCloudBackupOnLoad && autoBackupEnabled) {
      const timer = setTimeout(() => {
        syncGoogleDriveOnLoad();
      }, 1500);
      return () => clearTimeout(timer);
    }
  }, [authUser, needsAuth, hasCheckedCloudBackupOnLoad, autoBackupEnabled]);

  const getCurrentPaymentAmount = () => {
    return supportAmount === 'custom' ? (parseFloat(supportCustomAmount) || 0) : parseFloat(supportAmount);
  };

  const handleSimulatePayment = (appName: string) => {
    const amount = getCurrentPaymentAmount();
    if (amount <= 0) return;
    
    const encodedPn = encodeURIComponent("Jidhin Raju");
    const encodedTn = encodeURIComponent("Trade Track Support");
    
    // Create direct app URLs
    let appUrl = '';
    const cleanApp = appName.toLowerCase();
    if (cleanApp.includes('google') || cleanApp.includes('gpay')) {
      appUrl = `tez://upi/pay?pa=jidhinraju@ybl&pn=${encodedPn}&am=${amount}&tn=${encodedTn}&cu=INR`;
    } else if (cleanApp.includes('phone') || cleanApp.includes('phonepe')) {
      appUrl = `phonepe://pay?pa=jidhinraju@ybl&pn=${encodedPn}&am=${amount}&tn=${encodedTn}&cu=INR`;
    } else if (cleanApp.includes('paytm')) {
      appUrl = `paytmmp://pay?pa=jidhinraju@ybl&pn=${encodedPn}&am=${amount}&tn=${encodedTn}&cu=INR`;
    } else {
      appUrl = `upi://pay?pa=jidhinraju@ybl&pn=${encodedPn}&am=${amount}&tn=${encodedTn}&cu=INR`;
    }
    
    const universalUrl = `upi://pay?pa=jidhinraju@ybl&pn=${encodedPn}&am=${amount}&tn=${encodedTn}&cu=INR`;
    
    setActiveAppName(appName);
    setActiveUpiUrl(appUrl);
    setUniversalUpiUrl(universalUrl);
    setIsProcessingPayment(true);
    setCopiedUpi(false);

    // Direct redirection attempt
    const targetUrl = appUrl;
    try {
      window.location.href = targetUrl;
    } catch (e) {
      console.error("Direct redirect error:", e);
    }
  };

  // Calendar Helpers
  const calendarData = useMemo(() => {
    const map: Record<string, number> = {};
    history.forEach(t => {
      const date = new Date(t.timestamp).toISOString().split('T')[0];
      map[date] = (map[date] || 0) + t.profit;
    });
    return map;
  }, [history]);

  const tradesOnSelectedDate = useMemo(() => {
    if (!selectedCalendarDate) return [];
    return history.filter(t => {
      return new Date(t.timestamp).toISOString().split('T')[0] === selectedCalendarDate;
    });
  }, [history, selectedCalendarDate]);

  const handleSaveQuickPnL = () => {
    if (!selectedCalendarDate) return;
    const amount = parseFloat(quickPnLAmount);
    if (isNaN(amount) || amount <= 0) return;

    const actualProfit = quickPnLType === 'profit' ? amount : -amount;

    if (editingTradeId) {
      handleUpdateJournal(editingTradeId, {
        profit: actualProfit,
        notes: quickPnLNotes || 'Quick calendar P&L entry',
        ticker: quickPnLTicker || 'Quick Entry',
        type: actualProfit >= 0 ? 'long' : 'short',
        exit: Math.abs(actualProfit)
      });
      setEditingTradeId(null);
    } else {
      const targetTimestamp = new Date(selectedCalendarDate + 'T12:00:00').getTime();
      const newTrade: SavedTrade = {
        id: crypto.randomUUID(),
        mode: 'equity',
        type: actualProfit >= 0 ? 'long' : 'short',
        entry: 0,
        exit: Math.abs(actualProfit),
        quantity: 1,
        leverage: 1,
        profit: actualProfit,
        timestamp: targetTimestamp,
        entryDate: selectedCalendarDate,
        exitDate: selectedCalendarDate,
        notes: quickPnLNotes || 'Quick calendar P&L entry',
        ticker: quickPnLTicker || 'Quick Entry',
        isManual: true,
        isFavorite: false
      };

      setHistory([newTrade, ...history]);
      if (totalFunds > 0) {
        setTotalFunds(prev => prev + actualProfit);
      }
    }

    setQuickPnLAmount('');
    setQuickPnLTicker('');
    setQuickPnLNotes('');
  };

  const handleDeleteEditingTrade = () => {
    if (editingTradeId) {
      deleteTrade(editingTradeId);
      setEditingTradeId(null);
      setQuickPnLAmount('');
      setQuickPnLTicker('');
      setQuickPnLNotes('');
    }
  };

  const handleCancelEditQuickPnL = () => {
    setEditingTradeId(null);
    setQuickPnLAmount('');
    setQuickPnLTicker('');
    setQuickPnLNotes('');
  };

  const daysInMonth = (month: number, year: number) => new Date(year, month + 1, 0).getDate();
  const firstDayOfMonth = (month: number, year: number) => new Date(year, month, 1).getDay();

  // Waterfall Chart data for P&L Calendar
  const monthlyWaterfallData = useMemo(() => {
    const list: Array<{
      dayNum: number;
      dateStr: string;
      pnl: number;
      lastCumulative: number;
      currentCumulative: number;
      range: [number, number];
      isPositive: boolean;
    }> = [];

    const numDays = new Date(currentYear, currentMonth + 1, 0).getDate();
    let lastCumulative = 0;

    for (let d = 1; d <= numDays; d++) {
      const dateStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const pnl = calendarData[dateStr];
      if (pnl !== undefined) {
        const currentCumulative = lastCumulative + pnl;
        list.push({
          dayNum: d,
          dateStr,
          pnl,
          lastCumulative,
          currentCumulative,
          range: [lastCumulative, currentCumulative],
          isPositive: pnl >= 0
        });
        lastCumulative = currentCumulative;
      }
    }

    if (list.length > 0) {
      list.push({
        dayNum: 99,
        dateStr: 'Total',
        pnl: lastCumulative,
        lastCumulative: 0,
        currentCumulative: lastCumulative,
        range: [0, lastCumulative],
        isPositive: lastCumulative >= 0
      });
    }

    return list;
  }, [calendarData, currentMonth, currentYear]);

  // Spline Data filtered for spline chart
  const monthlySplineData = useMemo(() => {
    return monthlyWaterfallData.filter(d => d.dayNum !== 99);
  }, [monthlyWaterfallData]);

  // Fund transaction dynamic summaries
  const totalDeposits = useMemo(() => {
    return fundTransactions
      .filter(tx => tx.type === 'deposit')
      .reduce((sum, tx) => sum + tx.amount, 0);
  }, [fundTransactions]);

  const totalWithdrawals = useMemo(() => {
    return fundTransactions
      .filter(tx => tx.type === 'withdrawal')
      .reduce((sum, tx) => sum + tx.amount, 0);
  }, [fundTransactions]);

  // Active month's analytics (for P&L calendar)
  const activeMonthStats = useMemo(() => {
    const activePrefix = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}`;
    const monthEntries = Object.entries(calendarData).filter(([date]) => date.startsWith(activePrefix));
    let profitDays = 0;
    let lossDays = 0;
    let totalProfits = 0;
    let totalLosses = 0;
    for (const [_, p] of monthEntries) {
      if ((p as number) > 0) {
        profitDays++;
        totalProfits += p as number;
      } else if ((p as number) < 0) {
        lossDays++;
        totalLosses += p as number;
      }
    }
    return {
      profitDays,
      lossDays,
      totalProfits,
      totalLosses,
    };
  }, [calendarData, currentMonth, currentYear]);

  // Statistics
  const stats = useMemo(() => {
    if (history.length === 0) return { totalTrades: 0, winRate: 0, netPnL: 0, winCount: 0 };
    const winCount = history.filter(t => t.profit >= 0).length;
    const netPnL = history.reduce((acc, t) => acc + t.profit, 0);
    return {
      totalTrades: history.length,
      winRate: (winCount / history.length) * 100,
      netPnL,
      winCount
    };
  }, [history]);

  // Derived Trend Data oldest-first for Cumulative PnL Graph
  const historyTrendData = useMemo(() => {
    if (history.length === 0) return [];
    const sortedChrono = [...history].sort((a, b) => a.timestamp - b.timestamp);
    let cumulative = 0;
    return sortedChrono.map((trade, idx) => {
      cumulative += trade.profit;
      const dateStr = trade.exitDate 
        ? trade.exitDate 
        : new Date(trade.timestamp).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' });
      return {
        index: idx,
        displayIndex: idx + 1,
        date: dateStr,
        timestamp: trade.timestamp,
        profit: trade.profit,
        cumulative: cumulative,
        ticker: trade.ticker || (trade.mode === 'options' ? 'Options' : 'Equity') + ' #' + (idx + 1)
      };
    });
  }, [history]);

  // Handle active zoom/filter range in chronological trades list
  const filteredHistory = useMemo(() => {
    let result = [...history].sort((a, b) => b.timestamp - a.timestamp); // default newest first
    
    if (historyZoomRange) {
      const [minIdx, maxIdx] = historyZoomRange;
      const sortedChrono = [...history].sort((a, b) => a.timestamp - b.timestamp);
      const minTrade = sortedChrono[minIdx];
      const maxTrade = sortedChrono[maxIdx];
      if (minTrade && maxTrade) {
        const minTs = Math.min(minTrade.timestamp, maxTrade.timestamp);
        const maxTs = Math.max(minTrade.timestamp, maxTrade.timestamp);
        result = result.filter(trade => trade.timestamp >= minTs && trade.timestamp <= maxTs);
      }
    }
    return result;
  }, [history, historyZoomRange]);

  // Dynamic Metrics for standard display/analytical updates within the selected range
  const filteredStats = useMemo(() => {
    if (filteredHistory.length === 0) {
      return { totalTrades: 0, winRate: 0, netPnL: 0, winCount: 0, avgProfit: 0 };
    }
    const winCount = filteredHistory.filter(t => t.profit >= 0).length;
    const netPnL = filteredHistory.reduce((acc, t) => acc + t.profit, 0);
    const avgProfit = netPnL / filteredHistory.length;
    return {
      totalTrades: filteredHistory.length,
      winRate: (winCount / filteredHistory.length) * 100,
      netPnL,
      winCount,
      avgProfit
    };
  }, [filteredHistory]);

  // Calculations
  const results = useMemo((): TradeResult | null => {
    const entry = parseFloat(entryPrice) || 0;
    const exit = parseFloat(exitPrice) || 0;
    const qtyInput = parsedQty;
    const feesPct = (parseFloat(feePercentage) || 0) / 100;
    const lev = leverage || 1;

    let totalQty = qtyInput;
    if (mode === 'options') {
      const lSize = parseFloat(lotSize) || 1;
      totalQty = qtyInput * lSize;
    }

    if (entry <= 0 || totalQty <= 0) {
      return null;
    }

    const initialMargin = (entry * totalQty) / lev;
    const entryTotal = entry * totalQty;
    const exitTotal = exit * totalQty;

    let grossProfit = 0;
    if (type === 'long') {
      grossProfit = exitTotal - entryTotal;
    } else {
      grossProfit = entryTotal - exitTotal;
    }

    const fees = (entryTotal + exitTotal) * feesPct;
    const netProfit = grossProfit - fees;
    const profitPercentage = (netProfit / entryTotal) * 100;
    const roe = (netProfit / initialMargin) * 100;

    return {
      profit: netProfit,
      profitPercentage,
      totalValue: exitTotal,
      fees,
      isProfit: netProfit >= 0,
      roe
    };
  }, [mode, type, entryPrice, exitPrice, quantity, lotSize, leverage, feePercentage]);

  // Chart Data Generation
  const chartData = useMemo(() => {
    const entry = parseFloat(entryPrice) || 0;
    const qtyInput = parsedQty;
    const feesPct = (parseFloat(feePercentage) || 0) / 100;
    
    if (entry <= 0 || qtyInput <= 0) return [];

    let totalQty = qtyInput;
    if (mode === 'options') {
      const lSize = parseFloat(lotSize) || 1;
      totalQty = qtyInput * lSize;
    }

    const data = [];
    const steps = 20;
    
    // Range based on timeframe
    let rangeMultiplier = 0.2; // Default Month
    if (graphTimeframe === 'Day') rangeMultiplier = 0.03;
    if (graphTimeframe === 'Week') rangeMultiplier = 0.07;
    if (graphTimeframe === 'Month') rangeMultiplier = 0.15;
    if (graphTimeframe === 'Year') rangeMultiplier = 0.40;

    const range = entry * rangeMultiplier;
    const start = Math.max(0, entry - range);
    const end = entry + range;
    const stepSize = (end - start) / steps;

    for (let i = 0; i <= steps; i++) {
      const currentPrice = start + (i * stepSize);
      const entryTotal = entry * totalQty;
      const currentTotal = currentPrice * totalQty;

      const priceChangePct = (currentPrice - entry) / entry;

      let grossProfit = 0;
      if (type === 'long') {
        grossProfit = currentTotal - entryTotal;
      } else {
        grossProfit = entryTotal - currentTotal;
      }

      const totalFees = (entryTotal + currentTotal) * feesPct;
      const netProfit = grossProfit - totalFees;

      // Simulated Index Movements (Capital * Index% Change)
      // We use a slight beta adjustment to make the lines look distinct but correlated
      const niftyProfit = entryTotal * (priceChangePct * 1.02); 
      const sensexProfit = entryTotal * (priceChangePct * 0.98);

      data.push({
        price: parseFloat(currentPrice.toFixed(2)),
        profit: parseFloat(netProfit.toFixed(0)),
        nifty: parseFloat(niftyProfit.toFixed(0)),
        sensex: parseFloat(sensexProfit.toFixed(0))
      });
    }
    return data;
  }, [mode, type, entryPrice, quantity, lotSize, feePercentage, graphTimeframe]);

  // Zoomed or active selection bounds of the projection chart indices
  const zoomedChartData = useMemo(() => {
    if (!projectionZoomRange) return chartData;
    const [minIdx, maxIdx] = projectionZoomRange;
    return chartData.filter((_, idx) => idx >= minIdx && idx <= maxIdx);
  }, [chartData, projectionZoomRange]);

  // Reset projection zoom range when parameters change
  useEffect(() => {
    setProjectionZoomRange(null);
  }, [mode, type, entryPrice, quantity, lotSize, feePercentage, graphTimeframe]);

  const clearHistory = () => {
    setConfirmDialog({
      show: true,
      title: 'Clear History',
      message: 'Are you sure you want to clear all trade history of this device permanently?',
      onConfirm: () => {
        setConfirmDialog(null);
        setHistory([]);
        showToastAlert('Trade history cleared successfully.', 'success');
      }
    });
  };

  const reset = () => {
    setEntryPrice('');
    setExitPrice('');
    setQuantity('');
    setNotes('');
    setTradeSetup('');
    setEntryExitRationale('');
    setLessonsLearned('');
    setLeverage(1);
    setFeePercentage('0.05');
  };

  const saveTrade = () => {
    if (!results) return;
    const newTrade: SavedTrade = {
      id: crypto.randomUUID(),
      mode,
      type,
      entry: parseFloat(entryPrice),
      exit: parseFloat(exitPrice),
      quantity: parsedQty,
      lotSize: mode === 'options' ? parseFloat(lotSize) : undefined,
      leverage,
      profit: results.profit,
      timestamp: Date.now(),
      entryDate,
      exitDate,
      strategy,
      notes,
      tradeSetup,
      entryExitRationale,
      lessonsLearned,
      isFavorite: false
    };
    setHistory([newTrade, ...history]);
    if (totalFunds > 0) {
      setTotalFunds(prev => prev + results.profit);
    }
  };

  const handleSaveManualTrade = () => {
    const entry = parseFloat(manualEntryPrice) || 0;
    const exit = parseFloat(manualExitPrice) || 0;
    const qty = parseFloat(manualQuantity) || 1;
    const lotSz = parseFloat(manualLotSize) || 1;

    let calculatedProfit = 0;
    if (manualProfitCustom) {
      calculatedProfit = parseFloat(manualProfitAmount) || 0;
    } else {
      if (manualMode === 'options') {
        const valueRatio = (exit - entry) * qty * lotSz;
        calculatedProfit = manualType === 'long' ? valueRatio : -valueRatio;
      } else {
        const valueRatio = (exit - entry) * qty;
        calculatedProfit = manualType === 'long' ? valueRatio : -valueRatio;
      }
    }

    const newTrade: SavedTrade = {
      id: crypto.randomUUID(),
      mode: manualMode,
      type: manualType,
      entry,
      exit,
      quantity: qty,
      lotSize: manualMode === 'options' ? lotSz : undefined,
      leverage: 1,
      profit: calculatedProfit,
      timestamp: Date.now(),
      entryDate: manualEntryDate,
      exitDate: manualExitDate,
      strategy: manualStrategy,
      notes: manualNotes,
      ticker: manualTicker || 'Manual Entry',
      tradeSetup: manualTradeSetup,
      entryExitRationale: manualEntryExitRationale,
      lessonsLearned: manualLessonsLearned,
      isManual: true,
      isFavorite: false
    };

    setHistory([newTrade, ...history]);
    if (totalFunds > 0) {
      setTotalFunds(prev => prev + calculatedProfit);
    }

    // Reset Form
    setManualTicker('');
    setManualEntryPrice('');
    setManualExitPrice('');
    setManualQuantity('');
    setManualNotes('');
    setManualTradeSetup('');
    setManualEntryExitRationale('');
    setManualLessonsLearned('');
    setIsManualJournalOpen(false);
  };

  const handleUpdateJournal = (id: string, updatedFields: Partial<SavedTrade>) => {
    setHistory(prevHistory => prevHistory.map(t => {
      if (t.id === id) {
        if (updatedFields.profit !== undefined && totalFunds > 0) {
          const diff = updatedFields.profit - t.profit;
          setTotalFunds(prev => prev + diff);
        }
        return { ...t, ...updatedFields };
      }
      return t;
    }));
    if (selectedJournalTrade && selectedJournalTrade.id === id) {
      setSelectedJournalTrade(prev => prev ? { ...prev, ...updatedFields } : null);
    }
  };

  const deleteTrade = (id: string) => {
    const tradeToDelete = history.find(t => t.id === id);
    if (tradeToDelete && totalFunds > 0) {
      setTotalFunds(prev => prev - tradeToDelete.profit);
    }
    setHistory(history.filter(t => t.id !== id));
    if (editingTradeId === id) {
      setEditingTradeId(null);
      setQuickPnLAmount('');
      setQuickPnLTicker('');
      setQuickPnLNotes('');
    }
  };

  const handleAddFundTransaction = (type: 'deposit' | 'withdrawal') => {
    const amount = parseFloat(fundInput);
    if (isNaN(amount) || amount <= 0) return;

    if (type === 'withdrawal' && amount > totalFunds) {
      showToastAlert('Insufficient funds!', 'error');
      return;
    }

    if (editingFundTxId) {
      // Updating an existing adjustment
      const oldTx = fundTransactions.find(t => t.id === editingFundTxId);
      if (!oldTx) return;

      // Reverse old tx effect
      let tempFunds = totalFunds;
      if (oldTx.type === 'deposit') {
        tempFunds -= oldTx.amount;
      } else {
        tempFunds += oldTx.amount;
      }

      if (type === 'withdrawal' && amount > tempFunds) {
        showToastAlert('Insufficient funds for this updated withdrawal!', 'error');
        return;
      }

      const finalPrice = type === 'deposit' ? tempFunds + amount : tempFunds - amount;
      setTotalFunds(finalPrice);

      setFundTransactions(fundTransactions.map(tx => 
        tx.id === editingFundTxId 
          ? { ...tx, type, amount, notes: fundNotesInput, timestamp: Date.now() }
          : tx
      ));
      
      setEditingFundTxId(null);
      setFundInput('');
      setFundNotesInput('');
    } else {
      // New transaction creation
      const newTx: FundTransaction = {
        id: crypto.randomUUID(),
        type,
        amount,
        timestamp: Date.now(),
        notes: fundNotesInput || `${type === 'deposit' ? 'Capital Added' : 'Capital Withdrawn'}`
      };

      setFundTransactions([newTx, ...fundTransactions]);
      setTotalFunds(prev => type === 'deposit' ? prev + amount : prev - amount);
      setFundInput('');
      setFundNotesInput('');
    }
  };

  const handleDeleteFundTransaction = (id: string) => {
    const txToDelete = fundTransactions.find(t => t.id === id);
    if (!txToDelete) return;

    if (txToDelete.type === 'deposit') {
      setTotalFunds(prev => prev - txToDelete.amount);
    } else {
      setTotalFunds(prev => prev + txToDelete.amount);
    }

    setFundTransactions(fundTransactions.filter(t => t.id !== id));
    if (editingFundTxId === id) {
      setEditingFundTxId(null);
      setFundInput('');
      setFundNotesInput('');
    }
  };

  const handleStartEditFundTransaction = (tx: FundTransaction) => {
    setEditingFundTxId(tx.id);
    setFundInput(tx.amount.toString());
    setFundNotesInput(tx.notes || '');
  };

  const handleCancelEditFundTransaction = () => {
    setEditingFundTxId(null);
    setFundInput('');
    setFundNotesInput('');
  };

  const toggleFavorite = (id: string) => {
    setHistory(history.map(t => t.id === id ? { ...t, isFavorite: !t.isFavorite } : t));
  };

  const exportHistoryExcel = (onlyFiltered: boolean = false) => {
    const targetHistory = onlyFiltered ? filteredHistory : history;
    if (targetHistory.length === 0) return;
    
    const data = targetHistory.map(t => ({
      'Date': new Date(t.timestamp).toLocaleString(),
      'Entry Date': t.entryDate || '',
      'Exit Date': t.exitDate || '',
      'Mode': t.mode.toUpperCase(),
      'Type': t.type.toUpperCase(),
      'Strategy': t.strategy || 'N/A',
      'Entry Price': t.entry,
      'Exit Price': t.exit,
      'Quantity': t.quantity,
      'Lot Size': t.lotSize || '-',
      'Leverage': t.leverage,
      'Net Profit (INR)': t.profit.toFixed(2),
      'Notes': t.notes || '',
      'Trade Setup': t.tradeSetup || '',
      'Entry/Exit Rationale': t.entryExitRationale || '',
      'Lessons Learned': t.lessonsLearned || ''
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'History');
    const rangeStr = onlyFiltered ? `_Filtered_${Date.now()}` : `_Full_${Date.now()}`;
    XLSX.writeFile(wb, `Trade_History${rangeStr}.xlsx`);
  };

  const exportHistoryCSV = (onlyFiltered: boolean = false) => {
    const targetHistory = onlyFiltered ? filteredHistory : history;
    if (targetHistory.length === 0) return;

    const headers = [
      'Date', 'Entry Date', 'Exit Date', 'Mode', 'Type', 'Strategy', 
      'Entry Price', 'Exit Price', 'Quantity', 'Lot Size', 'Leverage', 
      'Net Profit (INR)', 'Notes', 'Trade Setup', 'Entry/Exit Rationale', 'Lessons Learned'
    ];
    
    const rows = targetHistory.map(t => [
      new Date(t.timestamp).toLocaleString(),
      t.entryDate || '',
      t.exitDate || '',
      t.mode.toUpperCase(),
      t.type.toUpperCase(),
      t.strategy || 'N/A',
      t.entry,
      t.exit,
      t.quantity,
      t.lotSize || '-',
      t.leverage,
      t.profit.toFixed(2),
      t.notes || '',
      t.tradeSetup || '',
      t.entryExitRationale || '',
      t.lessonsLearned || ''
    ]);

    const csvString = [
      headers.join(','),
      ...rows.map(row => row.map(val => {
        const str = String(val).replace(/"/g, '""');
        return `"${str}"`;
      }).join(','))
    ].join('\r\n');

    const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    const rangeStr = onlyFiltered ? `_Filtered_${Date.now()}` : `_Full_${Date.now()}`;
    link.setAttribute("download", `Trade_History${rangeStr}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportHistoryPDF = (onlyFiltered: boolean = false) => {
    const targetHistory = onlyFiltered ? filteredHistory : history;
    if (targetHistory.length === 0) return;
    const doc = new jsPDF();
    
    doc.setFontSize(20);
    doc.text('Trade History - Trade track P&L', 14, 22);
    
    doc.setFontSize(10);
    const subtitle = onlyFiltered ? `Filtered Range | Generated: ${new Date().toLocaleString()}` : `Full History | Generated: ${new Date().toLocaleString()}`;
    doc.text(subtitle, 14, 30);
    
    const tableData = targetHistory.map(t => [
      new Date(t.timestamp).toLocaleDateString(),
      t.mode.toUpperCase(),
      t.type.toUpperCase(),
      t.strategy || '-',
      `INR ${t.entry}`,
      `INR ${t.exit}`,
      t.quantity,
      `INR ${t.profit.toFixed(0)}`,
      t.notes || ''
    ]);

    autoTable(doc, {
      startY: 35,
      head: [['Date', 'Mode', 'Type', 'Strategy', 'Entry', 'Exit', 'Qty', 'PnL', 'Notes']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: '#4f46e5' },
      styles: { fontSize: 7 }
    });

    const rangeStr = onlyFiltered ? `_Filtered_${Date.now()}` : `_Full_${Date.now()}`;
    doc.save(`Trade_History${rangeStr}.pdf`);
  };

  const downloadPDF = () => {
    if (!results) return;
    const doc = new jsPDF();
    
    doc.setFontSize(20);
    doc.text('Trade Report - Trade track P&L', 14, 22);
    
    doc.setFontSize(12);
    doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 30);
    
    const tableData = [
      ['Mode', mode.toUpperCase()],
      ['Type', (type === 'long' ? (mode === 'equity' ? 'Buy' : 'Call') : (mode === 'equity' ? 'Sell' : 'Put'))],
      ['Strategy', strategy],
      ['Entry Date', entryDate],
      ['Exit Date', exitDate],
      ['Entry Price', `INR ${entryPrice}`],
      ['Exit Price', `INR ${exitPrice}`],
      ['Quantity/Lots', quantity],
      ['Lot Size', mode === 'options' ? lotSize : 'N/A'],
      ['Leverage', `${leverage}x`],
      ['Fees/Taxes (%)', `${feePercentage}%`],
      ['Total Fees', `INR ${results.fees.toFixed(2)}`],
      ['Net Profit/Loss', `INR ${results.profit.toFixed(2)}`],
      ['Return %', `${results.profitPercentage.toFixed(2)}%`],
      ['ROE %', `${results.roe.toFixed(2)}%`],
      ['Notes', notes || 'N/A'],
    ];

    autoTable(doc, {
      startY: 40,
      head: [['Parameter', 'Value']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: '#4f46e5' }
    });

    doc.save(`Trade_Report_${Date.now()}.pdf`);
  };

  const handleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setAuthUser(result.user);
        setNeedsAuth(false);
        if (result.user?.email) {
          setReportRecipient(result.user.email);
        }
        showToastAlert('Successfully signed in with Google!', 'success');
      }
    } catch (err) {
      console.error('Login failed:', err);
      showToastAlert('Login failed. Please check your internet connection or try again.', 'error');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    setConfirmDialog({
      show: true,
      title: 'Sign Out',
      message: 'Are you sure you want to sign out of your Google account?',
      onConfirm: async () => {
        setConfirmDialog(null);
        try {
          await logout();
          setAuthUser(null);
          setNeedsAuth(true);
          showToastAlert('Logged out successfully.', 'success');
        } catch (err) {
          console.error('Logout failed:', err);
          showToastAlert('Failed to sign out. Please try again.', 'error');
        }
      }
    });
  };

  const handleBackup = async () => {
    const token = getAccessToken();
    if (!token) {
      showToastAlert('Please sign in again to backup your data.', 'error');
      setNeedsAuth(true);
      return;
    }

    setIsBackingUp(true);
    try {
      await createOrUpdateBackup(token, {
        history,
        userName,
        timestamp: Date.now(),
        totalFunds,
        openingBalance,
        fundTransactions
      });
      const nowStr = new Date().toLocaleString('en-IN');
      setDriveLastBackup(nowStr);
      showToastAlert('Data successfully backed up to Google Drive (App Data Folder)!', 'success');
    } catch (err) {
      console.error('Backup failed:', err);
      showToastAlert('Failed to backup data. Please try again.', 'error');
    } finally {
      setIsBackingUp(false);
    }
  };

  const handleRestore = async () => {
    const token = getAccessToken();
    if (!token) {
      showToastAlert('Please sign in again to restore your data.', 'error');
      setNeedsAuth(true);
      return;
    }

    setConfirmDialog({
      show: true,
      title: 'Restore Backup from Drive',
      message: 'This will overwrite your current local trade history, funds, and transaction logs with the data from Google Drive. Continue?',
      onConfirm: async () => {
        setConfirmDialog(null);
        setIsRestoring(true);
        try {
          const data = await restoreBackup(token);
          if (data) {
            setHistory(data.history || []);
            if (data.userName) {
              setUserName(data.userName);
              localStorage.setItem('user_name', data.userName);
            }
            if (data.totalFunds !== undefined) {
              setTotalFunds(data.totalFunds);
              localStorage.setItem('trade_track_funds', String(data.totalFunds));
            }
            if (data.openingBalance !== undefined) {
              setOpeningBalance(data.openingBalance);
              localStorage.setItem('trade_track_opening_balance', String(data.openingBalance));
            }
            if (data.fundTransactions) {
              setFundTransactions(data.fundTransactions);
              localStorage.setItem('trade_fund_transactions', JSON.stringify(data.fundTransactions));
            }
            const backupDateStr = new Date(data.timestamp).toLocaleString('en-IN');
            setDriveLastBackup(backupDateStr);
            showToastAlert('Data successfully restored from Google Drive!', 'success');
          } else {
            showToastAlert('No backup found in your Google Drive App Data folder.', 'info');
          }
        } catch (err) {
          console.error('Restore failed:', err);
          showToastAlert('Failed to restore data. Please try again.', 'error');
        } finally {
          setIsRestoring(false);
        }
      }
    });
  };

  const handleScanGmail = async () => {
    setIsScanningGmail(true);
    setGmailError(null);
    try {
      const token = getAccessToken();
      if (!token) {
        setGmailError('Authentication required. Please sign in with Google again.');
        showToastAlert('Please sign in to Google to scan Gmail inbox.', 'error');
        return;
      }
      const emails = await listRecentEmails(token);
      setGmailEmails(emails);
      if (emails.length === 0) {
        showToastAlert('No broker or transaction emails found in your recent messages.', 'info');
      } else {
        showToastAlert(`Successfully scanned ${emails.length} broker-related email(s)!`, 'success');
      }
    } catch (err: any) {
      console.error('Failed to scan Gmail inbox:', err);
      setGmailError(err.message || 'Failed to sync with Gmail API.');
      showToastAlert('Failed to scan Gmail inbox.', 'error');
    } finally {
      setIsScanningGmail(false);
    }
  };

  const handleSendPerformanceEmail = () => {
    if (!reportRecipient.trim()) {
      showToastAlert('Recipient email address cannot be empty.', 'error');
      return;
    }
    
    setConfirmDialog({
      show: true,
      title: 'Send Performance Report',
      message: `Are you sure you want to send your professional HTML Performance Report via Gmail to ${reportRecipient}?`,
      onConfirm: async () => {
        setConfirmDialog(null);
        setIsSendingReport(true);
        try {
          const token = getAccessToken();
          if (!token) {
            showToastAlert('Authentication required. Please sign in with Google.', 'error');
            return;
          }
          
          const winPercent = filteredStats.winRate.toFixed(1);
          const totalTr = filteredStats.totalTrades;
          const pnlText = filteredStats.netPnL >= 0 
            ? `<span style="color: #10b981; font-weight: bold;">+₹${filteredStats.netPnL.toLocaleString('en-IN')}</span>` 
            : `<span style="color: #ef4444; font-weight: bold;">-₹${Math.abs(filteredStats.netPnL).toLocaleString('en-IN')}</span>`;
            
          const tradesListHtml = filteredHistory.slice(0, 5).map(t => {
            const dateStr = t.entryDate || new Date(t.timestamp).toLocaleDateString();
            const changeP = t.profit >= 0 ? 'color: #10b981;' : 'color: #ef4444;';
            const actionLabel = t.type ? t.type.toUpperCase() : 'LONG';
            return `
              <tr style="border-bottom: 1px solid #e1e8ed;">
                <td style="padding: 8px 12px; font-size: 13px;">${dateStr}</td>
                <td style="padding: 8px 12px; font-weight: bold; font-size: 13px;">${t.ticker || 'N/A'}</td>
                <td style="padding: 8px 12px; font-size: 13px;">${actionLabel} (${t.mode === 'options' ? 'Options' : 'Equity'})</td>
                <td style="padding: 8px 12px; font-size: 13px;">Qty: ${t.quantity} @ Entry ₹${t.entry}</td>
                <td style="padding: 8px 12px; font-size: 13px; font-weight: bold; ${changeP}">₹${t.profit >= 0 ? '+' : ''}${t.profit.toFixed(2)}</td>
              </tr>
            `;
          }).join('') || '<tr><td colspan="5" style="padding: 16px; text-align: center; color: #738a9c;">No trades recorded in this range.</td></tr>';

          const personalizedNote = reportCustomMessage.trim() 
            ? `<div style="margin: 20px 0; padding: 16px; background-color: #f3f4f6; border-radius: 8px; border-left: 4px solid #4f46e5; font-style: italic;">
                 <strong>Journal Note:</strong> "${reportCustomMessage}"
               </div>`
            : '';

          const htmlContent = `
            <!DOCTYPE html>
            <html>
              <head>
                <meta charset="utf-8">
                <title>${reportSubject}</title>
              </head>
              <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; background-color: #f8fafc; padding: 24px;">
                <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); overflow: hidden; border: 1px solid #e2e8f0;">
                  <div style="background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%); padding: 32px 24px; text-align: center; color: #ffffff;">
                    <h1 style="margin: 0; font-size: 24px; font-weight: 800; letter-spacing: -0.5px;">PRO TRADER JOURNAL</h1>
                    <p style="margin: 4px 0 0 0; font-size: 14px; opacity: 0.9; font-weight: 500;">Gmail Integration - Performance Report</p>
                  </div>
                  
                  <div style="padding: 24px;">
                    <p style="margin-top: 0; font-size: 15px; line-height: 1.5;">Hello,</p>
                    <p style="font-size: 15px; line-height: 1.5; margin-bottom: 24px;">Here is the compiled trading performance metric from the private database, generated on <strong>${new Date().toLocaleDateString('en-IN')}</strong>.</p>
                    
                    ${personalizedNote}

                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 28px;">
                      <tr>
                        <td style="width: 50%; padding: 12px; background-color: #f1f5f9; border-radius: 8px; border-right: 2px solid #ffffff;">
                          <div style="font-size: 10px; font-weight: bold; color: #64748b; text-transform: uppercase;">Win Ratio</div>
                          <div style="font-size: 20px; font-weight: 800; color: #4F46E5; margin-top: 4px;">${winPercent}%</div>
                        </td>
                        <td style="width: 50%; padding: 12px; background-color: #f1f5f9; border-radius: 8px;">
                          <div style="font-size: 10px; font-weight: bold; color: #64748b; text-transform: uppercase;">Total Executed</div>
                          <div style="font-size: 20px; font-weight: 800; color: #1e293b; margin-top: 4px;">${totalTr} Trades</div>
                        </td>
                      </tr>
                      <tr style="height: 8px;"></tr>
                      <tr>
                        <td style="width: 100%; padding: 12px; background-color: #f1f5f9; border-radius: 8px; text-align: center;" colspan="2">
                          <div style="font-size: 10px; font-weight: bold; color: #64748b; text-transform: uppercase;">Net Profit & Loss</div>
                          <div style="font-size: 26px; font-weight: 900; margin-top: 4px;">${pnlText}</div>
                        </td>
                      </tr>
                    </table>

                    <h3 style="font-size: 14px; font-weight: 700; color: #0f172a; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Recent Trades Highlight</h3>
                    <div style="overflow-x: auto; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px;">
                      <table style="width: 100%; border-collapse: collapse; text-align: left;">
                        <thead>
                          <tr style="background-color: #f8fafc; border-bottom: 1px solid #e1e8ed;">
                            <th style="padding: 10px 12px; font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase;">Date</th>
                            <th style="padding: 10px 12px; font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase;">Symbol</th>
                            <th style="padding: 10px 12px; font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase;">Action</th>
                            <th style="padding: 10px 12px; font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase;">Details</th>
                            <th style="padding: 10px 12px; font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase;">P&L</th>
                          </tr>
                        </thead>
                        <tbody>
                          ${tradesListHtml}
                        </tbody>
                      </table>
                    </div>
                  </div>
                  
                  <div style="padding: 20px; text-align: center; background-color: #f1f5f9; border-top: 1px solid #e2e8f0; font-size: 11px; color: #94a3b8;">
                    This report was automatically compiled by the Pro Trader Journal Application.<br>
                    Keep up the great risk management!
                  </div>
                </div>
              </body>
            </html>
          `;

          await sendTradingReport(token, reportRecipient, reportSubject, htmlContent);
          showToastAlert(`Trading report successfully sent to ${reportRecipient}!`, 'success');
          setReportCustomMessage('');
        } catch (err: any) {
          console.error('Failed to send trading report email:', err);
          showToastAlert(`Failed to send report: ${err.message || err.toString()}`, 'error');
        } finally {
          setIsSendingReport(false);
        }
      }
    });
  };

  const handleImportGmailTrade = (gmailTrade: any, emailSubject: string) => {
    if (!gmailTrade) return;
    
    setConfirmDialog({
      show: true,
      title: 'Import Trade from Gmail',
      message: `Do you want to import this trade from Gmail?
Trade: ${gmailTrade.action} ${gmailTrade.quantity} shares of ${gmailTrade.ticker} at Avg Price ₹${gmailTrade.price}.`,
      onConfirm: async () => {
        setConfirmDialog(null);
        
        const newTrade: SavedTrade = {
          id: `trade_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
          mode: gmailTrade.type || 'equity',
          type: gmailTrade.action === 'BUY' ? 'long' : 'short',
          entry: gmailTrade.price,
          exit: gmailTrade.price,
          quantity: gmailTrade.quantity,
          leverage: 1,
          profit: 0,
          timestamp: Date.now(),
          entryDate: new Date().toISOString().split('T')[0],
          exitDate: new Date().toISOString().split('T')[0],
          ticker: gmailTrade.ticker,
          strategy: 'Gmail Import',
          notes: `Imported from Gmail broker email: "${emailSubject}"`,
          tradeSetup: 'Imported',
          lotSize: 1,
          isFavorite: false,
          entryExitRationale: 'Imported from Gmail',
          lessonsLearned: 'Parsed automatically'
        };

        const updatedHistory = [newTrade, ...history];
        setHistory(updatedHistory);
        localStorage.setItem('trade_history', JSON.stringify(updatedHistory));
        
        // Auto trigger cloud sync if enabled!
        const authedToken = getAccessToken();
        if (authUser && authedToken && autoBackupEnabled) {
          try {
            await createOrUpdateBackup(authedToken, {
              history: updatedHistory,
              userName,
              timestamp: Date.now(),
              totalFunds,
              openingBalance,
              fundTransactions
            });
            setDriveLastBackup(new Date().toLocaleString('en-IN'));
            localStorage.setItem('trade_track_drive_last_backup', new Date().toLocaleString('en-IN'));
            setDriveSyncStatus('success');
          } catch (e) {
            console.error('Auto cloud backup failed post-import:', e);
          }
        }

        showToastAlert(`Successfully imported ${gmailTrade.ticker} trade into local journal!`, 'success');
      }
    });
  };

  const handleLocalBackup = () => {
    try {
      const backupData = {
        history,
        userName,
        timestamp: Date.now(),
        version: "1.0",
        app: "Trade Track P&L"
      };
      
      const jsonString = JSON.stringify(backupData, null, 2);
      const blob = new Blob([jsonString], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const dateStr = new Date().toISOString().split('T')[0];
      link.download = `trade_track_backup_${dateStr}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Local backup failed:', err);
      showToastAlert('Failed to generate backup file.', 'error');
    }
  };

  const handleLocalRestore = (event: ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    const file = event.target.files?.[0];
    if (!file) return;

    fileReader.onload = (e) => {
      try {
        const parsedData = JSON.parse(e.target?.result as string);
        if (parsedData && Array.isArray(parsedData.history)) {
          setConfirmDialog({
            show: true,
            title: 'Restore Local Backup',
            message: 'This will overwrite your current local trade history with the backup file data. Continue?',
            onConfirm: () => {
              setConfirmDialog(null);
              setHistory(parsedData.history);
              if (parsedData.userName) {
                setUserName(parsedData.userName);
                localStorage.setItem('user_name', parsedData.userName);
              }
              showToastAlert('Data successfully restored from backup file!', 'success');
            }
          });
        } else {
          showToastAlert('Invalid file format. Ensure you selected a correct backup JSON file.', 'error');
        }
      } catch (err) {
        console.error('Failed to parse backup file:', err);
        showToastAlert('Failed to read file. Please ensure it is a valid backup.json file.', 'error');
      }
    };
    fileReader.readAsText(file);
    event.target.value = ''; // Reset
  };

  const downloadExcel = () => {
    if (!results) return;
    const reportData = [
      { Parameter: 'Report Name', Value: 'Trade Analysis' },
      { Parameter: 'Date', Value: new Date().toLocaleString() },
      { Parameter: 'Mode', Value: mode },
      { Parameter: 'Type', Value: type },
      { Parameter: 'Strategy', Value: strategy },
      { Parameter: 'Entry Date', Value: entryDate },
      { Parameter: 'Exit Date', Value: exitDate },
      { Parameter: 'Entry Price', Value: entryPrice },
      { Parameter: 'Exit Price', Value: exitPrice },
      { Parameter: 'Quantity', Value: quantity },
      { Parameter: 'Lot Size', Value: mode === 'options' ? lotSize : '' },
      { Parameter: 'Leverage', Value: leverage },
      { Parameter: 'Fees %', Value: feePercentage },
      { Parameter: 'Total Fees', Value: results.fees },
      { Parameter: 'Net Profit', Value: results.profit },
      { Parameter: 'Return %', Value: results.profitPercentage },
      { Parameter: 'ROE %', Value: results.roe },
      { Parameter: 'Notes', Value: notes },
    ];

    const ws = XLSX.utils.json_to_sheet(reportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Trade Report');
    XLSX.writeFile(wb, `Trade_Report_${Date.now()}.xlsx`);
  };

  return (
    <>
      {/* Floating Cloud Sync Status Banner */}
      <AnimatePresence>
        {cloudNotification?.show && (
          <motion.div
            initial={{ opacity: 0, y: -50, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, scale: 0.95, y: -20 }}
            className="fixed top-4 left-4 right-4 sm:left-auto sm:right-4 z-[9999] max-w-sm mx-auto sm:mx-0"
          >
            <div className={`p-4 rounded-2xl shadow-2xl border flex items-start gap-4 transition-colors ${
              theme === 'dark' 
                ? 'bg-slate-900/95 border-slate-800 text-white backdrop-blur-md' 
                : 'bg-white/95 border-slate-200 text-slate-800 backdrop-blur-md'
            }`}>
              <div className={`p-1.5 rounded-xl shrink-0 ${
                cloudNotification.type === 'error' ? 'bg-rose-500/10 text-rose-500' :
                cloudNotification.type === 'restore_success' ? 'bg-emerald-500/10 text-emerald-500' :
                'bg-indigo-500/10 text-indigo-550'
              }`}>
                {cloudNotification.type === 'error' ? (
                  <Cloud className="w-5 h-5 text-rose-500" />
                ) : (
                  <Cloud className="w-5 h-5 text-emerald-500" />
                )}
              </div>
              <div className="flex-1 text-left">
                <p className="text-[10px] font-black uppercase tracking-wider text-indigo-500 dark:text-indigo-400 font-mono">
                  {cloudNotification.type === 'restore_success' ? 'Cloud Restored' : 'Cloud Sync Active'}
                </p>
                <p className={`text-[11px] font-bold leading-relaxed mt-0.5 ${theme === 'dark' ? 'text-slate-300' : 'text-slate-650'}`}>
                  {cloudNotification.message}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setCloudNotification(null)}
                className={`text-[10px] font-bold uppercase tracking-widest px-2 py-1 rounded-md transition-colors cursor-pointer ${
                  theme === 'dark' ? 'text-slate-500 hover:text-slate-300' : 'text-slate-400 hover:text-slate-600'
                }`}
              >
                Close
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Custom Confirmation Dialog Overlay */}
      <AnimatePresence>
        {confirmDialog?.show && (
          <div className="fixed inset-0 z-[10000] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 15 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              className={`w-full max-w-sm rounded-[2rem] p-6 shadow-2xl border transition-colors ${
                theme === 'dark' 
                  ? 'bg-slate-900 border-slate-800 text-white' 
                  : 'bg-white border-slate-200 text-slate-800'
              }`}
            >
              <h3 className={`text-sm font-black uppercase tracking-wider mb-2 ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                {confirmDialog.title}
              </h3>
              <p className={`text-xs font-bold leading-relaxed mb-6 transition-colors ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                {confirmDialog.message}
              </p>
              <div className="flex gap-3 justify-end">
                <button
                  type="button"
                  onClick={() => setConfirmDialog(null)}
                  className={`px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                    theme === 'dark' 
                      ? 'bg-slate-800 text-slate-300 hover:bg-slate-700' 
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={confirmDialog.onConfirm}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all cursor-pointer active:scale-95"
                >
                  Confirm
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Custom Toast Alerts */}
      <AnimatePresence>
        {toastMessage?.show && (
          <motion.div
            initial={{ opacity: 0, y: 50, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, scale: 0.95, y: 20 }}
            className="fixed bottom-4 left-4 right-4 sm:left-auto sm:right-4 z-[9999] max-w-sm mx-auto sm:mx-0"
          >
            <div className={`p-4 rounded-2xl shadow-2xl border flex items-center gap-3 transition-colors ${
              theme === 'dark' 
                ? 'bg-slate-900/95 border-slate-800 text-white backdrop-blur-md' 
                : 'bg-white/95 border-slate-200 text-slate-850 backdrop-blur-md'
            }`}>
              <div className={`p-1.5 rounded-xl shrink-0 ${
                toastMessage.type === 'error' ? 'bg-rose-500/10 text-rose-500' :
                toastMessage.type === 'success' ? 'bg-emerald-500/10 text-emerald-500' :
                'bg-indigo-500/10 text-indigo-500'
              }`}>
                {toastMessage.type === 'error' ? (
                  <Cloud className="w-4 h-4 text-rose-500" />
                ) : toastMessage.type === 'success' ? (
                  <Cloud className="w-4 h-4 text-emerald-500" />
                ) : (
                  <Cloud className="w-4 h-4 text-indigo-500" />
                )}
              </div>
              <p className={`flex-1 text-left text-xs font-bold leading-normal transition-colors ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                {toastMessage.message}
              </p>
              <button
                type="button"
                onClick={() => setToastMessage(null)}
                className={`text-[9.5px] font-bold uppercase tracking-wider px-1.5 py-1 rounded transition-colors cursor-pointer ${
                  theme === 'dark' ? 'text-slate-500 hover:text-slate-300' : 'text-slate-400 hover:text-slate-600'
                }`}
              >
                Dismiss
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <div className={`min-h-screen font-sans selection:bg-indigo-100 pb-20 transition-colors duration-300 ${
        theme === 'dark' 
          ? 'bg-slate-950 text-slate-100' 
          : 'bg-slate-50 text-slate-900'
      }`}>
      {/* Header */}
      <header className={`border-b sticky top-0 z-10 px-4 py-4 flex items-center justify-between shadow-sm transition-colors duration-300 ${
        theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'
      }`}>
        <div className="flex items-center gap-2">
          <div className="bg-indigo-600 p-2 rounded-lg">
            <TrendingUp className="w-6 h-6 text-white" />
          </div>
          <h1 className={`text-xl font-black tracking-tighter ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Trade track P&L</h1>
        </div>
        <div className="flex items-center gap-1">
          {totalFunds > 0 && (
            <div 
              onClick={() => setShowSettings(true)}
              className="flex items-center gap-2 px-3 py-1.5 rounded-xl border cursor-pointer transition-all bg-slate-50 border-slate-100 text-slate-600 dark:bg-slate-800 dark:border-slate-700 dark:text-slate-300"
            >
              <Wallet className="w-3.5 h-3.5 text-indigo-500" />
              <span className="text-xs font-bold font-mono">₹{totalFunds.toLocaleString('en-IN')}</span>
            </div>
          )}
          <button 
            onClick={toggleTheme}
            className={`p-2 rounded-full transition-colors ${
              theme === 'dark' ? 'hover:bg-slate-800 text-amber-400' : 'hover:bg-slate-100 text-slate-600'
            }`}
            title="Toggle Theme"
          >
            {theme === 'dark' ? <Sun className="w-6 h-6" /> : <Moon className="w-6 h-6" />}
          </button>

          <button 
            onClick={() => setShowSettings(true)}
            className={`hidden sm:block p-2 rounded-full transition-colors ${
              theme === 'dark' ? 'hover:bg-slate-800 text-slate-400' : 'hover:bg-slate-100 text-slate-600'
            }`}
            title="Settings"
          >
            <Settings className="w-6 h-6" />
          </button>
        </div>
      </header>

      <main className="max-w-md mx-auto p-4 space-y-6">
        <div className="space-y-6">


              {/* Dates */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest ml-1">Entry Date</label>
                  <input
                    type="date"
                    value={entryDate}
                    onChange={(e) => setEntryDate(e.target.value)}
                    className={`w-full border rounded-xl py-4 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white color-scheme-dark' : 'bg-white border-slate-200 text-slate-900'
                    }`}
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest ml-1">Exit Date</label>
                  <input
                    type="date"
                    value={exitDate}
                    onChange={(e) => setExitDate(e.target.value)}
                    className={`w-full border rounded-xl py-4 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white color-scheme-dark' : 'bg-white border-slate-200 text-slate-900'
                    }`}
                  />
                </div>
              </div>

              {/* Strategy Selection */}
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest ml-1">Trading Strategy</label>
                <div className="flex flex-wrap gap-2">
                  {['Intraday', 'Swing', 'Scalping'].map((s) => (
                    <button
                      key={s}
                      onClick={() => setStrategy(s)}
                      className={`px-4 py-2 rounded-full text-xs font-bold border transition-all ${
                        strategy === s 
                          ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' 
                          : `${theme === 'dark' ? 'bg-slate-800 border-slate-700 text-slate-400 hover:border-indigo-800' : 'bg-white border-slate-200 text-slate-500 hover:border-indigo-300'}`
                      }`}
                    >
                      {s}
                    </button>
                  ))}
                </div>
              </div>

              {/* Notes Field */}
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest ml-1">Trade Notes / Insights</label>
                <textarea
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  placeholder="Why did you take this trade? Any psychological insights?"
                  rows={2}
                  className={`w-full border rounded-xl py-3 px-4 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none ${
                    theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                  }`}
                />
              </div>

              {/* Collapsible Trade Journal Fields */}
              <div className={`p-4 rounded-2xl border transition-all ${
                theme === 'dark' ? 'border-slate-800 bg-slate-900/30' : 'border-slate-100 bg-slate-50/50'
              }`}>
                <button
                  type="button"
                  onClick={() => setShowAdvancedJournal(!showAdvancedJournal)}
                  className="w-full flex items-center justify-between text-left"
                >
                  <div className="flex items-center gap-2">
                    <BookOpen className="w-4 h-4 text-indigo-500 animate-pulse" />
                    <span className="text-[10px] font-bold text-slate-550 uppercase tracking-widest leading-none">Detailed Trade Journal</span>
                  </div>
                  <span className={`text-[9px] font-bold px-2 py-0.5 rounded-full transition-colors ${
                    showAdvancedJournal 
                      ? 'bg-rose-500/10 text-rose-500' 
                      : 'bg-indigo-500/10 text-indigo-500'
                  }`}>
                    {showAdvancedJournal ? 'Hide Details' : 'Add Details'}
                  </span>
                </button>

                {showAdvancedJournal && (
                  <div className="mt-4 space-y-4">
                    <div className="space-y-1">
                      <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest ml-1">Trade Setup</label>
                      <textarea
                        value={tradeSetup}
                        onChange={(e) => setTradeSetup(e.target.value)}
                        placeholder="e.g., Symmetrical triangle breakout with 1.5x average daily volume"
                        rows={2}
                        className={`w-full border rounded-xl py-2 px-3 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>

                    <div className="space-y-1">
                      <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest ml-1">Entry & Exit Rationale</label>
                      <textarea
                        value={entryExitRationale}
                        onChange={(e) => setEntryExitRationale(e.target.value)}
                        placeholder="e.g., Retest limit order filled. Exited half at R1 and remainder at target."
                        rows={2}
                        className={`w-full border rounded-xl py-2 px-3 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>

                    <div className="space-y-1">
                      <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest ml-1">Lessons Learned</label>
                      <textarea
                        value={lessonsLearned}
                        onChange={(e) => setLessonsLearned(e.target.value)}
                        placeholder="e.g., Don't trail stop too closely on high volatility events."
                        rows={2}
                        className={`w-full border rounded-xl py-2 px-3 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>
                  </div>
                )}
              </div>

              {/* Position Pricing Setup Section Separator */}
              <div className={`p-4 rounded-2xl border transition-all ${
                theme === 'dark' ? 'border-slate-800 bg-slate-900/10' : 'border-slate-200 bg-slate-50/30'
              }`}>
                <div className="flex items-center gap-2">
                  <TrendingUp className="w-4 h-4 text-indigo-500 animate-pulse" />
                  <span className="text-[10px] font-black text-slate-450 uppercase tracking-widest leading-none">Position Pricing Setup</span>
                </div>
              </div>

              {/* Inputs */}
              <div className="space-y-4">
                <div className="space-y-1.5" id="entry-input-group">
                  <label className={`text-sm font-medium ml-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>{mode === 'equity' ? 'Entry Price' : 'Premium Bought'} (₹)</label>
                  <div className="relative">
                    <IndianRupee className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="number"
                      inputMode="decimal"
                      value={entryPrice}
                      onChange={(e) => setEntryPrice(e.target.value)}
                      placeholder="0.00"
                      className={`w-full border rounded-xl py-4 pl-10 pr-4 focus:ring-2 focus:ring-indigo-500 outline-none transition-all ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                      }`}
                    />
                  </div>
                </div>

                <div className="space-y-1.5" id="exit-input-group">
                  <label className={`text-sm font-medium ml-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>{mode === 'equity' ? 'Exit Price' : 'Premium Sold'} (₹)</label>
                  <div className="relative">
                    <IndianRupee className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="number"
                      inputMode="decimal"
                      value={exitPrice}
                      onChange={(e) => setExitPrice(e.target.value)}
                      placeholder="0.00"
                      className={`w-full border rounded-xl py-4 pl-10 pr-4 focus:ring-2 focus:ring-indigo-500 outline-none transition-all ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                      }`}
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1.5" id="quantity-input-group">
                    <label className={`text-sm font-medium ml-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>{mode === 'equity' ? 'Quantity' : 'Lots'} <span className="text-xs opacity-60">(Optional)</span></label>
                    <input
                      type="number"
                      inputMode="decimal"
                      value={quantity}
                      onChange={(e) => setQuantity(e.target.value)}
                      placeholder="1"
                      className={`w-full border rounded-xl py-4 px-4 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-lg font-medium ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                      }`}
                    />
                  </div>
                  {mode === 'options' && (
                    <div className="space-y-1.5" id="lot-size-group">
                      <label className={`text-sm font-medium ml-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Lot Size</label>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={lotSize}
                        onChange={(e) => setLotSize(e.target.value)}
                        placeholder="65"
                        className={`w-full border rounded-xl py-4 px-4 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-lg font-medium ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                        }`}
                      />
                    </div>
                  )}
                </div>

                {/* Advanced Section Toggle */}
                <button 
                  onClick={() => setShowAdvanced(!showAdvanced)}
                  className="flex items-center gap-1.5 text-indigo-600 font-medium text-sm pt-2 ml-1"
                  id="toggle-advanced"
                >
                  {showAdvanced ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                  Advanced Options (Leverage & Fees)
                </button>

                <AnimatePresence>
                  {showAdvanced && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: 'auto', opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      className="overflow-hidden space-y-4"
                    >
                      <div className="space-y-1.5" id="leverage-slider-group">
                        <div className={`flex justify-between items-center text-sm font-medium px-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                          <label>Leverage</label>
                          <span className="text-indigo-600 font-bold">{leverage}x</span>
                        </div>
                        <input
                          type="range"
                          min="1"
                          max="100"
                          step="1"
                          value={leverage}
                          onChange={(e) => setLeverage(parseInt(e.target.value))}
                          className={`w-full h-2 rounded-lg appearance-none cursor-pointer accent-indigo-600 ${
                            theme === 'dark' ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        />
                        <div className="flex justify-between text-[10px] text-slate-400 font-mono px-1">
                          <span>1x</span>
                          <span>25x</span>
                          <span>50x</span>
                          <span>75x</span>
                          <span>100x</span>
                        </div>
                      </div>

                      <div className="space-y-1.5" id="fees-input-group">
                        <label className={`text-sm font-medium ml-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Brokerage & Taxes (%)</label>
                        <div className="relative">
                          <Percent className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                          <input
                            type="number"
                            step="0.001"
                            value={feePercentage}
                            onChange={(e) => setFeePercentage(e.target.value)}
                            placeholder="e.g. 0.05"
                            className={`w-full border rounded-xl py-3 px-4 focus:ring-2 focus:ring-indigo-500 outline-none transition-all ${
                              theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-600' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-300'
                            }`}
                          />
                        </div>
                        <p className="text-[10px] text-slate-400 px-1 italic">Includes STT, GST, SEBI charges, etc.</p>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Results Card (Visible in Step 1) */}
              <AnimatePresence mode="wait">
                {results ? (
                  <div className="space-y-6">
                    <motion.div
                      key="results"
                      initial={{ scale: 0.95, opacity: 0, y: 10 }}
                      animate={{ scale: 1, opacity: 1, y: 0 }}
                      exit={{ scale: 0.95, opacity: 0, y: 10 }}
                      className={`p-6 rounded-2xl shadow-xl border-t-4 transition-colors duration-300 ${
                        results.isProfit 
                          ? theme === 'dark' ? 'bg-green-950/20 border-green-500' : 'bg-green-50 border-green-500' 
                          : theme === 'dark' ? 'bg-red-950/20 border-red-500' : 'bg-red-50 border-red-500'
                      }`}
                      id="results-card"
                    >
                      <div className="flex justify-between items-start mb-6">
                        <div>
                          <h3 className={`text-sm font-bold uppercase tracking-wider mb-1 ${
                            results.isProfit ? 'text-green-500' : 'text-red-500'
                          }`}>
                            {results.isProfit ? 'Profit' : 'Loss'}
                          </h3>
                          <p className={`text-4xl font-extrabold tracking-tight ${
                            results.isProfit 
                              ? theme === 'dark' ? 'text-green-400' : 'text-green-700' 
                              : theme === 'dark' ? 'text-red-400' : 'text-red-700'
                          }`}>
                            {results.profit >= 0 ? '+' : ''}₹{results.profit.toLocaleString('en-IN', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                          </p>
                        </div>
                        <div className="flex gap-2">
                          <button 
                            onClick={downloadPDF}
                            className={`p-2.5 rounded-full shadow-sm transition-transform active:scale-90 ${
                              theme === 'dark' ? 'bg-slate-800 text-slate-400 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                            }`}
                            title="Download PDF"
                          >
                            <FileText className="w-5 h-5" />
                          </button>
                          <button 
                            onClick={downloadExcel}
                            className={`p-2.5 rounded-full shadow-sm transition-transform active:scale-90 ${
                              theme === 'dark' ? 'bg-slate-800 text-slate-400 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                            }`}
                            title="Download Excel"
                          >
                            <TableIcon className="w-5 h-5" />
                          </button>
                          <button 
                            onClick={() => {
                              saveTrade();
                              const btn = document.getElementById('save-confirm-toast');
                              if (btn) btn.classList.remove('opacity-0');
                              setTimeout(() => {
                                if (btn) btn.classList.add('opacity-0');
                              }, 2000);
                            }}
                            className={`p-2.5 rounded-full shadow-md transition-transform active:scale-90 ${
                              results.isProfit ? 'bg-green-500 text-white' : 'bg-red-500 text-white'
                            }`}
                            id="save-trade"
                            title="Save to History"
                          >
                            <Plus className="w-5 h-5" />
                          </button>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-4">
                        <div className={`${theme === 'dark' ? 'bg-slate-900/50' : 'bg-white/50'} p-3 rounded-xl transition-colors`}>
                          <p className="text-[10px] font-bold text-slate-400 mb-1 leading-none uppercase">Return (%)</p>
                          <p className={`text-lg font-bold ${results.isProfit ? 'text-green-600' : 'text-red-600'}`}>
                            {results.profitPercentage > 0 ? '+' : ''}{results.profitPercentage.toFixed(2)}%
                          </p>
                        </div>
                        <div className={`${theme === 'dark' ? 'bg-slate-900/50' : 'bg-white/50'} p-3 rounded-xl transition-colors`}>
                          <p className="text-[10px] font-bold text-slate-400 mb-1 leading-none uppercase">ROE (Leveraged)</p>
                          <p className={`text-lg font-bold ${results.isProfit ? 'text-green-600' : 'text-red-600'}`}>
                            {results.roe > 0 ? '+' : ''}{results.roe.toFixed(2)}%
                          </p>
                        </div>
                        <div className={`${theme === 'dark' ? 'bg-slate-900/50' : 'bg-white/50'} p-3 rounded-xl transition-colors relative group`}>
                          <p className="text-[10px] font-bold text-slate-400 mb-1 leading-none uppercase">Margin Used</p>
                          <p className={`text-lg font-bold ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                            ₹{( (parseFloat(entryPrice) * (mode === 'options' ? parsedQty * parseFloat(lotSize) : parsedQty)) / leverage).toLocaleString('en-IN', { maximumFractionDigits: 0 })}
                          </p>
                          {totalFunds > 0 && (parseFloat(entryPrice) * (mode === 'options' ? parsedQty * parseFloat(lotSize) : parsedQty)) / leverage > totalFunds && (
                            <div className="absolute -top-2 -right-1">
                              <span className="flex h-3 w-3 relative">
                                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                                <span className="relative inline-flex rounded-full h-3 w-3 bg-rose-500"></span>
                              </span>
                            </div>
                          )}
                        </div>
                        <div className={`${theme === 'dark' ? 'bg-slate-900/50' : 'bg-white/50'} p-3 rounded-xl transition-colors`}>
                          <p className="text-[10px] font-bold text-slate-400 mb-1 leading-none uppercase">Total Fees/Tax</p>
                          <p className={`text-lg font-bold ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                            ₹{results.fees.toLocaleString('en-IN', { maximumFractionDigits: 0 })}
                          </p>
                        </div>
                      </div>

                      {/* PnL Graph */}
                      <div className={`mt-8 pt-6 border-t ${theme === 'dark' ? 'border-slate-800' : 'border-slate-200/50'}`}>
                        <div className="flex flex-col gap-4 mb-6">
                          <div className="flex justify-between items-center">
                            <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2">
                              <TrendingUp className="w-3 h-3 text-indigo-400" />
                              PnL Projection
                            </h4>
                            <div className="flex items-center gap-2">
                              {projectionZoomRange && (
                                <button
                                  onClick={() => setProjectionZoomRange(null)}
                                  className="px-2.5 py-1 text-[9px] font-black uppercase tracking-widest rounded bg-indigo-600 text-white transition-all hover:bg-indigo-700 active:scale-95 cursor-pointer shadow-md shadow-indigo-650/10"
                                >
                                  Reset View
                                </button>
                              )}
                              <div className={`${theme === 'dark' ? 'bg-slate-800' : 'bg-slate-100'} p-0.5 rounded-lg transition-colors flex`}>
                                {(['Day', 'Week', 'Month', 'Year'] as const).map((tf) => (
                                  <button
                                    key={tf}
                                    onClick={() => setGraphTimeframe(tf)}
                                    className={`px-2 py-1 text-[9px] font-bold rounded-md transition-all ${
                                      graphTimeframe === tf 
                                        ? `${theme === 'dark' ? 'bg-slate-700 text-white shadow-sm' : 'bg-white text-indigo-600 shadow-sm'}` 
                                        : 'text-slate-400 hover:text-slate-600'
                                    }`}
                                  >
                                    {tf}
                                  </button>
                                ))}
                              </div>
                            </div>
                          </div>

                          <div className={`flex justify-between items-center p-2 rounded-xl border transition-colors ${
                            theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-slate-50 border-slate-200/50'
                          }`}>
                            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Benchmark Index</span>
                            <div className="flex gap-1">
                              {(['Nifty', 'Sensex', 'Both', 'None'] as const).map((b) => (
                                <button
                                  key={b}
                                  onClick={() => setSelectedBenchmark(b)}
                                  className={`px-2 py-1 text-[9px] font-bold rounded-md border transition-all ${
                                    selectedBenchmark === b 
                                      ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm' 
                                      : `${theme === 'dark' ? 'bg-slate-800 border-slate-700 text-slate-500 hover:border-slate-600' : 'bg-white border-slate-200 text-slate-400 hover:border-slate-300'}`
                                  }`}
                                >
                                  {b}
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>

                        <div className="h-48 w-full -ml-4">
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart 
                              data={zoomedChartData}
                              onMouseDown={(e) => {
                                if (e && e.activeTooltipIndex !== undefined) {
                                  setProjectionRefLeft(e.activeTooltipIndex);
                                }
                              }}
                              onMouseMove={(e) => {
                                if (projectionRefLeft !== null && e && e.activeTooltipIndex !== undefined) {
                                  setProjectionRefRight(e.activeTooltipIndex);
                                }
                              }}
                              onMouseUp={() => {
                                if (projectionRefLeft !== null && projectionRefRight !== null && projectionRefLeft !== projectionRefRight) {
                                  const minIdx = Math.min(projectionRefLeft, projectionRefRight);
                                  const maxIdx = Math.max(projectionRefLeft, projectionRefRight);
                                  setProjectionZoomRange([minIdx, maxIdx]);
                                }
                                setProjectionRefLeft(null);
                                setProjectionRefRight(null);
                              }}
                              style={{ cursor: 'crosshair' }}
                            >
                              <defs>
                                <linearGradient id="colorProfit" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="#22c55e" stopOpacity={0.3}/>
                                  <stop offset="95%" stopColor="#22c55e" stopOpacity={0}/>
                                </linearGradient>
                                <linearGradient id="colorLoss" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="#ef4444" stopOpacity={0.3}/>
                                  <stop offset="95%" stopColor="#ef4444" stopOpacity={0}/>
                                </linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                              <XAxis 
                                dataKey="price" 
                                hide 
                                type="number" 
                                domain={['dataMin', 'dataMax']} 
                              />
                              <YAxis 
                                hide 
                                domain={([dataMin, dataMax]) => {
                                  const absMax = Math.max(Math.abs(dataMin), Math.abs(dataMax));
                                  return [-absMax * 1.1, absMax * 1.1];
                                }}
                              />
                              <Tooltip 
                                cursor={{ stroke: theme === 'dark' ? '#475569' : '#cbd5e1', strokeWidth: 1.5, strokeDasharray: '4 4' }}
                                content={({ active, payload }) => {
                                  if (active && payload && payload.length) {
                                    const dataPoint = payload[0].payload;
                                    const val = dataPoint.profit;
                                    const showNifty = selectedBenchmark === 'Nifty' || selectedBenchmark === 'Both';
                                    const showSensex = selectedBenchmark === 'Sensex' || selectedBenchmark === 'Both';
                                    
                                    return (
                                      <div className={`px-4 py-3 shadow-2xl border rounded-xl text-xs backdrop-blur-md transition-all duration-200 ${
                                        theme === 'dark' 
                                          ? 'bg-slate-900/95 border-slate-700 text-slate-100' 
                                          : 'bg-white/95 border-slate-200 text-slate-800'
                                      }`}>
                                        <div className="border-b border-slate-200/20 dark:border-slate-700/50 pb-2 mb-2">
                                          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-0.5">Price Point</p>
                                          <p className="font-extrabold text-base font-mono">
                                            ₹{dataPoint.price.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                                          </p>
                                        </div>
                                        <div className="space-y-1.5 font-medium">
                                          <div className="flex items-center justify-between gap-8">
                                            <span className="text-slate-400 text-[11px] flex items-center gap-1.5 font-bold uppercase tracking-wider">
                                              <span className={`w-2 h-2 rounded-full ${val >= 0 ? 'bg-green-500' : 'bg-red-500'}`}></span>
                                              Your PnL:
                                            </span>
                                            <span className={`font-mono font-bold ${val >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                              {val >= 0 ? '+' : ''}₹{val.toLocaleString('en-IN')}
                                            </span>
                                          </div>
                                          {showNifty && (
                                            <div className="flex items-center justify-between gap-8">
                                              <span className="text-slate-400 text-[11px] flex items-center gap-1.5 font-bold uppercase tracking-wider">
                                                <span className="w-2 h-0.5 bg-indigo-500"></span>
                                                Nifty 50:
                                              </span>
                                              <span className={`font-mono font-bold ${dataPoint.nifty >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                                {dataPoint.nifty >= 0 ? '+' : ''}₹{dataPoint.nifty.toLocaleString('en-IN')}
                                              </span>
                                            </div>
                                          )}
                                          {showSensex && (
                                            <div className="flex items-center justify-between gap-8">
                                              <span className="text-slate-400 text-[11px] flex items-center gap-1.5 font-bold uppercase tracking-wider">
                                                <span className="w-2 h-0.5 bg-amber-500"></span>
                                                Sensex:
                                              </span>
                                              <span className={`font-mono font-bold ${dataPoint.sensex >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                                {dataPoint.sensex >= 0 ? '+' : ''}₹{dataPoint.sensex.toLocaleString('en-IN')}
                                              </span>
                                            </div>
                                          )}
                                        </div>
                                      </div>
                                    );
                                  }
                                  return null;
                                }}
                              />
                              <ReferenceLine y={0} stroke="#94A3B8" strokeWidth={1} strokeDasharray="3 3" />
                              <ReferenceLine x={parseFloat(entryPrice)} stroke="#6366F1" strokeWidth={2} label={{ value: 'Entry', position: 'top', fill: '#6366F1', fontSize: 10, fontWeight: 'bold' }} />
                              <Area
                                type="monotone"
                                dataKey="profit"
                                name="Your Trade"
                                stroke={results.isProfit ? "#22c55e" : "#ef4444"}
                                fillOpacity={1}
                                fill={results.isProfit ? "url(#colorProfit)" : "url(#colorLoss)"}
                                strokeWidth={3}
                                activeDot={{ r: 6, strokeWidth: 0, fill: results.isProfit ? '#22c55e' : '#ef4444' }}
                              />
                              {(selectedBenchmark === 'Nifty' || selectedBenchmark === 'Both') && (
                                <Line 
                                  type="monotone" 
                                  dataKey="nifty" 
                                  name="Nifty 50" 
                                  stroke="#6366f1" 
                                  strokeWidth={2} 
                                  strokeDasharray="5 5" 
                                  dot={false}
                                  activeDot={{ r: 4 }}
                                />
                              )}
                              {(selectedBenchmark === 'Sensex' || selectedBenchmark === 'Both') && (
                                <Line 
                                  type="monotone" 
                                  dataKey="sensex" 
                                  name="Sensex" 
                                  stroke="#f59e0b" 
                                  strokeWidth={2} 
                                  strokeDasharray="5 5" 
                                  dot={false}
                                  activeDot={{ r: 4 }}
                                />
                              )}
                              {projectionRefLeft !== null && projectionRefRight !== null && (
                                <ReferenceArea
                                  {...({
                                    x1: zoomedChartData[projectionRefLeft]?.price,
                                    x2: zoomedChartData[projectionRefRight]?.price,
                                    fill: "#6366f1",
                                    fillOpacity: 0.16
                                  } as any)}
                                />
                              )}
                              {projectionZoomRange && (
                                <ReferenceArea
                                  {...({
                                    x1: zoomedChartData[0]?.price,
                                    x2: zoomedChartData[zoomedChartData.length - 1]?.price,
                                    fill: "#6366f1",
                                    fillOpacity: 0.05,
                                    stroke: "#6366f1",
                                    strokeDasharray: "3 3",
                                    strokeOpacity: 0.4
                                  } as any)}
                                />
                              )}
                            </AreaChart>
                          </ResponsiveContainer>
                        </div>
                        <div className={`flex flex-wrap justify-center gap-4 mt-6 pt-4 border-t transition-colors ${theme === 'dark' ? 'border-slate-800' : 'border-slate-100'}`}>
                          <div className="flex items-center gap-1.5">
                            <div className="w-3 h-3 rounded-full bg-green-500"></div>
                            <span className={`text-[10px] font-bold uppercase tracking-wider transition-colors ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Your Trade</span>
                          </div>
                          {(selectedBenchmark === 'Nifty' || selectedBenchmark === 'Both') && (
                            <div className="flex items-center gap-1.5">
                              <div className="w-3 h-0.5 bg-indigo-500 border-t-2 border-dashed border-indigo-500"></div>
                              <span className={`text-[10px] font-bold uppercase tracking-wider transition-colors ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Nifty 50</span>
                            </div>
                          )}
                          {(selectedBenchmark === 'Sensex' || selectedBenchmark === 'Both') && (
                            <div className="flex items-center gap-1.5">
                              <div className="w-3 h-0.5 bg-amber-500 border-t-2 border-dashed border-amber-500"></div>
                              <span className={`text-[10px] font-bold uppercase tracking-wider transition-colors ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Sensex</span>
                            </div>
                          )}
                        </div>
                        <div className="flex justify-between mt-4 px-1 text-[9px] font-bold text-slate-400 uppercase tracking-tighter">
                          <span>{graphTimeframe} Bear Case</span>
                          <span>Exit Potential Range</span>
                          <span>{graphTimeframe} Bull Case</span>
                        </div>
                      </div>
                    </motion.div>

                    <div className="grid grid-cols-2 gap-3">
                      <button
                        onClick={reset}
                        className={`py-4 rounded-2xl font-bold flex items-center justify-center gap-2 active:scale-[0.98] transition-all border cursor-pointer ${
                          theme === 'dark' ? 'bg-slate-800 text-slate-300 border-slate-700 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200'
                        }`}
                        id="previous-button"
                      >
                        <RefreshCcw className="w-4 h-4" />
                        Clear Form
                      </button>

                      <button
                        onClick={() => {
                          saveTrade();
                          const btn = document.getElementById('save-confirm-toast');
                          if (btn) btn.classList.remove('opacity-0');
                          setTimeout(() => {
                            if (btn) btn.classList.add('opacity-0');
                          }, 2000);
                        }}
                        className={`py-4 rounded-2xl font-bold shadow-md flex items-center justify-center gap-2 active:scale-[0.98] transition-all cursor-pointer ${
                          theme === 'dark' ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'bg-slate-900 text-white hover:bg-slate-800'
                        }`}
                        id="last-save-button"
                      >
                        <Plus className="w-5 h-5" />
                        Save Trade
                      </button>
                    </div>
                    
                    <div 
                      id="save-confirm-toast" 
                      className="opacity-0 transition-opacity duration-300 text-center text-xs font-bold text-green-600"
                    >
                      ✓ Trade saved successfully!
                    </div>

                    <div className="pt-2 space-y-2">
                      <button
                        onClick={downloadExcel}
                        className={`w-full py-3 rounded-2xl font-bold border flex items-center justify-center gap-2 transition-all active:scale-[0.98] ${
                          theme === 'dark' ? 'bg-emerald-900/20 text-emerald-400 border-emerald-900/30 hover:bg-emerald-900/30' : 'bg-emerald-50 text-emerald-700 border-emerald-100 hover:bg-emerald-100'
                        }`}
                      >
                        <TableIcon className="w-5 h-5" />
                        Export Result to Excel
                      </button>
                      <button
                        onClick={downloadPDF}
                        className={`w-full py-3 rounded-2xl font-bold border flex items-center justify-center gap-2 transition-all active:scale-[0.98] ${
                          theme === 'dark' ? 'bg-indigo-900/20 text-indigo-400 border-indigo-900/30 hover:bg-indigo-900/30' : 'bg-indigo-50 text-indigo-700 border-indigo-100 hover:bg-indigo-100'
                        }`}
                      >
                        <FileText className="w-5 h-5" />
                        Export Result to PDF
                      </button>
                    </div>
                  </div>
                ) : (
                  <motion.div
                    key="placeholder"
                    className={`border-2 border-dashed p-12 rounded-2xl flex flex-col items-center justify-center text-center transition-colors ${
                      theme === 'dark' ? 'bg-slate-800/50 border-slate-700 text-slate-500' : 'bg-slate-100 border-slate-300 text-slate-400'
                    }`}
                    id="placeholder-card"
                  >
                    <CalcIcon className="w-12 h-12 mb-3 opacity-20" />
                    <p className="font-medium">Enter values to see results</p>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
      </main>

      {/* Profile Modal Overlay */}
      <AnimatePresence>

      </AnimatePresence>

      <AnimatePresence>
        {showCalendar && (
          <div className="fixed inset-0 z-50 flex items-end justify-center sm:items-center p-0 sm:p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              onClick={() => setShowCalendar(false)}
            />
            <motion.div
              initial={{ y: '100%' }}
              animate={{ y: 0 }}
              exit={{ y: '100%' }}
              className={`relative w-full max-w-lg rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] flex flex-col transition-colors duration-300 ${
                theme === 'dark' ? 'bg-slate-900' : 'bg-white'
              }`}
              id="calendar-panel"
            >
              <div className={`h-1.5 w-12 rounded-full mx-auto my-4 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />
              
              <div className={`px-4 pb-3 pt-1 border-b flex items-center gap-2.5 shrink-0 transition-colors ${
                theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
              }`}>
                <button
                  onClick={() => setShowCalendar(false)}
                  className={`p-1.5 rounded-xl border transition-all active:scale-95 ${
                    theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-755' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                  }`}
                  aria-label="Back to Trading"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <div className="flex flex-col flex-1 min-w-0">
                  <h2 className={`text-base font-extrabold truncate transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>P&L Calendar</h2>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest truncate">
                    Monthly Performance
                  </p>
                </div>
                <div className={`flex items-center gap-1 p-0.5 rounded-lg transition-colors shrink-0 ${theme === 'dark' ? 'bg-slate-800' : 'bg-slate-100'}`}>
                  <button 
                    onClick={() => {
                      if (currentMonth === 0) {
                        setCurrentMonth(11);
                        setCurrentYear(y => y - 1);
                      } else {
                        setCurrentMonth(m => m - 1);
                      }
                    }}
                    className={`p-1 rounded-md transition-all ${theme === 'dark' ? 'hover:bg-slate-700 text-slate-300' : 'hover:bg-white text-slate-600'}`}
                  >
                    <ChevronLeft className="w-3.5 h-3.5" />
                  </button>
                  <span className={`text-[11px] font-bold min-w-[75px] text-center transition-colors ${theme === 'dark' ? 'text-slate-200' : 'text-slate-700'}`}>
                    {new Date(currentYear, currentMonth).toLocaleString('default', { month: 'short', year: 'numeric' })}
                  </span>
                  <button 
                    onClick={() => {
                      if (currentMonth === 11) {
                        setCurrentMonth(0);
                        setCurrentYear(y => y + 1);
                      } else {
                        setCurrentMonth(m => m + 1);
                      }
                    }}
                    className={`p-1 rounded-md transition-all ${theme === 'dark' ? 'hover:bg-slate-700 text-slate-300' : 'hover:bg-white text-slate-600'}`}
                  >
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto p-4">
                    <div className="grid grid-cols-7 gap-1 mb-2">
                      {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(d => (
                        <div key={d} className="text-center text-[10px] font-bold text-slate-400 uppercase py-2">
                          {d}
                        </div>
                      ))}
                    </div>
                    
                    <div className="grid grid-cols-7 gap-1">
                      {Array.from({ length: firstDayOfMonth(currentMonth, currentYear) }).map((_, i) => (
                        <div key={`empty-${i}`} className="aspect-square" />
                      ))}
                      
                      {Array.from({ length: daysInMonth(currentMonth, currentYear) }).map((_, i) => {
                        const day = i + 1;
                        const dateStr = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                        const pnl = calendarData[dateStr];
                        const isToday = new Date().toISOString().split('T')[0] === dateStr;
                        const isSelected = selectedCalendarDate === dateStr;
                        const holiday = getHolidayName(dateStr);
                        const isWeekend = holiday === 'Weekend';
                        const isMarketHoliday = holiday && !isWeekend;

                        // Custom Theme Holiday Styling
                        const isThemedHoliday = isMarketHoliday || (isWeekend && includeWeekendsInHolidayTheme);

                        // Holiday background color style resolver
                        let holidayBackgroundClass = '';
                        if (isThemedHoliday) {
                          if (holidayColorTheme === 'rose') {
                            holidayBackgroundClass = theme === 'dark' ? 'bg-rose-950/20 border-rose-900/40 text-rose-300' : 'bg-rose-50/50 border-rose-200 text-rose-600';
                          } else if (holidayColorTheme === 'amber') {
                            holidayBackgroundClass = theme === 'dark' ? 'bg-amber-950/20 border-amber-900/35 text-amber-300' : 'bg-amber-50/50 border-amber-200 text-amber-600';
                          } else if (holidayColorTheme === 'indigo') {
                            holidayBackgroundClass = theme === 'dark' ? 'bg-indigo-950/20 border-indigo-900/40 text-indigo-300' : 'bg-indigo-50/50 border-indigo-200 text-indigo-650';
                          } else if (holidayColorTheme === 'emerald') {
                            holidayBackgroundClass = theme === 'dark' ? 'bg-emerald-950/20 border-emerald-900/40 text-emerald-300' : 'bg-emerald-50/50 border-emerald-200 text-emerald-600';
                          } else if (holidayColorTheme === 'violet') {
                            holidayBackgroundClass = theme === 'dark' ? 'bg-violet-950/20 border-violet-900/40 text-violet-300' : 'bg-violet-50/50 border-violet-200 text-violet-650';
                          } else { // 'slate'
                            holidayBackgroundClass = theme === 'dark' ? 'bg-slate-900/25 border-dashed border-slate-800 text-slate-400' : 'bg-slate-100/50 border-dashed border-slate-205 text-slate-505';
                          }
                        } else if (isWeekend) {
                          holidayBackgroundClass = theme === 'dark' ? 'bg-slate-900/15 border-dashed border-slate-850 text-slate-400' : 'bg-slate-50/40 border-dashed border-slate-200/80 text-slate-500';
                        } else {
                          holidayBackgroundClass = theme === 'dark' ? 'bg-slate-800/40 text-slate-400' : 'bg-white text-slate-600';
                        }

                        return (
                          <button 
                            key={day}
                            onClick={() => handleSelectCalendarDate(isSelected ? null : dateStr)}
                            className={`aspect-square rounded-xl border flex flex-col items-center justify-center relative transition-all cursor-pointer ${
                              isSelected
                                ? 'border-indigo-500 ring-2 ring-indigo-500/40 scale-[1.04] z-10'
                                : isToday 
                                  ? theme === 'dark' ? 'border-amber-450 shadow-sm shadow-amber-900/40' : 'border-amber-400 shadow-sm' 
                                  : theme === 'dark' ? 'border-slate-800 hover:border-slate-700' : 'border-slate-100 hover:border-slate-200'
                            } ${
                              pnl !== undefined 
                                ? pnl >= 0 
                                  ? theme === 'dark' ? 'bg-green-900/25 border-green-800' : 'bg-green-50 border-green-200' 
                                  : theme === 'dark' ? 'bg-red-900/25 border-red-800' : 'bg-red-50 border-red-200'
                                : holidayBackgroundClass
                            }`}
                            title={`${dateStr}${holiday ? ` (${holiday})` : ''}: ${pnl !== undefined ? `₹${pnl.toFixed(2)}` : 'No Trades logged'}`}
                          >
                            <span className={`text-[10px] font-bold ${
                              isSelected 
                                ? 'text-indigo-400' 
                                : isToday 
                                  ? 'text-amber-550 dark:text-amber-405' 
                                  : isThemedHoliday
                                    ? holidayColorTheme === 'rose' ? 'text-rose-500 dark:text-rose-400 font-extrabold' :
                                      holidayColorTheme === 'amber' ? 'text-amber-600 dark:text-amber-400 font-extrabold' :
                                      holidayColorTheme === 'indigo' ? 'text-indigo-600 dark:text-indigo-400 font-extrabold' :
                                      holidayColorTheme === 'emerald' ? 'text-emerald-600 dark:text-emerald-450 font-extrabold' :
                                      holidayColorTheme === 'violet' ? 'text-violet-605 dark:text-violet-400 font-extrabold' :
                                      'text-slate-500 dark:text-slate-400 font-extrabold'
                                    : theme === 'dark' ? 'text-slate-400' : 'text-slate-500'
                            }`}>
                              {day}
                            </span>
                            
                            {holiday && pnl === undefined && (
                              <span 
                                class-details="holiday-marker"
                                className={`absolute top-1 right-1 w-1.5 h-1.5 rounded-full ${
                                  isThemedHoliday
                                    ? holidayColorTheme === 'rose' ? 'bg-rose-500' :
                                      holidayColorTheme === 'amber' ? 'bg-amber-500' :
                                      holidayColorTheme === 'indigo' ? 'bg-indigo-500' :
                                      holidayColorTheme === 'emerald' ? 'bg-emerald-500' :
                                      holidayColorTheme === 'violet' ? 'bg-violet-500' :
                                      'bg-slate-400'
                                    : 'bg-slate-400/35'
                                }`} 
                              />
                            )}

                            {pnl !== undefined && (
                              <span className={`text-[8px] font-black leading-none mt-1 ${pnl >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                {pnl >= 0 ? '+' : ''}{Math.abs(pnl) >= 1000 ? `${(pnl/1000).toFixed(1)}k` : pnl.toFixed(0)}
                              </span>
                            )}
                          </button>
                        );
                      })}
                    </div>

                {/* Selected Calendar Date Quick-Add / Detail Focus Panel */}
                {selectedCalendarDate && (
                  <div className={`mt-6 p-4 rounded-2xl border transition-all animate-in fade-in slide-in-from-bottom-2 ${
                    theme === 'dark' ? 'bg-slate-900/80 border-slate-800 text-white' : 'bg-slate-50 border-slate-200/80'
                  }`}>
                    {/* Header */}
                    <div className="flex justify-between items-center mb-4">
                      <div>
                        <p className="text-[9px] font-black text-indigo-500 uppercase tracking-widest leading-none mb-1">Calendar Focus</p>
                        <h3 className={`text-xs font-black uppercase tracking-wide ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                          {new Date(selectedCalendarDate).toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' })}
                        </h3>
                      </div>
                      <button 
                        onClick={() => handleSelectCalendarDate(null)}
                        className={`text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded cursor-pointer transition-all ${
                          theme === 'dark' ? 'bg-slate-800 hover:bg-slate-700 text-slate-300' : 'bg-white border border-slate-200 hover:bg-slate-100 text-slate-600'
                        }`}
                      >
                        Clear Focus
                      </button>
                    </div>

                    {/* Holiday Status Section */}
                    {(() => {
                      const holidayName = getHolidayName(selectedCalendarDate);
                      const isCustomHoliday = !!customHolidays[selectedCalendarDate];
                      const isWeekend = holidayName === 'Weekend';
                      const isMarketHoliday = holidayName && !isWeekend;

                      return (
                        <div className={`p-3 rounded-xl border mb-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 transition-colors ${
                          isMarketHoliday
                            ? theme === 'dark' ? 'bg-amber-950/20 border-amber-900/30' : 'bg-amber-50 border-amber-100'
                            : isWeekend
                              ? theme === 'dark' ? 'bg-slate-950/20 border-slate-800' : 'bg-slate-100/50 border-slate-200/60'
                              : theme === 'dark' ? 'bg-slate-950/10 border-slate-800/60' : 'bg-white border-slate-100'
                        }`}>
                          <div className="flex items-center gap-2">
                            <Palmtree className={`w-4 h-4 shrink-0 ${isMarketHoliday ? 'text-amber-500' : isWeekend ? 'text-slate-400' : 'text-slate-350'}`} />
                            <div className="flex flex-col text-left">
                              <span className="text-[9px] font-black uppercase tracking-widest text-slate-400 leading-none mb-1">Trading Day Status</span>
                              <span className={`text-[11px] font-extrabold ${isMarketHoliday ? 'text-amber-600 dark:text-amber-400' : isWeekend ? 'text-slate-500' : theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                                {holidayName ? `Market Closed • ${holidayName}` : 'Market Open (Operational)'}
                              </span>
                            </div>
                          </div>

                          <div className="flex items-center gap-1.5 shrink-0">
                            {isCustomHoliday ? (
                              <button
                                type="button"
                                onClick={() => handleRemoveCustomHoliday(selectedCalendarDate)}
                                className="text-[9px] font-black uppercase tracking-widest bg-red-500/10 hover:bg-red-500/20 text-red-550 px-2.5 py-1.5 rounded-lg border border-red-500/20 cursor-pointer transition-all"
                              >
                                Remove Holiday
                              </button>
                            ) : !holidayName && (
                              <div className="flex items-center gap-1.5 w-full sm:w-auto">
                                <input
                                  type="text"
                                  id="custom-holiday-input"
                                  placeholder="New Custom Holiday"
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') {
                                      const inputEl = e.currentTarget;
                                      const val = inputEl.value.trim();
                                      if (val) {
                                        handleSaveCustomHoliday(selectedCalendarDate, val);
                                        inputEl.value = '';
                                      }
                                    }
                                  }}
                                  className={`px-3 py-1.5 rounded-lg border text-[10px] outline-none w-full sm:max-w-[140px] focus:ring-1 focus:ring-indigo-500 transition-all ${
                                    theme === 'dark' 
                                      ? 'bg-slate-950 border-slate-800 text-white placeholder:text-slate-650' 
                                      : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                                  }`}
                                />
                                <button
                                  type="button"
                                  onClick={() => {
                                    const input = document.getElementById('custom-holiday-input') as HTMLInputElement;
                                    const val = input?.value.trim();
                                    if (val) {
                                      handleSaveCustomHoliday(selectedCalendarDate, val);
                                      if (input) input.value = '';
                                    }
                                  }}
                                  className="text-[9px] font-black uppercase tracking-widest bg-indigo-650 hover:bg-indigo-700 text-white px-2.5 py-1.5 rounded-lg cursor-pointer transition-all shrink-0 font-bold"
                                >
                                  Mark
                                </button>
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })()}

                    {/* Trades list for this date */}
                    <div className="mb-4">
                      <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-2">Recorded Trades ({tradesOnSelectedDate.length})</p>
                      {tradesOnSelectedDate.length === 0 ? (
                        <p className="text-[10px] text-slate-400 font-bold italic py-1">No trades recorded for this date yet.</p>
                      ) : (
                        <div className="space-y-2 max-h-40 overflow-y-auto pr-1">
                          {tradesOnSelectedDate.map(trade => (
                            <div 
                              key={trade.id} 
                              className={`flex justify-between items-center p-2.5 rounded-xl text-xs border group relative transition-all ${
                                editingTradeId === trade.id
                                  ? 'border-indigo-500 ring-1 ring-indigo-500 bg-indigo-500/5'
                                  : theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80 text-white' : 'bg-white border-slate-100/80 text-slate-850'
                              }`}
                            >
                              <div className="flex items-center gap-2">
                                <span className={`w-2 h-2 rounded-full shrink-0 ${trade.profit >= 0 ? 'bg-green-500' : 'bg-red-500'}`} />
                                <div className="flex flex-col text-left">
                                  <span className="font-extrabold">{trade.ticker || 'Trade Entry'}</span>
                                  {trade.notes && trade.notes !== 'Quick calendar P&L entry' && (
                                    <span className="text-[10px] text-slate-400 dark:text-slate-500 font-medium truncate max-w-[140px]">{trade.notes}</span>
                                  )}
                                </div>
                                <span className={`text-[8px] px-1.5 py-0.5 font-black rounded uppercase shrink-0 ${
                                  trade.type === 'long' 
                                    ? 'bg-green-500/10 text-green-500' 
                                    : 'bg-red-500/10 text-red-500'
                                }`}>
                                  {trade.type}
                                </span>
                              </div>
                              
                              <div className="flex items-center gap-2.5 shrink-0">
                                <span className={`font-mono font-bold ${trade.profit >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                  {trade.profit >= 0 ? '+' : ''}₹{trade.profit.toLocaleString('en-IN')}
                                </span>
                                
                                <div className="flex items-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setEditingTradeId(trade.id);
                                      setQuickPnLAmount(Math.abs(trade.profit).toString());
                                      setQuickPnLType(trade.profit >= 0 ? 'profit' : 'loss');
                                      setQuickPnLTicker(trade.ticker || '');
                                      setQuickPnLNotes(trade.notes || '');
                                    }}
                                    className={`p-1 rounded-md transition-colors cursor-pointer ${
                                      theme === 'dark' ? 'hover:bg-slate-850 text-slate-400 hover:text-indigo-400' : 'hover:bg-slate-100 text-slate-500 hover:text-indigo-650'
                                    }`}
                                    title="Edit this entry in quick form"
                                  >
                                    <Edit3 className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => deleteTrade(trade.id)}
                                    className={`p-1 rounded-md transition-colors cursor-pointer ${
                                      theme === 'dark' ? 'hover:bg-slate-850 text-slate-400 hover:text-rose-400' : 'hover:bg-slate-100 text-slate-500 hover:text-rose-650'
                                    }`}
                                    title="Delete entry"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Quick Add Form Section */}
                    <div className={`p-5 rounded-2xl border transition-all ${
                      editingTradeId ? 'ring-2 ring-indigo-500/30' : ''
                    } ${
                      theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-white border-slate-200/50'
                    }`}>
                      <div className="flex items-center justify-between mb-3.5">
                        <span className="text-xs font-black text-indigo-500 uppercase tracking-widest leading-none">
                          {editingTradeId ? '✏️ Edit P&L Entry' : '⚡ Quick P&L Entry Form'}
                        </span>
                        {editingTradeId && (
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                            Editing Mode
                          </span>
                        )}
                      </div>

                      <div className="space-y-4">
                        <div className="grid grid-cols-2 gap-3.5">
                          <div>
                            <label className="text-[10px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest block mb-1.5">Entry Type</label>
                            <div className={`p-1 rounded-lg border flex gap-1 ${
                              theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-slate-50 border-slate-200'
                            }`}>
                              <button
                                type="button"
                                onClick={() => setQuickPnLType('profit')}
                                className={`flex-1 py-1.5 text-[10px] font-black uppercase tracking-widest rounded transition-all cursor-pointer ${
                                  quickPnLType === 'profit' 
                                    ? 'bg-green-500 text-white shadow-sm' 
                                    : 'text-slate-400 hover:text-green-500'
                                }`}
                              >
                                Profit
                              </button>
                              <button
                                type="button"
                                onClick={() => setQuickPnLType('loss')}
                                className={`flex-1 py-1.5 text-[10px] font-black uppercase tracking-widest rounded transition-all cursor-pointer ${
                                  quickPnLType === 'loss' 
                                    ? 'bg-red-500 text-white shadow-sm' 
                                    : 'text-slate-400 hover:text-red-500'
                                }`}
                              >
                                Loss
                              </button>
                            </div>
                          </div>

                          <div>
                            <label className="text-[10px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest block mb-1.5">Asset/Ticker Name</label>
                            <input
                              type="text"
                              value={quickPnLTicker}
                              onChange={(e) => setQuickPnLTicker(e.target.value)}
                              placeholder="e.g. RELIANCE CE, NIFTY"
                              className={`w-full py-2 px-3 rounded-lg border text-xs focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                                theme === 'dark' 
                                  ? 'bg-slate-900 border-slate-800 text-white placeholder:text-slate-655' 
                                  : 'bg-slate-50 border-slate-200 text-slate-900 placeholder:text-slate-350'
                              }`}
                            />
                          </div>
                        </div>

                        <div className="grid grid-cols-2 gap-3.5">
                          <div>
                            <label className="text-[10px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest block mb-1.5">PnL Amount (₹)</label>
                            <input
                              type="number"
                              min="0"
                              value={quickPnLAmount}
                              onChange={(e) => setQuickPnLAmount(e.target.value)}
                              placeholder="Sum of P&L"
                              className={`w-full py-2 px-3 rounded-lg border text-xs font-mono font-bold focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                                theme === 'dark' 
                                  ? 'bg-slate-900 border-slate-800 text-white' 
                                  : 'bg-slate-50 border-slate-200 text-slate-900'
                              }`}
                            />
                          </div>

                          <div>
                            <label className="text-[10px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest block mb-1.5">Brief Notes/Verdict</label>
                            <input
                              type="text"
                              value={quickPnLNotes}
                              onChange={(e) => setQuickPnLNotes(e.target.value)}
                              placeholder="Trade logic or thoughts"
                              className={`w-full py-2 px-3 rounded-lg border text-xs focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                                theme === 'dark' 
                                  ? 'bg-slate-900 border-slate-800 text-white placeholder:text-slate-655' 
                                  : 'bg-slate-50 border-slate-200 text-slate-900 placeholder:text-slate-350'
                              }`}
                            />
                          </div>
                        </div>

                        {/* Submissions */}
                        <div className="flex flex-col gap-2 pt-1.5">
                          <div className="flex gap-2 w-full">
                            <button
                              type="button"
                              onClick={handleSaveQuickPnL}
                              disabled={!quickPnLAmount || parseFloat(quickPnLAmount) <= 0}
                              className={`flex-1 py-3 rounded-lg font-black text-[10px] uppercase tracking-widest transition-all shadow-md active:scale-95 cursor-pointer ${
                                !quickPnLAmount || parseFloat(quickPnLAmount) <= 0
                                  ? 'bg-slate-300 text-slate-455 dark:bg-slate-800 dark:text-slate-605 cursor-not-allowed opacity-40'
                                  : quickPnLType === 'profit'
                                    ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-emerald-600/10'
                                    : 'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-600/10'
                              }`}
                            >
                              {editingTradeId ? 'Update Entry' : 'Save Quick Entry'}
                            </button>

                            {editingTradeId ? (
                              <button
                                type="button"
                                onClick={handleDeleteEditingTrade}
                                className="py-3 px-4 bg-rose-600 hover:bg-rose-700 text-white rounded-lg font-black text-[10px] uppercase tracking-widest transition-all active:scale-95 cursor-pointer flex items-center justify-center gap-1.5 shadow-md"
                                title="Delete this entry completely"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                                Delete
                              </button>
                            ) : (
                              <button
                                type="button"
                                onClick={() => {
                                  setManualEntryDate(selectedCalendarDate);
                                  setManualExitDate(selectedCalendarDate);
                                  setIsManualJournalOpen(true);
                                }}
                                className={`px-4 py-3 border rounded-lg font-black text-[10px] uppercase tracking-widest transition-all cursor-pointer ${
                                  theme === 'dark' 
                                    ? 'border-slate-800 hover:bg-slate-800 text-slate-300' 
                                    : 'border-slate-200 hover:bg-slate-150 text-slate-700'
                                }`}
                                title="Detailed form (prices, lots, strategies)"
                              >
                                Advanced Form
                              </button>
                            )}
                          </div>

                          {editingTradeId && (
                            <button
                              type="button"
                              onClick={handleCancelEditQuickPnL}
                              className={`w-full py-1.5 rounded-lg font-black text-[9px] uppercase tracking-widest transition-all hover:underline ${
                                theme === 'dark' ? 'text-slate-400 hover:text-white' : 'text-slate-500 hover:text-slate-850'
                              }`}
                            >
                              Cancel Edit
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                <div className="mt-8 grid grid-cols-2 gap-3.5">
                  <div className={`p-3.5 rounded-2xl border transition-colors ${
                    theme === 'dark' ? 'bg-green-950/15 border-green-900/30' : 'bg-emerald-50/55 border-emerald-100'
                  }`}>
                    <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest mb-1">Total Profits</p>
                    <p className="text-base font-black text-emerald-500">
                      ₹{activeMonthStats.totalProfits.toLocaleString('en-IN')}
                    </p>
                  </div>
                  <div className={`p-3.5 rounded-2xl border transition-colors ${
                    theme === 'dark' ? 'bg-red-950/15 border-red-900/30' : 'bg-rose-50/55 border-rose-100'
                  }`}>
                    <p className="text-[10px] font-bold text-rose-600 uppercase tracking-widest mb-1">Total Losses</p>
                    <p className="text-base font-black text-rose-500">
                      ₹{Math.abs(activeMonthStats.totalLosses).toLocaleString('en-IN')}
                    </p>
                  </div>

                  <div className={`p-3.5 rounded-2xl border transition-colors ${
                    theme === 'dark' ? 'bg-green-950/10 border-green-900/20' : 'bg-emerald-50/20 border-emerald-100/50'
                  }`}>
                    <div className="flex items-center gap-1 mb-1">
                      <TrendingUp className="w-3.5 h-3.5 text-emerald-500" />
                      <p className="text-[10px] font-bold text-emerald-650 dark:text-emerald-500 uppercase tracking-widest">Profitable Days</p>
                    </div>
                    <p className="text-base font-black text-emerald-600 dark:text-emerald-400">
                      {activeMonthStats.profitDays} {activeMonthStats.profitDays === 1 ? 'Day' : 'Days'}
                    </p>
                  </div>
                  <div className={`p-3.5 rounded-2xl border transition-colors ${
                    theme === 'dark' ? 'bg-red-950/10 border-red-900/20' : 'bg-rose-50/20 border-rose-100/50'
                  }`}>
                    <div className="flex items-center gap-1 mb-1">
                      <TrendingDown className="w-3.5 h-3.5 text-rose-500" />
                      <p className="text-[10px] font-bold text-rose-650 dark:text-rose-500 uppercase tracking-widest">Losing Days</p>
                    </div>
                    <p className="text-base font-black text-rose-600 dark:text-rose-400">
                      {activeMonthStats.lossDays} {activeMonthStats.lossDays === 1 ? 'Day' : 'Days'}
                    </p>
                  </div>
                </div>

                {/* PnL Growth Trend Chart */}
                <div className="mt-6">
                  <div className="flex justify-between items-center mb-3">
                    <div className="flex items-center gap-1.5">
                      <TrendingUp className="w-4 h-4 text-indigo-500" />
                      <div>
                        <h3 className={`text-xs font-black uppercase tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                          {!historyZoomRange ? 'Historical Performance Map' : 'Selected Range Focus'}
                        </h3>
                        <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                          PnL Growth Trend
                        </p>
                      </div>
                    </div>
                    {historyZoomRange && (
                      <button
                        onClick={() => setHistoryZoomRange(null)}
                        className="flex items-center gap-1.5 px-2.5 py-1 text-[9px] font-extrabold uppercase tracking-widest rounded bg-indigo-600 text-white hover:bg-indigo-700 transition-all shadow-md shadow-indigo-650/10 active:scale-95 cursor-pointer"
                      >
                        Reset Range
                      </button>
                    )}
                  </div>

                  <div className={`p-4 rounded-3xl border transition-colors ${
                    theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-slate-50 border-slate-100'
                  }`}>
                    {history.length === 0 ? (
                      <div className="flex flex-col items-center justify-center py-10 px-6 text-center">
                        <TrendingUp className="w-10 h-10 text-slate-400 mb-3 opacity-60" />
                        <span className={`text-xs font-bold mb-1 transition-colors ${theme === 'dark' ? 'text-slate-300' : 'text-slate-750'}`}>No Saved Trades</span>
                        <p className="text-[9px] text-slate-455 max-w-[200px] uppercase tracking-wider font-extrabold leading-relaxed">Trades logged in calendar and dashboard will build your trend graph automatically</p>
                      </div>
                    ) : (
                      <>
                        <div className="h-32 w-full relative">
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart
                              data={historyTrendData}
                              onMouseDown={(e) => {
                                if (e && e.activeTooltipIndex !== undefined) {
                                  setHistoryRefLeft(e.activeTooltipIndex);
                                }
                              }}
                              onMouseMove={(e) => {
                                if (historyRefLeft !== null && e && e.activeTooltipIndex !== undefined) {
                                  setHistoryRefRight(e.activeTooltipIndex);
                                }
                              }}
                              onMouseUp={() => {
                                if (historyRefLeft !== null && historyRefRight !== null && historyRefLeft !== historyRefRight) {
                                  const minIdx = Math.min(historyRefLeft, historyRefRight);
                                  const maxIdx = Math.max(historyRefLeft, historyRefRight);
                                  setHistoryZoomRange([minIdx, maxIdx]);
                                }
                                setHistoryRefLeft(null);
                                setHistoryRefRight(null);
                              }}
                              style={{ cursor: 'crosshair' }}
                            >
                              <defs>
                                <linearGradient id="colorCalendarHistoryGrow" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="#6366f1" stopOpacity={0.35}/>
                                  <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                                </linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={theme === 'dark' ? '#1e293b' : '#f1f5f9'} />
                              <XAxis 
                                dataKey="displayIndex" 
                                hide 
                                domain={['dataMin', 'dataMax']} 
                              />
                              <YAxis 
                                hide 
                                domain={['dataMin - 10', 'dataMax + 10']} 
                              />
                              <Tooltip
                                content={({ active, payload }) => {
                                  if (active && payload && payload.length) {
                                    const pt = payload[0].payload;
                                    return (
                                      <div className={`p-3 rounded-xl border text-[11px] font-medium shadow-xl backdrop-blur-md ${
                                        theme === 'dark' ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-800'
                                      }`}>
                                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">
                                          {pt.ticker} ({pt.date})
                                        </p>
                                        <p className="flex justify-between items-center gap-4 mb-0.5">
                                          <span className="text-slate-405 dark:text-slate-400">Trade PnL:</span>
                                          <span className={`font-mono font-bold ${pt.profit >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                            {pt.profit >= 0 ? '+' : ''}₹{pt.profit.toLocaleString('en-IN')}
                                          </span>
                                        </p>
                                        <p className="flex justify-between items-center gap-4">
                                          <span className="text-slate-405 dark:text-slate-400">Total Profit:</span>
                                          <span className={`font-mono font-bold ${pt.cumulative >= 0 ? 'text-indigo-400' : 'text-red-400'}`}>
                                            {pt.cumulative >= 0 ? '+' : ''}₹{pt.cumulative.toLocaleString('en-IN')}
                                          </span>
                                        </p>
                                      </div>
                                    );
                                  }
                                  return null;
                                }}
                              />
                              <ReferenceLine y={0} stroke={theme === 'dark' ? '#334155' : '#e2e8f0'} strokeWidth={1} strokeDasharray="3 3" />
                              <Area
                                type="monotone"
                                dataKey="cumulative"
                                stroke="#6366f1"
                                fillOpacity={1}
                                fill="url(#colorCalendarHistoryGrow)"
                                strokeWidth={3}
                                activeDot={{ r: 5 }}
                              />
                              {historyRefLeft !== null && historyRefRight !== null && (
                                <ReferenceArea
                                  {...({
                                    x1: historyTrendData[historyRefLeft]?.displayIndex,
                                    x2: historyTrendData[historyRefRight]?.displayIndex,
                                    fill: "#6366f1",
                                    fillOpacity: 0.15
                                  } as any)}
                                />
                              )}
                              {historyZoomRange && (
                                <ReferenceArea
                                  {...({
                                    x1: historyTrendData[historyZoomRange[0]]?.displayIndex,
                                    x2: historyTrendData[historyZoomRange[1]]?.displayIndex,
                                    fill: "#6366f1",
                                    fillOpacity: 0.06,
                                    stroke: "#6366f1",
                                    strokeDasharray: "4 4",
                                    strokeOpacity: 0.4
                                  } as any)}
                                />
                              )}
                            </AreaChart>
                          </ResponsiveContainer>
                        </div>
                        <p className="text-center text-[8px] font-bold tracking-wider text-slate-500 uppercase mt-2 leading-none">
                          💡 Drag horizontally across the performance points to set range focus
                        </p>
                      </>
                    )}
                  </div>
                </div>

                {/* Waterfall Chart Section under scroll option */}
                <div className="mt-6">
                  <div className="flex items-center gap-1.5 mb-3">
                    <BarChart3 className="w-4 h-4 text-indigo-500" />
                    <div>
                      <h3 className={`text-xs font-black uppercase tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Cumulative Waterfall</h3>
                      <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">Monthly Cumulative Flow</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-3xl border transition-colors ${
                    theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-slate-50 border-slate-100'
                  }`}>
                    {monthlyWaterfallData.length <= 1 ? (
                      <div className="flex flex-col items-center justify-center py-10 px-6 text-center">
                        <BarChart3 className="w-10 h-10 text-slate-400 mb-3 opacity-60" />
                        <span className={`text-xs font-bold mb-1 transition-colors ${theme === 'dark' ? 'text-slate-300' : 'text-slate-750'}`}>No Month Trades Logged</span>
                        <p className="text-[9px] text-slate-455 max-w-[200px] uppercase tracking-wider font-extrabold">Log trades inside the calendar to generate cumulative visual waterfalls</p>
                      </div>
                    ) : (
                      <div className="py-2 pr-3 h-[255px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart
                            data={monthlyWaterfallData}
                            margin={{ top: 15, right: 5, left: -20, bottom: 0 }}
                          >
                            <CartesianGrid 
                              strokeDasharray="3 3" 
                              stroke={theme === 'dark' ? '#1e293b' : '#f1f5f9'} 
                            />
                            <XAxis 
                              dataKey="dayNum" 
                              tickFormatter={(value) => value === 99 ? 'Total' : `D${value}`}
                              stroke={theme === 'dark' ? '#475569' : '#94a3b8'} 
                              fontSize={9}
                              fontWeight="bold"
                            />
                            <YAxis 
                              stroke={theme === 'dark' ? '#475569' : '#94a3b8'} 
                              fontSize={9}
                              fontWeight="bold"
                              tickFormatter={(val) => {
                                const absVal = Math.abs(val);
                                const sign = val < 0 ? '-' : '';
                                if (absVal >= 1000) return `${sign}₹${(absVal / 1000).toFixed(0)}k`;
                                return `${sign}₹${absVal}`;
                              }}
                            />
                            <Tooltip
                              content={({ active, payload }) => {
                                if (active && payload && payload.length) {
                                  const data = payload[0].payload;
                                  const isTot = data.dateStr === 'Total';
                                  return (
                                    <div className={`p-3 rounded-2xl shadow-xl border text-xs font-semibold ${
                                      theme === 'dark' ? 'bg-slate-900 border-slate-800 text-white' : 'bg-white border-slate-200 text-slate-800'
                                    }`}>
                                      <p className="text-[9px] uppercase font-black tracking-widest text-slate-450 mb-2">
                                        {isTot ? 'Month Conclusion' : `Trade Day ${data.dayNum}`}
                                      </p>
                                      <div className="space-y-1">
                                        {!isTot && (
                                          <div className="flex justify-between gap-6">
                                            <span className="text-slate-400 uppercase text-[9px] font-bold">Prev Cumulative:</span>
                                            <span className="font-mono">₹{data.lastCumulative.toLocaleString('en-IN')}</span>
                                          </div>
                                        )}
                                        <div className="flex justify-between gap-6">
                                          <span className="text-slate-400 uppercase text-[9px] font-bold">{isTot ? 'Net Profit:' : 'Day P&L:'}</span>
                                          <span className={`font-mono font-extrabold ${data.pnl >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                            {data.pnl >= 0 ? '+' : ''}₹{data.pnl.toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                        <div className="flex justify-between gap-6 border-t pt-1 mt-1 border-slate-800/25 dark:border-slate-800/60">
                                          <span className="text-slate-400 uppercase text-[9px] font-bold">{isTot ? 'Total Gain:' : 'Cumulative P&L:'}</span>
                                          <span className={`font-mono font-extrabold ${data.currentCumulative >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                            ₹{data.currentCumulative.toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                      </div>
                                    </div>
                                  );
                                }
                                return null;
                              }}
                            />
                            <Bar 
                              dataKey="range"
                              radius={[4, 4, 4, 4]}
                            >
                              {monthlyWaterfallData.map((entry, index) => {
                                let color = '#4f46e5';
                                if (entry.dateStr !== 'Total') {
                                  color = entry.pnl >= 0 ? '#10b981' : '#ef4444';
                                } else {
                                  color = entry.currentCumulative >= 0 ? '#6366f1' : '#f43f5e';
                                }
                                return (
                                  <Cell 
                                    key={`cell-${index}`} 
                                    fill={color}
                                  />
                                );
                              })}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                  </div>
                </div>

                {/* Spline Chart Section */}
                <div className="mt-6">
                  <div className="flex items-center gap-1.5 mb-3">
                    <TrendingUp className="w-4 h-4 text-indigo-500 animate-pulse" />
                    <div>
                      <h3 className={`text-xs font-black uppercase tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Cumulative Spline Curve</h3>
                      <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest font-mono">Monthly Growth Curve</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-3xl border transition-colors ${
                    theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-slate-50 border-slate-100'
                  }`}>
                    {monthlySplineData.length === 0 ? (
                      <div className="flex flex-col items-center justify-center py-10 px-6 text-center">
                        <TrendingUp className="w-10 h-10 text-slate-400 mb-3 opacity-60" />
                        <span className={`text-xs font-bold mb-1 transition-colors ${theme === 'dark' ? 'text-slate-300' : 'text-slate-750'}`}>No Month Trades Logged</span>
                        <p className="text-[9px] text-slate-455 max-w-[200px] uppercase tracking-wider font-extrabold">Log trades inside the calendar to generate cumulative visual splines</p>
                      </div>
                    ) : (
                      <div className="py-2 pr-4 h-[220px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart
                            data={monthlySplineData}
                            margin={{ top: 15, right: 5, left: -20, bottom: 0 }}
                          >
                            <defs>
                              <linearGradient id="colorMonthlySplineGradient" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#6366f1" stopOpacity={0.35}/>
                                <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid 
                              strokeDasharray="3 3" 
                              stroke={theme === 'dark' ? '#1e293b' : '#f1f5f9'}
                              vertical={false}
                            />
                            <XAxis 
                              dataKey="dayNum" 
                              tickFormatter={(value) => `D${value}`}
                              stroke={theme === 'dark' ? '#475569' : '#94a3b8'} 
                              fontSize={9}
                              fontWeight="bold"
                            />
                            <YAxis 
                              stroke={theme === 'dark' ? '#475569' : '#94a3b8'} 
                              fontSize={9}
                              fontWeight="bold"
                              tickFormatter={(val) => {
                                const absVal = Math.abs(val);
                                const sign = val < 0 ? '-' : '';
                                if (absVal >= 1000) return `${sign}₹${(absVal / 1000).toFixed(0)}k`;
                                return `${sign}₹${absVal}`;
                              }}
                            />
                            <Tooltip
                              content={({ active, payload }) => {
                                if (active && payload && payload.length) {
                                  const data = payload[0].payload;
                                  return (
                                    <div className={`p-3 rounded-2xl shadow-xl border text-xs font-semibold ${
                                      theme === 'dark' ? 'bg-slate-900 border-slate-800 text-white' : 'bg-white border-slate-200 text-slate-800'
                                    }`}>
                                      <p className="text-[9px] uppercase font-black tracking-widest text-[#6366f1] mb-2 font-mono">
                                        Day {data.dayNum} Performance
                                      </p>
                                      <div className="space-y-1">
                                        <div className="flex justify-between gap-6">
                                          <span className="text-slate-400 uppercase text-[9px] font-bold">Day Change:</span>
                                          <span className={`font-mono font-extrabold ${data.pnl >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                            {data.pnl >= 0 ? '+' : ''}₹{data.pnl.toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                        <div className="flex justify-between gap-6 border-t pt-1 mt-1 border-slate-800/25 dark:border-slate-800/60 font-mono">
                                          <span className="text-slate-400 uppercase text-[9px] font-bold">Month Cumulative:</span>
                                          <span className={`font-extrabold ${data.currentCumulative >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                            ₹{data.currentCumulative.toLocaleString('en-IN')}
                                          </span>
                                        </div>
                                      </div>
                                    </div>
                                  );
                                }
                                return null;
                              }}
                            />
                            <ReferenceLine y={0} stroke={theme === 'dark' ? '#334155' : '#e2e8f0'} strokeWidth={1} strokeDasharray="3 3" />
                            <Area
                              type="monotone"
                              dataKey="currentCumulative"
                              stroke="#6366f1"
                              strokeWidth={2.5}
                              fillOpacity={1}
                              fill="url(#colorMonthlySplineGradient)"
                              activeDot={{ r: 5 }}
                            />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                  </div>
                </div>
              </div>


            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* History Modal Overlay */}
        <AnimatePresence>
          {showHistory && (
            <div className="fixed inset-0 z-50 flex items-end justify-center sm:items-center p-0 sm:p-4">
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
                onClick={() => setShowHistory(false)}
              />
              <motion.div
                initial={{ y: '100%' }}
                animate={{ y: 0 }}
                exit={{ y: '100%' }}
                className={`relative w-full max-w-md rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[85vh] flex flex-col transition-colors duration-300 ${
                  theme === 'dark' ? 'bg-slate-900' : 'bg-white'
                }`}
                id="history-panel"
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-4 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />
                <div className={`px-6 pb-4 border-b flex items-center gap-3 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  <button
                    onClick={() => setShowHistory(false)}
                    className={`p-2 rounded-xl border transition-all active:scale-95 ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                    aria-label="Back to Trading"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <div className="flex flex-col flex-1">
                    <h2 className={`text-xl font-bold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Trade History</h2>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Historical Logs</p>
                  </div>
                  <div className="flex items-center gap-2">
                    {history.length > 0 && (
                      <button 
                        onClick={() => setShowOnlyFavorites(!showOnlyFavorites)}
                        className={`p-2 rounded-lg border transition-all ${
                          showOnlyFavorites 
                            ? 'bg-rose-50 border-rose-200 text-rose-500 shadow-sm' 
                            : theme === 'dark' ? 'bg-slate-800 border-slate-700 text-slate-500 hover:border-slate-600' : 'bg-white border-slate-200 text-slate-400 hover:border-slate-300'
                        }`}
                        title={showOnlyFavorites ? "Show All Trades" : "Show Favorites Only"}
                      >
                        <Heart className={`w-4 h-4 ${showOnlyFavorites ? 'fill-rose-500' : ''}`} />
                      </button>
                    )}
                    <div className="flex gap-2 items-center">
                      {history.length > 0 && (
                      <>
                        <div className="relative">
                          <button
                            onClick={() => setIsExportMenuOpen(!isExportMenuOpen)}
                            className={`text-xs font-bold px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1.5 border cursor-pointer ${
                              theme === 'dark' 
                                ? 'text-indigo-400 bg-slate-800 border-indigo-900/30 hover:bg-slate-700/60' 
                                : 'text-indigo-600 bg-white border-indigo-100 hover:bg-indigo-50/60'
                            }`}
                            title="Export options (CSV, Excel, PDF)"
                          >
                            <Download className="w-3.5 h-3.5" />
                            <span>Export Data</span>
                            <ChevronDown className="w-3 h-3 opacity-60" />
                          </button>
                          
                          {isExportMenuOpen && (
                            <>
                              <div 
                                className="fixed inset-0 z-40" 
                                onClick={() => setIsExportMenuOpen(false)}
                              />
                              <div className={`absolute right-0 mt-1.5 w-56 rounded-xl border shadow-xl z-50 p-2 py-2 flex flex-col space-y-1 transition-all animate-in fade-in slide-in-from-top-1 ${
                                theme === 'dark' 
                                  ? 'bg-slate-900 border-slate-800 text-slate-200' 
                                  : 'bg-white border-slate-200/80 text-slate-700'
                              }`}>
                                <div className="px-2.5 py-1 text-[10px] font-black uppercase tracking-widest text-slate-400">
                                  Full History ({history.length})
                                </div>
                                <button
                                  onClick={() => {
                                    exportHistoryExcel(false);
                                    setIsExportMenuOpen(false);
                                  }}
                                  className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                    theme === 'dark' ? 'hover:bg-slate-800 text-white' : 'hover:bg-slate-50 text-slate-700'
                                  }`}
                                >
                                  <TableIcon className="w-3.5 h-3.5 text-emerald-500" />
                                  Excel Spreadsheet (.xlsx)
                                </button>
                                <button
                                  onClick={() => {
                                    exportHistoryCSV(false);
                                    setIsExportMenuOpen(false);
                                  }}
                                  className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                    theme === 'dark' ? 'hover:bg-slate-800 text-white' : 'hover:bg-slate-50 text-slate-700'
                                  }`}
                                >
                                  <Download className="w-3.5 h-3.5 text-indigo-500" />
                                  CSV Data Format (.csv)
                                </button>
                                <button
                                  onClick={() => {
                                    exportHistoryPDF(false);
                                    setIsExportMenuOpen(false);
                                  }}
                                  className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                    theme === 'dark' ? 'hover:bg-slate-800 text-white' : 'hover:bg-slate-50 text-slate-700'
                                  }`}
                                >
                                  <FileText className="w-3.5 h-3.5 text-red-500" />
                                  PDF Report Document (.pdf)
                                </button>

                                {historyZoomRange && (
                                  <>
                                    <div className={`h-px my-1 ${theme === 'dark' ? 'bg-slate-800' : 'bg-slate-100'}`} />
                                    <div className="px-2.5 py-1 text-[10px] font-black uppercase tracking-widest text-indigo-500 flex items-center justify-between">
                                      <span>Selected Range</span>
                                      <span className="font-mono bg-indigo-500/10 text-indigo-400 px-1.5 py-0.5 rounded text-[9px]">
                                        {filteredHistory.length} Trades
                                      </span>
                                    </div>
                                    <button
                                      onClick={() => {
                                        exportHistoryExcel(true);
                                        setIsExportMenuOpen(false);
                                      }}
                                      className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                        theme === 'dark' ? 'hover:bg-slate-800 text-indigo-300' : 'hover:bg-indigo-50/50 text-indigo-700'
                                      }`}
                                    >
                                      <TableIcon className="w-3.5 h-3.5 text-emerald-500" />
                                      Excel (Selected Range)
                                    </button>
                                    <button
                                      onClick={() => {
                                        exportHistoryCSV(true);
                                        setIsExportMenuOpen(false);
                                      }}
                                      className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                        theme === 'dark' ? 'hover:bg-slate-800 text-indigo-300' : 'hover:bg-indigo-50/50 text-indigo-700'
                                      }`}
                                    >
                                      <Download className="w-3.5 h-3.5 text-indigo-400" />
                                      CSV (Selected Range)
                                    </button>
                                    <button
                                      onClick={() => {
                                        exportHistoryPDF(true);
                                        setIsExportMenuOpen(false);
                                      }}
                                      className={`text-left text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-2 cursor-pointer transition-colors ${
                                        theme === 'dark' ? 'hover:bg-slate-800 text-indigo-300' : 'hover:bg-indigo-50/50 text-indigo-700'
                                      }`}
                                    >
                                      <FileText className="w-3.5 h-3.5 text-red-500" />
                                      PDF (Selected Range)
                                    </button>
                                  </>
                                )}
                              </div>
                            </>
                          )}
                        </div>
                        <button 
                          onClick={clearHistory}
                          className="text-xs font-bold text-red-500 hover:bg-red-50/60 dark:hover:bg-red-950/20 px-3 py-1.5 rounded-lg transition-colors cursor-pointer"
                        >
                          Clear
                        </button>
                      </>
                    )}
                  </div>
                </div>
              </div>
              
              <div className="overflow-y-auto px-4 py-6 space-y-3 flex-1 pb-10">
                  {history.length === 0 ? (
                    <div className="text-center py-20 text-slate-400 flex flex-col items-center">
                      <History className="w-12 h-12 mb-4 opacity-10" />
                      <p>{t('noSavedTradesYet')}</p>
                    </div>
                  ) : (
                    <>
                      {/* Interactive Metrics Dashboard updates dynamically within selected range */}
                      <div className="grid grid-cols-2 gap-3 mb-4">
                        <div className={`p-3 rounded-2xl border flex flex-col justify-between transition-colors ${
                          theme === 'dark' ? 'bg-slate-800/40 border-slate-700/60' : 'bg-slate-50 border-slate-100'
                        }`}>
                          <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-1">
                            {t('netProfitLoss')}
                          </span>
                          <span className={`text-base font-black font-mono leading-none ${
                            filteredStats.netPnL >= 0 ? 'text-green-500' : 'text-red-500'
                          }`}>
                            {filteredStats.netPnL >= 0 ? '+' : ''}₹{filteredStats.netPnL.toLocaleString('en-IN')}
                          </span>
                        </div>

                        <div className={`p-3 rounded-2xl border flex flex-col justify-between transition-colors ${
                          theme === 'dark' ? 'bg-slate-800/40 border-slate-700/60' : 'bg-slate-50 border-slate-100'
                        }`}>
                          <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-1">
                            {t('winRateVol')}
                          </span>
                          <span className={`text-[13px] font-extrabold uppercase leading-none ${
                            theme === 'dark' ? 'text-white' : 'text-slate-800'
                          }`}>
                            {filteredStats.winRate.toFixed(1)}% <span className="text-[10px] text-slate-400 font-mono font-bold">({filteredStats.totalTrades})</span>
                          </span>
                        </div>
                      </div>

                      {filteredHistory
                        .filter(t => showOnlyFavorites ? t.isFavorite : true)
                        .map((trade) => (
                        <div 
                          key={trade.id} 
                          className={`border p-4 rounded-2xl flex items-center gap-4 group relative transition-colors ${
                            theme === 'dark' ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'
                          }`}
                        >
                          <div className={`p-2 rounded-xl shrink-0 transition-colors ${
                            trade.profit >= 0 
                              ? theme === 'dark' ? 'bg-green-900/30 text-green-400' : 'bg-green-100 text-green-600' 
                              : theme === 'dark' ? 'bg-red-900/30 text-red-400' : 'bg-red-100 text-red-600'
                          }`}>
                            {trade.type === 'long' ? <TrendingUp className="w-5 h-5" /> : <TrendingDown className="w-5 h-5" />}
                          </div>
                          <div className="flex-1">
                            <div className="flex justify-between items-start mb-0.5">
                              <div className="flex flex-col">
                                <div className="flex items-center gap-2 mb-1">
                                  <span className="text-[10px] font-black uppercase tracking-wider text-indigo-550 leading-none">
                                    {trade.ticker ? `${trade.ticker} [${trade.mode}]` : trade.mode} • {trade.type === 'long' ? (trade.mode === 'equity' ? 'Buy' : 'Call') : (trade.mode === 'equity' ? 'Sell' : 'Put')}
                                  </span>
                                  {trade.strategy && (
                                    <span className={`text-[9px] font-bold uppercase px-1.5 py-0.5 rounded leading-none transition-colors ${
                                      theme === 'dark' ? 'bg-slate-700 text-slate-300' : 'bg-slate-200 text-slate-600'
                                    }`}>
                                      {trade.strategy}
                                    </span>
                                  )}
                                </div>
                                <span className={`text-[10px] font-bold transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>
                                  {trade.entryDate ? `${trade.entryDate} to ${trade.exitDate}` : new Date(trade.timestamp).toLocaleDateString()}
                                </span>
                              </div>
                              <span className={`font-bold ${trade.profit >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                                {trade.profit >= 0 ? '+' : ''}₹{trade.profit.toLocaleString('en-IN')}
                              </span>
                            </div>
                            <div className={`text-sm font-medium transition-colors ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                              {trade.quantity} {trade.mode === 'equity' ? 'units' : `lots (x${trade.lotSize})`} @ ₹{trade.entry.toLocaleString('en-IN')}
                            </div>
                            {trade.notes && (
                              <p className={`text-[10px] mt-2 italic border-l-2 pl-2 transition-colors ${
                                theme === 'dark' ? 'text-slate-500 border-slate-700' : 'text-slate-400 border-slate-200'
                              }`}>
                                "{trade.notes}"
                              </p>
                            )}

                            {/* Trade Journal Fields Previews and Trigger */}
                            <div className="mt-2.5 flex flex-wrap gap-2 items-center">
                              <button
                                onClick={() => setSelectedJournalTrade(trade)}
                                className={`text-[9.5px] font-extrabold tracking-wide uppercase px-2.5 py-1 rounded-lg border-2 flex items-center gap-1.5 transition-all ${
                                  theme === 'dark' 
                                    ? 'text-indigo-400 bg-indigo-950/20 border-indigo-900/30 hover:bg-indigo-900/30' 
                                    : 'text-indigo-600 bg-indigo-50 border-indigo-100 hover:bg-indigo-100/50'
                                }`}
                              >
                                <BookOpen className="w-3.5 h-3.5" />
                                { (trade.tradeSetup || trade.entryExitRationale || trade.lessonsLearned) ? 'View/Update Journal' : 'Write Journal' }
                              </button>
                              {trade.isManual && (
                                <span className={`text-[9px] font-bold uppercase tracking-wider px-2 py-0.5 rounded transition-colors ${
                                  theme === 'dark' ? 'bg-amber-950/40 text-amber-500 border border-amber-900/40' : 'bg-amber-50 text-amber-600 border border-amber-100'
                                }`}>
                                  Manual Entry
                                </span>
                              )}
                              {(trade.tradeSetup || trade.entryExitRationale || trade.lessonsLearned) && (
                                <span className={`text-[9px] font-medium uppercase tracking-wider px-2 py-0.5 rounded transition-colors ${
                                  theme === 'dark' ? 'bg-green-950/40 text-green-400 border border-green-900/40' : 'bg-green-50 text-green-600 border border-green-100'
                                }`}>
                                  Journal On
                                </span>
                              )}
                            </div>
                          </div>
                          <div className="flex flex-col gap-1">
                            <button
                              onClick={() => toggleFavorite(trade.id)}
                              className={`p-2 rounded-full transition-all ${
                                trade.isFavorite 
                                  ? 'bg-rose-50 text-rose-500' 
                                  : 'text-slate-300 hover:text-rose-400 opacity-0 group-hover:opacity-100'
                              }`}
                            >
                              <Heart className={`w-4 h-4 ${trade.isFavorite ? 'fill-rose-500' : ''}`} />
                            </button>
                            <button 
                              onClick={() => deleteTrade(trade.id)}
                              className="p-2 text-slate-300 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"
                              id={`delete-trade-${trade.id}`}
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      ))}
                      {history.filter(t => showOnlyFavorites ? t.isFavorite : true).length === 0 && (
                        <div className="text-center py-20 text-slate-400 flex flex-col items-center">
                          <Heart className="w-12 h-12 mb-4 opacity-10" />
                          <p>No favorite trades found</p>
                        </div>
                      )}
                    </>
                  )}
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Funds Panel */}
        <AnimatePresence>
          {showFunds && (
            <div 
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setShowFunds(false);
                }
              }}
              className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 bg-black/50 backdrop-blur-sm"
            >
              <motion.div 
                initial={{ y: '100%' }}
                animate={{ y: 0 }}
                exit={{ y: '100%' }}
                className={`relative w-full max-w-lg rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[85vh] flex flex-col transition-colors duration-300 ${
                  theme === 'dark' ? 'bg-slate-900' : 'bg-white'
                }`}
                id="funds-panel"
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-4 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />
                
                <div className={`px-6 pb-4 border-b flex items-center gap-3 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  <button
                    onClick={() => setShowFunds(false)}
                    className={`p-2 rounded-xl border transition-all active:scale-95 ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                    aria-label="Back to Trading"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <div className="flex flex-col flex-1">
                    <h2 className={`text-xl font-bold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Fund Management</h2>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Adjust Wallet Balance Ledger</p>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-6">
                  {/* Balances Display Card */}
                  <div className={`p-5 rounded-[2rem] border transition-all ${
                    theme === 'dark' ? 'bg-slate-800/40 border-slate-700/80' : 'bg-slate-50/70 border-slate-200/60'
                  }`}>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="flex flex-col items-center text-center">
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Opening Balance</span>
                        <h4 className={`text-lg font-black transition-colors ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                          ₹{openingBalance.toLocaleString('en-IN')}
                        </h4>
                        <button 
                          onClick={() => {
                            setOpeningBalance(totalFunds);
                          }}
                          className="mt-2 text-[8px] font-black uppercase bg-indigo-600/10 text-indigo-505 px-3 py-1.5 rounded-full border border-indigo-500/20 active:scale-95 transition-all cursor-pointer"
                        >
                          Set Today
                        </button>
                      </div>
                      <div className="flex flex-col items-center text-center border-l border-slate-200/50 dark:border-slate-800 pl-4">
                        <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-wider mb-1">Current Balance</span>
                        <h4 className={`text-xl font-black transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                          ₹{totalFunds.toLocaleString('en-IN')}
                        </h4>
                        <div className={`mt-2 text-[8px] font-black uppercase px-3 py-1 rounded-full ${
                          totalFunds - openingBalance >= 0 ? 'bg-green-500/10 text-green-500' : 'bg-rose-500/10 text-rose-500'
                        }`}>
                          {totalFunds - openingBalance >= 0 ? '+' : ''}₹{Math.abs(totalFunds - openingBalance).toLocaleString('en-IN')}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Fund Transaction Entry Form */}
                  <div className={`p-5 rounded-2xl border transition-all ${
                    editingFundTxId 
                      ? 'border-indigo-500 ring-2 ring-indigo-500/15 bg-indigo-500/5' 
                      : theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-white border-slate-200/50'
                  }`}>
                    <div className="flex items-center justify-between mb-3">
                      <span className="text-[10px] font-black text-indigo-550 uppercase tracking-widest leading-none">
                        {editingFundTxId ? '✏️ Edit Capital Entry' : '⚡ Fund Transaction Entry'}
                      </span>
                      {editingFundTxId && (
                        <span className="text-[8px] font-bold text-slate-400 uppercase tracking-wider">
                          Editing Mode
                        </span>
                      )}
                    </div>

                    <div className="space-y-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                        {/* Amount Input */}
                        <div>
                          <label className="text-[10px] font-bold text-slate-450 dark:text-slate-500 uppercase tracking-wider block mb-1.5">Amount (₹)</label>
                          <div className="relative">
                            <IndianRupee className="absolute left-3.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                            <input 
                              type="number"
                              min="0"
                              value={fundInput}
                              onChange={(e) => setFundInput(e.target.value)}
                              placeholder="Amount to Adjust"
                              className={`w-full py-2 pl-9 pr-3 rounded-xl border outline-none transition-all text-sm font-bold font-mono focus:ring-1 focus:ring-indigo-500 ${
                                theme === 'dark' ? 'bg-slate-900 border-slate-700 text-white' : 'bg-slate-50 border-slate-200 text-slate-900'
                              }`}
                            />
                          </div>
                        </div>

                        {/* Note/Label Input */}
                        <div>
                          <label className="text-[10px] font-bold text-slate-455 dark:text-slate-500 uppercase tracking-wider block mb-1.5">Note / Description</label>
                          <input 
                            type="text"
                            value={fundNotesInput}
                            onChange={(e) => setFundNotesInput(e.target.value)}
                            placeholder="e.g. Deposit, Salary, Bonus payout"
                            className={`w-full py-2 px-3 rounded-xl border outline-none transition-all text-sm focus:ring-1 focus:ring-indigo-500 ${
                              theme === 'dark' ? 'bg-slate-900 border-slate-700 text-white placeholder:text-slate-650' : 'bg-slate-50 border-slate-200 text-slate-900 placeholder:text-slate-350'
                            }`}
                          />
                        </div>
                      </div>

                      {/* Action buttons */}
                      <div className="flex flex-col gap-2">
                        <div className="grid grid-cols-2 gap-3">
                          <button 
                            type="button"
                            onClick={() => handleAddFundTransaction('deposit')}
                            disabled={!fundInput || parseFloat(fundInput) <= 0}
                            className={`py-3 rounded-xl font-black text-[10px] uppercase tracking-widest flex items-center justify-center gap-2 active:scale-[0.98] transition-all shadow-md cursor-pointer ${
                              !fundInput || parseFloat(fundInput) <= 0
                                ? 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed opacity-40'
                                : 'bg-green-600 hover:bg-green-700 text-white shadow-green-900/10'
                            }`}
                          >
                            <Plus className="w-4 h-4 shrink-0" />
                            {editingFundTxId ? 'Update Deposit' : 'Deposit'}
                          </button>
                          <button 
                            type="button"
                            onClick={() => handleAddFundTransaction('withdrawal')}
                            disabled={!fundInput || parseFloat(fundInput) <= 0}
                            className={`py-3 rounded-xl font-black text-[10px] uppercase tracking-widest flex items-center justify-center gap-2 active:scale-[0.98] transition-all shadow-md cursor-pointer ${
                              !fundInput || parseFloat(fundInput) <= 0
                                ? 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed opacity-40'
                                : 'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-900/10'
                            }`}
                          >
                            <ArrowDownLeft className="w-4 h-4 shrink-0" />
                            {editingFundTxId ? 'Update Withdraw' : 'Withdraw'}
                          </button>
                        </div>

                        {editingFundTxId && (
                          <button
                            type="button"
                            onClick={handleCancelEditFundTransaction}
                            className={`w-full py-1.5 rounded-lg font-black text-[9px] uppercase tracking-widest transition-all hover:underline ${
                              theme === 'dark' ? 'text-slate-400 hover:text-white' : 'text-slate-500 hover:text-slate-850'
                            }`}
                          >
                            Cancel Edit
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Fund Transaction ledger items */}
                  <div className="space-y-4">
                    {/* Deposit & Withdrawal Histories Summary Row */}
                    <div className="grid grid-cols-2 gap-3">
                      <div className={`p-3 rounded-xl border flex flex-col transition-all ${
                        theme === 'dark' ? 'bg-green-950/10 border-green-900/20' : 'bg-emerald-50/20 border-emerald-100/50'
                      }`}>
                        <span className="text-[9px] font-black uppercase tracking-widest text-emerald-600 dark:text-emerald-450">Total Deposited</span>
                        <span className="text-sm font-black font-mono text-emerald-500 mt-1">
                          +₹{totalDeposits.toLocaleString('en-IN')}
                        </span>
                      </div>
                      <div className={`p-3 rounded-xl border flex flex-col transition-all ${
                        theme === 'dark' ? 'bg-red-950/10 border-red-900/20' : 'bg-rose-50/20 border-rose-100/50'
                      }`}>
                        <span className="text-[9px] font-black uppercase tracking-widest text-rose-600 dark:text-rose-450">Total Withdrawn</span>
                        <span className="text-sm font-black font-mono text-rose-500 mt-1">
                          -₹{totalWithdrawals.toLocaleString('en-IN')}
                        </span>
                      </div>
                    </div>

                    <div className="flex items-center justify-between border-b pb-2 border-slate-200/50 dark:border-slate-800/60">
                      <h3 className={`text-xs font-black uppercase tracking-widest ${
                        theme === 'dark' ? 'text-slate-400' : 'text-slate-505'
                      }`}>
                        Capital History Logs
                      </h3>
                      
                      {/* Filter tabs */}
                      <div className={`flex items-center gap-1 p-0.5 rounded-lg border ${
                        theme === 'dark' ? 'bg-slate-950/60 border-slate-800' : 'bg-slate-100/70 border-slate-200/50'
                      }`}>
                        {(['all', 'deposit', 'withdrawal'] as const).map((tab) => (
                          <button
                            key={tab}
                            type="button"
                            onClick={() => setFundTab(tab)}
                            className={`px-2 py-1 rounded-md text-[9px] font-extrabold uppercase tracking-wider transition-all cursor-pointer ${
                              fundTab === tab
                                ? (theme === 'dark' ? 'bg-indigo-650 text-white shadow-sm' : 'bg-white text-indigo-600 shadow-sm')
                                : (theme === 'dark' ? 'text-slate-500 hover:text-slate-300' : 'text-slate-500 hover:text-slate-700')
                            }`}
                          >
                            {tab === 'all' ? 'All' : tab === 'deposit' ? 'Deposits' : 'Withdrawals'}
                          </button>
                        ))}
                      </div>
                    </div>
                    
                    {fundTransactions.length === 0 ? (
                      <div className={`p-6 rounded-2xl text-center border border-dashed transition-all ${
                        theme === 'dark' ? 'border-slate-850 bg-slate-900/10 text-slate-500' : 'border-slate-200 bg-slate-50/50 text-slate-400'
                      }`}>
                        <IndianRupee className="w-7 h-7 mx-auto mb-2 opacity-30 text-indigo-500" />
                        <p className="text-[11px] font-bold tracking-tight">No manual transactions logged yet.</p>
                        <p className="text-[9px] opacity-85 mt-0.5">Deposits and withdrawals will appear here as adjustable items.</p>
                      </div>
                    ) : fundTransactions.filter(tx => fundTab === 'all' || tx.type === fundTab).length === 0 ? (
                      <div className={`p-6 rounded-2xl text-center border border-dashed transition-all ${
                        theme === 'dark' ? 'border-slate-850 bg-slate-900/10 text-slate-500' : 'border-slate-200 bg-slate-50/50 text-slate-400'
                      }`}>
                        <IndianRupee className="w-7 h-7 mx-auto mb-2 opacity-30 text-indigo-550" />
                        <p className="text-[11px] font-bold tracking-tight">No transactions match your selection</p>
                        <p className="text-[9px] opacity-85 mt-0.5">There are no {fundTab === 'deposit' ? 'deposit' : 'withdrawal'} transactions recorded.</p>
                      </div>
                    ) : (
                      <div className="space-y-2 max-h-56 overflow-y-auto pr-1">
                        {fundTransactions
                          .filter(tx => fundTab === 'all' || tx.type === fundTab)
                          .map(tx => (
                            <div 
                              key={tx.id}
                              className={`flex justify-between items-center p-3 rounded-xl text-xs border transition-all ${
                                editingFundTxId === tx.id
                                  ? 'border-indigo-500 ring-1 ring-indigo-500 bg-indigo-500/5'
                                  : theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80 text-white hover:bg-slate-950/60' : 'bg-white border-slate-100 hover:bg-slate-50/50'
                              }`}
                            >
                              <div className="flex items-center gap-2.5">
                                <span className={`w-2 h-2 rounded-full shrink-0 ${
                                  tx.type === 'deposit' ? 'bg-green-500' : 'bg-rose-500'
                                }`} />
                                <div className="flex flex-col text-left">
                                  <span className="font-extrabold text-[12px]">{tx.notes || (tx.type === 'deposit' ? 'Capital Added' : 'Capital Withdrawn')}</span>
                                  <span className="text-[8px] text-slate-400 dark:text-slate-505 font-medium mt-0.5 font-mono">
                                    {new Date(tx.timestamp).toLocaleString('en-IN', {
                                      day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
                                    })}
                                  </span>
                                </div>
                              </div>

                              <div className="flex items-center gap-3 shrink-0">
                                <span className={`font-mono font-black text-sm ${
                                  tx.type === 'deposit' ? 'text-green-500' : 'text-rose-500'
                                }`}>
                                  {tx.type === 'deposit' ? '+' : '-'}₹{tx.amount.toLocaleString('en-IN')}
                                </span>

                                <div className="flex items-center gap-1.5">
                                  <button
                                    type="button"
                                    onClick={() => handleStartEditFundTransaction(tx)}
                                    className={`p-1.5 rounded-md transition-colors cursor-pointer ${
                                      theme === 'dark' ? 'hover:bg-slate-850 text-slate-400 hover:text-indigo-400' : 'hover:bg-slate-100 text-slate-505 hover:text-indigo-650'
                                    }`}
                                    title="Edit this log"
                                  >
                                    <Edit3 className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteFundTransaction(tx.id)}
                                    className={`p-1.5 rounded-md transition-colors cursor-pointer ${
                                      theme === 'dark' ? 'hover:bg-slate-850 text-slate-400 hover:text-rose-400' : 'hover:bg-slate-100 text-slate-505 hover:text-rose-650'
                                    }`}
                                    title="Delete this log"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            </div>
                          ))}
                      </div>
                    )}
                  </div>
                </div>


              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Settings Panel */}
        <AnimatePresence>
          {showSettings && (
            <div 
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setShowSettings(false);
                }
              }}
              className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 bg-black/50 backdrop-blur-sm"
            >
              <motion.div 
                initial={{ y: '100%' }}
                animate={{ y: 0 }}
                exit={{ y: '100%' }}
                className={`relative w-full max-w-md rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[85vh] flex flex-col transition-colors duration-300 ${
                  theme === 'dark' ? 'bg-slate-900' : 'bg-white'
                }`}
                id="settings-panel"
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-4 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />
                
                <div className={`px-6 pb-4 border-b flex items-center gap-3 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  {supportStep === 'menu' ? (
                    <button
                      onClick={() => setShowSettings(false)}
                      className={`p-2 rounded-xl border transition-all active:scale-95 ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                      }`}
                      aria-label="Back to Trading"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                  ) : (
                    <button 
                      onClick={() => {
                        if (supportStep === 'info') setSupportStep('menu');
                        else if (supportStep === 'pay') setSupportStep('info');
                        else if (supportStep === 'success') {
                          setSupportStep('menu');
                        }
                      }}
                      className={`p-2 rounded-xl border transition-all active:scale-95 ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                      }`}
                      title="Back to Settings"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                  )}
                  <div className="flex flex-col flex-1">
                    <h2 className={`text-xl font-bold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                      {supportStep === 'menu' && t('appSettings')}
                      {supportStep === 'info' && "Founder Profile"}
                      {supportStep === 'pay' && "Support Founder"}
                      {supportStep === 'success' && "Thank You! ❤️"}
                    </h2>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                      {supportStep === 'menu' && t('configSecurity')}
                      {supportStep === 'info' && "Step 1 of 2: Your Details"}
                      {supportStep === 'pay' && "Step 2 of 2: Payment"}
                      {supportStep === 'success' && "Contribution Complete"}
                    </p>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-8">
                  {supportStep === 'menu' ? (
                    <>
                      {/* Account / Profile Section */}
                      <div className="space-y-4">
                        <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>{t('userProfile')}</h3>
                        <div className={`p-4 rounded-2xl border flex flex-col items-center transition-all ${
                          theme === 'dark' ? 'bg-slate-800/30 border-slate-800' : 'bg-slate-50 border-slate-100'
                        }`}>
                          <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-2 relative transition-colors ${theme === 'dark' ? 'bg-slate-800' : 'bg-indigo-100'}`}>
                            {authUser?.photoURL ? (
                              <img src={authUser.photoURL} alt="Profile" className="w-full h-full rounded-full object-cover" referrerPolicy="no-referrer" />
                            ) : (
                              <User className={`w-6 h-6 transition-colors ${theme === 'dark' ? 'text-indigo-400' : 'text-indigo-600'}`} />
                            )}
                            {authUser && <div className={`absolute bottom-0 right-0 w-3.5 h-3.5 bg-green-500 border-2 rounded-full transition-colors ${theme === 'dark' ? 'border-slate-800' : 'border-white'}`}></div>}
                          </div>
                          
                          <div className="w-full space-y-1">
                            <input
                              type="text"
                              value={userName}
                              onChange={(e) => setUserName(e.target.value)}
                              className={`w-full text-base font-bold text-center bg-transparent border-b border-transparent focus:border-indigo-500 outline-none px-2 transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}
                              placeholder="Your Name"
                            />
                            <p className="text-[9px] text-center text-slate-400 font-bold uppercase tracking-widest">{t('verifiedTrader')}</p>
                          </div>

                          {/* AI Chat Box Button - placed near the profile picture/name icon */}
                          <button
                            type="button"
                            onClick={() => setShowAIChat(true)}
                            className={`w-full mt-3 flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl font-bold text-xs uppercase tracking-wider transition-all cursor-pointer active:scale-95 shadow-md ${
                              theme === 'dark'
                                ? 'bg-indigo-600/30 hover:bg-indigo-600/40 text-indigo-300 border border-indigo-500/20'
                                : 'bg-indigo-55 bg-indigo-50 hover:bg-indigo-100 text-indigo-600'
                            }`}
                          >
                            <Sparkles className="w-3.5 h-3.5 animate-pulse shrink-0" />
                            <span>{t('askJournalAi')}</span>
                          </button>

                          <div className="w-full mt-3">
                            {!authUser ? (
                              <button
                                onClick={handleLogin}
                                disabled={isLoggingIn}
                                className={`w-full flex items-center justify-center gap-2 border py-2 rounded-xl transition-all active:scale-[0.98] ${
                                  theme === 'dark' ? 'bg-slate-700 border-slate-600 hover:bg-slate-600 text-white' : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                                }`}
                              >
                                {isLoggingIn ? (
                                  <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-600" />
                                ) : (
                                  <LogIn className="w-3.5 h-3.5 text-indigo-600" />
                                )}
                                <span className="text-xs font-bold">{t('signInGoogle')}</span>
                              </button>
                            ) : (
                              <div className="space-y-2">
                                 <div className="flex items-center justify-between px-1">
                                    <span className={`text-[10px] font-bold truncate transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-500'}`}>{authUser.email}</span>
                                    <button onClick={handleLogout} className="text-[10px] font-bold text-rose-500 hover:underline shrink-0">Sign Out</button>
                                 </div>
                                 <div className={`h-px w-full transition-colors ${theme === 'dark' ? 'bg-slate-800' : 'bg-slate-100'}`} />
                              </div>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* App & AI Language Section */}
                      <div className="space-y-4">
                        <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>
                          {t('responseLanguage')}
                        </h3>
                        <div className={`p-4 rounded-2xl border transition-all ${
                          theme === 'dark' ? 'bg-slate-800/30 border-slate-800' : 'bg-slate-50 border-slate-100'
                        }`}>
                          <div className="flex items-center justify-between gap-4">
                            <div className="flex items-center gap-2.5">
                              <div className={`p-2 rounded-xl shrink-0 ${theme === 'dark' ? 'bg-indigo-900/30 text-indigo-400' : 'bg-indigo-100 text-indigo-650'}`}>
                                <Globe className="w-4 h-4" />
                              </div>
                              <div className="flex flex-col text-left">
                                <span className={`text-xs font-extrabold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                                  App & AI Language
                                </span>
                                <span className="text-[9px] text-slate-400 font-bold uppercase tracking-wider font-mono">
                                  Set app interface & chatbot response
                                </span>
                              </div>
                            </div>

                            <select
                              value={aiLanguage}
                              onChange={(e) => {
                                const newLang = e.target.value;
                                setAiLanguage(newLang);
                                const selectedLang = INDIAN_LANGUAGES.find(l => l.code === newLang);
                                const langLabel = selectedLang ? `${selectedLang.name} (${selectedLang.native})` : newLang;
                                setAiMessages(prev => [
                                  ...prev,
                                  { role: 'assistant', content: `🌐 App and AI response language changed to ${langLabel}.` }
                                ]);
                              }}
                              className={`text-xs font-bold py-2 px-3 rounded-xl border outline-none cursor-pointer transition-all max-w-[150px] truncate ${
                                theme === 'dark' 
                                  ? 'bg-slate-900 border-slate-700 text-slate-100 focus:border-indigo-505' 
                                  : 'bg-white border-slate-205 text-slate-800 focus:border-indigo-600'
                              }`}
                            >
                              {INDIAN_LANGUAGES.map(lang => (
                                <option key={lang.code} value={lang.code}>
                                  {lang.name} ({lang.native})
                                </option>
                              ))}
                            </select>
                          </div>
                        </div>
                      </div>

                      {/* Google Drive Cloud Backup Section */}
                      <div className="space-y-4">
                        <div className="flex items-center justify-between">
                          <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>
                            Google Drive Backup
                          </h3>
                          {authUser && (
                            <span className="flex items-center gap-1 font-mono text-[8px] uppercase tracking-wider font-extrabold px-1.5 py-0.5 rounded-full bg-indigo-500/10 text-indigo-500 dark:text-indigo-400 animate-pulse">
                              Active
                            </span>
                          )}
                        </div>

                        {!authUser ? (
                          <div className={`p-4 rounded-[2rem] border text-center transition-all ${
                            theme === 'dark' ? 'bg-slate-950/40 border-slate-800 text-slate-400' : 'bg-slate-50 border-slate-100 text-slate-500'
                          }`}>
                            <Cloud className="w-8 h-8 mx-auto text-indigo-500 mb-2.5 opacity-80" />
                            <p className={`text-xs font-black transition-all mb-1 ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Cloud Sync Offline</p>
                            <p className="text-[9.5px] leading-relaxed max-w-[240px] mx-auto text-slate-400 font-bold">
                              SIGN IN WITH GOOGLE UPPER ABOVE TO ACTIVATE MODERN SERVERLESS CLOUD AUTOMATION ON YOUR PRIVATE STORAGE DRIVES.
                            </p>
                          </div>
                        ) : (
                          <div className={`p-4 rounded-[2rem] border space-y-4 transition-all ${
                            theme === 'dark' ? 'bg-slate-950/20 border-slate-900' : 'bg-slate-50 border-slate-100'
                          }`}>
                            {/* Sync Status Status Monitor */}
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-2.5 text-left">
                                <div className={`p-2 rounded-xl shrink-0 ${
                                  driveSyncStatus === 'error' ? 'bg-rose-500/10 text-rose-500' : 
                                  driveSyncStatus === 'success' ? 'bg-emerald-500/10 text-emerald-500' : 
                                  'bg-indigo-500/10 text-indigo-550'
                                }`}>
                                  {driveSyncStatus === 'checking' || driveSyncStatus === 'backing-up' || driveSyncStatus === 'restoring' ? (
                                    <Loader2 className="w-4 h-4 animate-spin text-indigo-500" />
                                  ) : driveSyncStatus === 'error' ? (
                                    <Cloud className="w-4 h-4 text-rose-500" />
                                  ) : (
                                    <Cloud className="w-4 h-4 text-indigo-500" />
                                  )}
                                </div>
                                <div className="flex flex-col">
                                  <span className={`text-[11px] font-extrabold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                                    {driveSyncStatus === 'idle' && "Cloud Storage Ready"}
                                    {driveSyncStatus === 'checking' && "Checking cloud updates..."}
                                    {driveSyncStatus === 'backing-up' && "Saving backup to Drive..."}
                                    {driveSyncStatus === 'restoring' && "Restoring cloud file..."}
                                    {driveSyncStatus === 'success' && "Cloud backup synced"}
                                    {driveSyncStatus === 'error' && "Sync connection issue"}
                                  </span>
                                  <span className="text-[8px] text-slate-404 uppercase tracking-widest font-mono font-black mt-0.5">
                                    Last Sync: {driveLastBackup || "Never"}
                                  </span>
                                </div>
                              </div>
                            </div>

                            {/* Auto Sync Toggle switch */}
                            <div className={`flex items-center justify-between p-3 rounded-2xl border transition-all ${
                              theme === 'dark' ? 'bg-slate-900/40 border-slate-800/60' : 'bg-white border-slate-100'
                            }`}>
                              <div className="flex flex-col text-left">
                                <span className={`text-[10px] font-bold ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                                  Sync on Open
                                </span>
                                <span className="text-[8px] text-slate-400 font-bold uppercase tracking-wider font-mono">
                                  Automated cloud check on app startup
                                </span>
                              </div>
                              <button
                                type="button"
                                onClick={() => setAutoBackupEnabled(prev => !prev)}
                                className={`w-8 h-4 rounded-full p-0.5 transition-colors cursor-pointer relative ${
                                  autoBackupEnabled ? 'bg-indigo-500' : (theme === 'dark' ? 'bg-slate-800' : 'bg-slate-200')
                                }`}
                              >
                                <div className={`w-3 h-3 bg-white rounded-full shadow-sm transition-transform ${
                                  autoBackupEnabled ? 'translate-x-4' : 'translate-x-0'
                                }`} />
                              </button>
                            </div>

                            {/* Operations button lists */}
                            <div className="grid grid-cols-2 gap-2.5">
                              <button
                                type="button"
                                onClick={handleBackup}
                                disabled={isBackingUp || isRestoring || driveSyncStatus === 'checking' || driveSyncStatus === 'backing-up' || driveSyncStatus === 'restoring'}
                                className={`py-2 px-3 rounded-xl text-[10px] uppercase tracking-widest font-black flex items-center justify-center gap-1.5 border transition-all active:scale-[0.96] cursor-pointer ${
                                  theme === 'dark' 
                                    ? 'bg-slate-800 border-slate-705 text-slate-300 hover:bg-slate-755 disabled:opacity-50' 
                                    : 'bg-white border-slate-150 text-slate-650 hover:bg-slate-100 disabled:opacity-50'
                                }`}
                              >
                                {isBackingUp ? (
                                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                ) : (
                                  <CloudUpload className="w-3.5 h-3.5 text-indigo-500" />
                                )}
                                <span className="truncate">Backup Now</span>
                              </button>

                              <button
                                type="button"
                                onClick={handleRestore}
                                disabled={isBackingUp || isRestoring || driveSyncStatus === 'checking' || driveSyncStatus === 'backing-up' || driveSyncStatus === 'restoring'}
                                className={`py-2 px-3 rounded-xl text-[10px] uppercase tracking-widest font-black flex items-center justify-center gap-1.5 border transition-all active:scale-[0.96] cursor-pointer ${
                                  theme === 'dark' 
                                    ? 'bg-slate-800 border-slate-705 text-slate-300 hover:bg-slate-755 disabled:opacity-50' 
                                    : 'bg-white border-slate-150 text-slate-650 hover:bg-slate-100 disabled:opacity-50'
                                }`}
                              >
                                {isRestoring ? (
                                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                ) : (
                                  <CloudDownload className="w-3.5 h-3.5 text-emerald-500" />
                                )}
                                <span className="truncate">Restore</span>
                              </button>
                            </div>
                          </div>
                        )}
                      </div>

                      {/* Internal Memory Backup Section */}
                      <div className="space-y-4">
                        <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>Internal Backup (Device Memory)</h3>
                        
                        <div className="space-y-3">
                          <div className="grid grid-cols-2 gap-3">
                            <button 
                              onClick={handleLocalBackup}
                              className={`py-2 px-3 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 border transition-all active:scale-[0.95] ${
                                theme === 'dark' ? 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                              }`}
                              title="Save backup file to internal memory"
                            >
                              <Download className="w-4 h-4 text-indigo-500" />
                              <span className="truncate">Save to Memory</span>
                            </button>
                            
                            <label 
                              className={`py-2 px-3 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 border transition-all active:scale-[0.95] cursor-pointer ${
                                theme === 'dark' ? 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                              }`}
                              title="Restore backup file from internal memory"
                            >
                              <Upload className="w-4 h-4 text-emerald-500" />
                              <span className="truncate">Restore Backup</span>
                              <input 
                                type="file" 
                                accept=".json" 
                                onChange={handleLocalRestore} 
                                className="hidden" 
                              />
                            </label>
                          </div>
                          <p className="text-[9px] text-center text-slate-400 uppercase tracking-widest font-bold">Offline manual backup stored purely on your device</p>
                        </div>
                      </div>

                      {/* Gmail Workspace Integration Section */}
                      <div className="space-y-4">
                        <div className="flex items-center justify-between">
                          <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>
                            {t('gmailIntegration')}
                          </h3>
                          {authUser && (
                            <span className="flex items-center gap-1 font-mono text-[8px] uppercase tracking-wider font-extrabold px-1.5 py-0.5 rounded-full bg-indigo-550/10 text-indigo-500 dark:text-indigo-400">
                              Connected
                            </span>
                          )}
                        </div>

                        {!authUser ? (
                          <div className={`p-4 rounded-[2rem] border text-center transition-all ${
                            theme === 'dark' ? 'bg-slate-950/40 border-slate-800 text-slate-400' : 'bg-slate-50 border-slate-100 text-slate-500'
                          }`}>
                            <LogIn className="w-8 h-8 mx-auto text-indigo-500 mb-2.5 opacity-80" />
                            <p className={`text-xs font-black transition-all mb-1 ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Gmail Offline</p>
                            <p className="text-[9.5px] leading-relaxed max-w-[240px] mx-auto text-slate-400 font-bold">
                              SIGN IN WITH GOOGLE UPPER ABOVE TO SCAN YOUR INBOX FOR BROKER TRANSACTION CONFIRMATIONS AND DISPATCH PERFORMANCE SUMMARIES.
                            </p>
                          </div>
                        ) : (
                          <div className={`p-4 rounded-[2rem] border space-y-5 transition-all ${
                            theme === 'dark' ? 'bg-slate-950/20 border-slate-900' : 'bg-slate-50 border-slate-100'
                          }`}>
                            
                            {/* Scanner Sub-Block */}
                            <div className="space-y-3">
                              <h4 className={`text-[10px] font-bold uppercase tracking-widest text-left ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                                {t('recentBrokerEmails')}
                              </h4>
                              
                              <button
                                type="button"
                                onClick={handleScanGmail}
                                disabled={isScanningGmail}
                                className={`w-full py-2.5 px-4 rounded-xl text-[10px] uppercase tracking-widest font-black flex items-center justify-center gap-1.5 border transition-all active:scale-[0.96] cursor-pointer ${
                                  theme === 'dark' 
                                    ? 'bg-indigo-600/10 border-indigo-500/20 text-indigo-400 hover:bg-indigo-600/20 disabled:opacity-50' 
                                    : 'bg-indigo-50 border-indigo-100 text-indigo-600 hover:bg-indigo-100 disabled:opacity-50'
                                }`}
                              >
                                {isScanningGmail ? (
                                  <>
                                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                    <span>Scanning Inbox...</span>
                                  </>
                                ) : (
                                  <>
                                    <RefreshCcw className="w-3.5 h-3.5" />
                                    <span>{t('scanBrokerTrades')}</span>
                                  </>
                                )}
                              </button>

                              {gmailError && (
                                <p className="text-[9px] text-rose-500 font-bold self-start text-left">{gmailError}</p>
                              )}

                              {gmailEmails.length > 0 ? (
                                <div className="space-y-2 max-h-[220px] overflow-y-auto pr-1">
                                  {gmailEmails.map((email) => (
                                    <div 
                                      key={email.id} 
                                      className={`p-3 rounded-2xl border text-left flex flex-col gap-1.5 transition-all ${
                                        theme === 'dark' ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-slate-150'
                                      }`}
                                    >
                                      <div className="flex items-start justify-between gap-1.5">
                                        <div className="flex flex-col min-w-0 flex-1">
                                          <span className="font-mono text-[8px] uppercase tracking-widest text-slate-400 font-bold max-w-[170px] truncate">
                                            {email.from}
                                          </span>
                                          <span className={`text-[10px] font-extrabold truncate mt-0.5 ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                                            {email.subject}
                                          </span>
                                        </div>
                                        <span className="text-[8px] font-mono text-slate-450 font-bold shrink-0">{email.date}</span>
                                      </div>
                                      <p className="text-[9px] text-slate-400 line-clamp-2 leading-normal">{email.snippet}</p>
                                      
                                      {/* Parse Trade Indicator */}
                                      {email.parsedTrade ? (
                                        <div className={`p-2.5 rounded-xl border flex items-center justify-between gap-1.5 mt-1 ${
                                          theme === 'dark' ? 'bg-slate-950/60 border-indigo-500/10' : 'bg-slate-50 border-indigo-150'
                                        }`}>
                                          <div className="flex flex-col text-left">
                                            <span className="text-[8px] font-bold text-indigo-500 uppercase tracking-widest">Detected Transaction</span>
                                            <span className={`text-[10px] font-black mt-0.5 ${theme === 'dark' ? 'text-slate-200' : 'text-slate-850'}`}>
                                              {email.parsedTrade.action === 'BUY' ? 'Buy' : 'Sell'} {email.parsedTrade.quantity} {email.parsedTrade.ticker} @ ₹{email.parsedTrade.price}
                                            </span>
                                          </div>
                                          <button
                                            type="button"
                                            onClick={() => handleImportGmailTrade(email.parsedTrade, email.subject)}
                                            className="px-2.5 py-1 text-[9px] uppercase tracking-wider font-extrabold rounded bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-500 transition-colors active:scale-[0.95]"
                                          >
                                            {t('importToJournal')}
                                          </button>
                                        </div>
                                      ) : (
                                        <div className="flex items-center gap-1 text-[8.5px] text-slate-400 font-bold uppercase mt-1">
                                          <Info className="w-3 h-3 text-slate-500 shrink-0" />
                                          <span>General Trader Alert (Inbox item check)</span>
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                !isScanningGmail && (
                                  <p className="text-[9.5px] text-slate-400 font-bold py-2 text-center">
                                    {t('noEmailsFound')}
                                  </p>
                                )
                              )}
                            </div>

                            <hr className={`border-dashed ${theme === 'dark' ? 'border-slate-800' : 'border-slate-200'}`} />

                            {/* Dispatcher Sub-Block */}
                            <div className="space-y-3.5">
                              <h4 className={`text-[10px] font-bold uppercase tracking-widest text-left ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                                {t('sendPerfReport')}
                              </h4>

                              <div className="space-y-3">
                                {/* Recipient address */}
                                <div className="text-left">
                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">
                                    {t('recipientEmail')}
                                  </label>
                                  <input
                                    type="email"
                                    value={reportRecipient}
                                    onChange={(e) => setReportRecipient(e.target.value)}
                                    placeholder="your-email@gmail.com"
                                    className={`w-full p-2 text-[11px] rounded-xl font-medium border focus:ring-1 focus:ring-indigo-500 focus:outline-none transition-colors ${
                                      theme === 'dark'
                                        ? 'bg-slate-900 border-slate-800 text-white'
                                        : 'bg-white border-slate-200 text-slate-800'
                                    }`}
                                  />
                                </div>

                                {/* Custom Note Subject */}
                                <div className="text-left">
                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">
                                    {t('customReportNote')}
                                  </label>
                                  <textarea
                                    value={reportCustomMessage}
                                    onChange={(e) => setReportCustomMessage(e.target.value)}
                                    placeholder="Add notes for your portfolio tracking..."
                                    rows={2}
                                    className={`w-full p-2 text-[11px] rounded-xl font-medium border focus:ring-1 focus:ring-indigo-500 focus:outline-none transition-colors resize-none ${
                                      theme === 'dark'
                                        ? 'bg-slate-900 border-slate-800 text-white'
                                        : 'bg-white border-slate-200 text-slate-800'
                                    }`}
                                  />
                                </div>

                                <button
                                  type="button"
                                  onClick={handleSendPerformanceEmail}
                                  disabled={isSendingReport}
                                  className={`w-full py-2.5 px-4 rounded-xl text-[10px] uppercase tracking-widest font-black flex items-center justify-center gap-1.5 border transition-colors active:scale-[0.96] cursor-pointer ${
                                    theme === 'dark' 
                                      ? 'bg-emerald-600/10 border-emerald-500/20 text-emerald-400 hover:bg-emerald-600/20 disabled:opacity-50' 
                                      : 'bg-emerald-50 border-emerald-100 text-emerald-600 hover:bg-emerald-100 disabled:opacity-50'
                                  }`}
                                >
                                  {isSendingReport ? (
                                    <>
                                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                      <span>Sending Report...</span>
                                    </>
                                  ) : (
                                    <>
                                      <Send className="w-3.5 h-3.5 text-emerald-500" />
                                      <span>{t('sendMimeReport')}</span>
                                    </>
                                  )}
                                </button>
                              </div>
                            </div>
                          </div>
                        )}
                      </div>

                      <hr className={`border-dashed ${theme === 'dark' ? 'border-slate-800' : 'border-slate-200'}`} />

                      {/* Calendar Holiday Settings Section */}
                      <div className="space-y-4">
                        <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>
                          Calendar Holiday Colors
                        </h3>
                        
                        <div className={`p-4 rounded-[2rem] border space-y-4 transition-all ${
                          theme === 'dark' ? 'bg-slate-800/20 border-slate-800' : 'bg-slate-50 border-slate-100'
                        }`}>
                          {/* Holiday Color Theme Choice */}
                          <div className="space-y-2">
                            <span className={`text-[10px] font-bold block text-left ${theme === 'dark' ? 'text-slate-300' : 'text-slate-750'}`}>
                              Holiday Color Theme Accent
                            </span>
                            
                            <div className="grid grid-cols-3 gap-2">
                              {[
                                { id: 'rose', label: 'Rose', color: 'bg-rose-500' },
                                { id: 'amber', label: 'Amber', color: 'bg-amber-500' },
                                { id: 'indigo', label: 'Indigo', color: 'bg-indigo-500' },
                                { id: 'emerald', label: 'Emerald', color: 'bg-emerald-500' },
                                { id: 'violet', label: 'Violet', color: 'bg-violet-500' },
                                { id: 'slate', label: 'Slate', color: 'bg-slate-500' }
                              ].map(scheme => (
                                <button
                                  key={scheme.id}
                                  type="button"
                                  onClick={() => setHolidayColorTheme(scheme.id as any)}
                                  className={`py-2 px-1.5 rounded-xl text-[10px] uppercase font-black tracking-widest flex items-center justify-center gap-1.5 border transition-all active:scale-[0.96] cursor-pointer ${
                                    holidayColorTheme === scheme.id
                                      ? 'border-indigo-500 ring-2 ring-indigo-550/20 bg-indigo-500/10 text-indigo-500 font-extrabold'
                                      : theme === 'dark'
                                        ? 'bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800'
                                        : 'bg-white border-slate-150 text-slate-605 hover:bg-slate-100'
                                  }`}
                                >
                                  <div className={`w-2 h-2 rounded-full shrink-0 ${scheme.color}`} />
                                  <span>{scheme.label}</span>
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Toggle Include Saturdays & Sundays in Holiday color change */}
                          <div className={`flex items-center justify-between p-3 rounded-2xl border transition-all ${
                            theme === 'dark' ? 'bg-slate-900/40 border-slate-800/60' : 'bg-white border-slate-100'
                          }`}>
                            <div className="flex flex-col text-left">
                              <span className={`text-[10px] font-bold ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                                Apply to Sat & Sun
                              </span>
                              <span className="text-[8px] text-slate-400 font-bold uppercase tracking-wider font-mono">
                                Match holiday colors with Saturdays & Sundays
                              </span>
                            </div>
                            <button
                              type="button"
                              onClick={() => setIncludeWeekendsInHolidayTheme(prev => !prev)}
                              className={`w-8 h-4 rounded-full p-0.5 transition-colors cursor-pointer relative ${
                                includeWeekendsInHolidayTheme ? 'bg-indigo-500' : (theme === 'dark' ? 'bg-slate-800' : 'bg-slate-200')
                              }`}
                            >
                              <div className={`w-3 h-3 bg-white rounded-full shadow-sm transition-transform ${
                                includeWeekendsInHolidayTheme ? 'translate-x-4' : 'translate-x-0'
                              }`} />
                            </button>
                          </div>
                        </div>
                      </div>

                      {/* Support Creator Block */}
                      <div className="space-y-4">
                        <h3 className={`text-[10px] font-bold uppercase tracking-widest transition-colors ${theme === 'dark' ? 'text-slate-500' : 'text-slate-400'}`}>Back the App</h3>
                        
                        <div className={`p-5 rounded-[2rem] border transition-all ${
                          theme === 'dark' 
                            ? 'bg-indigo-950/20 border-indigo-900/40 hover:border-indigo-800' 
                            : 'bg-indigo-50/40 border-indigo-100 hover:border-indigo-200'
                        }`}>
                          <div className="flex items-start gap-4">
                            <div className={`p-3 rounded-2xl shrink-0 ${theme === 'dark' ? 'bg-indigo-900/30 text-indigo-455 font-extrabold' : 'bg-indigo-100 text-indigo-600 font-extrabold'}`}>
                              <Heart className="w-6 h-6 fill-indigo-500/10 text-indigo-550" />
                            </div>
                            <div className="flex-1 text-left">
                              <h4 className={`text-sm font-black transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Support Founder</h4>
                              <p className={`text-[11px] leading-relaxed mt-1 mb-3.5 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                                Help us maintain and build state-of-the-art features for the trading community.
                              </p>
                              <button
                                type="button"
                                onClick={() => {
                                  if (!supportName) setSupportName(userName || '');
                                  if (!supportEmail) setSupportEmail(authUser?.email || '');
                                  setSupportStep('info');
                                }}
                                className="text-[10px] font-black uppercase tracking-widest bg-green-600 hover:bg-green-700 text-white px-4 py-2.5 rounded-xl cursor-pointer transition-all shadow-md shadow-green-600/15 flex items-center gap-1.5"
                              >
                                <span>Support Founder</span>
                                <ChevronRight className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </>
                  ) : supportStep === 'info' ? (
                    <div className="space-y-6 text-left">
                      <div className="text-center pb-2">
                        <span className="text-[9px] font-black text-indigo-500 uppercase tracking-widest block mb-1">Step 1 of 2</span>
                        <h3 className={`text-base font-extrabold ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Your Profile Details</h3>
                        <p className={`text-[11px] leading-relaxed mt-1 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                          Give us some details so the founder knows who backed this workspace.
                        </p>
                      </div>

                      <div className="space-y-4">
                        <div>
                          <label className={`text-[10px] font-bold uppercase tracking-wider block mb-1.5 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Your Name</label>
                          <input
                            type="text"
                            value={supportName}
                            onChange={(e) => setSupportName(e.target.value)}
                            placeholder="e.g. Rahul Sharma"
                            className={`w-full py-2.5 px-4 rounded-xl border text-xs focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                              theme === 'dark' 
                                ? 'bg-slate-950 border-slate-800 text-white placeholder:text-slate-700' 
                                : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                            }`}
                          />
                        </div>

                        <div>
                          <label className={`text-[10px] font-bold uppercase tracking-wider block mb-1.5 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Your Location</label>
                          <input
                            type="text"
                            value={supportLocation}
                            onChange={(e) => setSupportLocation(e.target.value)}
                            placeholder="e.g. Mumbai, India"
                            className={`w-full py-2.5 px-4 rounded-xl border text-xs focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                              theme === 'dark' 
                                ? 'bg-slate-950 border-slate-800 text-white placeholder:text-slate-700' 
                                : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                            }`}
                          />
                        </div>

                        <div>
                          <label className={`text-[10px] font-bold uppercase tracking-wider block mb-1.5 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Email Address</label>
                          <input
                            type="email"
                            value={supportEmail}
                            onChange={(e) => setSupportEmail(e.target.value)}
                            placeholder="e.g. rahul@gmail.com"
                            className={`w-full py-2.5 px-4 rounded-xl border text-xs focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                              theme === 'dark' 
                                ? 'bg-slate-950 border-slate-800 text-white placeholder:text-slate-700' 
                                : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                            }`}
                          />
                        </div>
                      </div>

                      <div className="pt-4 flex gap-3">
                        <button
                          type="button"
                          onClick={() => setSupportStep('menu')}
                          className={`flex-1 py-3 rounded-xl font-bold text-xs transition-colors ${
                            theme === 'dark' ? 'bg-slate-800 text-slate-400 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                          }`}
                        >
                          Cancel
                        </button>
                        <button
                          type="button"
                          onClick={() => setSupportStep('pay')}
                          disabled={!supportName.trim() || !supportLocation.trim() || !supportEmail.trim()}
                          className={`flex-1 py-3 rounded-xl font-black text-xs uppercase tracking-widest transition-all shadow-md ${
                            (!supportName.trim() || !supportLocation.trim() || !supportEmail.trim())
                              ? 'bg-slate-300 text-slate-400 dark:bg-slate-800 dark:text-slate-700 cursor-not-allowed opacity-50'
                              : 'bg-green-600 text-white hover:bg-green-700 shadow-green-600/10 active:scale-[0.98]'
                          }`}
                        >
                          Next: Support
                        </button>
                      </div>
                    </div>
                  ) : supportStep === 'pay' ? (
                    <div className="space-y-6 text-left">
                      {/* Form Details Header */}
                      <div className={`p-4 rounded-2xl border flex items-center justify-between text-xs transition-colors ${
                        theme === 'dark' ? 'bg-slate-800/40 border-slate-800' : 'bg-green-50/20 border-green-105'
                      }`}>
                        <div>
                          <p className="font-bold text-green-500 uppercase tracking-wider text-[8px] mb-0.5">Contributor Details</p>
                          <p className={`font-semibold ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>{supportName}</p>
                          <p className="text-[10px] text-slate-400">{supportLocation}</p>
                        </div>
                        <button 
                          onClick={() => setSupportStep('info')}
                          className="text-[9px] text-green-500 font-black uppercase tracking-wider hover:underline"
                        >
                          Edit Info
                        </button>
                      </div>

                      {/* Select Amount Section */}
                      <div className="space-y-3">
                        <label className={`text-[10px] font-bold uppercase tracking-wider block ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                          Select Support Amount
                        </label>
                        <div className="grid grid-cols-4 gap-2">
                          {['100', '250', '500', 'custom'].map((val) => {
                            const isSelected = supportAmount === val;
                            return (
                              <button
                                key={val}
                                type="button"
                                onClick={() => setSupportAmount(val)}
                                className={`py-3 rounded-xl border font-mono font-black text-xs uppercase tracking-widest transition-all cursor-pointer ${
                                  isSelected
                                    ? 'bg-green-600 border-green-600 text-white shadow-md'
                                    : theme === 'dark'
                                      ? 'bg-slate-800 border-slate-700 text-slate-300 hover:border-slate-500'
                                      : 'bg-white border-slate-200 text-slate-700 hover:border-slate-300'
                                }`}
                              >
                                {val === 'custom' ? 'Custom' : `₹${val}`}
                              </button>
                            );
                          })}
                        </div>

                        {supportAmount === 'custom' && (
                          <div className="pt-1.5 transition-all">
                            <input
                              type="number"
                              min="1"
                              value={supportCustomAmount}
                              onChange={(e) => setSupportCustomAmount(e.target.value)}
                              placeholder="Enter Custom Amount (₹)"
                              className={`w-full py-2.5 px-4 rounded-xl border text-xs font-mono font-bold focus:ring-1 focus:ring-indigo-500 outline-none transition-all ${
                                theme === 'dark' 
                                  ? 'bg-slate-950 border-slate-800 text-white' 
                                  : 'bg-white border-slate-200 text-slate-900'
                              }`}
                            />
                          </div>
                        )}
                      </div>

                      {/* Pay UPI Section */}
                      <div className="space-y-3">
                        <span className={`text-[10px] font-bold uppercase tracking-wider block ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                          Choose UPI Application
                        </span>
                        
                        <div className="space-y-2">
                          {[
                            { id: 'gpay', name: 'Google Pay (GPay)', color: 'border-emerald-500/30 bg-emerald-550/5 text-emerald-500' },
                            { id: 'phonepe', name: 'PhonePe', color: 'border-indigo-500/30 bg-indigo-550/5 text-indigo-550' },
                            { id: 'paytm', name: 'Paytm / BHIM UPI', color: 'border-blue-550/30 bg-blue-550/5 text-blue-500' }
                          ].map((app) => (
                            <button
                              key={app.id}
                              type="button"
                              onClick={() => {
                                handleSimulatePayment(app.name);
                              }}
                              disabled={isProcessingPayment || (supportAmount === 'custom' && !supportCustomAmount)}
                              className={`w-full p-3.5 rounded-xl border flex items-center justify-between text-xs transition-all active:scale-[0.98] ${
                                (supportAmount === 'custom' && !supportCustomAmount)
                                  ? 'opacity-40 cursor-not-allowed'
                                  : theme === 'dark' 
                                    ? 'bg-slate-800/10 border-slate-800 hover:border-slate-700 hover:bg-slate-800/30' 
                                    : 'bg-white border-slate-150 hover:bg-slate-50'
                              } ${isProcessingPayment ? 'opacity-50' : 'cursor-pointer'}`}
                            >
                              <div className="flex items-center gap-3">
                                <span className={`w-8 h-8 rounded-lg border font-mono font-black text-[9px] uppercase flex items-center justify-center ${app.color}`}>
                                  UPI
                                </span>
                                <div className="text-left">
                                  <p className={`font-extrabold ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>{app.name}</p>
                                  <p className="text-[10px] text-slate-400">Instant UPI payment portal</p>
                                </div>
                              </div>
                              <ChevronRight className="w-4 h-4 text-slate-450" />
                            </button>
                          ))}
                        </div>

                        {/* Copy UPI ID as fallback */}
                        <div className={`p-3 rounded-xl border text-center transition-colors ${
                          theme === 'dark' ? 'bg-slate-950/40 border-slate-800/80' : 'bg-slate-50 border-slate-200'
                        }`}>
                          <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Direct Merchant UPI ID</p>
                          <p className="text-xs font-mono font-black text-green-500 select-all">jidhinraju@ybl</p>
                          <p className="text-[9px] text-slate-400 mt-1 uppercase tracking-wider">Payments are fully simulated for security validation</p>
                        </div>
                      </div>

                      {/* Display loading screen if processing */}
                      {isProcessingPayment && (
                        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-[60] flex flex-col items-center justify-center p-4 sm:p-6 text-center overflow-y-auto animate-in fade-in transition-all">
                          <div className={`w-full max-w-md p-6 rounded-[2rem] border shadow-2xl transition-all ${
                            theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-100'
                          }`}>
                            <div className="flex justify-between items-center mb-4">
                              <h3 className={`text-sm font-black uppercase tracking-widest ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                                UPI Payment Portal
                              </h3>
                              <button 
                                onClick={() => setIsProcessingPayment(false)}
                                className={`p-1.5 rounded-full transition-colors ${theme === 'dark' ? 'hover:bg-slate-800 text-slate-400' : 'hover:bg-slate-100 text-slate-500'}`}
                              >
                                ✕
                              </button>
                            </div>

                            <p className={`text-xs mb-5 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                              We tried opening your chosen UPI app. If it didn't launch automatically, tap the link or scan the QR code to pay <b>₹{getCurrentPaymentAmount()}</b>.
                            </p>

                            {/* Mobile Launch Section */}
                            <div className="space-y-3 mb-6">
                              <div className="flex flex-col sm:flex-row gap-2.5">
                                <a 
                                  href={activeUpiUrl}
                                  className="flex-1 py-3 px-4 rounded-xl bg-green-600 hover:bg-green-700 text-white font-extrabold text-[10px] uppercase tracking-wider text-center transition-all flex items-center justify-center gap-1.5 shadow-md shadow-green-650/15"
                                >
                                  <span>{activeAppName || 'Launch Chosen App'}</span>
                                </a>
                                <a 
                                  href={universalUpiUrl}
                                  className={`flex-1 py-3 px-4 rounded-xl font-extrabold text-[10px] uppercase tracking-wider text-center transition-all border flex items-center justify-center gap-1.5 ${
                                    theme === 'dark' 
                                      ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700' 
                                      : 'bg-slate-50 border-slate-250 text-slate-700 hover:bg-slate-100'
                                  }`}
                                >
                                  <span>Any UPI App</span>
                                </a>
                              </div>
                            </div>

                            {/* Desktop QR Scan Section */}
                            <div className="space-y-3 mb-6">
                              <p className="text-[9px] font-black uppercase tracking-widest text-slate-400">Scan QR to pay directly</p>
                              <div className="flex justify-center">
                                <div className="bg-white p-3 rounded-2xl inline-block shadow-inner border border-slate-200">
                                  <img 
                                    src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(universalUpiUrl)}`}
                                    alt="UPI Payment QR Code"
                                    className="w-36 h-36 mx-auto bg-white"
                                    referrerPolicy="no-referrer"
                                  />
                                </div>
                              </div>
                              <p className="text-[9px] text-slate-450 italic">Scan with Google Pay, PhonePe, Paytm, or BHIM</p>
                            </div>

                            {/* Manual UPI ID copy block */}
                            <div className={`p-3.5 rounded-2xl border text-center transition-colors mb-6 ${
                              theme === 'dark' ? 'bg-slate-950/40 border-slate-800/85' : 'bg-slate-50 border-slate-200'
                            }`}>
                              <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest mb-1">Direct Merchant UPI ID</p>
                              <div className="flex items-center justify-center gap-2">
                                <span className="text-xs font-mono font-black text-green-600 select-all">jidhinraju@ybl</span>
                                <button 
                                  type="button"
                                  onClick={() => {
                                    navigator.clipboard.writeText('jidhinraju@ybl');
                                    setCopiedUpi(true);
                                    setTimeout(() => setCopiedUpi(false), 2000);
                                  }}
                                  className={`text-[9px] font-black uppercase tracking-wider py-1 px-2.5 rounded-lg border transition-all ${
                                    copiedUpi
                                      ? 'bg-green-500/10 border-green-500 text-green-500'
                                      : theme === 'dark'
                                        ? 'border-slate-700 text-slate-300 hover:bg-slate-800'
                                        : 'border-slate-300 text-slate-600 hover:bg-white shadow-sm'
                                  }`}
                                >
                                  {copiedUpi ? 'Copied!' : 'Copy'}
                                </button>
                              </div>
                            </div>

                            {/* Finish/Back actions */}
                            <div className="flex gap-2.5 pt-1 border-t border-slate-200/20">
                              <button
                                type="button"
                                onClick={() => setIsProcessingPayment(false)}
                                className={`flex-1 py-3 rounded-xl font-bold text-xs uppercase tracking-wider transition-all ${
                                  theme === 'dark'
                                    ? 'bg-slate-800 text-slate-400 hover:bg-slate-700'
                                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                                }`}
                              >
                                Cancel
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  setIsProcessingPayment(false);
                                  setSupportStep('success');
                                }}
                                className="flex-1 py-3 rounded-xl bg-indigo-650 hover:bg-indigo-700 text-white font-extrabold text-xs uppercase tracking-wider transition-all shadow-md shadow-indigo-600/15"
                              >
                                Done / I've Paid
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="space-y-6 text-center py-6">
                      <div className="flex justify-center">
                        <motion.div 
                          initial={{ scale: 0 }}
                          animate={{ scale: [0, 1.2, 1] }}
                          transition={{ duration: 0.5 }}
                          className="w-16 h-16 rounded-full bg-green-500/10 border border-green-500/30 flex items-center justify-center text-green-555"
                        >
                          <Heart className="w-8 h-8 fill-green-500 text-green-500" />
                        </motion.div>
                      </div>

                      <div className="space-y-2">
                        <h4 className={`text-lg font-black uppercase tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                          Support Received! 🎉
                        </h4>
                        <p className={`text-xs px-2 leading-relaxed ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                          Thank you <span className="font-extrabold text-green-600">{supportName}</span> from <span className="font-extrabold text-green-600">{supportLocation}</span>! Your donation of <span className="font-mono font-bold text-green-600">₹{getCurrentPaymentAmount()}</span> has been simulated successfully.
                        </p>
                        <p className={`text-[11px] leading-relaxed p-3.5 rounded-xl border text-left italic ${
                          theme === 'dark' ? 'bg-slate-950/40 border-slate-850 text-slate-400' : 'bg-slate-50 border-slate-150 text-slate-500'
                        }`}>
                          "Independent tracking tools thrive because of amazing traders like yourself. I appreciate your support in making trading simpler!"<br />
                          <span className="font-black pt-1 block text-[10px] uppercase tracking-wider text-green-500 not-italic">— App Founder</span>
                        </p>
                      </div>

                      <div className="pt-4 px-4">
                        <button
                          type="button"
                          onClick={() => {
                            setSupportStep('menu');
                            setShowSettings(false);
                          }}
                          className="w-full py-3.5 rounded-xl bg-green-600 hover:bg-green-700 text-white font-black text-xs uppercase tracking-widest transition-all shadow-md shadow-green-600/15 cursor-pointer"
                        >
                          You're Welcome
                        </button>
                      </div>
                    </div>
                  )}
                </div>


              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Manual Trade Journal Entry Modal */}
        <AnimatePresence>
          {isManualJournalOpen && (
            <div 
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setIsManualJournalOpen(false);
                }
              }}
              className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 bg-black/50 backdrop-blur-sm"
            >
              <motion.div 
                initial={{ y: '100%', opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                exit={{ y: '100%', opacity: 0 }}
                className={`relative w-full max-w-lg rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] flex flex-col transition-colors duration-300 ${
                  theme === 'dark' ? 'bg-slate-900' : 'bg-white'
                }`}
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-3 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />
                
                <div className={`px-6 pb-4 border-b flex items-center gap-3 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  <button
                    onClick={() => setIsManualJournalOpen(false)}
                    className={`p-2 rounded-xl border transition-all active:scale-95 ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                    aria-label="Back to Trading"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <div className="flex flex-col flex-1">
                    <h2 className={`text-lg font-bold transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Manual Journal Entry</h2>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Log trades that can't be computed automatically</p>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-6">
                  {/* Ticker & Mode & Type */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Ticker / Ticker Name</label>
                      <input
                        type="text"
                        placeholder="e.g. RELIANCE, NIFTY 22000 CE"
                        value={manualTicker}
                        onChange={(e) => setManualTicker(e.target.value)}
                        className={`w-full border rounded-xl py-3 px-3.5 focus:ring-2 focus:ring-indigo-555 outline-none transition-all text-sm ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Strategy Used</label>
                      <input
                        type="text"
                        placeholder="e.g. Breakout, Support Retest"
                        value={manualStrategy}
                        onChange={(e) => setManualStrategy(e.target.value)}
                        className={`w-full border rounded-xl py-3 px-3.5 focus:ring-2 focus:ring-indigo-555 outline-none transition-all text-sm ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Trade Mode</label>
                      <div className={`p-1.5 rounded-xl border flex gap-1 ${theme === 'dark' ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                        <button
                          onClick={() => setManualMode('equity')}
                          className={`flex-1 py-1 px-3 text-xs font-bold rounded-lg transition-all ${
                            manualMode === 'equity' 
                              ? 'bg-indigo-650 text-white shadow-sm' 
                              : `text-slate-500 hover:text-slate-800`
                          }`}
                        >
                          Equity
                        </button>
                        <button
                          onClick={() => setManualMode('options')}
                          className={`flex-1 py-1 px-3 text-xs font-bold rounded-lg transition-all ${
                            manualMode === 'options' 
                              ? 'bg-indigo-650 text-white shadow-sm' 
                              : `text-slate-500 hover:text-slate-800`
                          }`}
                        >
                          Options
                        </button>
                      </div>
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Trade Direction</label>
                      <div className={`p-1.5 rounded-xl border flex gap-1 ${theme === 'dark' ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                        <button
                          onClick={() => setManualType('long')}
                          className={`flex-1 py-1 px-3 text-xs font-bold rounded-lg transition-all flex items-center justify-center gap-1 ${
                            manualType === 'long' 
                              ? 'bg-green-500 text-white shadow-sm' 
                              : `text-slate-500 hover:text-green-500`
                          }`}
                        >
                          <TrendingUp className="w-3.5 h-3.5" />
                          Long
                        </button>
                        <button
                          onClick={() => setManualType('short')}
                          className={`flex-1 py-1 px-3 text-xs font-bold rounded-lg transition-all flex items-center justify-center gap-1 ${
                            manualType === 'short' 
                              ? 'bg-red-500 text-white shadow-sm' 
                              : `text-slate-500 hover:text-red-500`
                          }`}
                        >
                          <TrendingDown className="w-3.5 h-3.5" />
                          Short
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Prices & Quantities */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Entry Price (₹)</label>
                      <input
                        type="number"
                        placeholder="0.00"
                        value={manualEntryPrice}
                        onChange={(e) => setManualEntryPrice(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-xs font-mono ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Exit Price (₹)</label>
                      <input
                        type="number"
                        placeholder="0.00"
                        value={manualExitPrice}
                        onChange={(e) => setManualExitPrice(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-xs font-mono ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">
                        {manualMode === 'equity' ? 'Units' : 'Lots'}
                      </label>
                      <input
                        type="number"
                        placeholder={manualMode === 'equity' ? '100' : '1'}
                        value={manualQuantity}
                        onChange={(e) => setManualQuantity(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-xs font-mono ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                    {manualMode === 'options' && (
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Lot Size</label>
                        <input
                          type="number"
                          placeholder="65"
                          value={manualLotSize}
                          onChange={(e) => setManualLotSize(e.target.value)}
                          className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-xs font-mono ${
                            theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                          }`}
                        />
                      </div>
                    )}
                  </div>

                  {/* Dates */}
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Entry Date</label>
                      <input
                        type="date"
                        value={manualEntryDate}
                        onChange={(e) => setManualEntryDate(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-505 outline-none transition-all text-xs ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white color-scheme-dark' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Exit Date</label>
                      <input
                        type="date"
                        value={manualExitDate}
                        onChange={(e) => setManualExitDate(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 focus:ring-2 focus:ring-indigo-505 outline-none transition-all text-xs ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white color-scheme-dark' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                  </div>

                  {/* Custom Profit/Loss field */}
                  <div className={`p-4 rounded-xl border space-y-3 ${
                    theme === 'dark' ? 'bg-slate-800/40 border-slate-700' : 'bg-slate-50 border-slate-100'
                  }`}>
                    <div className="flex items-center justify-between">
                      <div className="flex flex-col">
                        <span className="text-xs font-extrabold tracking-wide uppercase">Custom Profit/Loss Value</span>
                        <span className="text-[9px] text-slate-400 uppercase tracking-widest">Override calculation with exact custom PnL</span>
                      </div>
                      <input 
                        type="checkbox"
                        checked={manualProfitCustom}
                        onChange={(e) => setManualProfitCustom(e.target.checked)}
                        className="w-4 h-4 cursor-pointer accent-indigo-600 rounded-md"
                      />
                    </div>

                    {manualProfitCustom && (
                      <div className="space-y-1.5 pt-1 animate-fadeIn">
                        <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Actual Profit/Loss Amount (₹)</label>
                        <input
                          type="number"
                          placeholder="e.g. 1500 for profit, -750 for loss"
                          value={manualProfitAmount}
                          onChange={(e) => setManualProfitAmount(e.target.value)}
                          className={`w-full border rounded-xl py-2 px-3.5 focus:ring-2 focus:ring-indigo-500 outline-none transition-all text-xs font-mono font-bold ${
                            theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                          }`}
                        />
                      </div>
                    )}
                  </div>

                  {/* Detailed Journal Fields */}
                  <div className="space-y-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Trade Setup Setup</label>
                      <textarea
                        rows={2}
                        placeholder="What indicators, volume patterns, chart setups or key levels aligned here?"
                        value={manualTradeSetup}
                        onChange={(e) => setManualTradeSetup(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Entry & Exit Rationale</label>
                      <textarea
                        rows={2}
                        placeholder="What triggered your entry & exit decisions? Any specific signal or target logic?"
                        value={manualEntryExitRationale}
                        onChange={(e) => setManualEntryExitRationale(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Lessons Learned</label>
                      <textarea
                        rows={2}
                        placeholder="What went well? What rules did you break? How can you improve next time?"
                        value={manualLessonsLearned}
                        onChange={(e) => setManualLessonsLearned(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Brief Verdict Notes</label>
                      <textarea
                        rows={2}
                        placeholder="Any additional thoughts or trading mindset comments..."
                        value={manualNotes}
                        onChange={(e) => setManualNotes(e.target.value)}
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                <div className={`p-6 border-t shrink-0 flex gap-3 transition-colors ${theme === 'dark' ? 'border-slate-800' : 'border-slate-100'}`}>
                  <button 
                    onClick={() => setIsManualJournalOpen(false)}
                    className={`flex-1 py-4 rounded-2xl font-bold transition-colors text-sm ${
                      theme === 'dark' ? 'bg-slate-800 text-slate-350 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                    }`}
                  >
                    Cancel
                  </button>
                  <button 
                    onClick={handleSaveManualTrade}
                    className="flex-1 py-4 bg-indigo-600 text-white rounded-2xl font-bold transition-all text-sm hover:bg-indigo-700 shadow-xl shadow-indigo-900/10 active:scale-[0.98]"
                  >
                    Log Trade Entry
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Existing Trade Journal Editor Modal */}
        <AnimatePresence>
          {selectedJournalTrade && (
            <div 
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setSelectedJournalTrade(null);
                }
              }}
              className="fixed inset-0 z-[60] flex items-end sm:items-center justify-center p-0 sm:p-4 bg-black/60 backdrop-blur-sm"
            >
              <motion.div 
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.95, opacity: 0 }}
                className={`relative w-full max-w-lg rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] flex flex-col transition-all duration-200 ${
                  theme === 'dark' ? 'bg-slate-900' : 'bg-white'
                }`}
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-3 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />

                <div className={`px-6 pb-4 border-b flex items-center gap-3 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  <button
                    onClick={() => setSelectedJournalTrade(null)}
                    className={`p-2 rounded-xl border transition-all active:scale-95 ${
                      theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                    aria-label="Back to History"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <div className="flex flex-col flex-1">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Trade Journal & Analysis</span>
                    <h2 className="text-base font-extrabold flex items-center gap-2">
                      <span className={`w-2.5 h-2.5 rounded-full ${selectedJournalTrade.profit >= 0 ? 'bg-green-500' : 'bg-red-500'}`}></span>
                      {selectedJournalTrade.ticker || selectedJournalTrade.strategy || 'Trade Detail'}
                    </h2>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-6">
                  {/* Summary card */}
                  <div className={`p-4 rounded-2xl flex justify-between items-center ${
                    theme === 'dark' ? 'bg-slate-800/40 text-slate-300 border border-slate-800' : 'bg-slate-50 text-slate-700 border border-slate-100'
                  }`}>
                    <div className="text-xs space-y-1">
                      <div>
                        Direction: <span className="font-extrabold uppercase text-indigo-500">{selectedJournalTrade.type} {selectedJournalTrade.mode}</span>
                      </div>
                      <div>
                        Entry price: <span className="font-mono font-bold">₹{selectedJournalTrade.entry.toLocaleString('en-IN')}</span> • Exit: <span className="font-mono font-bold">₹{selectedJournalTrade.exit.toLocaleString('en-IN')}</span>
                      </div>
                      <div>
                        Units: <span className="font-bold">{selectedJournalTrade.quantity} {selectedJournalTrade.mode === 'equity' ? 'units' : `lots`}</span>
                      </div>
                    </div>
                    <div className="text-right">
                      <div className="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Net Result</div>
                      <div className={`text-lg font-black font-mono ${selectedJournalTrade.profit >= 0 ? 'text-green-500' : 'text-red-500'}`}>
                        {selectedJournalTrade.profit >= 0 ? '+' : ''}₹{selectedJournalTrade.profit.toLocaleString('en-IN')}
                      </div>
                    </div>
                  </div>

                  {/* Edits */}
                  <div className="space-y-4">
                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Modify Asset/Ticker Name</label>
                      <input
                        type="text"
                        value={selectedJournalTrade.ticker || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { ticker: e.target.value })}
                        placeholder="e.g. RELIANCE, INFOSYS, NIFTY Call Option"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-550 outline-none transition-all ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-655' : 'bg-white border-slate-200 text-slate-900 placeholder:text-slate-350'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Update Net Profit/Loss (₹)</label>
                      <input
                        type="number"
                        value={selectedJournalTrade.profit || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { profit: parseFloat(e.target.value) || 0 })}
                        placeholder="Trade Result"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-550 outline-none transition-all font-mono font-bold ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className={`h-px transition-colors ${theme === 'dark' ? 'bg-slate-800' : 'bg-slate-100'}`} />

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Trade Setup Details</label>
                      <textarea
                        rows={3}
                        value={selectedJournalTrade.tradeSetup || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { tradeSetup: e.target.value })}
                        placeholder="Describe the trade setup e.g. Fibonacci 61.8% support retest with RSI bullish divergence"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Entry/Exit Rationale</label>
                      <textarea
                        rows={3}
                        value={selectedJournalTrade.entryExitRationale || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { entryExitRationale: e.target.value })}
                        placeholder="What triggers and stops were defined? Did you execute on rules or emotional impulses?"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Lessons Learned & Takeaways</label>
                      <textarea
                        rows={3}
                        value={selectedJournalTrade.lessonsLearned || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { lessonsLearned: e.target.value })}
                        placeholder="How can you optimize this entry? Any changes needed in size management or stop trailing?"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest pl-1">Verdict Notes / General Insights</label>
                      <textarea
                        rows={2}
                        value={selectedJournalTrade.notes || ''}
                        onChange={(e) => handleUpdateJournal(selectedJournalTrade.id, { notes: e.target.value })}
                        placeholder="Notes or trading psychology observations"
                        className={`w-full border rounded-xl py-2.5 px-3 text-xs focus:ring-2 focus:ring-indigo-555 outline-none transition-all resize-none ${
                          theme === 'dark' ? 'bg-slate-800 border-slate-700 text-white placeholder:text-slate-650' : 'bg-white border-slate-200 text-slate-900'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                <div className={`p-6 border-t shrink-0 transition-colors ${theme === 'dark' ? 'border-slate-800' : 'border-slate-100'}`}>
                  <button 
                    onClick={() => setSelectedJournalTrade(null)}
                    className="w-full py-4 bg-indigo-600 text-white rounded-2xl font-bold transition-all text-sm hover:bg-indigo-700 shadow-xl active:scale-[0.98]"
                  >
                    Save & Sync Journal Details
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Trading Journal AI Chat Panel Overlay */}
        <AnimatePresence>
          {showAIChat && (
            <div 
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setShowAIChat(false);
                }
              }}
              className="fixed inset-0 z-[70] flex items-end sm:items-center justify-center p-0 sm:p-4 bg-black/60 backdrop-blur-sm"
            >
              <motion.div 
                initial={{ y: '100%', opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                exit={{ y: '100%', opacity: 0 }}
                className={`relative w-full max-w-lg rounded-t-[2.5rem] sm:rounded-3xl shadow-2xl overflow-hidden h-[90vh] sm:h-[80vh] flex flex-col transition-colors duration-300 ${
                  theme === 'dark' ? 'bg-slate-900 border border-slate-800' : 'bg-white border border-slate-100'
                }`}
              >
                <div className={`h-1.5 w-12 rounded-full mx-auto my-3 shrink-0 transition-colors ${theme === 'dark' ? 'bg-slate-700' : 'bg-slate-300'}`} />

                <div className={`px-6 pb-4 border-b flex items-center justify-between shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800' : 'border-slate-100'
                }`}>
                  <div className="flex items-center gap-2.5">
                    <button
                      onClick={() => setShowAIChat(false)}
                      className={`p-2 rounded-xl border transition-all active:scale-95 ${
                        theme === 'dark' ? 'bg-slate-800 border-slate-700/60 text-slate-300 hover:bg-slate-750' : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                      }`}
                      aria-label="Back to Settings"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                    <div className="flex flex-col text-left">
                      <h2 className={`text-base font-extrabold flex items-center gap-1.5 transition-colors ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                        <Bot className="w-4 h-4 text-indigo-550 shrink-0" />
                        <span>Journal AI Assistant</span>
                      </h2>
                      <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none mt-1">Real-Time Trade Analyzer</p>
                    </div>
                  </div>

                  <button
                    onClick={() => setAiMessages([{ role: 'assistant', content: 'Hi there! I am your Trading Journal AI Assistant. Ask me anything about your trade stats, strategies, risk rules, or trading psychology.' }])}
                    className="text-[10px] font-bold text-indigo-500 hover:underline cursor-pointer"
                  >
                    Clear Chat
                  </button>
                </div>

                {/* Indian Language Selector Bar */}
                <div className={`px-6 py-2.5 border-b flex items-center justify-between gap-4 shrink-0 transition-all ${
                  theme === 'dark' ? 'border-slate-800 bg-slate-950/45' : 'border-slate-100 bg-slate-50/60'
                }`}>
                  <div className="flex items-center gap-2 shrink-0">
                    <Globe className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                    <span className={`text-[10px] font-extrabold uppercase tracking-wider ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                      Response Language
                    </span>
                  </div>
                  
                  <select
                    value={aiLanguage}
                    onChange={(e) => {
                      const newLang = e.target.value;
                      setAiLanguage(newLang);
                      const selectedLang = INDIAN_LANGUAGES.find(l => l.code === newLang);
                      const langLabel = selectedLang ? `${selectedLang.name} (${selectedLang.native})` : newLang;
                      setAiMessages(prev => [
                        ...prev,
                        { role: 'assistant', content: `🌐 Response language updated to ${langLabel}. Ask me anything, and I'll reply in this language!` }
                      ]);
                    }}
                    className={`text-[11px] font-black py-1 px-3 rounded-lg border outline-none cursor-pointer transition-all max-w-[180px] truncate ${
                      theme === 'dark' 
                        ? 'bg-slate-800 border-slate-700 text-slate-100 focus:border-indigo-500' 
                        : 'bg-white border-slate-200 text-slate-800 focus:border-indigo-650'
                    }`}
                  >
                    {INDIAN_LANGUAGES.map(lang => (
                      <option key={lang.code} value={lang.code}>
                        {lang.name} ({lang.native})
                      </option>
                    ))}
                  </select>
                </div>

                {/* Messages Panel */}
                <div className="flex-1 overflow-y-auto p-6 space-y-4 flex flex-col min-h-0 bg-slate-50/50 dark:bg-slate-950/20">
                  {aiMessages.map((msg, index) => (
                    <div 
                      key={index} 
                      className={`flex flex-col max-w-[85%] ${
                        msg.role === 'user' ? 'self-end items-end' : 'self-start items-start'
                      }`}
                    >
                      <div className={`p-3.5 rounded-2xl text-xs leading-relaxed ${
                        msg.role === 'user'
                          ? 'bg-indigo-600 text-white rounded-tr-none'
                          : theme === 'dark'
                            ? 'bg-slate-800 text-slate-100 rounded-tl-none border border-slate-750'
                            : 'bg-white text-slate-850 rounded-tl-none border border-slate-150'
                      }`}>
                        {msg.content}
                      </div>
                      <span className="text-[8px] font-black uppercase font-mono tracking-widest text-slate-400 mt-1 px-1">
                        {msg.role === 'user' ? (userName || 'You') : 'Journal AI'}
                      </span>
                    </div>
                  ))}
                  {aiSending && (
                    <div className="flex items-center gap-2 text-indigo-500 text-xs self-start p-3 bg-indigo-500/5 rounded-2xl dark:bg-indigo-500/10">
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span className="font-bold tracking-wider uppercase text-[9px]">AI is analyzing your trades...</span>
                    </div>
                  )}
                </div>

                {/* Quick prompts chips bar */}
                <div className={`p-3 border-t overflow-x-auto whitespace-nowrap flex gap-2 shrink-0 transition-colors ${
                  theme === 'dark' ? 'border-slate-800 bg-slate-900/60' : 'border-slate-100 bg-slate-50/40'
                }`}>
                  {[
                    "Analyze my trading stats",
                    "How can I manage psychological discipline?",
                    "Calculate position size for ₹10k risk",
                    "Critique my trading strategy"
                  ].map((chip) => (
                    <button
                      key={chip}
                      type="button"
                      disabled={aiSending}
                      onClick={() => handleSendAIMessage(chip)}
                      className={`py-1.5 px-3 rounded-full text-[10px] font-bold tracking-tight border cursor-pointer shrink-0 transition-all active:scale-[0.97] ${
                        theme === 'dark'
                          ? 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-750'
                          : 'bg-white border-slate-200 text-slate-650 hover:bg-slate-100'
                      }`}
                    >
                      {chip}
                    </button>
                  ))}
                </div>

                {/* Message Input Bar */}
                <form 
                  onSubmit={(e) => {
                    e.preventDefault();
                    handleSendAIMessage();
                  }}
                  className={`p-4 border-t shrink-0 flex gap-2 transition-colors ${
                    theme === 'dark' ? 'border-slate-850 bg-slate-900' : 'border-slate-100 bg-white'
                  }`}
                >
                  <input
                    type="text"
                    value={aiInput}
                    onChange={(e) => setAiInput(e.target.value)}
                    placeholder="Type message or trading query..."
                    disabled={aiSending}
                    className={`flex-1 border rounded-xl py-3 px-3.5 text-xs focus:ring-2 focus:ring-indigo-550 outline-none transition-all ${
                      theme === 'dark'
                        ? 'bg-slate-800 border-slate-705 text-white placeholder:text-slate-650'
                        : 'bg-slate-50 border-slate-200 text-slate-900 placeholder:text-slate-400'
                    }`}
                  />
                  <button
                    type="submit"
                    disabled={!aiInput.trim() || aiSending}
                    className={`p-3 rounded-xl flex items-center justify-center transition-all bg-indigo-600 text-white hover:bg-indigo-700 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer`}
                    title="Send Message"
                  >
                    <Send className="w-4 h-4" />
                  </button>
                </form>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

      {/* Android-style Navigation Bar Placeholder */}
      <nav className={`fixed bottom-0 left-0 right-0 h-16 border-t px-6 flex items-center justify-evenly sm:hidden transition-colors duration-300 ${
        theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'
      }`}>
         <div 
          onClick={() => {
            if (showCalendar) {
              setShowCalendar(false);
            } else {
              setShowCalendar(true);
              setShowHistory(false);
              setShowFunds(false);
              setShowSettings(false);
            }
          }}
          className={`flex flex-col items-center gap-1 ${showCalendar ? 'text-indigo-600' : (theme === 'dark' ? 'text-slate-500' : 'text-slate-400')}`}
        >
            <Calendar className="w-6 h-6" />
            <span className="text-[10px] font-bold uppercase tracking-wider">{t('calendar')}</span>
         </div>
         <div 
          onClick={() => {
            setShowHistory(false);
            setShowCalendar(false);
            setShowFunds(false);
            setShowSettings(false);
          }}
          className={`flex flex-col items-center gap-1 ${(!showHistory && !showCalendar && !showFunds && !showSettings) ? 'text-indigo-600' : (theme === 'dark' ? 'text-slate-500' : 'text-slate-400')}`}
        >
            <TrendingUp className="w-6 h-6" />
            <span className="text-[10px] font-bold uppercase tracking-wider">{t('journal')}</span>
         </div>
         <div 
          onClick={() => {
            setShowFunds(true);
            setShowSettings(false);
            setShowHistory(false);
            setShowCalendar(false);
          }}
          className={`flex flex-col items-center gap-1 transition-colors cursor-pointer ${showFunds ? 'text-indigo-600' : (theme === 'dark' ? 'text-slate-600 hover:text-slate-400' : 'text-slate-400 hover:text-slate-600')}`}
        >
            <Wallet className="w-6 h-6" />
            <span className="text-[10px] font-bold uppercase tracking-wider">{t('funds')}</span>
          </div>
          <div 
          onClick={() => {
            if (showHistory) {
              setShowHistory(false);
            } else {
              setShowHistory(true);
              setShowCalendar(false);
              setShowFunds(false);
              setShowSettings(false);
            }
          }}
          className={`flex flex-col items-center gap-1 ${showHistory ? 'text-indigo-600' : (theme === 'dark' ? 'text-slate-500' : 'text-slate-400')}`}
        >
            <History className="w-6 h-6" />
            <span className="text-[10px] font-bold uppercase tracking-wider">{t('analytics')}</span>
         </div>
          <div 
          onClick={() => {
            setShowSettings(true);
            setShowFunds(false);
            setShowHistory(false);
            setShowCalendar(false);
          }}
          className={`flex flex-col items-center gap-1 transition-colors cursor-pointer ${showSettings ? 'text-indigo-600' : (theme === 'dark' ? 'text-slate-600 hover:text-slate-400' : 'text-slate-400 hover:text-slate-600')}`}
        >
            <Settings className="w-6 h-6" />
            <span className="text-[10px] font-bold uppercase tracking-wider">{t('settings')}</span>
          </div>
      </nav>
    </div>
    </>
  );
}
