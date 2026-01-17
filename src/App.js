import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Calendar, Clock, Users, Plus, Edit2, Trash2, Save, X, Upload, Download, Search, Printer, AlertCircle, Moon, Sun, Globe, BarChart3, Award, RefreshCw, Repeat, Zap, CheckCircle, Settings, TrendingUp, Shield, Gift, FileText } from 'lucide-react';
import * as XLSX from 'xlsx';
import { HebrewCalendar, HDate, Event } from 'hebcal';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// הגדרת סוגי תורנויות עם נקודות בסיס
const DUTY_TYPES = {
  'מטוס': { name: 'מטוס', basePoints: 2, description: 'תורנות מטוס', color: '#3b82f6' },
  'גלגלת': { name: 'גלגלת', basePoints: 1, description: 'תורנות גלגלת', color: '#10b981' },
  'ירושלים': { name: 'ירושלים', basePoints: 1.5, description: 'תורנות ירושלים', color: '#f59e0b' },
};

// הגדרת סוגי ימים ובונוס נקודות
const DATE_TYPES = {
  'חול': { name: 'חול', bonus: 0, description: 'יום חול רגיל', color: '#6b7280' },
  'שבת': { name: 'שבת', bonus: 1, description: 'יום שבת', color: '#8b5cf6' },
  'חג': { name: 'חג', bonus: 3, description: 'חג', color: '#ef4444' },
};

// מקדמי סטטוס/תפקיד (ככל שהתפקיד גבוה יותר, המקדם נמוך יותר)
const DEFAULT_STATUS_MULTIPLIERS = {
  'חייל רגיל': 1.0,
  'מפקד': 0.7,
  'קצין': 0.5,
};

// פטורים ברירת מחדל - כל פטור מגדיר מאילו סוגי תורנויות הוא פוטר
const DEFAULT_EXEMPTIONS = {
  'פטור אבק': { name: 'פטור אבק', exemptFromDutyTypes: ['מטוס'], description: 'פטור מתורנות מטוס בגלל אלרגיה לאבק' },
  'פטור רפואי': { name: 'פטור רפואי', exemptFromDutyTypes: [], description: 'פטור רפואי כללי' },
};

// חגים ברירת מחדל - כל חג עם משקל (קושי) שלו ומזהה לוח עברי
const DEFAULT_HOLIDAYS = {
  "פסח א'": {
    name: "פסח א'",
    weight: 3,
    description: "פסח יום ראשון-שני - 2 ימים",
    hebrewId: 'Pesach',
    durationDays: 2,
    dayOffset: 0
  },
  "פסח ב'": {
    name: "פסח ב'",
    weight: 3,
    description: "פסח יום שביעי-שמיני - 2 ימים",
    hebrewId: 'Pesach',
    durationDays: 2,
    dayOffset: 6
  },
  'סוכות': {
    name: 'סוכות',
    weight: 3,
    description: 'חג הסוכות - 7 ימים',
    hebrewId: 'Sukkot',
    durationDays: 7
  },
  'ראש השנה': {
    name: 'ראש השנה',
    weight: 2,
    description: 'ראש השנה - 2 ימים',
    hebrewId: 'Rosh Hashana',
    durationDays: 2
  },
  'יום כיפור': {
    name: 'יום כיפור',
    weight: 2.5,
    description: 'יום כיפור - יום אחד קשה',
    hebrewId: 'Yom Kippur',
    durationDays: 1
  },
  'שבועות': {
    name: 'שבועות',
    weight: 1.5,
    description: 'חג השבועות - יום אחד',
    hebrewId: 'Shavuot',
    durationDays: 1
  },
  'שמיני עצרת': {
    name: 'שמיני עצרת',
    weight: 1,
    description: 'שמיני עצרת - יום אחד',
    hebrewId: 'Shmini Atzeret',
    durationDays: 1
  },
};

const translations = {
  en: {
    appTitle: 'Shift Manager', subtitle: 'Advanced Duty Scheduling System',
    employees: 'Employees', schedule: 'Schedule', reports: 'Reports',
    employeeManagement: 'Employee Management', reportsTitle: 'Reports & Statistics',
    template: 'Template', import: 'Import', add: 'Add',
    searchPlaceholder: 'Search by name or ID...', allDepartments: 'All Departments',
    allStatuses: 'All Statuses', newEmployee: 'New Employee', name: 'Name',
    personalNumber: 'Personal Number', sex: 'Sex', status: 'Status',
    department: 'Department', select: 'Select', male: 'Male', female: 'Female',
    other: 'Other', addEmployee: 'Add Employee', noEmployees: 'No employees yet',
    importOrAdd: 'Import from Excel or add manually', noMatching: 'No matching employees',
    adjustFilters: 'Try adjusting filters', shifts: 'Shifts', actions: 'Actions',
    shiftSchedule: 'Shift Schedule', list: 'List', week: 'Week', month: 'Month',
    export: 'Export', print: 'Print', addShift: 'Add Shift',
    conflictDetected: 'shift conflict(s) detected! (Less than 21 days gap)', addEmployeesFirst: 'Add employees first',
    needEmployees: 'You need to add employees before creating shifts', newShift: 'New Shift',
    employee: 'Employee', date: 'Date',
    shiftType: 'Shift Type', shiftTypePlaceholder: 'Select shift type',
    noShifts: 'No shifts scheduled', clickToAdd: 'Click "Add Shift" to create',
    previousWeek: '← Previous', nextWeek: 'Next →', weekOf: 'Week of',
    sun: 'Sun', mon: 'Mon', tue: 'Tue', wed: 'Wed', thu: 'Thu', fri: 'Fri', sat: 'Sat',
    imported: 'Successfully imported', employeesText: 'employees!',
    errorReading: 'Error reading file', conflictWarning: '⚠️ Conflict! Employee has another shift within 21 days',
    noDept: 'No Dept', noDepartment: 'No Department', selectEmployee: 'Select Employee',
    deptPlaceholder: 'e.g., Sales, IT', totalShifts: 'Total Shifts',
    weekdayShifts: 'Weekday Shifts', shabbatShifts: 'Shabbat Shifts',
    justicePoints: 'Justice Points', points: 'Points', count: 'Count',
    noData: 'No data available', employeeStats: 'Employee Statistics',
    editEmployee: 'Edit Employee', deleteConfirm: 'Are you sure?',
    deleteEmployeeMsg: 'Delete employee', deleteShiftMsg: 'Delete shift',
    cancel: 'Cancel', delete: 'Delete', repeating: 'Repeating',
    repeatType: 'Repeat Type', none: 'None', daily: 'Daily', weekly: 'Weekly',
    monthly: 'Monthly', repeatUntil: 'Repeat Until', filterByEmployee: 'Filter by Employee',
    filterByShiftType: 'Filter by Shift Type', allEmployees: 'All Employees',
    allShiftTypes: 'All Types', swapShift: 'Swap Shift', swapWith: 'Swap With',
    swapShiftTitle: 'Swap Shifts', selectShiftToSwap: 'Select shift to swap',
    importShifts: 'Import Shifts', shiftsTemplate: 'Shifts Template',
    saving: 'Saving...', saved: 'Saved', autoDistribute: 'Auto Distribute',
    unassigned: 'Unassigned', optional: 'Optional', noEmployee: 'No Employee',
    dutyType: 'Duty Type', dateType: 'Date Type', statusMultiplier: 'Status Multiplier',
    manageStatuses: 'Manage Statuses', basePoints: 'Base Points', bonusPoints: 'Bonus',
    finalPoints: 'Final Points', multiplier: 'Multiplier', settingsTitle: 'Settings',
    manageStatusMultipliers: 'Manage Status Multipliers', addNewStatus: 'Add New Status',
    statusName: 'Status Name', manualOverride: 'Manual Override',
    resetSystem: 'Reset System', resetConfirm: 'Reset All Data?',
    resetWarning: 'This will delete all employees, shifts, and settings. This action cannot be undone!',
    resetButton: 'Reset Everything', year: 'Year', allYears: 'All Years',
    managePoints: 'Manage Points', dutyTypePoints: 'Duty Type Points', dateTypePoints: 'Date Type Points',
    pointsSettings: 'Points Settings', dutyTypeName: 'Type Name', dateTypeName: 'Type Name',
    dutyTypeDescription: 'Description', addDutyType: 'Add Type', addDateType: 'Add Type',
    exemptions: 'Exemptions', manageExemptions: 'Manage Exemptions', exemptionSettings: 'Exemption Settings',
    exemptionName: 'Exemption Name', exemptFrom: 'Exempt From', addExemption: 'Add Exemption',
    exemptionDescription: 'Description', noExemptions: 'No Exemptions',
    holidayJustice: 'Holiday Justice', manageHolidays: 'Manage Holidays', holidayName: 'Holiday Name',
    holidayWeight: 'Holiday Weight', addHoliday: 'Add Holiday', holidayDescription: 'Description',
    noHolidays: 'No Holidays', holidayHistory: 'Holiday History', lastYear: 'Last Year',
    employeeHolidayHistory: 'Employee Holiday History', assignedHoliday: 'Assigned Holiday',
    dryRun: 'Dry Run', previewDistribution: 'Preview Distribution', acceptDistribution: 'Accept & Apply',
    distributionPreview: 'Distribution Preview', distributed: 'Distributed', unassigned: 'Unassigned',
    willBeAssigned: 'will be assigned to', previewBeforeApply: 'Preview the automatic distribution before applying',
    exportPDF: 'Export PDF', pdfReport: 'Shift Schedule Report', generatedOn: 'Generated on',
    yearlyExport: 'Yearly Export', yearlyReport: 'Yearly Report', totalShiftsInYear: 'Total Shifts in Year'
  },
  he: {
    appTitle: 'מנהל תורנויות', subtitle: 'מערכת מתקדמת לניהול תורנויות',
    employees: 'עובדים', schedule: 'לוח תורנויות', reports: 'דוחות',
    employeeManagement: 'ניהול עובדים', reportsTitle: 'דוחות וסטטיסטיקה',
    template: 'תבנית', import: 'ייבוא', add: 'הוסף',
    searchPlaceholder: 'חיפוש לפי שם...', allDepartments: 'כל המחלקות',
    allStatuses: 'כל הסטטוסים', newEmployee: 'עובד חדש', name: 'שם',
    personalNumber: 'מספר אישי', sex: 'מין', status: 'סטטוס/תפקיד',
    department: 'מחלקה', select: 'בחר', male: 'זכר', female: 'נקבה',
    other: 'אחר', addEmployee: 'הוסף עובד', noEmployees: 'אין עובדים',
    importOrAdd: 'ייבא מאקסל', noMatching: 'אין תואמים',
    adjustFilters: 'נסה מסננים', shifts: 'תורנויות', actions: 'פעולות',
    shiftSchedule: 'לוח תורנויות', list: 'רשימה', week: 'שבוע', month: 'חודש',
    export: 'ייצוא', print: 'הדפסה', addShift: 'הוסף תורנות',
    conflictDetected: 'התנגשויות זוהו! (פחות מ-21 יום בין תורנויות)', addEmployeesFirst: 'הוסף עובדים תחילה',
    needEmployees: 'הוסף עובדים לפני תורנויות', newShift: 'תורנות חדשה',
    employee: 'עובד', date: 'תאריך',
    shiftType: 'סוג תורנות', shiftTypePlaceholder: 'בחר סוג תורנות',
    noShifts: 'אין תורנויות', clickToAdd: 'לחץ "הוסף תורנות"',
    previousWeek: 'קודם ←', nextWeek: '→ הבא', weekOf: 'שבוע של',
    sun: "א'", mon: "ב'", tue: "ג'", wed: "ד'", thu: "ה'", fri: "ו'", sat: "ש'",
    imported: 'יובאו', employeesText: 'עובדים!',
    errorReading: 'שגיאה', conflictWarning: '⚠️ התנגשות! העובד כבר משובץ לתורנות בתוך 21 יום',
    noDept: 'ללא מחלקה', noDepartment: 'ללא מחלקה', selectEmployee: 'בחר עובד',
    deptPlaceholder: 'מכירות, IT', totalShifts: 'סה"כ תורנויות',
    weekdayShifts: 'תורנויות חול', shabbatShifts: 'תורנויות שבת',
    justicePoints: 'ניקוד צדק', points: 'ניקוד', count: 'כמות',
    noData: 'אין נתונים זמינים', employeeStats: 'סטטיסטיקות עובדים',
    editEmployee: 'ערוך עובד', deleteConfirm: 'האם אתה בטוח?',
    deleteEmployeeMsg: 'למחוק את', deleteShiftMsg: 'למחוק תורנות זו',
    cancel: 'ביטול', delete: 'מחק', repeating: 'חוזרת',
    repeatType: 'סוג חזרה', none: 'ללא', daily: 'יומי', weekly: 'שבועי',
    monthly: 'חודשי', repeatUntil: 'חזור עד', filterByEmployee: 'סינון לפי עובד',
    filterByShiftType: 'סינון לפי סוג', allEmployees: 'כל העובדים',
    allShiftTypes: 'כל הסוגים', swapShift: 'החלף תורנות', swapWith: 'החלף עם',
    swapShiftTitle: 'החלפת תורנויות', selectShiftToSwap: 'בחר תורנות להחלפה',
    importShifts: 'ייבוא תורנויות', shiftsTemplate: 'תבנית תורנויות',
    saving: 'שומר...', saved: '✓ נשמר', autoDistribute: 'חלוקה אוטומטית',
    unassigned: 'לא משובץ', optional: 'אופציונלי', noEmployee: 'ללא עובד',
    dutyType: 'סוג תורנות', dateType: 'סוג יום', statusMultiplier: 'מקדם תפקיד',
    manageStatuses: 'ניהול תפקידים', basePoints: 'נקודות בסיס', bonusPoints: 'בונוס',
    finalPoints: 'נקודות סופיות', multiplier: 'מקדם', settingsTitle: 'הגדרות',
    manageStatusMultipliers: 'ניהול מקדמי תפקידים', addNewStatus: 'הוסף תפקיד חדש',
    statusName: 'שם תפקיד', manualOverride: 'עקיפה ידנית',
    resetSystem: 'איפוס מערכת', resetConfirm: 'לאפס את כל הנתונים?',
    resetWarning: 'פעולה זו תמחק את כל העובדים, התורנויות וההגדרות. לא ניתן לבטל פעולה זו!',
    resetButton: 'אפס הכל', year: 'שנה', allYears: 'כל השנים',
    managePoints: 'ניהול נקודות', dutyTypePoints: 'נקודות סוגי תורנויות', dateTypePoints: 'נקודות סוגי ימים',
    pointsSettings: 'הגדרות נקודות', dutyTypeName: 'שם סוג', dateTypeName: 'שם סוג',
    dutyTypeDescription: 'תיאור', addDutyType: 'הוסף סוג', addDateType: 'הוסף סוג',
    exemptions: 'פטורים', manageExemptions: 'ניהול פטורים', exemptionSettings: 'הגדרות פטורים',
    exemptionName: 'שם פטור', exemptFrom: 'פוטר מ', addExemption: 'הוסף פטור',
    exemptionDescription: 'תיאור', noExemptions: 'ללא פטורים',
    holidayJustice: 'צדק חגים', manageHolidays: 'ניהול חגים', holidayName: 'שם חג',
    holidayWeight: 'משקל חג', addHoliday: 'הוסף חג', holidayDescription: 'תיאור',
    noHolidays: 'ללא חגים', holidayHistory: 'היסטוריית חגים', lastYear: 'שנה שעברה',
    employeeHolidayHistory: 'היסטוריית חגים של עובדים', assignedHoliday: 'חג משובץ',
    dryRun: 'ניסיון יבש', previewDistribution: 'תצוגה מקדימה של החלוקה', acceptDistribution: 'אישור וביצוע',
    distributionPreview: 'תצוגה מקדימה של החלוקה', distributed: 'חולק', unassigned: 'לא משובץ',
    willBeAssigned: 'ישובץ ל', previewBeforeApply: 'צפה בחלוקה אוטומטית לפני ביצוע',
    exportPDF: 'ייצוא PDF', pdfReport: 'דוח לוח תורנויות', generatedOn: 'נוצר בתאריך',
    yearlyExport: 'ייצוא שנתי', yearlyReport: 'דוח שנתי', totalShiftsInYear: 'סה"כ תורנויות בשנה'
  }
};

export default function App() {
  const [employees, setEmployees] = useState([]);
  const [shifts, setShifts] = useState([]);
  const [activeTab, setActiveTab] = useState('employees');
  const [showAddEmployee, setShowAddEmployee] = useState(false);
  const [showAddShift, setShowAddShift] = useState(false);
  const [editingShift, setEditingShift] = useState(null);
  const [editingEmployee, setEditingEmployee] = useState(null);
  const [language, setLanguage] = useState('he');
  const [darkMode, setDarkMode] = useState(false);
  const [uploadMessage, setUploadMessage] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [filterDepartment, setFilterDepartment] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [calendarView, setCalendarView] = useState('list');
  const [currentWeekStart, setCurrentWeekStart] = useState(new Date());
  const [statusMultipliers, setStatusMultipliers] = useState(DEFAULT_STATUS_MULTIPLIERS);
  const [showStatusSettings, setShowStatusSettings] = useState(false);
  const [newStatusName, setNewStatusName] = useState('');
  const [newStatusMultiplier, setNewStatusMultiplier] = useState(1.0);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [resetConfirm, setResetConfirm] = useState(false);
  const [filterEmployee, setFilterEmployee] = useState('');
  const [filterShiftType, setFilterShiftType] = useState('');
  const [swappingShift, setSwappingShift] = useState(null);
  const [saveStatus, setSaveStatus] = useState('saved');
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
  const [dutyTypePoints, setDutyTypePoints] = useState(DUTY_TYPES);
  const [dateTypePoints, setDateTypePoints] = useState(DATE_TYPES);
  const [showPointsSettings, setShowPointsSettings] = useState(false);
  const [newDutyType, setNewDutyType] = useState({ name: '', basePoints: 1, description: '' });
  const [newDateType, setNewDateType] = useState({ name: '', bonus: 0, description: '' });
  const [exemptionTypes, setExemptionTypes] = useState(DEFAULT_EXEMPTIONS);
  const [showExemptionSettings, setShowExemptionSettings] = useState(false);
  const [newExemption, setNewExemption] = useState({ name: '', exemptFromDutyTypes: [], description: '' });

  const [holidays, setHolidays] = useState(DEFAULT_HOLIDAYS);
  const [showHolidaySettings, setShowHolidaySettings] = useState(false);
  const [newHoliday, setNewHoliday] = useState({ name: '', weight: 1, description: '' });
  const [holidayHistory, setHolidayHistory] = useState({});
  const [dryRunPreview, setDryRunPreview] = useState(null);

  const [newEmployee, setNewEmployee] = useState({
    name: '', personalNumber: '', sex: '', status: '', department: '', exemptions: []
  });

  const [newShift, setNewShift] = useState({
    employeeId: '', startDate: '', endDate: '',
    dutyType: '', dateType: '', manualPoints: null,
    holidayName: ''
  });

  const t = translations[language];

  // מטמון לתאריכי חגים ואירועי לוח עברי
  const holidayDatesCache = useRef({});
  const hebrewCalendarCache = useRef({});

  // פונקציה לזיהוי אוטומטי של יום בשבוע
  const getDayOfWeek = (dateStr) => {
    const date = new Date(dateStr);
    return date.getDay(); // 0 = Sunday, 6 = Saturday
  };

  // חישוב אוטומטי של תאריך סיום לפי סוג התורנות ותאריך ההתחלה
  const calculateEndDate = (startDate, dateType) => {
    if (!startDate || !dateType) return '';

    const start = new Date(startDate);

    if (dateType === 'חול') {
      // חול: בדרך כלל שני→חמישי (4 ימים)
      // אם מתחיל ביום שני (1), סיום ביום חמישי (4) = +3 ימים
      const end = new Date(start);
      end.setDate(start.getDate() + 3); // 4 ימים כולל יום ההתחלה = +3
      return end.toISOString().split('T')[0];
    } else if (dateType === 'שבת') {
      // שבת: חמישי→שני (4 ימים, עובר את השבת)
      // אם מתחיל ביום חמישי (4), סיום ביום שני (1) = +4 ימים
      const end = new Date(start);
      end.setDate(start.getDate() + 4); // עובר דרך שישי+שבת+ראשון
      return end.toISOString().split('T')[0];
    } else if (dateType === 'חג') {
      // חג: נניח אותו משך כמו שבת
      const end = new Date(start);
      end.setDate(start.getDate() + 3);
      return end.toISOString().split('T')[0];
    }

    return '';
  };

  // קבלת רשימת שנים זמינות מהתורנויות
  const availableYears = useMemo(() => {
    const years = new Set();
    shifts.forEach(shift => {
      if (shift.startDate) {
        const year = new Date(shift.startDate).getFullYear();
        years.add(year);
      }
    });
    return Array.from(years).sort((a, b) => b - a); // מיון יורד (החדש ביותר ראשון)
  }, [shifts]);

  // פונקציה לחישוב נקודות תורנות
  const calculateShiftPoints = React.useCallback((shift, employeeStatus) => {
    // אם יש עקיפה ידנית, השתמש בה
    if (shift.manualPoints !== null && shift.manualPoints !== undefined) {
      return parseFloat(shift.manualPoints);
    }

    const dutyType = dutyTypePoints[shift.dutyType] || dutyTypePoints['גלגלת'];
    const dateType = dateTypePoints[shift.dateType] || dateTypePoints['חול'];
    const statusMultiplier = statusMultipliers[employeeStatus] || 1.0;

    let basePoints = dutyType.basePoints;
    let bonusPoints = dateType.bonus;

    // זיהוי אוטומטי: אם זה חג שחל בשבת, הוסף בונוס נוסף
    const dayOfWeek = getDayOfWeek(shift.date);
    const isSaturday = dayOfWeek === 6;
    const isHoliday = shift.dateType === 'חג';

    if (isHoliday && isSaturday) {
      bonusPoints += dateTypePoints['שבת'].bonus; // הוסף גם את בונוס השבת
    }

    // נוסחה: (נקודות בסיס + בונוס) × מקדם סטטוס
    const totalPoints = (basePoints + bonusPoints) * statusMultiplier;

    return totalPoints;
  }, [statusMultipliers, dutyTypePoints, dateTypePoints]);

  // Load & Save
  useEffect(() => {
    try {
      const saved = localStorage.getItem('employees');
      if (saved) setEmployees(JSON.parse(saved));
      const savedShifts = localStorage.getItem('shifts');
      if (savedShifts) setShifts(JSON.parse(savedShifts));
      const savedLang = localStorage.getItem('language');
      if (savedLang) setLanguage(savedLang);
      const savedDark = localStorage.getItem('darkMode');
      if (savedDark) setDarkMode(savedDark === 'true');
      const savedYear = localStorage.getItem('selectedYear');
      if (savedYear) setSelectedYear(parseInt(savedYear));
      const savedMultipliers = localStorage.getItem('statusMultipliers');
      if (savedMultipliers) {
        setStatusMultipliers(JSON.parse(savedMultipliers));
      }
      const savedDutyPoints = localStorage.getItem('dutyTypePoints');
      if (savedDutyPoints) {
        setDutyTypePoints(JSON.parse(savedDutyPoints));
      }
      const savedDatePoints = localStorage.getItem('dateTypePoints');
      if (savedDatePoints) {
        setDateTypePoints(JSON.parse(savedDatePoints));
      }
      const savedExemptions = localStorage.getItem('exemptionTypes');
      if (savedExemptions) {
        setExemptionTypes(JSON.parse(savedExemptions));
      }
      const savedHolidays = localStorage.getItem('holidays');
      if (savedHolidays) {
        setHolidays(JSON.parse(savedHolidays));
      }
      const savedHolidayHistory = localStorage.getItem('holidayHistory');
      if (savedHolidayHistory) {
        setHolidayHistory(JSON.parse(savedHolidayHistory));
      }
    } catch (e) {
      console.error('Error loading data:', e);
    }
  }, []);

  // Auto-save with debounce
  const saveTimeoutRef = useRef(null);
  const isFirstRenderEmployees = useRef(true);

  useEffect(() => {
    if (isFirstRenderEmployees.current) {
      isFirstRenderEmployees.current = false;
      return;
    }

    setSaveStatus('saving');
    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    saveTimeoutRef.current = setTimeout(() => {
      localStorage.setItem('employees', JSON.stringify(employees));
      setSaveStatus('saved');
      setTimeout(() => setSaveStatus(''), 2000);
    }, 1000);
  }, [employees]);

  const isFirstRenderShifts = useRef(true);
  useEffect(() => {
    if (isFirstRenderShifts.current) {
      isFirstRenderShifts.current = false;
      return;
    }

    setSaveStatus('saving');
    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    saveTimeoutRef.current = setTimeout(() => {
      localStorage.setItem('shifts', JSON.stringify(shifts));
      setSaveStatus('saved');
      setTimeout(() => setSaveStatus(''), 2000);
    }, 1000);
  }, [shifts]);

  useEffect(() => {
    localStorage.setItem('language', language);
  }, [language]);

  useEffect(() => {
    localStorage.setItem('darkMode', String(darkMode));
  }, [darkMode]);

  useEffect(() => {
    if (selectedYear) {
      localStorage.setItem('selectedYear', String(selectedYear));
    } else {
      localStorage.removeItem('selectedYear');
    }
  }, [selectedYear]);

  const isFirstRenderMultipliers = useRef(true);
  useEffect(() => {
    if (isFirstRenderMultipliers.current) {
      isFirstRenderMultipliers.current = false;
      return;
    }
    localStorage.setItem('statusMultipliers', JSON.stringify(statusMultipliers));
  }, [statusMultipliers]);

  const isFirstRenderDutyPoints = useRef(true);
  useEffect(() => {
    if (isFirstRenderDutyPoints.current) {
      isFirstRenderDutyPoints.current = false;
      return;
    }
    localStorage.setItem('dutyTypePoints', JSON.stringify(dutyTypePoints));
  }, [dutyTypePoints]);

  const isFirstRenderDatePoints = useRef(true);
  useEffect(() => {
    if (isFirstRenderDatePoints.current) {
      isFirstRenderDatePoints.current = false;
      return;
    }
    localStorage.setItem('dateTypePoints', JSON.stringify(dateTypePoints));
  }, [dateTypePoints]);

  const isFirstRenderExemptions = useRef(true);
  useEffect(() => {
    if (isFirstRenderExemptions.current) {
      isFirstRenderExemptions.current = false;
      return;
    }
    localStorage.setItem('exemptionTypes', JSON.stringify(exemptionTypes));
  }, [exemptionTypes]);

  const isFirstRenderHolidays = useRef(true);
  useEffect(() => {
    if (isFirstRenderHolidays.current) {
      isFirstRenderHolidays.current = false;
      return;
    }
    localStorage.setItem('holidays', JSON.stringify(holidays));
  }, [holidays]);

  const isFirstRenderHolidayHistory = useRef(true);
  useEffect(() => {
    if (isFirstRenderHolidayHistory.current) {
      isFirstRenderHolidayHistory.current = false;
      return;
    }
    localStorage.setItem('holidayHistory', JSON.stringify(holidayHistory));
  }, [holidayHistory]);

  const departments = useMemo(() => [...new Set(employees.map(e => e.department).filter(Boolean))], [employees]);
  const statuses = useMemo(() => [...new Set(employees.map(e => e.status).filter(Boolean))], [employees]);

  const filteredEmployees = useMemo(() => {
    return employees.filter(emp => {
      const matchesSearch = emp.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           emp.personalNumber.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesDept = !filterDepartment || emp.department === filterDepartment;
      const matchesStatus = !filterStatus || emp.status === filterStatus;
      return matchesSearch && matchesDept && matchesStatus;
    });
  }, [employees, searchTerm, filterDepartment, filterStatus]);

  const filteredShifts = useMemo(() => {
    return shifts.filter(shift => {
      const matchesEmployee = !filterEmployee || shift.employeeId === parseInt(filterEmployee);
      const matchesShiftType = !filterShiftType || shift.dutyType === filterShiftType;

      // סינון לפי שנה - אם selectedYear הוא null, מציג הכל
      let matchesYear = true;
      if (selectedYear && shift.startDate) {
        const shiftYear = new Date(shift.startDate).getFullYear();
        matchesYear = shiftYear === selectedYear;
      }

      return matchesEmployee && matchesShiftType && matchesYear;
    });
  }, [shifts, filterEmployee, filterShiftType, selectedYear]);

  const employeeStats = useMemo(() => {
    return employees.map(emp => {
      // סינון תורנויות לפי עובד ושנה
      const empShifts = shifts.filter(s => {
        if (s.employeeId !== emp.id) return false;

        // סינון לפי שנה נבחרת
        if (selectedYear && s.startDate) {
          const shiftYear = new Date(s.startDate).getFullYear();
          return shiftYear === selectedYear;
        }

        return true;
      });
      const dutyTypeCounts = {};
      const dateTypeCounts = {};
      let totalPoints = 0;

      Object.keys(dutyTypePoints).forEach(type => {
        dutyTypeCounts[type] = 0;
      });

      Object.keys(dateTypePoints).forEach(type => {
        dateTypeCounts[type] = 0;
      });

      empShifts.forEach(shift => {
        const dutyType = shift.dutyType || 'גלגלת';
        const dateType = shift.dateType || 'חול';

        if (dutyTypeCounts[dutyType] !== undefined) {
          dutyTypeCounts[dutyType]++;
        }

        if (dateTypeCounts[dateType] !== undefined) {
          dateTypeCounts[dateType]++;
        }

        totalPoints += calculateShiftPoints(shift, emp.status);
      });

      return {
        ...emp,
        totalShifts: empShifts.length,
        dutyTypeCounts,
        dateTypeCounts,
        justicePoints: totalPoints
      };
    });
  }, [employees, shifts, calculateShiftPoints, selectedYear, dutyTypePoints, dateTypePoints]);

  // בדיקת רווח של 21 יום בין תורנויות של אותו עובד
  const hasConflict = (employeeId, startDate, endDate, excludeShiftId = null) => {
    if (!employeeId || !startDate) return false;

    const currentStart = new Date(startDate);
    const currentEnd = endDate ? new Date(endDate) : currentStart;
    const employeeShifts = shifts.filter(s =>
      s.employeeId === employeeId && s.id !== excludeShiftId
    );

    return employeeShifts.some(shift => {
      const shiftStart = new Date(shift.startDate);
      const shiftEnd = shift.endDate ? new Date(shift.endDate) : shiftStart;

      // בדיקת רווח: המרחק בין סוף תורנות אחת לתחילת השנייה צריך להיות לפחות 21 יום
      const gapAfter = Math.abs((shiftStart - currentEnd) / (1000 * 60 * 60 * 24));
      const gapBefore = Math.abs((currentStart - shiftEnd) / (1000 * 60 * 60 * 24));
      const minGap = Math.min(gapAfter, gapBefore);

      return minGap < 21;
    });
  };

  const conflictingShifts = useMemo(() => {
    const conflicts = new Set();
    shifts.forEach(shift => {
      if (!shift.employeeId || !shift.startDate) return;

      const currentStart = new Date(shift.startDate);
      const currentEnd = shift.endDate ? new Date(shift.endDate) : currentStart;
      const employeeShifts = shifts.filter(s =>
        s.employeeId === shift.employeeId && s.id !== shift.id
      );

      const isConflict = employeeShifts.some(s => {
        const shiftStart = new Date(s.startDate);
        const shiftEnd = s.endDate ? new Date(s.endDate) : shiftStart;

        const gapAfter = Math.abs((shiftStart - currentEnd) / (1000 * 60 * 60 * 24));
        const gapBefore = Math.abs((currentStart - shiftEnd) / (1000 * 60 * 60 * 24));
        const minGap = Math.min(gapAfter, gapBefore);

        return minGap < 21;
      });

      if (isConflict) {
        conflicts.add(shift.id);
      }
    });
    return conflicts;
  }, [shifts]);

  const generateRepeatingShifts = (baseShift, repeatType, repeatUntil) => {
    const shifts = [];
    const startDate = new Date(baseShift.startDate);
    const endDate = new Date(repeatUntil);
    let currentDate = new Date(startDate);
    let id = Date.now();

    while (currentDate <= endDate) {
      const startDateStr = currentDate.toISOString().split('T')[0];
      const endDateStr = baseShift.endDate ?
        new Date(new Date(currentDate).setDate(currentDate.getDate() +
          (new Date(baseShift.endDate) - new Date(baseShift.startDate)) / (1000 * 60 * 60 * 24)
        )).toISOString().split('T')[0]
        : startDateStr;

      shifts.push({
        id: id++,
        employeeId: baseShift.employeeId,
        startDate: startDateStr,
        endDate: endDateStr,
        dutyType: baseShift.dutyType,
        dateType: baseShift.dateType,
        manualPoints: baseShift.manualPoints,
        repeatType: repeatType
      });

      if (repeatType === 'daily') {
        currentDate.setDate(currentDate.getDate() + 1);
      } else if (repeatType === 'weekly') {
        currentDate.setDate(currentDate.getDate() + 7);
      } else if (repeatType === 'monthly') {
        currentDate.setMonth(currentDate.getMonth() + 1);
      }
    }

    return shifts;
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);
        const imported = json.map((r, i) => {
          const exemptionsStr = r.Exemptions || r.exemptions || r['פטורים'] || '';
          const exemptionsArray = exemptionsStr ? exemptionsStr.split(',').map(ex => ex.trim()).filter(ex => ex && exemptionTypes[ex]) : [];

          return {
            id: Date.now() + i,
            name: r.Name || r.name || r['שם'] || '',
            personalNumber: r['Personal Number'] || r.personalNumber || r['מספר אישי'] || '',
            sex: r.Sex || r.sex || r['מין'] || '',
            status: r.Status || r.status || r['סטטוס'] || r['תפקיד'] || '',
            department: r.Department || r.department || r['מחלקה'] || '',
            exemptions: exemptionsArray
          };
        });
        setEmployees([...employees, ...imported]);
        setUploadMessage(`${t.imported} ${imported.length} ${t.employeesText}`);
        setTimeout(() => setUploadMessage(''), 3000);
      } catch (error) {
        setUploadMessage(t.errorReading);
        setTimeout(() => setUploadMessage(''), 3000);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleShiftsFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);

        const imported = json.map((r, i) => {
          const empId = r['Employee ID'] || r.employeeId || r['מזהה עובד'] || null;
          const empName = r['Employee Name'] || r.employeeName || r['שם עובד'];

          let finalEmpId = null;
          if (empId) {
            finalEmpId = parseInt(empId);
          } else if (empName) {
            const foundEmp = employees.find(e => e.name === empName);
            if (foundEmp) finalEmpId = foundEmp.id;
          }

          const startDate = r['Start Date'] || r.startDate || r['תאריך התחלה'] || r.Date || r.date || r['תאריך'] || '';
          const dateType = r['Date Type'] || r.dateType || r['סוג יום'] || 'חול';
          const endDate = r['End Date'] || r.endDate || r['תאריך סיום'] || '';

          // חישוב אוטומטי של תאריך סיום אם לא צוין
          const calculatedEndDate = endDate || calculateEndDate(startDate, dateType);

          return {
            id: Date.now() + i,
            employeeId: finalEmpId,
            startDate: startDate,
            endDate: calculatedEndDate,
            dutyType: r['Duty Type'] || r.dutyType || r['סוג תורנות'] || 'גלגלת',
            dateType: dateType,
            manualPoints: r['Manual Points'] || r.manualPoints || null,
            repeatType: 'none'
          };
        });

        setShifts([...shifts, ...imported]);
        setUploadMessage(`✓ יובאו ${imported.length} תורנויות!`);
        setTimeout(() => setUploadMessage(''), 3000);
      } catch (error) {
        setUploadMessage('❌ שגיאה בקריאת קובץ התורנויות');
        setTimeout(() => setUploadMessage(''), 3000);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleDownloadTemplate = () => {
    const template = [{ Name: 'John', 'Personal Number': '123', Sex: 'Male', Status: 'חייל רגיל', Department: 'Sales', Exemptions: 'פטור אבק, פטור רפואי' }];
    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Employees');
    XLSX.writeFile(wb, 'employees_template.xlsx');
  };

  const handleDownloadShiftsTemplate = () => {
    const template = [{
      'Employee ID': employees[0]?.id || '',
      'Employee Name': employees[0]?.name || 'John Doe',
      'Start Date': '2024-01-15',
      'End Date': '2024-01-18',
      'Duty Type': 'מטוס',
      'Date Type': 'חול',
      'Manual Points': '',
      'Note': 'End Date is optional - will be calculated automatically if empty'
    }];
    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shifts');
    XLSX.writeFile(wb, 'shifts_template.xlsx');
  };

  const handleExportShifts = () => {
    const exportData = filteredShifts.map(shift => {
      const emp = employees.find(e => e.id === shift.employeeId);
      const points = emp ? calculateShiftPoints(shift, emp.status) : 0;

      return {
        'Employee ID': shift.employeeId || '',
        'Employee Name': emp?.name || t.unassigned,
        'Personal Number': emp?.personalNumber || '',
        'Department': emp?.department || '',
        'Status': emp?.status || '',
        'Start Date': shift.startDate,
        'End Date': shift.endDate,
        'Duty Type': shift.dutyType,
        'Date Type': shift.dateType,
        'Manual Points': shift.manualPoints || '',
        'Calculated Points': points.toFixed(2)
      };
    });
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shifts');
    XLSX.writeFile(wb, `shifts_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleAddEmployee = () => {
    if (newEmployee.name && newEmployee.personalNumber) {
      setEmployees([...employees, { id: Date.now(), ...newEmployee }]);
      setNewEmployee({ name: '', personalNumber: '', sex: '', status: '', department: '', exemptions: [] });
      setShowAddEmployee(false);
    }
  };

  const handleEditEmployee = () => {
    if (editingEmployee && editingEmployee.name && editingEmployee.personalNumber) {
      setEmployees(employees.map(e => e.id === editingEmployee.id ? editingEmployee : e));
      setEditingEmployee(null);
    }
  };

  const handleDeleteEmployee = (id) => {
    setEmployees(employees.filter(e => e.id !== id));
    setShifts(shifts.filter(s => s.employeeId !== id));
    setDeleteConfirm(null);
  };

  const handleResetSystem = () => {
    // מחיקת כל הנתונים
    localStorage.clear();

    // איפוס כל ה-state
    setEmployees([]);
    setShifts([]);
    setStatusMultipliers(DEFAULT_STATUS_MULTIPLIERS);
    setSearchTerm('');
    setFilterDepartment('');
    setFilterStatus('');
    setFilterEmployee('');
    setFilterShiftType('');
    setShowAddEmployee(false);
    setShowAddShift(false);
    setEditingEmployee(null);
    setEditingShift(null);
    setDeleteConfirm(null);
    setSwappingShift(null);
    setShowStatusSettings(false);
    setResetConfirm(false);
    setUploadMessage('');

    // איפוס לשפה עברית ומצב בהיר
    setLanguage('he');
    setDarkMode(false);
  };

  const handleAddShift = () => {
    if (newShift.startDate && newShift.dutyType && newShift.dateType) {
      // חישוב אוטומטי של תאריך סיום אם לא הוזן ידנית
      const finalEndDate = newShift.endDate || calculateEndDate(newShift.startDate, newShift.dateType);

      if (newShift.employeeId && hasConflict(parseInt(newShift.employeeId), newShift.startDate, finalEndDate)) {
        setUploadMessage(t.conflictWarning);
        setTimeout(() => setUploadMessage(''), 3000);
        return;
      }

      // זיהוי אוטומטי של חג אם לא צוין ידנית
      const detectedHoliday = !newShift.holidayName ? getHolidayForDate(newShift.startDate) : null;
      const effectiveHolidayName = newShift.holidayName || detectedHoliday || '';

      const newShiftObj = {
        id: Date.now(),
        employeeId: newShift.employeeId ? parseInt(newShift.employeeId) : null,
        startDate: newShift.startDate,
        endDate: finalEndDate,
        dutyType: newShift.dutyType,
        dateType: newShift.dateType,
        manualPoints: newShift.manualPoints,
        holidayName: effectiveHolidayName
      };

      if (newShiftObj.employeeId && effectiveHolidayName) {
        updateHolidayHistory(newShiftObj.employeeId, effectiveHolidayName, newShiftObj.startDate);
      }

      setShifts([...shifts, newShiftObj]);
      setNewShift({ employeeId: '', startDate: '', endDate: '', dutyType: '', dateType: '', manualPoints: null, holidayName: '' });
      setShowAddShift(false);
    }
  };

  const handleUpdateShift = () => {
    if (editingShift && (!editingShift.employeeId || !hasConflict(editingShift.employeeId, editingShift.startDate, editingShift.endDate, editingShift.id))) {
      setShifts(shifts.map(s => s.id === editingShift.id ? editingShift : s));
      setEditingShift(null);
    }
  };

  const handleDeleteShift = (id) => {
    setShifts(shifts.filter(s => s.id !== id));
    setDeleteConfirm(null);
  };

  const handleSwapShifts = (shift1Id, shift2Id) => {
    const shift1 = shifts.find(s => s.id === shift1Id);
    const shift2 = shifts.find(s => s.id === shift2Id);

    if (shift1 && shift2) {
      const updatedShifts = shifts.map(s => {
        if (s.id === shift1Id) {
          return { ...s, employeeId: shift2.employeeId };
        }
        if (s.id === shift2Id) {
          return { ...s, employeeId: shift1.employeeId };
        }
        return s;
      });
      setShifts(updatedShifts);
      setSwappingShift(null);
      setUploadMessage('✅ תורנויות הוחלפו בהצלחה!');
      setTimeout(() => setUploadMessage(''), 3000);
    }
  };

  const handleAutoDistribute = (dryRun = true) => {
    const unassignedShifts = shifts.filter(s => !s.employeeId);
    if (unassignedShifts.length === 0) {
      setUploadMessage('❌ אין תורנויות לא משובצות');
      setTimeout(() => setUploadMessage(''), 3000);
      return;
    }

    if (employees.length === 0) {
      setUploadMessage('❌ אין עובדים במערכת');
      setTimeout(() => setUploadMessage(''), 3000);
      return;
    }

    const employeePointsMap = {};
    employees.forEach(emp => {
      const stat = employeeStats.find(s => s.id === emp.id);
      employeePointsMap[emp.id] = stat ? stat.justicePoints : 0;
    });

    const updatedShifts = [...shifts];
    let distributedCount = 0;
    const assignmentDetails = [];

    unassignedShifts.forEach(shift => {
      // זיהוי אוטומטי של חג לפי תאריך
      const detectedHoliday = getHolidayForDate(shift.startDate);
      const effectiveHolidayName = shift.holidayName || detectedHoliday;

      const sortedEmployees = [...employees].sort((a, b) =>
        employeePointsMap[a.id] - employeePointsMap[b.id]
      );

      for (let emp of sortedEmployees) {
        const hasConflictCheck = !hasConflict(emp.id, shift.startDate, shift.endDate);
        const hasExemptionCheck = !isEmployeeExemptFromDuty(emp, shift.role);
        const hasHolidayConflict = effectiveHolidayName && didEmployeeDoHolidayLastYear(emp.id, effectiveHolidayName);

        if (hasConflictCheck && hasExemptionCheck && !hasHolidayConflict) {
          const shiftIndex = updatedShifts.findIndex(s => s.id === shift.id);
          if (shiftIndex !== -1) {
            // עדכון התורנות עם החג שזוהה אוטומטית
            const updatedShift = {
              ...shift,
              employeeId: emp.id,
              holidayName: effectiveHolidayName || shift.holidayName
            };

            updatedShifts[shiftIndex] = updatedShift;
            const shiftPoints = calculateShiftPoints(shift, emp.status);
            employeePointsMap[emp.id] += shiftPoints;

            assignmentDetails.push({
              shift: shift,
              employee: emp,
              holidayName: effectiveHolidayName
            });

            if (!dryRun && effectiveHolidayName) {
              updateHolidayHistory(emp.id, effectiveHolidayName, shift.startDate);
            }

            distributedCount++;
            break;
          }
        }
      }
    });

    if (dryRun) {
      // Show preview modal
      setDryRunPreview({
        updatedShifts: updatedShifts,
        distributedCount: distributedCount,
        unassignedCount: unassignedShifts.length - distributedCount,
        assignmentDetails: assignmentDetails
      });
    } else {
      // Apply directly
      setShifts(updatedShifts);
      assignmentDetails.forEach(detail => {
        if (detail.holidayName) {
          updateHolidayHistory(detail.employee.id, detail.holidayName, detail.shift.startDate);
        }
      });
      setUploadMessage(`✅ שובצו ${distributedCount} תורנויות אוטומטית!`);
      setTimeout(() => setUploadMessage(''), 3000);
    }
  };

  const applyDryRunPreview = () => {
    if (dryRunPreview) {
      setShifts(dryRunPreview.updatedShifts);
      dryRunPreview.assignmentDetails.forEach(detail => {
        if (detail.holidayName) {
          updateHolidayHistory(detail.employee.id, detail.holidayName, detail.shift.startDate);
        }
      });
      setUploadMessage(`✅ שובצו ${dryRunPreview.distributedCount} תורנויות אוטומטית!`);
      setTimeout(() => setUploadMessage(''), 3000);
      setDryRunPreview(null);
    }
  };

  const handleExportPDF = () => {
    const doc = new jsPDF();

    // RTL support
    const isHebrew = language === 'he';

    // Title
    doc.setFontSize(20);
    doc.text(isHebrew ? 'דוח לוח תורנויות' : 'Shift Schedule Report', isHebrew ? 200 : 10, 15, { align: isHebrew ? 'right' : 'left' });

    // Date
    doc.setFontSize(10);
    const dateStr = new Date().toLocaleDateString(isHebrew ? 'he-IL' : 'en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    doc.text(`${isHebrew ? 'נוצר בתאריך' : 'Generated on'}: ${dateStr}`, isHebrew ? 200 : 10, 25, { align: isHebrew ? 'right' : 'left' });

    // Prepare table data
    const tableData = filteredShifts.map(shift => {
      const emp = getEmployee(shift.employeeId);
      const startDate = new Date(shift.startDate).toLocaleDateString(isHebrew ? 'he-IL' : 'en-US', { year: 'numeric', month: 'short', day: 'numeric' });
      const endDate = shift.endDate && shift.endDate !== shift.startDate
        ? new Date(shift.endDate).toLocaleDateString(isHebrew ? 'he-IL' : 'en-US', { year: 'numeric', month: 'short', day: 'numeric' })
        : '';

      return [
        emp?.name || (isHebrew ? 'לא משובץ' : 'Unassigned'),
        startDate,
        endDate || '-',
        shift.dutyType || '-',
        shift.dateType || '-',
        shift.holidayName || '-'
      ];
    });

    // Table headers
    const headers = isHebrew
      ? [['שם עובד', 'תאריך התחלה', 'תאריך סיום', 'סוג תורנות', 'סוג יום', 'חג']]
      : [['Employee', 'Start Date', 'End Date', 'Duty Type', 'Date Type', 'Holiday']];

    // Generate table
    autoTable(doc, {
      head: headers,
      body: tableData,
      startY: 35,
      styles: {
        fontSize: 10,
        cellPadding: 5,
        halign: isHebrew ? 'right' : 'left'
      },
      headStyles: {
        fillColor: [59, 130, 246],
        textColor: 255,
        fontStyle: 'bold'
      },
      alternateRowStyles: {
        fillColor: [245, 247, 250]
      },
      margin: { top: 35, right: 10, bottom: 10, left: 10 }
    });

    // Add statistics at the bottom
    const finalY = doc.previousAutoTable.finalY + 15;
    doc.setFontSize(12);
    doc.text(isHebrew ? 'סטטיסטיקה:' : 'Statistics:', isHebrew ? 200 : 10, finalY, { align: isHebrew ? 'right' : 'left' });

    doc.setFontSize(10);
    doc.text(`${isHebrew ? 'סה"כ תורנויות' : 'Total shifts'}: ${shifts.length}`, isHebrew ? 200 : 10, finalY + 8, { align: isHebrew ? 'right' : 'left' });
    doc.text(`${isHebrew ? 'משובצות' : 'Assigned'}: ${shifts.filter(s => s.employeeId).length}`, isHebrew ? 200 : 10, finalY + 16, { align: isHebrew ? 'right' : 'left' });
    doc.text(`${isHebrew ? 'לא משובצות' : 'Unassigned'}: ${shifts.filter(s => !s.employeeId).length}`, isHebrew ? 200 : 10, finalY + 24, { align: isHebrew ? 'right' : 'left' });

    // Save PDF
    const filename = isHebrew ? `לוח_תורנויות_${new Date().toISOString().split('T')[0]}.pdf` : `shift_schedule_${new Date().toISOString().split('T')[0]}.pdf`;
    doc.save(filename);

    setUploadMessage(`✅ ${isHebrew ? 'PDF יוצא בהצלחה!' : 'PDF exported successfully!'}`);
    setTimeout(() => setUploadMessage(''), 3000);
  };

  const handleYearlyExport = () => {
    const isHebrew = language === 'he';
    const year = selectedYear;

    // Filter shifts for the selected year
    const yearShifts = shifts.filter(shift => {
      const shiftYear = new Date(shift.startDate).getFullYear();
      return shiftYear === year && shift.employeeId;
    });

    // Group by employee
    const employeeYearlyData = {};
    yearShifts.forEach(shift => {
      const empId = shift.employeeId;
      if (!employeeYearlyData[empId]) {
        const emp = employees.find(e => e.id === empId);
        employeeYearlyData[empId] = {
          name: emp?.name || 'Unknown',
          status: emp?.status || '',
          shifts: [],
          totalPoints: 0,
          dutyTypeCounts: {}
        };
      }
      employeeYearlyData[empId].shifts.push(shift);

      // Count duty types
      const dutyType = shift.dutyType;
      if (!employeeYearlyData[empId].dutyTypeCounts[dutyType]) {
        employeeYearlyData[empId].dutyTypeCounts[dutyType] = 0;
      }
      employeeYearlyData[empId].dutyTypeCounts[dutyType]++;

      // Calculate points
      const points = calculateShiftPoints(shift, employeeYearlyData[empId].status);
      employeeYearlyData[empId].totalPoints += points;
    });

    // Create Excel workbook
    const wb = XLSX.utils.book_new();

    // Summary sheet
    const summaryData = Object.values(employeeYearlyData).map(emp => {
      const row = {
        [isHebrew ? 'שם' : 'Name']: emp.name,
        [isHebrew ? 'סטטוס' : 'Status']: emp.status,
        [isHebrew ? 'סה"כ תורנויות' : 'Total Shifts']: emp.shifts.length,
        [isHebrew ? 'ניקוד צדק' : 'Justice Points']: emp.totalPoints.toFixed(1)
      };

      // Add duty type counts
      Object.keys(dutyTypePoints).forEach(type => {
        row[type] = emp.dutyTypeCounts[type] || 0;
      });

      return row;
    });

    const summarySheet = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summarySheet, isHebrew ? `סיכום ${year}` : `Summary ${year}`);

    // Individual employee sheets (up to 10 employees)
    const topEmployees = Object.values(employeeYearlyData).slice(0, 10);
    topEmployees.forEach(emp => {
      const sheetData = emp.shifts.map(shift => ({
        [isHebrew ? 'תאריך התחלה' : 'Start Date']: new Date(shift.startDate).toLocaleDateString(isHebrew ? 'he-IL' : 'en-US'),
        [isHebrew ? 'תאריך סיום' : 'End Date']: shift.endDate && shift.endDate !== shift.startDate
          ? new Date(shift.endDate).toLocaleDateString(isHebrew ? 'he-IL' : 'en-US')
          : '-',
        [isHebrew ? 'סוג תורנות' : 'Duty Type']: shift.dutyType,
        [isHebrew ? 'סוג יום' : 'Date Type']: shift.dateType,
        [isHebrew ? 'חג' : 'Holiday']: shift.holidayName || '-'
      }));

      const empSheet = XLSX.utils.json_to_sheet(sheetData);
      const sheetName = emp.name.substring(0, 30); // Excel sheet name limit
      XLSX.utils.book_append_sheet(wb, empSheet, sheetName);
    });

    // Save file
    const filename = isHebrew ? `דוח_שנתי_${year}.xlsx` : `yearly_report_${year}.xlsx`;
    XLSX.writeFile(wb, filename);

    setUploadMessage(`✅ ${isHebrew ? `דוח שנתי ${year} יוצא בהצלחה!` : `Yearly report ${year} exported successfully!`}`);
    setTimeout(() => setUploadMessage(''), 3000);
  };

  const handleAddStatusMultiplier = () => {
    if (newStatusName.trim() && !statusMultipliers[newStatusName.trim()]) {
      setStatusMultipliers({...statusMultipliers, [newStatusName.trim()]: parseFloat(newStatusMultiplier)});
      setNewStatusName('');
      setNewStatusMultiplier(1.0);
    }
  };

  const handleDeleteStatusMultiplier = (statusName) => {
    const newMultipliers = {...statusMultipliers};
    delete newMultipliers[statusName];
    setStatusMultipliers(newMultipliers);
  };

  const handleAddDutyType = () => {
    if (newDutyType.name.trim() && !dutyTypePoints[newDutyType.name.trim()]) {
      const colors = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#14b8a6'];
      const randomColor = colors[Math.floor(Math.random() * colors.length)];

      setDutyTypePoints({
        ...dutyTypePoints,
        [newDutyType.name.trim()]: {
          name: newDutyType.name.trim(),
          basePoints: parseFloat(newDutyType.basePoints) || 1,
          description: newDutyType.description.trim() || newDutyType.name.trim(),
          color: randomColor
        }
      });
      setNewDutyType({ name: '', basePoints: 1, description: '' });
    }
  };

  const handleDeleteDutyType = (typeName) => {
    const newTypes = {...dutyTypePoints};
    delete newTypes[typeName];
    setDutyTypePoints(newTypes);
  };

  const handleAddDateType = () => {
    if (newDateType.name.trim() && !dateTypePoints[newDateType.name.trim()]) {
      const colors = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#14b8a6'];
      const randomColor = colors[Math.floor(Math.random() * colors.length)];

      setDateTypePoints({
        ...dateTypePoints,
        [newDateType.name.trim()]: {
          name: newDateType.name.trim(),
          bonus: parseFloat(newDateType.bonus) || 0,
          description: newDateType.description.trim() || newDateType.name.trim(),
          color: randomColor
        }
      });
      setNewDateType({ name: '', bonus: 0, description: '' });
    }
  };

  const handleDeleteDateType = (typeName) => {
    const newTypes = {...dateTypePoints};
    delete newTypes[typeName];
    setDateTypePoints(newTypes);
  };

  const handleAddExemption = () => {
    if (newExemption.name.trim() && !exemptionTypes[newExemption.name.trim()]) {
      setExemptionTypes({
        ...exemptionTypes,
        [newExemption.name.trim()]: {
          name: newExemption.name.trim(),
          exemptFromDutyTypes: newExemption.exemptFromDutyTypes,
          description: newExemption.description.trim() || newExemption.name.trim()
        }
      });
      setNewExemption({ name: '', exemptFromDutyTypes: [], description: '' });
    }
  };

  const handleDeleteExemption = (exemptionName) => {
    const newExemptions = {...exemptionTypes};
    delete newExemptions[exemptionName];
    setExemptionTypes(newExemptions);
  };

  // חישוב תאריכי חג לשנה הנוכחית
  const getHolidayDatesForYear = (holidayData, year = new Date().getFullYear()) => {
    try {
      if (!holidayData.hebrewId) return null;

      const events = HebrewCalendar.calendar({
        year: year,
        isHebrewYear: false,
        candlelighting: false,
        sedrot: false,
        omer: false
      });

      for (const event of events) {
        if (event.getDesc().includes(holidayData.hebrewId)) {
          const baseDate = new Date(event.getDate().greg());
          const offset = holidayData.dayOffset || 0;
          const startDate = new Date(baseDate);
          startDate.setDate(startDate.getDate() + offset);

          const endDate = new Date(startDate);
          endDate.setDate(endDate.getDate() + (holidayData.durationDays || 1) - 1);

          return {
            start: startDate.toLocaleDateString('he-IL'),
            end: endDate.toLocaleDateString('he-IL'),
            startISO: startDate.toISOString().split('T')[0],
            endISO: endDate.toISOString().split('T')[0]
          };
        }
      }
      return null;
    } catch (error) {
      console.error('Error calculating holiday dates:', error);
      return null;
    }
  };

  // בדיקה אם עובד פטור מסוג תורנות מסוים
  const isEmployeeExemptFromDuty = (employee, dutyType) => {
    if (!employee.exemptions || employee.exemptions.length === 0) return false;

    return employee.exemptions.some(exemptionName => {
      const exemption = exemptionTypes[exemptionName];
      return exemption && exemption.exemptFromDutyTypes.includes(dutyType);
    });
  };

  // זיהוי חג לפי תאריך
  const getHolidayForDate = (dateString) => {
    try {
      const date = new Date(dateString);
      const year = date.getFullYear();

      // קבלת כל החגים של השנה
      const events = HebrewCalendar.calendar({
        year: year,
        isHebrewYear: false,
        candlelighting: false,
        sedrot: false,
        omer: false
      });

      // המרת התאריך לפורמט להשוואה
      const targetDate = date.toISOString().split('T')[0];

      // חיפוש חג שמתאים לתאריך
      for (const event of events) {
        const eventDate = event.getDate().greg().toISOString().split('T')[0];

        // בדיקה לכל חג מוגדר אם יש התאמה
        for (const [holidayName, holidayData] of Object.entries(holidays)) {
          if (holidayData.hebrewId && event.getDesc().includes(holidayData.hebrewId)) {
            // בדיקה אם התאריך בטווח החג (כולל offset ו-duration)
            const eventDateObj = new Date(eventDate);
            const offset = holidayData.dayOffset || 0;

            const holidayStartDate = new Date(eventDateObj);
            holidayStartDate.setDate(holidayStartDate.getDate() + offset);

            const holidayEndDate = new Date(holidayStartDate);
            holidayEndDate.setDate(holidayEndDate.getDate() + (holidayData.durationDays || 1) - 1);

            const targetDateObj = new Date(targetDate);

            // בדיקה אם התאריך בטווח
            if (targetDateObj >= holidayStartDate && targetDateObj <= holidayEndDate) {
              return holidayName;
            }
          }
        }
      }

      return null;
    } catch (error) {
      console.error('Error detecting holiday:', error);
      return null;
    }
  };

  // ניהול חגים
  const handleAddHoliday = () => {
    if (newHoliday.name.trim() && !holidays[newHoliday.name.trim()]) {
      setHolidays({
        ...holidays,
        [newHoliday.name.trim()]: {
          name: newHoliday.name.trim(),
          weight: parseFloat(newHoliday.weight) || 1,
          description: newHoliday.description.trim() || newHoliday.name.trim()
        }
      });
      setNewHoliday({ name: '', weight: 1, description: '' });
    }
  };

  const handleDeleteHoliday = (holidayName) => {
    const newHolidays = {...holidays};
    delete newHolidays[holidayName];
    setHolidays(newHolidays);
  };

  // בדיקה אם עובד עשה חג בשנה שעברה
  const didEmployeeDoHolidayLastYear = (employeeId, holidayName) => {
    const currentYear = new Date().getFullYear();
    const lastYear = currentYear - 1;
    const key = `${employeeId}_${lastYear}`;

    if (!holidayHistory[key]) return false;
    return holidayHistory[key].includes(holidayName);
  };

  // עדכון היסטוריית חגים כאשר משבצים תורנות חג
  const updateHolidayHistory = (employeeId, holidayName, shiftDate) => {
    const year = new Date(shiftDate).getFullYear();
    const key = `${employeeId}_${year}`;

    setHolidayHistory(prev => {
      const current = prev[key] || [];
      if (!current.includes(holidayName)) {
        return {
          ...prev,
          [key]: [...current, holidayName]
        };
      }
      return prev;
    });
  };

  const getEmployee = (employeeId) => employees.find(e => e.id === employeeId);

  // תצוגה לפי תורנויות - כל תורנות מוצגת פעם אחת בלבד
  const getAllShifts = () => filteredShifts
    .filter(s => s.startDate) // סנן תורנויות ישנות ללא startDate
    .sort((a, b) => a.startDate.localeCompare(b.startDate));

  // תצוגת שבוע/חודש - מחזיר תורנויות שפעילות בתאריך מסוים
  const getShiftsByDate = (date) => filteredShifts.filter(s => {
    if (!s.startDate) return false; // דלג על תורנויות ישנות
    const shiftStart = new Date(s.startDate);
    const shiftEnd = s.endDate ? new Date(s.endDate) : shiftStart;
    const currentDate = new Date(date);
    return currentDate >= shiftStart && currentDate <= shiftEnd;
  }).sort((a, b) => a.startDate.localeCompare(b.startDate));

  const getWeekDates = (startDate) => {
    const dates = [];
    const start = new Date(startDate);
    start.setDate(start.getDate() - start.getDay());
    for (let i = 0; i < 7; i++) {
      const date = new Date(start);
      date.setDate(start.getDate() + i);
      dates.push(date.toISOString().split('T')[0]);
    }
    return dates;
  };

  const getMonthDates = () => {
    const now = new Date();
    const dates = [];
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
      dates.push(new Date(d).toISOString().split('T')[0]);
    }
    return dates;
  };

  const navigateWeek = (dir) => {
    const newDate = new Date(currentWeekStart);
    newDate.setDate(newDate.getDate() + (dir * 7));
    setCurrentWeekStart(newDate);
  };

  const styles = {
    container: { minHeight: '100vh', padding: '16px', background: darkMode ? '#111827' : 'linear-gradient(to bottom right, #eff6ff, #e0e7ff)', direction: language === 'he' ? 'rtl' : 'ltr' },
    card: { maxWidth: '1280px', margin: '0 auto', background: darkMode ? '#1f2937' : 'white', borderRadius: '12px', boxShadow: '0 10px 15px -3px rgba(0,0,0,0.1)', overflow: 'hidden' },
    header: { padding: '24px', background: darkMode ? 'linear-gradient(to right, #4338ca, #7e22ce)' : 'linear-gradient(to right, #3b82f6, #6366f1)', color: 'white' },
    headerFlex: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' },
    headerLeft: { display: 'flex', alignItems: 'center', gap: '12px' },
    title: { margin: 0, fontSize: '30px', fontWeight: 'bold' },
    subtitle: { fontSize: '14px', opacity: 0.9, marginTop: '4px' },
    headerRight: { display: 'flex', gap: '12px', alignItems: 'center' },
    iconBtn: { padding: '8px', background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '8px', cursor: 'pointer', color: 'white', transition: 'all 0.2s' },
    saveIndicator: {
      display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 12px',
      background: 'rgba(255,255,255,0.2)', borderRadius: '8px', fontSize: '13px',
      opacity: saveStatus ? 1 : 0, transition: 'opacity 0.3s'
    },
    message: { padding: '16px', borderLeft: '4px solid' },
    tabs: { display: 'flex', borderBottom: darkMode ? '1px solid #374151' : '1px solid #e5e7eb' },
    tab: (active) => ({
      flex: 1, padding: '16px 24px', fontWeight: 500, border: 'none', cursor: 'pointer', transition: 'all 0.2s',
      background: active ? (darkMode ? '#374151' : '#eff6ff') : 'transparent',
      color: active ? (darkMode ? '#60a5fa' : '#2563eb') : (darkMode ? '#9ca3af' : '#6b7280'),
      borderBottom: active ? '2px solid ' + (darkMode ? '#60a5fa' : '#2563eb') : 'none'
    }),
    content: { padding: '24px' },
    flexBetween: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', flexWrap: 'wrap', gap: '16px' },
    h2: { fontSize: '24px', fontWeight: 'bold', color: darkMode ? 'white' : '#1f2937', margin: 0 },
    btnGroup: { display: 'flex', gap: '12px', flexWrap: 'wrap' },
    btn: (color) => ({
      padding: '10px 20px', borderRadius: '8px', border: 'none', cursor: 'pointer',
      display: 'flex', alignItems: 'center', gap: '8px', fontSize: '14px', fontWeight: 500,
      background: darkMode ? (color === 'green' ? '#065f46' : color === 'indigo' ? '#4338ca' : color === 'blue' ? '#1e40af' : color === 'purple' ? '#6b21a8' : color === 'red' ? '#991b1b' : color === 'orange' ? '#c2410c' : '#374151') :
                            (color === 'green' ? '#16a34a' : color === 'indigo' ? '#6366f1' : color === 'blue' ? '#3b82f6' : color === 'purple' ? '#9333ea' : color === 'red' ? '#dc2626' : color === 'orange' ? '#f97316' : '#6b7280'),
      color: 'white', transition: 'all 0.2s'
    }),
    grid3: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px', marginBottom: '24px' },
    inputWrapper: { position: 'relative' },
    input: {
      width: '100%', padding: '10px', paddingLeft: language === 'he' ? '10px' : '40px', paddingRight: language === 'he' ? '40px' : '10px',
      border: '2px solid ' + (darkMode ? '#4b5563' : '#d1d5db'), borderRadius: '8px',
      background: darkMode ? '#1f2937' : 'white', color: darkMode ? 'white' : 'black', fontSize: '14px',
      outline: 'none', transition: 'border-color 0.2s'
    },
    dateInput: {
      width: '100%', padding: '14px 12px', minHeight: '48px',
      border: '2px solid ' + (darkMode ? '#4b5563' : '#d1d5db'), borderRadius: '8px',
      background: darkMode ? '#1f2937' : 'white', color: darkMode ? 'white' : 'black', fontSize: '15px',
      outline: 'none', transition: 'all 0.2s', cursor: 'pointer', fontWeight: 500,
      colorScheme: darkMode ? 'dark' : 'light',
      boxShadow: darkMode ? '0 1px 3px rgba(0,0,0,0.2)' : '0 1px 2px rgba(0,0,0,0.05)'
    },
    select: {
      width: '100%', padding: '14px 12px', minHeight: '48px', border: '2px solid ' + (darkMode ? '#4b5563' : '#d1d5db'), borderRadius: '8px',
      background: darkMode ? '#1f2937' : 'white', color: darkMode ? 'white' : 'black', fontSize: '15px',
      outline: 'none', transition: 'all 0.2s', cursor: 'pointer', fontWeight: 500,
      boxShadow: darkMode ? '0 1px 3px rgba(0,0,0,0.2)' : '0 1px 2px rgba(0,0,0,0.05)'
    },
    modal: {
      marginBottom: '24px', padding: '20px', borderRadius: '12px',
      border: '2px solid ' + (darkMode ? '#4b5563' : '#3b82f6'),
      background: darkMode ? '#1f2937' : '#eff6ff',
      boxShadow: darkMode ? '0 4px 6px rgba(0,0,0,0.3)' : '0 4px 6px rgba(0,0,0,0.1)'
    },
    modalHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' },
    modalTitle: { fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', margin: 0 },
    grid2: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' },
    label: { display: 'block', fontSize: '14px', marginBottom: '6px', color: darkMode ? '#d1d5db' : '#374151', fontWeight: 500 },
    emptyState: {
      textAlign: 'center', padding: '48px', borderRadius: '12px',
      border: '2px dashed ' + (darkMode ? '#4b5563' : '#d1d5db'),
      background: darkMode ? '#1f2937' : '#f9fafb'
    },
    table: {
      width: '100%',
      borderCollapse: 'separate',
      borderSpacing: 0,
      border: darkMode ? '2px solid #4b5563' : '1px solid #e5e7eb',
      borderRadius: '12px',
      overflow: 'hidden'
    },
    thead: {
      background: darkMode ? '#374151' : '#f3f4f6',
      borderBottom: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')
    },
    th: {
      padding: '14px 16px',
      textAlign: language === 'he' ? 'right' : 'left',
      fontSize: '14px',
      fontWeight: 600,
      color: darkMode ? '#e5e7eb' : '#374151'
    },
    tr: {
      borderBottom: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'),
      transition: 'background 0.2s',
      background: darkMode ? '#1f2937' : 'transparent'
    },
    td: {
      padding: '12px 16px',
      color: darkMode ? '#e5e7eb' : '#6b7280'
    },
    avatar: {
      width: '40px', height: '40px', borderRadius: '50%',
      background: darkMode ? 'linear-gradient(to bottom right, #6366f1, #9333ea)' : 'linear-gradient(to bottom right, #3b82f6, #6366f1)',
      color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold'
    },
    badge: (type) => ({
      padding: '4px 12px', borderRadius: '9999px', fontSize: '12px', fontWeight: 500,
      background: type === 'green' ? '#dcfce7' : type === 'red' ? '#fee2e2' : type === 'yellow' ? '#fef3c7' : type === 'blue' ? '#dbeafe' : type === 'gray' ? '#f3f4f6' : '#e5e7eb',
      color: type === 'green' ? '#166534' : type === 'red' ? '#991b1b' : type === 'yellow' ? '#854d0e' : type === 'blue' ? '#1e40af' : type === 'gray' ? '#6b7280' : '#374151'
    }),
    calViewBtns: {
      display: 'flex',
      background: darkMode ? '#1f2937' : '#f3f4f6',
      borderRadius: '10px',
      padding: '6px',
      border: darkMode ? '2px solid #4b5563' : 'none'
    },
    calViewBtn: (active) => ({
      padding: '10px 18px',
      border: 'none',
      borderRadius: '8px',
      cursor: 'pointer',
      fontSize: '14px',
      fontWeight: 500,
      background: active ? (darkMode ? '#3b82f6' : 'white') : 'transparent',
      color: active ? 'white' : (darkMode ? '#9ca3af' : '#1f2937'),
      boxShadow: active ? '0 2px 4px rgba(0,0,0,0.2)' : 'none',
      transition: 'all 0.2s'
    }),
    dateCard: {
      border: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'),
      borderRadius: '12px',
      overflow: 'hidden',
      background: darkMode ? '#1f2937' : 'white',
      boxShadow: darkMode ? '0 2px 4px rgba(0,0,0,0.2)' : '0 1px 3px rgba(0,0,0,0.1)'
    },
    dateHeader: {
      background: darkMode ? '#374151' : '#f3f4f6',
      padding: '12px 16px',
      borderBottom: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')
    },
    shiftRow: (hasConflict) => ({
      padding: '16px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      background: hasConflict ? (darkMode ? '#7f1d1d' : '#fee2e2') : (darkMode ? '#1f2937' : 'transparent'),
      borderBottom: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'),
      transition: 'background 0.2s'
    }),
    statsCard: {
      padding: '20px',
      borderRadius: '12px',
      background: darkMode ? '#1f2937' : 'white',
      border: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'),
      marginBottom: '16px',
      boxShadow: darkMode ? '0 2px 4px rgba(0,0,0,0.2)' : '0 1px 3px rgba(0,0,0,0.1)'
    },
    statsGrid: {
      display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px'
    },
    statBox: {
      padding: '16px',
      borderRadius: '10px',
      background: darkMode ? '#374151' : '#f9fafb',
      border: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'),
      transition: 'all 0.2s'
    },
    confirmDialog: {
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
    },
    confirmBox: {
      background: darkMode ? '#1f2937' : 'white',
      padding: '24px',
      borderRadius: '16px',
      maxWidth: '400px',
      width: '90%',
      border: darkMode ? '2px solid #4b5563' : 'none',
      boxShadow: '0 20px 25px -5px rgba(0,0,0,0.3)'
    },
    pointsBreakdown: {
      display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px',
      padding: '4px 8px', borderRadius: '6px', background: darkMode ? '#1f2937' : '#f3f4f6'
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <div style={styles.header}>
          <div style={styles.headerFlex}>
            <div style={styles.headerLeft}>
              <Calendar size={32} />
              <div>
                <h1 style={styles.title}>
                  {t.appTitle}
                  {selectedYear && (
                    <span style={{
                      fontSize: '16px', fontWeight: 500, marginRight: language === 'he' ? '12px' : '0',
                      marginLeft: language === 'en' ? '12px' : '0',
                      padding: '4px 12px', borderRadius: '6px',
                      background: darkMode ? '#3b82f6' : '#dbeafe',
                      color: darkMode ? 'white' : '#1e40af'
                    }}>
                      {selectedYear}
                    </span>
                  )}
                </h1>
                <p style={styles.subtitle}>{t.subtitle}</p>
              </div>
            </div>
            <div style={styles.headerRight}>
              {saveStatus && (
                <div style={styles.saveIndicator}>
                  {saveStatus === 'saving' ? (
                    <>
                      <RefreshCw size={16} style={{animation: 'spin 1s linear infinite'}} />
                      <span>{t.saving}</span>
                    </>
                  ) : (
                    <>
                      <CheckCircle size={16} />
                      <span>{t.saved}</span>
                    </>
                  )}
                </div>
              )}
              <select
                value={selectedYear || ''}
                onChange={(e) => setSelectedYear(e.target.value ? parseInt(e.target.value) : null)}
                style={{
                  padding: '8px 12px', borderRadius: '8px', fontSize: '14px', fontWeight: 600,
                  border: '2px solid ' + (darkMode ? '#4b5563' : '#d1d5db'),
                  background: darkMode ? '#1f2937' : 'white',
                  color: darkMode ? 'white' : '#1f2937',
                  cursor: 'pointer', outline: 'none',
                  minWidth: '100px'
                }}
              >
                <option value="">{t.allYears}</option>
                {availableYears.map(year => (
                  <option key={year} value={year}>{year}</option>
                ))}
              </select>
              <button onClick={() => setShowStatusSettings(!showStatusSettings)} style={styles.iconBtn} title={t.manageStatuses}>
                <Settings size={20} />
              </button>
              <button onClick={() => setShowPointsSettings(!showPointsSettings)} style={{...styles.iconBtn, color: '#10b981'}} title={t.managePoints}>
                <Award size={20} />
              </button>
              <button onClick={() => setShowExemptionSettings(!showExemptionSettings)} style={{...styles.iconBtn, color: '#f59e0b'}} title={t.manageExemptions}>
                <Shield size={20} />
              </button>
              <button onClick={() => setShowHolidaySettings(!showHolidaySettings)} style={{...styles.iconBtn, color: '#8b5cf6'}} title={t.manageHolidays}>
                <Gift size={20} />
              </button>
              <button onClick={() => setResetConfirm(true)} style={{...styles.iconBtn, color: '#ef4444'}} title={t.resetSystem}>
                <Trash2 size={20} />
              </button>
              <button onClick={() => setLanguage(language === 'en' ? 'he' : 'en')} style={styles.iconBtn}>
                <Globe size={20} />
              </button>
              <button onClick={() => setDarkMode(!darkMode)} style={styles.iconBtn}>
                {darkMode ? <Sun size={20} /> : <Moon size={20} />}
              </button>
            </div>
          </div>
        </div>

        {uploadMessage && (
          <div style={{...styles.message, borderColor: uploadMessage.includes('⚠️') || uploadMessage.includes('❌') ? '#f59e0b' : '#10b981', background: uploadMessage.includes('⚠️') || uploadMessage.includes('❌') ? '#fef3c7' : '#d1fae5', color: uploadMessage.includes('⚠️') || uploadMessage.includes('❌') ? '#92400e' : '#065f46'}}>
            {uploadMessage}
          </div>
        )}

        {showStatusSettings && (
          <div style={{...styles.message, borderColor: '#3b82f6', background: darkMode ? '#1e293b' : '#eff6ff', color: darkMode ? 'white' : '#1e293b'}}>
            <h3 style={{margin: '0 0 16px 0', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <TrendingUp size={20} />
              {t.manageStatusMultipliers}
            </h3>
            <div style={{marginBottom: '16px'}}>
              <div style={{display: 'grid', gridTemplateColumns: '2fr 1fr auto', gap: '8px', marginBottom: '8px'}}>
                <input
                  type="text"
                  value={newStatusName}
                  onChange={(e) => setNewStatusName(e.target.value)}
                  placeholder={t.statusName}
                  style={styles.input}
                />
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="2"
                  value={newStatusMultiplier}
                  onChange={(e) => setNewStatusMultiplier(e.target.value)}
                  placeholder={t.multiplier}
                  style={styles.input}
                />
                <button onClick={handleAddStatusMultiplier} style={styles.btn('green')}>
                  <Plus size={16} />
                  {t.add}
                </button>
              </div>
            </div>
            <div style={{display: 'flex', flexWrap: 'wrap', gap: '8px'}}>
              {Object.entries(statusMultipliers).map(([status, multiplier]) => (
                <div key={status} style={{
                  display: 'inline-flex', alignItems: 'center', gap: '8px',
                  padding: '8px 12px', borderRadius: '8px',
                  background: darkMode ? '#374151' : 'white',
                  border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')
                }}>
                  <span style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{status}</span>
                  <span style={{padding: '2px 8px', borderRadius: '4px', background: darkMode ? '#1f2937' : '#f3f4f6', fontSize: '12px', fontWeight: 600}}>
                    ×{multiplier.toFixed(1)}
                  </span>
                  <button
                    onClick={() => handleDeleteStatusMultiplier(status)}
                    style={{background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: '#ef4444'}}
                  >
                    <X size={16} />
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        {showPointsSettings && (
          <div style={{...styles.message, borderColor: '#10b981', background: darkMode ? '#1e293b' : '#ecfdf5', color: darkMode ? 'white' : '#1e293b'}}>
            <h3 style={{margin: '0 0 16px 0', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <Award size={20} />
              {t.pointsSettings}
            </h3>

            <div style={{marginBottom: '24px'}}>
              <h4 style={{margin: '0 0 12px 0', fontSize: '16px', fontWeight: 600}}>{t.dutyTypePoints}</h4>

              {/* טופס להוספת סוג תורנות חדש */}
              <div style={{marginBottom: '16px', padding: '12px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb', border: '2px dashed ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
                <div style={{display: 'grid', gridTemplateColumns: '2fr 1fr 2fr auto', gap: '8px', alignItems: 'end'}}>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.dutyTypeName}</label>
                    <input
                      type="text"
                      value={newDutyType.name}
                      onChange={(e) => setNewDutyType({...newDutyType, name: e.target.value})}
                      placeholder="לדוגמא: משמר"
                      style={styles.input}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>נקודות</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="10"
                      value={newDutyType.basePoints}
                      onChange={(e) => setNewDutyType({...newDutyType, basePoints: e.target.value})}
                      style={{...styles.input, textAlign: 'center'}}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.dutyTypeDescription}</label>
                    <input
                      type="text"
                      value={newDutyType.description}
                      onChange={(e) => setNewDutyType({...newDutyType, description: e.target.value})}
                      placeholder="תיאור קצר"
                      style={styles.input}
                    />
                  </div>
                  <button onClick={handleAddDutyType} style={{...styles.btn('green'), whiteSpace: 'nowrap'}}>
                    <Plus size={16} />
                    {t.addDutyType}
                  </button>
                </div>
              </div>

              <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                {Object.entries(dutyTypePoints).map(([key, type]) => (
                  <div key={key} style={{display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : 'white', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')}}>
                    <div style={{flex: 1}}>
                      <span style={{fontWeight: 600, fontSize: '16px'}}>{type.name}</span>
                      <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{type.description}</p>
                    </div>
                    <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                      <label style={{fontSize: '14px', color: darkMode ? '#d1d5db' : '#6b7280'}}>נקודות:</label>
                      <input
                        type="number"
                        step="0.5"
                        min="0"
                        max="10"
                        value={type.basePoints}
                        onChange={(e) => {
                          const newValue = parseFloat(e.target.value);
                          setDutyTypePoints({
                            ...dutyTypePoints,
                            [key]: {...type, basePoints: newValue}
                          });
                        }}
                        style={{...styles.input, width: '80px', padding: '8px', textAlign: 'center'}}
                      />
                      <button
                        onClick={() => handleDeleteDutyType(key)}
                        style={{background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: '#ef4444'}}
                        title="מחק"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <h4 style={{margin: '0 0 12px 0', fontSize: '16px', fontWeight: 600}}>{t.dateTypePoints}</h4>

              {/* טופס להוספת סוג יום חדש */}
              <div style={{marginBottom: '16px', padding: '12px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb', border: '2px dashed ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
                <div style={{display: 'grid', gridTemplateColumns: '2fr 1fr 2fr auto', gap: '8px', alignItems: 'end'}}>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.dateTypeName}</label>
                    <input
                      type="text"
                      value={newDateType.name}
                      onChange={(e) => setNewDateType({...newDateType, name: e.target.value})}
                      placeholder="לדוגמא: ערב חג"
                      style={styles.input}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>בונוס</label>
                    <input
                      type="number"
                      step="0.5"
                      min="0"
                      max="10"
                      value={newDateType.bonus}
                      onChange={(e) => setNewDateType({...newDateType, bonus: e.target.value})}
                      style={{...styles.input, textAlign: 'center'}}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.dutyTypeDescription}</label>
                    <input
                      type="text"
                      value={newDateType.description}
                      onChange={(e) => setNewDateType({...newDateType, description: e.target.value})}
                      placeholder="תיאור קצר"
                      style={styles.input}
                    />
                  </div>
                  <button onClick={handleAddDateType} style={{...styles.btn('green'), whiteSpace: 'nowrap'}}>
                    <Plus size={16} />
                    {t.addDateType}
                  </button>
                </div>
              </div>

              <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                {Object.entries(dateTypePoints).map(([key, type]) => (
                  <div key={key} style={{display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : 'white', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')}}>
                    <div style={{flex: 1}}>
                      <span style={{fontWeight: 600, fontSize: '16px'}}>{type.name}</span>
                      <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{type.description}</p>
                    </div>
                    <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                      <label style={{fontSize: '14px', color: darkMode ? '#d1d5db' : '#6b7280'}}>בונוס:</label>
                      <input
                        type="number"
                        step="0.5"
                        min="0"
                        max="10"
                        value={type.bonus}
                        onChange={(e) => {
                          const newValue = parseFloat(e.target.value);
                          setDateTypePoints({
                            ...dateTypePoints,
                            [key]: {...type, bonus: newValue}
                          });
                        }}
                        style={{...styles.input, width: '80px', padding: '8px', textAlign: 'center'}}
                      />
                      <button
                        onClick={() => handleDeleteDateType(key)}
                        style={{background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: '#ef4444'}}
                        title="מחק"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {showExemptionSettings && (
          <div style={{...styles.message, borderColor: '#f59e0b', background: darkMode ? '#1e293b' : '#fff7ed', color: darkMode ? 'white' : '#1e293b'}}>
            <h3 style={{margin: '0 0 16px 0', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <Shield size={20} />
              {t.exemptionSettings}
            </h3>
            <p style={{margin: '0 0 16px 0', fontSize: '14px', color: darkMode ? '#d1d5db' : '#6b7280'}}>
              Define exemptions and specify which duty types they exclude
            </p>

            <div style={{marginBottom: '16px', padding: '12px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb', border: '2px dashed ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
              <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                <div style={{display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px'}}>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.exemptionName}</label>
                    <input
                      type="text"
                      value={newExemption.name}
                      onChange={(e) => setNewExemption({...newExemption, name: e.target.value})}
                      placeholder="e.g., Dust allergy"
                      style={styles.input}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.exemptionDescription}</label>
                    <input
                      type="text"
                      value={newExemption.description}
                      onChange={(e) => setNewExemption({...newExemption, description: e.target.value})}
                      placeholder="Description"
                      style={styles.input}
                    />
                  </div>
                </div>
                <div>
                  <label style={{...styles.label, marginBottom: '4px'}}>{t.exemptFrom}</label>
                  <div style={{display: 'flex', flexWrap: 'wrap', gap: '8px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : 'white', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')}}>
                    {Object.values(dutyTypePoints).map(type => (
                      <label key={type.name} style={{display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '6px 12px', borderRadius: '6px', background: newExemption.exemptFromDutyTypes.includes(type.name) ? (darkMode ? '#3b82f6' : '#dbeafe') : 'transparent', border: '1px solid ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
                        <input
                          type="checkbox"
                          checked={newExemption.exemptFromDutyTypes.includes(type.name)}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setNewExemption({...newExemption, exemptFromDutyTypes: [...newExemption.exemptFromDutyTypes, type.name]});
                            } else {
                              setNewExemption({...newExemption, exemptFromDutyTypes: newExemption.exemptFromDutyTypes.filter(t => t !== type.name)});
                            }
                          }}
                          style={{cursor: 'pointer'}}
                        />
                        <span style={{fontWeight: 500}}>{type.name}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <button onClick={handleAddExemption} style={{...styles.btn('green'), width: 'fit-content'}}>
                  <Plus size={16} />
                  {t.addExemption}
                </button>
              </div>
            </div>

            <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
              {Object.entries(exemptionTypes).map(([key, exemption]) => (
                <div key={key} style={{display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : 'white', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')}}>
                  <Shield size={20} color="#f59e0b" />
                  <div style={{flex: 1}}>
                    <span style={{fontWeight: 600, fontSize: '16px'}}>{exemption.name}</span>
                    <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{exemption.description}</p>
                    {exemption.exemptFromDutyTypes.length > 0 && (
                      <div style={{marginTop: '8px', display: 'flex', flexWrap: 'wrap', gap: '6px'}}>
                        {exemption.exemptFromDutyTypes.map(dutyType => (
                          <span key={dutyType} style={{
                            fontSize: '12px', padding: '4px 8px', borderRadius: '4px',
                            background: darkMode ? '#ef4444' : '#fee2e2',
                            color: darkMode ? 'white' : '#991b1b',
                            fontWeight: 600
                          }}>
                            {dutyType}
                          </span>
                        ))}
                      </div>
                    )}
                  </div>
                  <button
                    onClick={() => handleDeleteExemption(key)}
                    style={{background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: '#ef4444'}}
                    title="Delete"
                  >
                    <Trash2 size={18} />
                  </button>
                </div>
              ))}
              {Object.keys(exemptionTypes).length === 0 && (
                <p style={{textAlign: 'center', color: darkMode ? '#9ca3af' : '#6b7280', padding: '24px'}}>
                  {t.noExemptions}
                </p>
              )}
            </div>
          </div>
        )}

        {showHolidaySettings && (
          <div style={{...styles.message, borderColor: '#8b5cf6', background: darkMode ? '#1e293b' : '#faf5ff', color: darkMode ? 'white' : '#1e293b'}}>
            <h3 style={{margin: '0 0 16px 0', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <Gift size={20} />
              {t.manageHolidays}
            </h3>
            <p style={{margin: '0 0 16px 0', fontSize: '14px', color: darkMode ? '#d1d5db' : '#6b7280'}}>
              {language === 'he' ? 'הגדר חגים ומשקל (קושי) לכל חג. המערכת תמנע מעובדים לקבל אותו חג שנתיים ברציפות' : 'Define holidays and their difficulty weight. The system will prevent employees from getting the same holiday two years in a row'}
            </p>

            <div style={{marginBottom: '16px', padding: '12px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb', border: '2px dashed ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
              <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                <div style={{display: 'grid', gridTemplateColumns: '2fr 1fr 2fr', gap: '8px'}}>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.holidayName}</label>
                    <input
                      type="text"
                      value={newHoliday.name}
                      onChange={(e) => setNewHoliday({...newHoliday, name: e.target.value})}
                      placeholder={language === 'he' ? 'לדוגמא: פסח' : 'e.g., Passover'}
                      style={styles.input}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.holidayWeight}</label>
                    <input
                      type="number"
                      step="0.1"
                      value={newHoliday.weight}
                      onChange={(e) => setNewHoliday({...newHoliday, weight: parseFloat(e.target.value)})}
                      placeholder="1.0"
                      style={styles.input}
                    />
                  </div>
                  <div>
                    <label style={{...styles.label, marginBottom: '4px'}}>{t.holidayDescription}</label>
                    <input
                      type="text"
                      value={newHoliday.description}
                      onChange={(e) => setNewHoliday({...newHoliday, description: e.target.value})}
                      placeholder={language === 'he' ? 'תיאור' : 'Description'}
                      style={styles.input}
                    />
                  </div>
                </div>
                <button onClick={handleAddHoliday} style={{...styles.btn('green'), width: 'fit-content'}}>
                  <Plus size={16} />
                  {t.addHoliday}
                </button>
              </div>
            </div>

            <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
              {Object.entries(holidays).map(([key, holiday]) => {
                const dates = getHolidayDatesForYear(holiday);
                return (
                  <div key={key} style={{display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : 'white', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb')}}>
                    <Gift size={20} color="#8b5cf6" />
                    <div style={{flex: 1}}>
                      <div style={{display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap'}}>
                        <span style={{fontWeight: 600, fontSize: '16px'}}>{holiday.name}</span>
                        <span style={{
                          fontSize: '12px', padding: '4px 8px', borderRadius: '4px',
                          background: darkMode ? '#8b5cf6' : '#ede9fe',
                          color: darkMode ? 'white' : '#6b21a8',
                          fontWeight: 600
                        }}>
                          {language === 'he' ? `משקל: ${holiday.weight}` : `Weight: ${holiday.weight}`}
                        </span>
                        {dates && (
                          <span style={{
                            fontSize: '12px', padding: '4px 8px', borderRadius: '4px',
                            background: darkMode ? '#3b82f6' : '#dbeafe',
                            color: darkMode ? 'white' : '#1e40af',
                            fontWeight: 600
                          }}>
                            📅 {dates.start} - {dates.end}
                          </span>
                        )}
                      </div>
                      <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{holiday.description}</p>
                    </div>
                    <button
                      onClick={() => handleDeleteHoliday(key)}
                      style={{background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: '#ef4444'}}
                      title="Delete"
                    >
                      <Trash2 size={18} />
                    </button>
                  </div>
                );
              })}
              {Object.keys(holidays).length === 0 && (
                <p style={{textAlign: 'center', color: darkMode ? '#9ca3af' : '#6b7280', padding: '24px'}}>
                  {t.noHolidays}
                </p>
              )}
            </div>
          </div>
        )}

        {deleteConfirm && (
          <div style={styles.confirmDialog}>
            <div style={styles.confirmBox}>
              <h3 style={{margin: '0 0 16px 0', color: darkMode ? 'white' : '#1f2937'}}>{t.deleteConfirm}</h3>
              <p style={{margin: '0 0 24px 0', color: darkMode ? '#d1d5db' : '#6b7280'}}>
                {deleteConfirm.type === 'employee'
                  ? `${t.deleteEmployeeMsg} ${deleteConfirm.name}?`
                  : t.deleteShiftMsg}
              </p>
              <div style={{display: 'flex', gap: '12px', justifyContent: 'flex-end'}}>
                <button onClick={() => setDeleteConfirm(null)} style={{...styles.btn('gray'), padding: '8px 16px'}}>
                  {t.cancel}
                </button>
                <button
                  onClick={() => deleteConfirm.type === 'employee'
                    ? handleDeleteEmployee(deleteConfirm.id)
                    : handleDeleteShift(deleteConfirm.id)}
                  style={{...styles.btn('red'), padding: '8px 16px'}}
                >
                  {t.delete}
                </button>
              </div>
            </div>
          </div>
        )}

        {resetConfirm && (
          <div style={styles.confirmDialog}>
            <div style={styles.confirmBox}>
              <h3 style={{margin: '0 0 16px 0', color: darkMode ? 'white' : '#1f2937', display: 'flex', alignItems: 'center', gap: '8px'}}>
                <AlertCircle size={24} color="#ef4444" />
                {t.resetConfirm}
              </h3>
              <p style={{margin: '0 0 24px 0', color: darkMode ? '#d1d5db' : '#6b7280', lineHeight: '1.5'}}>
                {t.resetWarning}
              </p>
              <div style={{display: 'flex', gap: '12px', justifyContent: 'flex-end'}}>
                <button onClick={() => setResetConfirm(false)} style={{...styles.btn('gray'), padding: '8px 16px'}}>
                  {t.cancel}
                </button>
                <button
                  onClick={handleResetSystem}
                  style={{...styles.btn('red'), padding: '8px 16px'}}
                >
                  <Trash2 size={16} style={{display: 'inline', marginLeft: language === 'he' ? '8px' : '0', marginRight: language === 'en' ? '8px' : '0'}} />
                  {t.resetButton}
                </button>
              </div>
            </div>
          </div>
        )}

        <div style={styles.tabs}>
          <button onClick={() => setActiveTab('employees')} style={styles.tab(activeTab === 'employees')}>
            <Users size={20} style={{display: 'inline', marginRight: '8px'}} />
            {t.employees}
          </button>
          <button onClick={() => setActiveTab('schedule')} style={styles.tab(activeTab === 'schedule')}>
            <Clock size={20} style={{display: 'inline', marginRight: '8px'}} />
            {t.schedule}
          </button>
          <button onClick={() => setActiveTab('reports')} style={styles.tab(activeTab === 'reports')}>
            <BarChart3 size={20} style={{display: 'inline', marginRight: '8px'}} />
            {t.reports}
          </button>
        </div>

        <div style={styles.content}>
          {activeTab === 'employees' && (
            <div className="tab-content">
              <div style={styles.flexBetween}>
                <h2 style={styles.h2}>{t.employeeManagement}</h2>
                <div style={styles.btnGroup}>
                  <button onClick={handleDownloadTemplate} style={styles.btn('green')}>
                    <Download size={20} />
                    {t.template}
                  </button>
                  <label style={{...styles.btn('indigo'), cursor: 'pointer'}}>
                    <Upload size={20} />
                    {t.import}
                    <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{display: 'none'}} />
                  </label>
                  <button onClick={() => setShowAddEmployee(true)} style={styles.btn('blue')}>
                    <Plus size={20} />
                    {t.add}
                  </button>
                </div>
              </div>

              <div style={styles.grid3}>
                <div style={styles.inputWrapper}>
                  <Search size={20} style={{position: 'absolute', [language === 'he' ? 'right' : 'left']: '12px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af'}} />
                  <input type="text" placeholder={t.searchPlaceholder} value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} style={styles.input} />
                </div>
                <select value={filterDepartment} onChange={(e) => setFilterDepartment(e.target.value)} style={styles.select}>
                  <option value="">{t.allDepartments}</option>
                  {departments.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
                <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} style={styles.select}>
                  <option value="">{t.allStatuses}</option>
                  {statuses.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              {(showAddEmployee || editingEmployee) && (
                <div style={styles.modal}>
                  <div style={styles.modalHeader}>
                    <h3 style={styles.modalTitle}>{editingEmployee ? t.editEmployee : t.newEmployee}</h3>
                    <button onClick={() => { setShowAddEmployee(false); setEditingEmployee(null); }} style={{background: 'none', border: 'none', cursor: 'pointer'}}>
                      <X size={20} color={darkMode ? 'white' : 'black'} />
                    </button>
                  </div>
                  <div style={styles.grid2}>
                    <div>
                      <label style={styles.label}>{t.name} *</label>
                      <input
                        type="text"
                        value={editingEmployee ? editingEmployee.name : newEmployee.name}
                        onChange={(e) => editingEmployee
                          ? setEditingEmployee({...editingEmployee, name: e.target.value})
                          : setNewEmployee({...newEmployee, name: e.target.value})}
                        style={styles.input}
                      />
                    </div>
                    <div>
                      <label style={styles.label}>{t.personalNumber} *</label>
                      <input
                        type="text"
                        value={editingEmployee ? editingEmployee.personalNumber : newEmployee.personalNumber}
                        onChange={(e) => editingEmployee
                          ? setEditingEmployee({...editingEmployee, personalNumber: e.target.value})
                          : setNewEmployee({...newEmployee, personalNumber: e.target.value})}
                        style={styles.input}
                      />
                    </div>
                    <div>
                      <label style={styles.label}>{t.sex}</label>
                      <select
                        value={editingEmployee ? editingEmployee.sex : newEmployee.sex}
                        onChange={(e) => editingEmployee
                          ? setEditingEmployee({...editingEmployee, sex: e.target.value})
                          : setNewEmployee({...newEmployee, sex: e.target.value})}
                        style={styles.select}
                      >
                        <option value="">{t.select}</option>
                        <option value="Male">{t.male}</option>
                        <option value="Female">{t.female}</option>
                        <option value="Other">{t.other}</option>
                      </select>
                    </div>
                    <div>
                      <label style={styles.label}>{t.status}</label>
                      <select
                        value={editingEmployee ? editingEmployee.status : newEmployee.status}
                        onChange={(e) => editingEmployee
                          ? setEditingEmployee({...editingEmployee, status: e.target.value})
                          : setNewEmployee({...newEmployee, status: e.target.value})}
                        style={styles.select}
                      >
                        <option value="">{t.select}</option>
                        {Object.keys(statusMultipliers).map(s => <option key={s} value={s}>{s} (×{statusMultipliers[s].toFixed(1)})</option>)}
                      </select>
                    </div>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={styles.label}>{t.department}</label>
                      <input
                        type="text"
                        value={editingEmployee ? editingEmployee.department : newEmployee.department}
                        onChange={(e) => editingEmployee
                          ? setEditingEmployee({...editingEmployee, department: e.target.value})
                          : setNewEmployee({...newEmployee, department: e.target.value})}
                        placeholder={t.deptPlaceholder}
                        style={styles.input}
                      />
                    </div>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <Shield size={16} />
                        {t.exemptions}
                      </label>
                      <div style={{display: 'flex', flexWrap: 'wrap', gap: '8px', padding: '12px', borderRadius: '8px', background: darkMode ? '#374151' : '#f9fafb', border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'), minHeight: '48px'}}>
                        {Object.keys(exemptionTypes).length === 0 ? (
                          <span style={{color: darkMode ? '#9ca3af' : '#6b7280', fontSize: '14px'}}>
                            {t.noExemptions} - Add exemptions in settings first
                          </span>
                        ) : (
                          Object.entries(exemptionTypes).map(([key, exemption]) => (
                            <label key={key} style={{display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '6px 12px', borderRadius: '6px', background: (editingEmployee ? editingEmployee.exemptions : newEmployee.exemptions)?.includes(key) ? (darkMode ? '#f59e0b' : '#fed7aa') : 'transparent', border: '1px solid ' + (darkMode ? '#4b5563' : '#d1d5db')}}>
                              <input
                                type="checkbox"
                                checked={(editingEmployee ? editingEmployee.exemptions : newEmployee.exemptions)?.includes(key) || false}
                                onChange={(e) => {
                                  const currentExemptions = (editingEmployee ? editingEmployee.exemptions : newEmployee.exemptions) || [];
                                  const newExemptions = e.target.checked
                                    ? [...currentExemptions, key]
                                    : currentExemptions.filter(ex => ex !== key);

                                  if (editingEmployee) {
                                    setEditingEmployee({...editingEmployee, exemptions: newExemptions});
                                  } else {
                                    setNewEmployee({...newEmployee, exemptions: newExemptions});
                                  }
                                }}
                                style={{cursor: 'pointer'}}
                              />
                              <span style={{fontWeight: 500}}>{exemption.name}</span>
                              {exemption.exemptFromDutyTypes.length > 0 && (
                                <span style={{fontSize: '11px', color: darkMode ? '#9ca3af' : '#6b7280'}}>
                                  ({exemption.exemptFromDutyTypes.join(', ')})
                                </span>
                              )}
                            </label>
                          ))
                        )}
                      </div>
                    </div>
                  </div>
                  <button
                    onClick={editingEmployee ? handleEditEmployee : handleAddEmployee}
                    style={{...styles.btn('blue'), width: '100%', marginTop: '16px', justifyContent: 'center'}}
                  >
                    {editingEmployee ? t.editEmployee : t.addEmployee}
                  </button>
                </div>
              )}

              {filteredEmployees.length === 0 ? (
                <div style={styles.emptyState}>
                  <Users size={64} color="#9ca3af" style={{margin: '0 auto 16px'}} />
                  <p style={{color: darkMode ? '#d1d5db' : '#6b7280', fontSize: '18px', margin: '8px 0'}}>{employees.length === 0 ? t.noEmployees : t.noMatching}</p>
                  <p style={{color: darkMode ? '#9ca3af' : '#9ca3af', fontSize: '14px'}}>{employees.length === 0 ? t.importOrAdd : t.adjustFilters}</p>
                </div>
              ) : (
                <table style={styles.table}>
                  <thead style={styles.thead}>
                    <tr>
                      <th style={styles.th}>{t.name}</th>
                      <th style={styles.th}>{t.personalNumber}</th>
                      <th style={styles.th}>{t.sex}</th>
                      <th style={styles.th}>{t.status}</th>
                      <th style={styles.th}>{t.department}</th>
                      <th style={{...styles.th, textAlign: 'center'}}>{t.actions}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredEmployees.map(emp => {
                      const multiplier = statusMultipliers[emp.status] || 1.0;
                      return (
                        <tr key={emp.id} style={styles.tr} onMouseEnter={(e) => e.currentTarget.style.background = darkMode ? '#374151' : '#f9fafb'} onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}>
                          <td style={styles.td}>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                              <div style={styles.avatar}>{emp.name.charAt(0)}</div>
                              <span style={{fontWeight: 500, color: darkMode ? 'white' : '#1f2937'}}>{emp.name}</span>
                            </div>
                          </td>
                          <td style={styles.td}>{emp.personalNumber}</td>
                          <td style={styles.td}>{emp.sex || '-'}</td>
                          <td style={styles.td}>
                            <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                              <span style={styles.badge('green')}>{emp.status || 'N/A'}</span>
                              {emp.status && <span style={{fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>×{multiplier.toFixed(1)}</span>}
                            </div>
                          </td>
                          <td style={styles.td}>{emp.department || '-'}</td>
                          <td style={{...styles.td, textAlign: 'center'}}>
                            <div style={{display: 'flex', gap: '8px', justifyContent: 'center'}}>
                              <button onClick={() => setEditingEmployee(emp)} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#3b82f6'}}>
                                <Edit2 size={18} />
                              </button>
                              <button onClick={() => setDeleteConfirm({type: 'employee', id: emp.id, name: emp.name})} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#ef4444'}}>
                                <Trash2 size={18} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {activeTab === 'reports' && (
            <div className="tab-content">
              <div style={styles.flexBetween}>
                <h2 style={styles.h2}>{t.reportsTitle}</h2>
                <button onClick={handleYearlyExport} style={styles.btn('green')} disabled={shifts.length === 0}>
                  <Download size={20} />
                  {t.yearlyExport}
                </button>
              </div>

              {employees.length === 0 || shifts.length === 0 ? (
                <div style={styles.emptyState}>
                  <BarChart3 size={64} color="#9ca3af" style={{margin: '0 auto 16px'}} />
                  <p style={{color: darkMode ? '#d1d5db' : '#6b7280', fontSize: '18px', margin: '8px 0'}}>{t.noData}</p>
                  <p style={{color: darkMode ? '#9ca3af' : '#9ca3af', fontSize: '14px'}}>
                    {employees.length === 0 ? t.addEmployeesFirst : 'הוסף תורנויות כדי לראות דוחות'}
                  </p>
                </div>
              ) : (
                <>
                  <div style={{marginBottom: '24px'}}>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px'}}>
                      טבלת סוגי תורנויות וניקוד
                    </h3>
                    <div style={styles.statsGrid}>
                      {Object.values(dutyTypePoints).map(type => (
                        <div key={type.name} style={{...styles.statBox, borderLeft: `4px solid ${type.color}`}}>
                          <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                            <div>
                              <h4 style={{margin: 0, fontSize: '16px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{type.name}</h4>
                              <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{type.description}</p>
                            </div>
                            <div style={{textAlign: 'center'}}>
                              <Award size={24} color={type.color} />
                              <p style={{margin: '4px 0 0 0', fontSize: '14px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{type.basePoints} נק׳ בסיס</p>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div style={{marginBottom: '24px'}}>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px'}}>
                      בונוסים לפי מועד
                    </h3>
                    <div style={styles.statsGrid}>
                      {Object.values(dateTypePoints).map(type => (
                        <div key={type.name} style={{...styles.statBox, borderLeft: `4px solid ${type.color}`}}>
                          <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                            <div>
                              <h4 style={{margin: 0, fontSize: '16px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{type.name}</h4>
                              <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{type.description}</p>
                            </div>
                            <div style={{textAlign: 'center'}}>
                              <p style={{margin: 0, fontSize: '14px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>+{type.bonus} נק׳</p>
                            </div>
                          </div>
                        </div>
                      ))}
                      <div style={{...styles.statBox, borderLeft: '4px solid #a855f7'}}>
                        <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                          <div>
                            <h4 style={{margin: 0, fontSize: '16px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>חג בשבת</h4>
                            <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>זיהוי אוטומטי</p>
                          </div>
                          <div style={{textAlign: 'center'}}>
                            <p style={{margin: 0, fontSize: '14px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>+4 נק׳</p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div style={{marginBottom: '24px'}}>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '8px'}}>
                      <TrendingUp size={20} />
                      {language === 'he' ? 'השוואת עומסים' : 'Load Comparison'}
                    </h3>
                    <div style={{display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '16px'}}>
                      {/* Most shifts */}
                      {(() => {
                        const sorted = [...employeeStats].sort((a, b) => b.totalShifts - a.totalShifts);
                        const top = sorted[0];
                        if (!top) return null;
                        return (
                          <div style={{background: darkMode ? '#1f2937' : '#fef3c7', padding: '20px', borderRadius: '12px', border: `2px solid ${darkMode ? '#f59e0b' : '#fbbf24'}`}}>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '12px'}}>
                              <Award size={24} color="#f59e0b" />
                              <span style={{fontSize: '16px', fontWeight: 600, color: darkMode ? '#fbbf24' : '#92400e'}}>
                                {language === 'he' ? 'הכי הרבה תורנויות' : 'Most Shifts'}
                              </span>
                            </div>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                              <div style={{...styles.avatar, width: '48px', height: '48px', fontSize: '20px'}}>{top.name.charAt(0)}</div>
                              <div>
                                <p style={{margin: 0, fontSize: '18px', fontWeight: 700, color: darkMode ? 'white' : '#1f2937'}}>{top.name}</p>
                                <p style={{margin: '4px 0 0 0', fontSize: '24px', fontWeight: 800, color: '#f59e0b'}}>{top.totalShifts} {language === 'he' ? 'תורנויות' : 'shifts'}</p>
                                <p style={{margin: '4px 0 0 0', fontSize: '13px', color: darkMode ? '#d1d5db' : '#78350f'}}>
                                  {language === 'he' ? `ניקוד צדק: ${top.justicePoints.toFixed(1)}` : `Justice points: ${top.justicePoints.toFixed(1)}`}
                                </p>
                              </div>
                            </div>
                          </div>
                        );
                      })()}

                      {/* Least shifts */}
                      {(() => {
                        const sorted = [...employeeStats].sort((a, b) => a.totalShifts - b.totalShifts);
                        const bottom = sorted[0];
                        if (!bottom) return null;
                        return (
                          <div style={{background: darkMode ? '#1f2937' : '#dbeafe', padding: '20px', borderRadius: '12px', border: `2px solid ${darkMode ? '#3b82f6' : '#60a5fa'}`}}>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '12px'}}>
                              <Shield size={24} color="#3b82f6" />
                              <span style={{fontSize: '16px', fontWeight: 600, color: darkMode ? '#60a5fa' : '#1e40af'}}>
                                {language === 'he' ? 'הכי מעט תורנויות' : 'Least Shifts'}
                              </span>
                            </div>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                              <div style={{...styles.avatar, width: '48px', height: '48px', fontSize: '20px'}}>{bottom.name.charAt(0)}</div>
                              <div>
                                <p style={{margin: 0, fontSize: '18px', fontWeight: 700, color: darkMode ? 'white' : '#1f2937'}}>{bottom.name}</p>
                                <p style={{margin: '4px 0 0 0', fontSize: '24px', fontWeight: 800, color: '#3b82f6'}}>{bottom.totalShifts} {language === 'he' ? 'תורנויות' : 'shifts'}</p>
                                <p style={{margin: '4px 0 0 0', fontSize: '13px', color: darkMode ? '#d1d5db' : '#1e3a8a'}}>
                                  {language === 'he' ? `ניקוד צדק: ${bottom.justicePoints.toFixed(1)}` : `Justice points: ${bottom.justicePoints.toFixed(1)}`}
                                </p>
                              </div>
                            </div>
                          </div>
                        );
                      })()}

                      {/* Average */}
                      {(() => {
                        const average = employeeStats.reduce((sum, emp) => sum + emp.totalShifts, 0) / employeeStats.length;
                        return (
                          <div style={{background: darkMode ? '#1f2937' : '#f3e8ff', padding: '20px', borderRadius: '12px', border: `2px solid ${darkMode ? '#a855f7' : '#c084fc'}`}}>
                            <div style={{display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '12px'}}>
                              <BarChart3 size={24} color="#a855f7" />
                              <span style={{fontSize: '16px', fontWeight: 600, color: darkMode ? '#c084fc' : '#6b21a8'}}>
                                {language === 'he' ? 'ממוצע תורנויות' : 'Average Shifts'}
                              </span>
                            </div>
                            <div>
                              <p style={{margin: 0, fontSize: '24px', fontWeight: 800, color: '#a855f7'}}>{average.toFixed(1)} {language === 'he' ? 'תורנויות' : 'shifts'}</p>
                              <p style={{margin: '8px 0 0 0', fontSize: '13px', color: darkMode ? '#d1d5db' : '#581c87'}}>
                                {language === 'he' ? `סה"כ ${shifts.filter(s => s.employeeId).length} תורנויות משובצות` : `Total ${shifts.filter(s => s.employeeId).length} assigned shifts`}
                              </p>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  </div>

                  <div style={{marginBottom: '24px'}}>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '8px'}}>
                      <BarChart3 size={20} />
                      {language === 'he' ? 'תרשים תורנויות לפי חודש' : 'Shifts by Month'}
                    </h3>
                    {(() => {
                      // Calculate shifts by month
                      const monthlyData = {};
                      shifts.filter(s => s.employeeId).forEach(shift => {
                        const date = new Date(shift.startDate);
                        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                        const monthName = date.toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { year: 'numeric', month: 'short' });

                        if (!monthlyData[monthKey]) {
                          monthlyData[monthKey] = { name: monthName, count: 0 };
                        }
                        monthlyData[monthKey].count++;
                      });

                      const sortedMonths = Object.values(monthlyData).sort((a, b) => a.name.localeCompare(b.name));
                      const maxCount = Math.max(...sortedMonths.map(m => m.count), 1);

                      if (sortedMonths.length === 0) {
                        return (
                          <div style={{padding: '40px', textAlign: 'center', background: darkMode ? '#1f2937' : '#f9fafb', borderRadius: '12px'}}>
                            <BarChart3 size={48} color="#9ca3af" style={{margin: '0 auto 12px'}} />
                            <p style={{color: darkMode ? '#9ca3af' : '#6b7280'}}>{language === 'he' ? 'אין נתונים להצגה' : 'No data to display'}</p>
                          </div>
                        );
                      }

                      return (
                        <div style={{background: darkMode ? '#1f2937' : '#f9fafb', padding: '24px', borderRadius: '12px', border: `1px solid ${darkMode ? '#374151' : '#e5e7eb'}`}}>
                          <div style={{display: 'flex', flexDirection: 'column', gap: '16px'}}>
                            {sortedMonths.map((month, index) => {
                              const barWidth = (month.count / maxCount) * 100;
                              const colors = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#14b8a6'];
                              const color = colors[index % colors.length];

                              return (
                                <div key={month.name} style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                                  <div style={{minWidth: '100px', fontWeight: 500, fontSize: '14px', color: darkMode ? 'white' : '#1f2937'}}>
                                    {month.name}
                                  </div>
                                  <div style={{flex: 1, height: '32px', background: darkMode ? '#374151' : '#e5e7eb', borderRadius: '6px', overflow: 'hidden', position: 'relative'}}>
                                    <div style={{
                                      width: `${barWidth}%`,
                                      height: '100%',
                                      background: `linear-gradient(90deg, ${color}, ${color}dd)`,
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'flex-end',
                                      paddingRight: '12px',
                                      transition: 'width 0.3s ease',
                                      borderRadius: '6px'
                                    }}>
                                      <span style={{fontSize: '14px', fontWeight: 700, color: 'white', textShadow: '0 1px 2px rgba(0,0,0,0.3)'}}>
                                        {month.count}
                                      </span>
                                    </div>
                                  </div>
                                  <div style={{minWidth: '80px', textAlign: 'right', fontSize: '13px', color: darkMode ? '#9ca3af' : '#6b7280'}}>
                                    {((month.count / shifts.filter(s => s.employeeId).length) * 100).toFixed(1)}%
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })()}
                  </div>

                  <div>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px'}}>
                      {t.employeeStats}
                    </h3>
                    <table style={styles.table}>
                      <thead style={styles.thead}>
                        <tr>
                          <th style={styles.th}>{t.name}</th>
                          <th style={styles.th}>{t.status}</th>
                          <th style={{...styles.th, textAlign: 'center'}}>{t.totalShifts}</th>
                          {Object.keys(dutyTypePoints).map(type => (
                            <th key={type} style={{...styles.th, textAlign: 'center'}}>{type}</th>
                          ))}
                          <th style={{...styles.th, textAlign: 'center'}}>{t.justicePoints}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {employeeStats.sort((a, b) => a.justicePoints - b.justicePoints).map(emp => (
                          <tr key={emp.id} style={styles.tr} onMouseEnter={(e) => e.currentTarget.style.background = darkMode ? '#374151' : '#f9fafb'} onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}>
                            <td style={styles.td}>
                              <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                                <div style={styles.avatar}>{emp.name.charAt(0)}</div>
                                <span style={{fontWeight: 500, color: darkMode ? 'white' : '#1f2937'}}>{emp.name}</span>
                              </div>
                            </td>
                            <td style={styles.td}>
                              <span style={styles.badge('green')}>{emp.status}</span>
                              {emp.status && <span style={{fontSize: '11px', marginRight: '4px', color: darkMode ? '#9ca3af' : '#6b7280'}}>×{(statusMultipliers[emp.status] || 1.0).toFixed(1)}</span>}
                            </td>
                            <td style={{...styles.td, textAlign: 'center', fontWeight: 600}}>{emp.totalShifts}</td>
                            {Object.keys(dutyTypePoints).map(type => (
                              <td key={type} style={{...styles.td, textAlign: 'center'}}>
                                {emp.dutyTypeCounts[type] || 0}
                              </td>
                            ))}
                            <td style={{...styles.td, textAlign: 'center'}}>
                              <span style={{...styles.badge('blue'), fontSize: '14px', fontWeight: 600}}>
                                {emp.justicePoints.toFixed(1)}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}

          {activeTab === 'schedule' && (
            <div className="tab-content">
              <div style={styles.flexBetween}>
                <h2 style={styles.h2}>{t.shiftSchedule}</h2>
                <div style={styles.btnGroup}>
                  <div style={styles.calViewBtns}>
                    <button onClick={() => setCalendarView('list')} style={styles.calViewBtn(calendarView === 'list')}>
                      {t.list}
                    </button>
                    <button onClick={() => setCalendarView('week')} style={styles.calViewBtn(calendarView === 'week')}>
                      {t.week}
                    </button>
                    <button onClick={() => setCalendarView('month')} style={styles.calViewBtn(calendarView === 'month')}>
                      {t.month}
                    </button>
                  </div>
                  <button onClick={handleDownloadShiftsTemplate} style={styles.btn('green')}>
                    <Download size={20} />
                    {t.shiftsTemplate}
                  </button>
                  <label style={{...styles.btn('indigo'), cursor: 'pointer'}}>
                    <Upload size={20} />
                    {t.importShifts}
                    <input type="file" accept=".xlsx,.xls" onChange={handleShiftsFileUpload} style={{display: 'none'}} />
                  </label>
                  <button onClick={() => handleAutoDistribute(true)} style={styles.btn('orange')} disabled={employees.length === 0}>
                    <Zap size={20} />
                    {t.autoDistribute}
                  </button>
                  <button onClick={handleExportShifts} style={styles.btn('green')} disabled={shifts.length === 0}>
                    <Download size={20} />
                    {t.export}
                  </button>
                  <button onClick={() => window.print()} style={styles.btn('purple')}>
                    <Printer size={20} />
                    {t.print}
                  </button>
                  <button onClick={handleExportPDF} style={styles.btn('red')} disabled={shifts.length === 0}>
                    <FileText size={20} />
                    {t.exportPDF}
                  </button>
                  <button onClick={() => setShowAddShift(true)} style={styles.btn('blue')}>
                    <Plus size={20} />
                    {t.addShift}
                  </button>
                </div>
              </div>

              <div style={styles.grid3}>
                <select value={filterEmployee} onChange={(e) => setFilterEmployee(e.target.value)} style={styles.select}>
                  <option value="">{t.allEmployees}</option>
                  {employees.map(emp => <option key={emp.id} value={emp.id}>{emp.name}</option>)}
                </select>
                <select value={filterShiftType} onChange={(e) => setFilterShiftType(e.target.value)} style={styles.select}>
                  <option value="">{t.allShiftTypes}</option>
                  {Object.keys(dutyTypePoints).map(type => <option key={type} value={type}>{type}</option>)}
                </select>
              </div>

              {conflictingShifts.size > 0 && (
                <div style={{...styles.message, borderColor: '#f59e0b', background: '#fef3c7', color: '#92400e', marginBottom: '16px'}}>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <AlertCircle size={20} />
                    {conflictingShifts.size} {t.conflictDetected}
                  </div>
                </div>
              )}

              {showAddShift && (
                <div style={styles.modal}>
                  <div style={styles.modalHeader}>
                    <h3 style={styles.modalTitle}>{t.newShift}</h3>
                    <button onClick={() => setShowAddShift(false)} style={{background: 'none', border: 'none', cursor: 'pointer'}}>
                      <X size={20} color={darkMode ? 'white' : 'black'} />
                    </button>
                  </div>
                  <div style={styles.grid2}>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={styles.label}>{t.employee} ({t.optional})</label>
                      <select value={newShift.employeeId} onChange={(e) => setNewShift({...newShift, employeeId: e.target.value})} style={styles.select}>
                        <option value="">{t.noEmployee}</option>
                        {employees.map(emp => (
                          <option key={emp.id} value={emp.id}>{emp.name} - {emp.status}</option>
                        ))}
                      </select>
                      {newShift.employeeId && newShift.role && (() => {
                        const selectedEmp = employees.find(e => e.id === parseInt(newShift.employeeId));
                        if (selectedEmp && isEmployeeExemptFromDuty(selectedEmp, newShift.role)) {
                          return (
                            <div style={{display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', marginTop: '8px', borderRadius: '6px', background: '#fef3c7', border: '1px solid #fbbf24', color: '#92400e'}}>
                              <AlertCircle size={18} style={{flexShrink: 0}} />
                              <span style={{fontSize: '13px', fontWeight: 500}}>
                                {language === 'he' ? 'אזהרה: לעובד יש פטור מתורנות מסוג זה' : 'Warning: Employee has exemption for this duty type'}
                              </span>
                            </div>
                          );
                        }
                        return null;
                      })()}
                    </div>
                    <div>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <Calendar size={16} />
                        תאריך התחלה
                        <span style={{color: '#ef4444', fontWeight: 700}}>*</span>
                      </label>
                      <input
                        type="date"
                        value={newShift.startDate}
                        onChange={(e) => {
                          const newStartDate = e.target.value;
                          const autoEndDate = calculateEndDate(newStartDate, newShift.dateType);
                          setNewShift({...newShift, startDate: newStartDate, endDate: autoEndDate});
                        }}
                        style={{
                          ...styles.dateInput,
                          borderColor: !newShift.startDate ? '#ef4444' : (darkMode ? '#4b5563' : '#d1d5db')
                        }}
                        required
                      />
                    </div>
                    <div>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <Calendar size={16} />
                        תאריך סיום ({t.optional})
                      </label>
                      <input
                        type="date"
                        value={newShift.endDate}
                        onChange={(e) => setNewShift({...newShift, endDate: e.target.value})}
                        placeholder="יחושב אוטומטית"
                        style={styles.dateInput}
                      />
                    </div>
                    <div>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        {t.dutyType}
                        <span style={{color: '#ef4444', fontWeight: 700}}>*</span>
                      </label>
                      <select
                        value={newShift.dutyType}
                        onChange={(e) => setNewShift({...newShift, dutyType: e.target.value})}
                        style={{
                          ...styles.select,
                          borderColor: !newShift.dutyType ? '#ef4444' : (darkMode ? '#4b5563' : '#d1d5db')
                        }}
                        required
                      >
                        <option value="">לחץ לבחירה 👇</option>
                        {Object.values(dutyTypePoints).map(type => (
                          <option key={type.name} value={type.name}>
                            {type.name} ({type.basePoints} נק׳ בסיס)
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        {t.dateType}
                        <span style={{color: '#ef4444', fontWeight: 700}}>*</span>
                      </label>
                      <select
                        value={newShift.dateType}
                        onChange={(e) => {
                          const newDateType = e.target.value;
                          const autoEndDate = newShift.startDate && !newShift.endDate
                            ? calculateEndDate(newShift.startDate, newDateType)
                            : newShift.endDate;
                          setNewShift({...newShift, dateType: newDateType, endDate: autoEndDate});
                        }}
                        style={{
                          ...styles.select,
                          borderColor: !newShift.dateType ? '#ef4444' : (darkMode ? '#4b5563' : '#d1d5db')
                        }}
                        required
                      >
                        <option value="">לחץ לבחירה 👇</option>
                        {Object.values(dateTypePoints).map(type => (
                          <option key={type.name} value={type.name}>
                            {type.name} (+{type.bonus} נק׳)
                          </option>
                        ))}
                      </select>
                    </div>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={{...styles.label, display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <Gift size={16} />
                        {t.assignedHoliday} ({t.optional})
                      </label>
                      <select value={newShift.holidayName} onChange={(e) => setNewShift({...newShift, holidayName: e.target.value})} style={styles.select}>
                        <option value="">{language === 'he' ? 'אין חג' : 'No Holiday'}</option>
                        {Object.values(holidays).map(holiday => (
                          <option key={holiday.name} value={holiday.name}>
                            {holiday.name} (משקל: {holiday.weight})
                          </option>
                        ))}
                      </select>
                      {newShift.startDate && (() => {
                        const detectedHoliday = getHolidayForDate(newShift.startDate);
                        if (detectedHoliday && !newShift.holidayName) {
                          return (
                            <div style={{display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', marginTop: '8px', borderRadius: '6px', background: '#dbeafe', border: '1px solid #3b82f6', color: '#1e40af'}}>
                              <Gift size={18} style={{flexShrink: 0}} />
                              <span style={{fontSize: '13px', fontWeight: 500}}>
                                {language === 'he' ? `זוהה אוטומטית: ${detectedHoliday}` : `Auto-detected: ${detectedHoliday}`}
                              </span>
                            </div>
                          );
                        }
                        return null;
                      })()}
                    </div>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={styles.label}>{t.manualOverride} ({t.optional})</label>
                      <input
                        type="number"
                        step="0.1"
                        value={newShift.manualPoints || ''}
                        onChange={(e) => setNewShift({...newShift, manualPoints: e.target.value ? parseFloat(e.target.value) : null})}
                        placeholder="השאר ריק לחישוב אוטומטי"
                        style={styles.input}
                      />
                    </div>
                  </div>
                  <button onClick={handleAddShift} style={{...styles.btn('blue'), width: '100%', marginTop: '16px', justifyContent: 'center', padding: '16px', fontSize: '16px', fontWeight: 600}}>
                    <Plus size={20} style={{marginLeft: language === 'he' ? '8px' : '0', marginRight: language === 'en' ? '8px' : '0'}} />
                    {t.addShift}
                  </button>
                </div>
              )}

              {swappingShift && (
                <div style={styles.modal}>
                  <div style={styles.modalHeader}>
                    <h3 style={styles.modalTitle}>{t.swapShiftTitle}</h3>
                    <button onClick={() => setSwappingShift(null)} style={{background: 'none', border: 'none', cursor: 'pointer'}}>
                      <X size={20} color={darkMode ? 'white' : 'black'} />
                    </button>
                  </div>
                  <p style={{marginBottom: '16px', color: darkMode ? '#d1d5db' : '#6b7280'}}>
                    {t.selectShiftToSwap}
                  </p>
                  <div style={{maxHeight: '400px', overflowY: 'auto'}}>
                    {shifts.filter(s => s.id !== swappingShift.id && s.date === swappingShift.date).map(shift => {
                      const emp = getEmployee(shift.employeeId);
                      const dutyType = dutyTypePoints[shift.dutyType] || dutyTypePoints['גלגלת'];
                      return (
                        <div
                          key={shift.id}
                          onClick={() => handleSwapShifts(swappingShift.id, shift.id)}
                          style={{
                            padding: '12px',
                            marginBottom: '8px',
                            borderRadius: '8px',
                            border: `1px solid ${dutyType.color}`,
                            background: dutyType.color + '10',
                            cursor: 'pointer',
                            transition: 'all 0.2s'
                          }}
                          onMouseEnter={(e) => e.currentTarget.style.background = dutyType.color + '30'}
                          onMouseLeave={(e) => e.currentTarget.style.background = dutyType.color + '10'}
                        >
                          <div style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{emp?.name || t.unassigned}</div>
                          <div style={{fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280', marginTop: '4px'}}>
                            {new Date(shift.date).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'short', day: 'numeric' })} • {shift.dutyType}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {dryRunPreview && (
                <div style={styles.modal}>
                  <div style={styles.modalHeader}>
                    <h3 style={{...styles.modalTitle, display: 'flex', alignItems: 'center', gap: '8px'}}>
                      <Zap size={24} color="#f97316" />
                      {t.distributionPreview}
                    </h3>
                    <button onClick={() => setDryRunPreview(null)} style={{background: 'none', border: 'none', cursor: 'pointer'}}>
                      <X size={20} color={darkMode ? 'white' : 'black'} />
                    </button>
                  </div>
                  <div style={{marginBottom: '20px'}}>
                    <p style={{color: darkMode ? '#d1d5db' : '#6b7280', marginBottom: '16px', fontSize: '14px'}}>
                      {t.previewBeforeApply}
                    </p>
                    <div style={{display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px', marginBottom: '20px'}}>
                      <div style={{padding: '16px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#ecfdf5', border: `2px solid ${darkMode ? '#10b981' : '#6ee7b7'}`}}>
                        <div style={{fontSize: '14px', color: darkMode ? '#9ca3af' : '#6b7280', marginBottom: '4px'}}>
                          {t.distributed}
                        </div>
                        <div style={{fontSize: '28px', fontWeight: 800, color: '#10b981'}}>
                          {dryRunPreview.distributedCount}
                        </div>
                      </div>
                      <div style={{padding: '16px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#fef3c7', border: `2px solid ${darkMode ? '#f59e0b' : '#fbbf24'}`}}>
                        <div style={{fontSize: '14px', color: darkMode ? '#9ca3af' : '#6b7280', marginBottom: '4px'}}>
                          {t.unassigned}
                        </div>
                        <div style={{fontSize: '28px', fontWeight: 800, color: '#f59e0b'}}>
                          {dryRunPreview.unassignedCount}
                        </div>
                      </div>
                      <div style={{padding: '16px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#dbeafe', border: `2px solid ${darkMode ? '#3b82f6' : '#60a5fa'}`}}>
                        <div style={{fontSize: '14px', color: darkMode ? '#9ca3af' : '#6b7280', marginBottom: '4px'}}>
                          {language === 'he' ? 'סה"כ' : 'Total'}
                        </div>
                        <div style={{fontSize: '28px', fontWeight: 800, color: '#3b82f6'}}>
                          {dryRunPreview.distributedCount + dryRunPreview.unassignedCount}
                        </div>
                      </div>
                    </div>
                    <div style={{maxHeight: '400px', overflowY: 'auto', marginBottom: '20px'}}>
                      <h4 style={{fontSize: '16px', fontWeight: 600, marginBottom: '12px', color: darkMode ? 'white' : '#1f2937'}}>
                        {language === 'he' ? 'שיוך מוצע:' : 'Proposed Assignments:'}
                      </h4>
                      {dryRunPreview.assignmentDetails.map((detail, index) => {
                        const dutyType = dutyTypePoints[detail.shift.dutyType] || dutyTypePoints['גלגלת'];
                        return (
                          <div key={index} style={{padding: '12px', marginBottom: '8px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb', border: `1px solid ${darkMode ? '#374151' : '#e5e7eb'}`}}>
                            <div style={{display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px'}}>
                              <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                                <div style={{...styles.avatar, width: '36px', height: '36px', fontSize: '14px'}}>{detail.employee.name.charAt(0)}</div>
                                <div>
                                  <div style={{fontWeight: 600, fontSize: '14px', color: darkMode ? 'white' : '#1f2937'}}>{detail.employee.name}</div>
                                  <div style={{fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>
                                    {new Date(detail.shift.startDate).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { year: 'numeric', month: 'short', day: 'numeric' })}
                                    {detail.shift.endDate && detail.shift.endDate !== detail.shift.startDate && (
                                      <> - {new Date(detail.shift.endDate).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { year: 'numeric', month: 'short', day: 'numeric' })}</>
                                    )}
                                  </div>
                                </div>
                              </div>
                              <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                                <span style={{...styles.badge('blue'), fontSize: '12px'}}>{detail.shift.dutyType}</span>
                                {detail.holidayName && (
                                  <span style={{...styles.badge('purple'), fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px'}}>
                                    <Gift size={14} />
                                    {detail.holidayName}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <div style={{display: 'flex', gap: '12px', justifyContent: 'flex-end'}}>
                      <button onClick={() => setDryRunPreview(null)} style={{...styles.btn('gray'), padding: '12px 24px'}}>
                        {t.cancel}
                      </button>
                      <button onClick={applyDryRunPreview} style={{...styles.btn('green'), padding: '12px 24px', display: 'flex', alignItems: 'center', gap: '8px'}}>
                        <CheckCircle size={20} />
                        {t.acceptDistribution}
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {calendarView === 'list' && filteredShifts.length === 0 && !showAddShift && (
                <div style={styles.emptyState}>
                  <Clock size={64} color="#9ca3af" style={{margin: '0 auto 16px'}} />
                  <p style={{color: darkMode ? '#d1d5db' : '#6b7280', fontSize: '18px', margin: '8px 0'}}>{t.noShifts}</p>
                  <p style={{color: darkMode ? '#9ca3af' : '#9ca3af', fontSize: '14px'}}>{t.clickToAdd}</p>
                </div>
              )}

              {calendarView === 'list' && filteredShifts.length > 0 && (
                <div style={{display: 'flex', flexDirection: 'column', gap: '16px'}}>
                  {getAllShifts().map(shift => {
                        const emp = getEmployee(shift.employeeId);
                        const isConflict = conflictingShifts.has(shift.id);
                        const dutyType = dutyTypePoints[shift.dutyType] || dutyTypePoints['גלגלת'];
                        const dateType = dateTypePoints[shift.dateType] || dateTypePoints['חול'];
                        const points = emp ? calculateShiftPoints(shift, emp.status) : 0;

                        // בדיקה אם זה חג בשבת
                        const dayOfWeek = getDayOfWeek(shift.startDate);
                        const isSaturday = dayOfWeek === 6;
                        const isHoliday = shift.dateType === 'חג';
                        const isHolidayOnSaturday = isHoliday && isSaturday;

                        // חישוב הרווח הקרוב ביותר לתורנות אחרת של אותו עובד
                        let closestGapDays = null;
                        if (shift.employeeId && isConflict) {
                          const currentStart = new Date(shift.startDate);
                          const currentEnd = shift.endDate ? new Date(shift.endDate) : currentStart;
                          const empShifts = shifts.filter(s => s.employeeId === shift.employeeId && s.id !== shift.id);
                          const gaps = empShifts.map(s => {
                            const shiftStart = new Date(s.startDate);
                            const shiftEnd = s.endDate ? new Date(s.endDate) : shiftStart;
                            const gapAfter = Math.abs((shiftStart - currentEnd) / (1000 * 60 * 60 * 24));
                            const gapBefore = Math.abs((currentStart - shiftEnd) / (1000 * 60 * 60 * 24));
                            return Math.min(gapAfter, gapBefore);
                          });
                          closestGapDays = gaps.length > 0 ? Math.min(...gaps) : null;
                        }

                        return editingShift?.id === shift.id ? (
                          <div key={shift.id} style={{...styles.shiftRow(false), background: darkMode ? '#374151' : '#eff6ff', flexDirection: 'column'}}>
                            <div style={{display: 'flex', gap: '8px', width: '100%'}}>
                              <div style={{flex: 1, display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '8px'}}>
                                <select value={editingShift.employeeId || ''} onChange={(e) => setEditingShift({...editingShift, employeeId: e.target.value ? parseInt(e.target.value) : null})} style={{...styles.select, padding: '6px'}}>
                                  <option value="">{t.noEmployee}</option>
                                  {employees.map(emp => <option key={emp.id} value={emp.id}>{emp.name}</option>)}
                                </select>
                                <input type="date" value={editingShift.startDate} onChange={(e) => setEditingShift({...editingShift, startDate: e.target.value})} style={{...styles.dateInput, padding: '8px'}} placeholder="התחלה" />
                                <input type="date" value={editingShift.endDate} onChange={(e) => setEditingShift({...editingShift, endDate: e.target.value})} style={{...styles.dateInput, padding: '8px'}} placeholder="סיום" />
                                <select value={editingShift.dutyType} onChange={(e) => setEditingShift({...editingShift, dutyType: e.target.value})} style={{...styles.select, padding: '6px'}}>
                                  {Object.values(dutyTypePoints).map(type => (
                                    <option key={type.name} value={type.name}>{type.name}</option>
                                  ))}
                                </select>
                                <select value={editingShift.dateType} onChange={(e) => setEditingShift({...editingShift, dateType: e.target.value})} style={{...styles.select, padding: '6px'}}>
                                  {Object.values(dateTypePoints).map(type => (
                                    <option key={type.name} value={type.name}>{type.name}</option>
                                  ))}
                                </select>
                              </div>
                              <div style={{display: 'flex', gap: '8px'}}>
                                <button onClick={handleUpdateShift} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#10b981'}}>
                                  <Save size={18} />
                                </button>
                                <button onClick={() => setEditingShift(null)} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#6b7280'}}>
                                  <X size={18} />
                                </button>
                              </div>
                            </div>
                            {editingShift.employeeId && editingShift.dutyType && (() => {
                              const selectedEmp = employees.find(e => e.id === editingShift.employeeId);
                              if (selectedEmp && isEmployeeExemptFromDuty(selectedEmp, editingShift.dutyType)) {
                                return (
                                  <div style={{display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', marginTop: '8px', borderRadius: '6px', background: '#fef3c7', border: '1px solid #fbbf24', color: '#92400e', width: '100%'}}>
                                    <AlertCircle size={18} style={{flexShrink: 0}} />
                                    <span style={{fontSize: '13px', fontWeight: 500}}>
                                      {language === 'he' ? 'אזהרה: לעובד יש פטור מתורנות מסוג זה' : 'Warning: Employee has exemption for this duty type'}
                                    </span>
                                  </div>
                                );
                              }
                              return null;
                            })()}
                          </div>
                        ) : (
                          <div key={shift.id} style={styles.shiftRow(isConflict)}>
                            <div style={{flex: 1}}>
                              <div style={{display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '8px'}}>
                                {shift.employeeId ? (
                                  <>
                                    <div style={styles.avatar}>{emp?.name.charAt(0)}</div>
                                    <div>
                                      <p style={{margin: 0, fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{emp?.name}</p>
                                      <p style={{margin: 0, fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>
                                        {emp?.status} - {emp?.department || t.noDept}
                                      </p>
                                    </div>
                                  </>
                                ) : (
                                  <>
                                    <div style={{...styles.avatar, background: '#9ca3af'}}>?</div>
                                    <div>
                                      <p style={{margin: 0, fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{t.unassigned}</p>
                                      <p style={{margin: 0, fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{t.noEmployee}</p>
                                    </div>
                                  </>
                                )}
                              </div>
                              <div style={{display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center'}}>
                                <span style={{fontSize: '13px', color: darkMode ? '#d1d5db' : '#6b7280', fontWeight: 500}}>
                                  {new Date(shift.startDate).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'short', day: 'numeric' })}
                                  {shift.endDate && shift.endDate !== shift.startDate && (
                                    <> → {new Date(shift.endDate).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'short', day: 'numeric' })}</>
                                  )}
                                </span>
                                <span style={{...styles.badge('blue'), background: dutyType.color + '20', color: dutyType.color, border: `1px solid ${dutyType.color}`, fontWeight: 600}}>
                                  {shift.dutyType}
                                </span>
                                <span style={{...styles.badge('gray'), background: dateType.color + '20', color: dateType.color, border: `1px solid ${dateType.color}`, fontWeight: 600}}>
                                  {shift.dateType}
                                </span>
                                {isHolidayOnSaturday && (
                                  <span style={{...styles.badge('yellow'), fontSize: '11px', fontWeight: 600}}>
                                    חג בשבת!
                                  </span>
                                )}
                                <span style={styles.pointsBreakdown}>
                                  {shift.manualPoints !== null && shift.manualPoints !== undefined ? (
                                    <span style={{color: '#f59e0b', fontWeight: 600}}>✓ {shift.manualPoints} נק׳</span>
                                  ) : (
                                    <>
                                      <Award size={14} color={dutyType.color} />
                                      <span style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{points.toFixed(1)} נק׳</span>
                                    </>
                                  )}
                                </span>
                                {!shift.employeeId && (
                                  <span style={{...styles.badge('gray'), fontSize: '11px'}}>
                                    {t.unassigned}
                                  </span>
                                )}
                                {shift.repeatType && shift.repeatType !== 'none' && (
                                  <span style={{...styles.badge('yellow'), fontSize: '11px'}}>
                                    <Repeat size={12} style={{display: 'inline', marginLeft: '2px'}} />
                                    {t[shift.repeatType]}
                                  </span>
                                )}
                                {isConflict && (
                                  <span style={{...styles.badge('red'), fontSize: '12px'}}>
                                    <AlertCircle size={14} style={{display: 'inline', marginLeft: '4px'}} />
                                    התנגשות {closestGapDays !== null && `(${Math.floor(closestGapDays)} ימים)`}
                                  </span>
                                )}
                              </div>
                            </div>
                            <div style={{display: 'flex', gap: '8px'}}>
                              <button onClick={() => setSwappingShift(shift)} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#f59e0b'}}>
                                <RefreshCw size={18} />
                              </button>
                              <button onClick={() => setEditingShift(shift)} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#3b82f6'}}>
                                <Edit2 size={18} />
                              </button>
                              <button onClick={() => setDeleteConfirm({type: 'shift', id: shift.id})} style={{background: 'none', border: 'none', cursor: 'pointer', padding: '8px', color: '#ef4444'}}>
                                <Trash2 size={18} />
                              </button>
                            </div>
                          </div>
                        );
                  })}
                </div>
              )}

              {calendarView === 'week' && (
                <div>
                  <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px'}}>
                    <button onClick={() => navigateWeek(-1)} style={styles.btn('gray')}>
                      {t.previousWeek}
                    </button>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>
                      {t.weekOf} {new Date(currentWeekStart).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'long', day: 'numeric', year: 'numeric' })}
                    </h3>
                    <button onClick={() => navigateWeek(1)} style={styles.btn('gray')}>
                      {t.nextWeek}
                    </button>
                  </div>
                  <div style={{display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '8px'}}>
                    {getWeekDates(currentWeekStart).map((date, idx) => {
                      const dayShifts = getShiftsByDate(date);
                      const dayNames = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
                      return (
                        <div key={date} style={{...styles.dateCard, minHeight: '200px'}}>
                          <div style={{...styles.dateHeader, textAlign: 'center'}}>
                            <div style={{fontSize: '12px', fontWeight: 600, color: darkMode ? '#9ca3af' : '#6b7280'}}>{t[dayNames[idx]]}</div>
                            <div style={{fontSize: '18px', fontWeight: 'bold', color: darkMode ? 'white' : '#1f2937'}}>{new Date(date).getDate()}</div>
                          </div>
                          <div style={{padding: '8px'}}>
                            {dayShifts.map(shift => {
                              const emp = getEmployee(shift.employeeId);
                              const dutyType = dutyTypePoints[shift.dutyType] || dutyTypePoints['גלגלת'];
                              const points = emp ? calculateShiftPoints(shift, emp.status) : 0;
                              return (
                                <div key={shift.id} style={{marginBottom: '8px', padding: '8px', borderRadius: '6px', background: dutyType.color + '15', border: `1px solid ${dutyType.color}40`}}>
                                  <div style={{fontSize: '12px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '4px'}}>
                                    {shift.employeeId ? emp?.name : t.unassigned}
                                  </div>
                                  <div style={{fontSize: '10px', fontWeight: 600, color: dutyType.color}}>
                                    {shift.dutyType} • {points.toFixed(1)} נק׳
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {calendarView === 'month' && (
                <div style={{display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '8px'}}>
                  {getMonthDates().map(date => {
                    const dayShifts = getShiftsByDate(date);
                    return (
                      <div key={date} style={{...styles.dateCard, minHeight: '120px'}}>
                        <div style={{...styles.dateHeader, textAlign: 'center', padding: '8px'}}>
                          <div style={{fontSize: '16px', fontWeight: 'bold', color: darkMode ? 'white' : '#1f2937'}}>{new Date(date).getDate()}</div>
                        </div>
                        <div style={{padding: '4px', fontSize: '10px'}}>
                          {dayShifts.slice(0, 3).map(shift => {
                            const emp = getEmployee(shift.employeeId);
                            const dutyType = dutyTypePoints[shift.dutyType] || dutyTypePoints['גלגלת'];
                            return (
                              <div key={shift.id} style={{marginBottom: '4px', padding: '4px', borderRadius: '4px', background: dutyType.color + '15', border: `1px solid ${dutyType.color}40`}}>
                                <div style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>
                                  {shift.employeeId ? emp?.name : t.unassigned}
                                </div>
                              </div>
                            );
                          })}
                          {dayShifts.length > 3 && (
                            <div style={{textAlign: 'center', color: darkMode ? '#9ca3af' : '#6b7280', fontWeight: 600}}>
                              +{dayShifts.length - 3}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
        @media print {
          button, .no-print { display: none !important; }
        }
      `}</style>
    </div>
  );
}
