import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Calendar, Clock, Users, Plus, Edit2, Trash2, Save, X, Upload, Download, Search, Printer, AlertCircle, Moon, Sun, Globe, BarChart3, Award, RefreshCw, Repeat, Zap, CheckCircle, Settings, TrendingUp } from 'lucide-react';
import * as XLSX from 'xlsx';

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
    statusName: 'Status Name', manualOverride: 'Manual Override'
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
    statusName: 'שם תפקיד', manualOverride: 'עקיפה ידנית'
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
  const [filterEmployee, setFilterEmployee] = useState('');
  const [filterShiftType, setFilterShiftType] = useState('');
  const [swappingShift, setSwappingShift] = useState(null);
  const [saveStatus, setSaveStatus] = useState('saved');

  const [newEmployee, setNewEmployee] = useState({
    name: '', personalNumber: '', sex: '', status: '', department: ''
  });

  const [newShift, setNewShift] = useState({
    employeeId: '', startDate: '', endDate: '',
    dutyType: '', dateType: '', manualPoints: null,
    repeatType: 'none', repeatUntil: ''
  });

  const t = translations[language];

  // פונקציה לזיהוי אוטומטי של יום בשבוע
  const getDayOfWeek = (dateStr) => {
    const date = new Date(dateStr);
    return date.getDay(); // 0 = Sunday, 6 = Saturday
  };

  // חישוב אוטומטי של תאריך סיום לפי סוג התורנות ותאריך ההתחלה
  const calculateEndDate = (startDate, dateType) => {
    if (!startDate || !dateType) return '';

    const start = new Date(startDate);
    const dayOfWeek = start.getDay(); // 0=ראשון, 1=שני, 2=שלישי... 6=שבת

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

  // פונקציה לחישוב נקודות תורנות
  const calculateShiftPoints = (shift, employeeStatus) => {
    // אם יש עקיפה ידנית, השתמש בה
    if (shift.manualPoints !== null && shift.manualPoints !== undefined) {
      return parseFloat(shift.manualPoints);
    }

    const dutyType = DUTY_TYPES[shift.dutyType] || DUTY_TYPES['גלגלת'];
    const dateType = DATE_TYPES[shift.dateType] || DATE_TYPES['חול'];
    const statusMultiplier = statusMultipliers[employeeStatus] || 1.0;

    let basePoints = dutyType.basePoints;
    let bonusPoints = dateType.bonus;

    // זיהוי אוטומטי: אם זה חג שחל בשבת, הוסף בונוס נוסף
    const dayOfWeek = getDayOfWeek(shift.date);
    const isSaturday = dayOfWeek === 6;
    const isHoliday = shift.dateType === 'חג';

    if (isHoliday && isSaturday) {
      bonusPoints += DATE_TYPES['שבת'].bonus; // הוסף גם את בונוס השבת
    }

    // נוסחה: (נקודות בסיס + בונוס) × מקדם סטטוס
    const totalPoints = (basePoints + bonusPoints) * statusMultiplier;

    return totalPoints;
  };

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
      const savedMultipliers = localStorage.getItem('statusMultipliers');
      if (savedMultipliers) {
        setStatusMultipliers(JSON.parse(savedMultipliers));
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

  const isFirstRenderMultipliers = useRef(true);
  useEffect(() => {
    if (isFirstRenderMultipliers.current) {
      isFirstRenderMultipliers.current = false;
      return;
    }
    localStorage.setItem('statusMultipliers', JSON.stringify(statusMultipliers));
  }, [statusMultipliers]);

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
      return matchesEmployee && matchesShiftType;
    });
  }, [shifts, filterEmployee, filterShiftType]);

  const employeeStats = useMemo(() => {
    return employees.map(emp => {
      const empShifts = shifts.filter(s => s.employeeId === emp.id);
      const dutyTypeCounts = {};
      const dateTypeCounts = {};
      let totalPoints = 0;

      Object.keys(DUTY_TYPES).forEach(type => {
        dutyTypeCounts[type] = 0;
      });

      Object.keys(DATE_TYPES).forEach(type => {
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
  }, [employees, shifts, statusMultipliers]);

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
        const imported = json.map((r, i) => ({
          id: Date.now() + i,
          name: r.Name || r.name || r['שם'] || '',
          personalNumber: r['Personal Number'] || r.personalNumber || r['מספר אישי'] || '',
          sex: r.Sex || r.sex || r['מין'] || '',
          status: r.Status || r.status || r['סטטוס'] || r['תפקיד'] || '',
          department: r.Department || r.department || r['מחלקה'] || ''
        }));
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
    const template = [{ Name: 'John', 'Personal Number': '123', Sex: 'Male', Status: 'חייל רגיל', Department: 'Sales' }];
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
      setNewEmployee({ name: '', personalNumber: '', sex: '', status: '', department: '' });
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

  const handleAddShift = () => {
    if (newShift.startDate && newShift.dutyType && newShift.dateType) {
      // חישוב אוטומטי של תאריך סיום אם לא הוזן ידנית
      const finalEndDate = newShift.endDate || calculateEndDate(newShift.startDate, newShift.dateType);

      if (newShift.repeatType !== 'none' && newShift.repeatUntil) {
        const repeatingShifts = generateRepeatingShifts(
          {
            employeeId: newShift.employeeId ? parseInt(newShift.employeeId) : null,
            startDate: newShift.startDate,
            endDate: finalEndDate,
            dutyType: newShift.dutyType,
            dateType: newShift.dateType,
            manualPoints: newShift.manualPoints
          },
          newShift.repeatType,
          newShift.repeatUntil
        );
        setShifts([...shifts, ...repeatingShifts]);
      } else {
        if (newShift.employeeId && hasConflict(parseInt(newShift.employeeId), newShift.startDate, finalEndDate)) {
          setUploadMessage(t.conflictWarning);
          setTimeout(() => setUploadMessage(''), 3000);
          return;
        }
        setShifts([...shifts, {
          id: Date.now(),
          employeeId: newShift.employeeId ? parseInt(newShift.employeeId) : null,
          startDate: newShift.startDate,
          endDate: finalEndDate,
          dutyType: newShift.dutyType,
          dateType: newShift.dateType,
          manualPoints: newShift.manualPoints,
          repeatType: 'none'
        }]);
      }
      setNewShift({ employeeId: '', startDate: '', endDate: '', dutyType: '', dateType: '', manualPoints: null, repeatType: 'none', repeatUntil: '' });
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

  const handleAutoDistribute = () => {
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

    unassignedShifts.forEach(shift => {
      const sortedEmployees = [...employees].sort((a, b) =>
        employeePointsMap[a.id] - employeePointsMap[b.id]
      );

      for (let emp of sortedEmployees) {
        if (!hasConflict(emp.id, shift.startDate, shift.endDate)) {
          const shiftIndex = updatedShifts.findIndex(s => s.id === shift.id);
          if (shiftIndex !== -1) {
            updatedShifts[shiftIndex] = { ...shift, employeeId: emp.id };
            const shiftPoints = calculateShiftPoints(shift, emp.status);
            employeePointsMap[emp.id] += shiftPoints;
            distributedCount++;
            break;
          }
        }
      }
    });

    setShifts(updatedShifts);
    setUploadMessage(`✅ שובצו ${distributedCount} תורנויות אוטומטית!`);
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

  const getEmployee = (employeeId) => employees.find(e => e.id === employeeId);
  const getShiftsByDate = (date) => filteredShifts.filter(s => {
    // תורנות מוצגת ביום אם היא מתחילה או נמצאת בטווח התאריכים שלה
    const shiftStart = new Date(s.startDate);
    const shiftEnd = s.endDate ? new Date(s.endDate) : shiftStart;
    const currentDate = new Date(date);
    return currentDate >= shiftStart && currentDate <= shiftEnd;
  }).sort((a, b) => a.startDate.localeCompare(b.startDate));
  const getDates = () => {
    // אסוף את כל התאריכים בטווח של כל תורנות
    const allDates = new Set();
    filteredShifts.forEach(s => {
      const start = new Date(s.startDate);
      const end = s.endDate ? new Date(s.endDate) : start;
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        allDates.add(new Date(d).toISOString().split('T')[0]);
      }
    });
    return Array.from(allDates).sort();
  };

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
      border: '1px solid ' + (darkMode ? '#374151' : '#d1d5db'), borderRadius: '8px',
      background: darkMode ? '#374151' : 'white', color: darkMode ? 'white' : 'black', fontSize: '14px'
    },
    select: {
      width: '100%', padding: '10px', border: '1px solid ' + (darkMode ? '#374151' : '#d1d5db'), borderRadius: '8px',
      background: darkMode ? '#374151' : 'white', color: darkMode ? 'white' : 'black', fontSize: '14px'
    },
    modal: {
      marginBottom: '24px', padding: '20px', borderRadius: '8px', border: '2px solid ' + (darkMode ? '#3b82f6' : '#3b82f6'),
      background: darkMode ? '#374151' : '#eff6ff'
    },
    modalHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' },
    modalTitle: { fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', margin: 0 },
    grid2: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' },
    label: { display: 'block', fontSize: '14px', marginBottom: '6px', color: darkMode ? '#d1d5db' : '#374151', fontWeight: 500 },
    emptyState: {
      textAlign: 'center', padding: '48px', borderRadius: '8px', border: '2px dashed ' + (darkMode ? '#374151' : '#d1d5db'),
      background: darkMode ? '#374151' : '#f9fafb'
    },
    table: { width: '100%', borderCollapse: 'collapse' },
    thead: { background: darkMode ? '#374151' : '#f3f4f6', borderBottom: '2px solid ' + (darkMode ? '#4b5563' : '#e5e7eb') },
    th: { padding: '12px 16px', textAlign: language === 'he' ? 'right' : 'left', fontSize: '14px', fontWeight: 600, color: darkMode ? '#d1d5db' : '#374151' },
    tr: { borderBottom: '1px solid ' + (darkMode ? '#374151' : '#e5e7eb'), transition: 'background 0.2s' },
    td: { padding: '12px 16px', color: darkMode ? '#d1d5db' : '#6b7280' },
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
    calViewBtns: { display: 'flex', background: darkMode ? '#374151' : '#f3f4f6', borderRadius: '8px', padding: '4px' },
    calViewBtn: (active) => ({
      padding: '8px 16px', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '14px',
      background: active ? (darkMode ? '#4b5563' : 'white') : 'transparent',
      color: darkMode ? 'white' : '#1f2937',
      boxShadow: active ? '0 1px 3px rgba(0,0,0,0.1)' : 'none'
    }),
    dateCard: { border: '1px solid ' + (darkMode ? '#374151' : '#e5e7eb'), borderRadius: '8px', overflow: 'hidden' },
    dateHeader: { background: darkMode ? '#374151' : '#f3f4f6', padding: '12px 16px', borderBottom: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb') },
    shiftRow: (hasConflict) => ({
      padding: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center',
      background: hasConflict ? (darkMode ? '#7f1d1d' : '#fee2e2') : 'transparent',
      borderBottom: '1px solid ' + (darkMode ? '#374151' : '#e5e7eb'),
      transition: 'background 0.2s'
    }),
    statsCard: {
      padding: '20px', borderRadius: '8px', background: darkMode ? '#374151' : 'white',
      border: '1px solid ' + (darkMode ? '#4b5563' : '#e5e7eb'), marginBottom: '16px'
    },
    statsGrid: {
      display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px'
    },
    statBox: {
      padding: '16px', borderRadius: '8px', background: darkMode ? '#1f2937' : '#f9fafb',
      border: '1px solid ' + (darkMode ? '#374151' : '#e5e7eb')
    },
    confirmDialog: {
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
    },
    confirmBox: {
      background: darkMode ? '#1f2937' : 'white', padding: '24px', borderRadius: '12px',
      maxWidth: '400px', width: '90%', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.3)'
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
                <h1 style={styles.title}>{t.appTitle}</h1>
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
              <button onClick={() => setShowStatusSettings(!showStatusSettings)} style={styles.iconBtn} title={t.manageStatuses}>
                <Settings size={20} />
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
            <div>
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
            <div>
              <div style={styles.flexBetween}>
                <h2 style={styles.h2}>{t.reportsTitle}</h2>
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
                      {Object.values(DUTY_TYPES).map(type => (
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
                      {Object.values(DATE_TYPES).map(type => (
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
                          {Object.keys(DUTY_TYPES).map(type => (
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
                            {Object.keys(DUTY_TYPES).map(type => (
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
            <div>
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
                  <button onClick={handleAutoDistribute} style={styles.btn('orange')} disabled={employees.length === 0}>
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
                  {Object.keys(DUTY_TYPES).map(type => <option key={type} value={type}>{type}</option>)}
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
                    </div>
                    <div>
                      <label style={styles.label}>תאריך התחלה *</label>
                      <input
                        type="date"
                        value={newShift.startDate}
                        onChange={(e) => {
                          const newStartDate = e.target.value;
                          const autoEndDate = calculateEndDate(newStartDate, newShift.dateType);
                          setNewShift({...newShift, startDate: newStartDate, endDate: autoEndDate});
                        }}
                        style={styles.input}
                      />
                    </div>
                    <div>
                      <label style={styles.label}>תאריך סיום ({t.optional})</label>
                      <input
                        type="date"
                        value={newShift.endDate}
                        onChange={(e) => setNewShift({...newShift, endDate: e.target.value})}
                        placeholder="יחושב אוטומטית"
                        style={styles.input}
                      />
                    </div>
                    <div>
                      <label style={styles.label}>{t.dutyType} *</label>
                      <select value={newShift.dutyType} onChange={(e) => setNewShift({...newShift, dutyType: e.target.value})} style={styles.select}>
                        <option value="">{t.shiftTypePlaceholder}</option>
                        {Object.values(DUTY_TYPES).map(type => (
                          <option key={type.name} value={type.name}>
                            {type.name} ({type.basePoints} נק׳ בסיס)
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label style={styles.label}>{t.dateType} *</label>
                      <select
                        value={newShift.dateType}
                        onChange={(e) => {
                          const newDateType = e.target.value;
                          const autoEndDate = newShift.startDate && !newShift.endDate
                            ? calculateEndDate(newShift.startDate, newDateType)
                            : newShift.endDate;
                          setNewShift({...newShift, dateType: newDateType, endDate: autoEndDate});
                        }}
                        style={styles.select}
                      >
                        <option value="">{t.select}</option>
                        {Object.values(DATE_TYPES).map(type => (
                          <option key={type.name} value={type.name}>
                            {type.name} (+{type.bonus} נק׳)
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label style={styles.label}>
                        <Repeat size={16} style={{display: 'inline', marginLeft: '4px'}} />
                        {t.repeatType}
                      </label>
                      <select value={newShift.repeatType} onChange={(e) => setNewShift({...newShift, repeatType: e.target.value})} style={styles.select}>
                        <option value="none">{t.none}</option>
                        <option value="daily">{t.daily}</option>
                        <option value="weekly">{t.weekly}</option>
                        <option value="monthly">{t.monthly}</option>
                      </select>
                    </div>
                    {newShift.repeatType !== 'none' && (
                      <div>
                        <label style={styles.label}>{t.repeatUntil} *</label>
                        <input type="date" value={newShift.repeatUntil} onChange={(e) => setNewShift({...newShift, repeatUntil: e.target.value})} style={styles.input} />
                      </div>
                    )}
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
                  <button onClick={handleAddShift} style={{...styles.btn('blue'), width: '100%', marginTop: '16px', justifyContent: 'center'}}>
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
                      const dutyType = DUTY_TYPES[shift.dutyType] || DUTY_TYPES['גלגלת'];
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

              {calendarView === 'list' && filteredShifts.length === 0 && !showAddShift && (
                <div style={styles.emptyState}>
                  <Clock size={64} color="#9ca3af" style={{margin: '0 auto 16px'}} />
                  <p style={{color: darkMode ? '#d1d5db' : '#6b7280', fontSize: '18px', margin: '8px 0'}}>{t.noShifts}</p>
                  <p style={{color: darkMode ? '#9ca3af' : '#9ca3af', fontSize: '14px'}}>{t.clickToAdd}</p>
                </div>
              )}

              {calendarView === 'list' && filteredShifts.length > 0 && (
                <div style={{display: 'flex', flexDirection: 'column', gap: '16px'}}>
                  {getDates().map(date => (
                    <div key={date} style={styles.dateCard}>
                      <div style={styles.dateHeader}>
                        <h3 style={{margin: 0, fontSize: '16px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>
                          {new Date(date).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}
                        </h3>
                      </div>
                      {getShiftsByDate(date).map(shift => {
                        const emp = getEmployee(shift.employeeId);
                        const isConflict = conflictingShifts.has(shift.id);
                        const dutyType = DUTY_TYPES[shift.dutyType] || DUTY_TYPES['גלגלת'];
                        const dateType = DATE_TYPES[shift.dateType] || DATE_TYPES['חול'];
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
                          <div key={shift.id} style={{...styles.shiftRow(false), background: darkMode ? '#374151' : '#eff6ff'}}>
                            <div style={{flex: 1, display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '8px'}}>
                              <select value={editingShift.employeeId || ''} onChange={(e) => setEditingShift({...editingShift, employeeId: e.target.value ? parseInt(e.target.value) : null})} style={{...styles.select, padding: '6px'}}>
                                <option value="">{t.noEmployee}</option>
                                {employees.map(emp => <option key={emp.id} value={emp.id}>{emp.name}</option>)}
                              </select>
                              <input type="date" value={editingShift.startDate} onChange={(e) => setEditingShift({...editingShift, startDate: e.target.value})} style={{...styles.input, padding: '6px'}} placeholder="התחלה" />
                              <input type="date" value={editingShift.endDate} onChange={(e) => setEditingShift({...editingShift, endDate: e.target.value})} style={{...styles.input, padding: '6px'}} placeholder="סיום" />
                              <select value={editingShift.dutyType} onChange={(e) => setEditingShift({...editingShift, dutyType: e.target.value})} style={{...styles.select, padding: '6px'}}>
                                {Object.values(DUTY_TYPES).map(type => (
                                  <option key={type.name} value={type.name}>{type.name}</option>
                                ))}
                              </select>
                              <select value={editingShift.dateType} onChange={(e) => setEditingShift({...editingShift, dateType: e.target.value})} style={{...styles.select, padding: '6px'}}>
                                {Object.values(DATE_TYPES).map(type => (
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
                  ))}
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
                              const dutyType = DUTY_TYPES[shift.dutyType] || DUTY_TYPES['גלגלת'];
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
                            const dutyType = DUTY_TYPES[shift.dutyType] || DUTY_TYPES['גלגלת'];
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
