import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Calendar, Clock, Users, Plus, Edit2, Trash2, Save, X, Upload, Download, Search, Printer, AlertCircle, Moon, Sun, Globe, BarChart3, Award, RefreshCw, Repeat } from 'lucide-react';
import * as XLSX from 'xlsx';

// הגדרת סוגי משמרות וניקוד צדק
const SHIFT_TYPES = {
  'חול': { name: 'חול', points: 1, description: 'שני-חמישי', color: '#3b82f6' },
  'שבת': { name: 'שבת', points: 2, description: 'חמישי-שני', color: '#8b5cf6' },
  'שישי': { name: 'שישי', points: 1.5, description: 'משמרת שישי', color: '#f59e0b' },
  'חג': { name: 'חג', points: 3, description: 'משמרת חג', color: '#ef4444' },
};

const translations = {
  en: {
    appTitle: 'Shift Manager', subtitle: 'Company Scheduling System', 
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
    conflictDetected: 'shift conflict(s) detected!', addEmployeesFirst: 'Add employees first', 
    needEmployees: 'You need to add employees before creating shifts', newShift: 'New Shift', 
    employee: 'Employee', date: 'Date', startTime: 'Start Date', endTime: 'End Date', 
    shiftType: 'Shift Type', shiftTypePlaceholder: 'Select shift type', 
    noShifts: 'No shifts scheduled', clickToAdd: 'Click "Add Shift" to create', 
    previousWeek: '← Previous', nextWeek: 'Next →', weekOf: 'Week of',
    sun: 'Sun', mon: 'Mon', tue: 'Tue', wed: 'Wed', thu: 'Thu', fri: 'Fri', sat: 'Sat',
    imported: 'Successfully imported', employeesText: 'employees!',
    errorReading: 'Error reading file', conflictWarning: '⚠️ Conflict detected!',
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
    swapShiftTitle: 'Swap Shifts', selectShiftToSwap: 'Select shift to swap'
  },
  he: {
    appTitle: 'מנהל משמרות', subtitle: 'מערכת ניהול משמרות', 
    employees: 'עובדים', schedule: 'לוח משמרות', reports: 'דוחות',
    employeeManagement: 'ניהול עובדים', reportsTitle: 'דוחות וסטטיסטיקה',
    template: 'תבנית', import: 'ייבוא', add: 'הוסף', 
    searchPlaceholder: 'חיפוש לפי שם...', allDepartments: 'כל המחלקות', 
    allStatuses: 'כל הסטטוסים', newEmployee: 'עובד חדש', name: 'שם', 
    personalNumber: 'מספר אישי', sex: 'מין', status: 'סטטוס', 
    department: 'מחלקה', select: 'בחר', male: 'זכר', female: 'נקבה', 
    other: 'אחר', addEmployee: 'הוסף עובד', noEmployees: 'אין עובדים', 
    importOrAdd: 'ייבא מאקסל', noMatching: 'אין תואמים', 
    adjustFilters: 'נסה מסננים', shifts: 'משמרות', actions: 'פעולות', 
    shiftSchedule: 'לוח משמרות', list: 'רשימה', week: 'שבוע', month: 'חודש', 
    export: 'ייצוא', print: 'הדפסה', addShift: 'הוסף משמרת', 
    conflictDetected: 'התנגשויות זוהו!', addEmployeesFirst: 'הוסף עובדים תחילה', 
    needEmployees: 'הוסף עובדים לפני משמרות', newShift: 'משמרת חדשה', 
    employee: 'עובד', date: 'תאריך', startTime: 'תאריך התחלה', endTime: 'תאריך סיום', 
    shiftType: 'סוג משמרת', shiftTypePlaceholder: 'בחר סוג משמרת', 
    noShifts: 'אין משמרות', clickToAdd: 'לחץ "הוסף משמרת"',
    previousWeek: 'קודם ←', nextWeek: '→ הבא', weekOf: 'שבוע של',
    sun: "א'", mon: "ב'", tue: "ג'", wed: "ד'", thu: "ה'", fri: "ו'", sat: "ש'",
    imported: 'יובאו', employeesText: 'עובדים!',
    errorReading: 'שגיאה', conflictWarning: '⚠️ התנגשות!',
    noDept: 'ללא מחלקה', noDepartment: 'ללא מחלקה', selectEmployee: 'בחר עובד',
    deptPlaceholder: 'מכירות, IT', totalShifts: 'סה"כ משמרות',
    weekdayShifts: 'משמרות חול', shabbatShifts: 'משמרות שבת',
    justicePoints: 'ניקוד צדק', points: 'ניקוד', count: 'כמות',
    noData: 'אין נתונים זמינים', employeeStats: 'סטטיסטיקות עובדים',
    editEmployee: 'ערוך עובד', deleteConfirm: 'האם אתה בטוח?',
    deleteEmployeeMsg: 'למחוק את', deleteShiftMsg: 'למחוק משמרת זו',
    cancel: 'ביטול', delete: 'מחק', repeating: 'חוזרת',
    repeatType: 'סוג חזרה', none: 'ללא', daily: 'יומי', weekly: 'שבועי',
    monthly: 'חודשי', repeatUntil: 'חזור עד', filterByEmployee: 'סינון לפי עובד',
    filterByShiftType: 'סינון לפי סוג', allEmployees: 'כל העובדים',
    allShiftTypes: 'כל הסוגים', swapShift: 'החלף משמרת', swapWith: 'החלף עם',
    swapShiftTitle: 'החלפת משמרות', selectShiftToSwap: 'בחר משמרת להחלפה'
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
  const [customStatuses, setCustomStatuses] = useState([]);
  const [showAddStatus, setShowAddStatus] = useState(false);
  const [newStatusName, setNewStatusName] = useState('');
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [filterEmployee, setFilterEmployee] = useState('');
  const [filterShiftType, setFilterShiftType] = useState('');
  const [swappingShift, setSwappingShift] = useState(null);

  const [newEmployee, setNewEmployee] = useState({
    name: '', personalNumber: '', sex: '', status: '', department: ''
  });

  const [newShift, setNewShift] = useState({
    employeeId: '', date: '', startTime: '', endTime: '', role: '',
    repeatType: 'none', repeatUntil: ''
  });

  const t = translations[language];

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
      const savedStatuses = localStorage.getItem('customStatuses');
      if (savedStatuses) {
        setCustomStatuses(JSON.parse(savedStatuses));
      } else {
        setCustomStatuses(['פעיל', 'לא פעיל', 'בחופשה']);
      }
    } catch (e) {
      setCustomStatuses(['פעיל', 'לא פעיל', 'בחופשה']);
    }
  }, []);

  const isFirstRenderEmployees = useRef(true);
  useEffect(() => {
    if (isFirstRenderEmployees.current) {
      isFirstRenderEmployees.current = false;
      return;
    }
    localStorage.setItem('employees', JSON.stringify(employees));
  }, [employees]);

  const isFirstRenderShifts = useRef(true);
  useEffect(() => {
    if (isFirstRenderShifts.current) {
      isFirstRenderShifts.current = false;
      return;
    }
    localStorage.setItem('shifts', JSON.stringify(shifts));
  }, [shifts]);

  useEffect(() => { 
    localStorage.setItem('language', language); 
  }, [language]);

  useEffect(() => { 
    localStorage.setItem('darkMode', String(darkMode)); 
  }, [darkMode]);

  const isFirstRenderStatuses = useRef(true);
  useEffect(() => {
    if (isFirstRenderStatuses.current) {
      isFirstRenderStatuses.current = false;
      return;
    }
    localStorage.setItem('customStatuses', JSON.stringify(customStatuses));
  }, [customStatuses]);

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
      const matchesShiftType = !filterShiftType || shift.role === filterShiftType;
      return matchesEmployee && matchesShiftType;
    });
  }, [shifts, filterEmployee, filterShiftType]);

  const employeeStats = useMemo(() => {
    return employees.map(emp => {
      const empShifts = shifts.filter(s => s.employeeId === emp.id);
      const shiftCounts = {};
      let totalPoints = 0;
      
      Object.keys(SHIFT_TYPES).forEach(type => {
        shiftCounts[type] = 0;
      });

      empShifts.forEach(shift => {
        const type = shift.role || 'חול';
        if (shiftCounts[type] !== undefined) {
          shiftCounts[type]++;
          totalPoints += SHIFT_TYPES[type]?.points || 0;
        }
      });

      return {
        ...emp,
        totalShifts: empShifts.length,
        shiftCounts,
        justicePoints: totalPoints
      };
    });
  }, [employees, shifts]);

  const hasConflict = (employeeId, date, startTime, endTime, excludeShiftId = null) => {
    const employeeShifts = shifts.filter(s => 
      s.employeeId === employeeId && s.date === date && s.id !== excludeShiftId
    );
    return employeeShifts.some(shift => (startTime < shift.endTime && endTime > shift.startTime));
  };

  const conflictingShifts = useMemo(() => {
    const conflicts = new Set();
    shifts.forEach(shift => {
      const employeeShifts = shifts.filter(s => 
        s.employeeId === shift.employeeId && s.date === shift.date && s.id !== shift.id
      );
      const isConflict = employeeShifts.some(s => 
        shift.startTime < s.endTime && shift.endTime > s.startTime
      );
      if (isConflict) {
        conflicts.add(shift.id);
      }
    });
    return conflicts;
  }, [shifts]);

  const generateRepeatingShifts = (baseShift, repeatType, repeatUntil) => {
    const shifts = [];
    const startDate = new Date(baseShift.date);
    const endDate = new Date(repeatUntil);
    let currentDate = new Date(startDate);
    let id = Date.now();

    while (currentDate <= endDate) {
      const dateStr = currentDate.toISOString().split('T')[0];
      const startTime = baseShift.startTime.replace(baseShift.date, dateStr);
      const endTime = baseShift.endTime.replace(baseShift.date, dateStr);
      
      shifts.push({
        id: id++,
        employeeId: baseShift.employeeId,
        date: dateStr,
        startTime,
        endTime,
        role: baseShift.role,
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
          name: r.Name || r.name || '',
          personalNumber: r['Personal Number'] || r.personalNumber || '',
          sex: r.Sex || r.sex || '',
          status: r.Status || r.status || '',
          department: r.Department || r.department || ''
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

  const handleDownloadTemplate = () => {
    const template = [{ Name: 'John', 'Personal Number': '123', Sex: 'Male', Status: 'Active', Department: 'Sales' }];
    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Employees');
    XLSX.writeFile(wb, 'template.xlsx');
  };

  const handleExportShifts = () => {
    const exportData = filteredShifts.map(shift => {
      const emp = employees.find(e => e.id === shift.employeeId);
      return {
        'Employee Name': emp?.name || 'Unknown',
        'Personal Number': emp?.personalNumber || '',
        'Department': emp?.department || '',
        'Date': shift.date,
        'Start Date': shift.startTime,
        'End Date': shift.endTime,
        'Shift Type': shift.role
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
    if (newShift.employeeId && newShift.date && newShift.startTime && newShift.endTime && newShift.role) {
      if (newShift.repeatType !== 'none' && newShift.repeatUntil) {
        const repeatingShifts = generateRepeatingShifts(
          {
            employeeId: parseInt(newShift.employeeId),
            date: newShift.date,
            startTime: newShift.startTime,
            endTime: newShift.endTime,
            role: newShift.role
          },
          newShift.repeatType,
          newShift.repeatUntil
        );
        setShifts([...shifts, ...repeatingShifts]);
      } else {
        if (hasConflict(parseInt(newShift.employeeId), newShift.date, newShift.startTime, newShift.endTime)) {
          setUploadMessage(t.conflictWarning);
          setTimeout(() => setUploadMessage(''), 3000);
          return;
        }
        setShifts([...shifts, {
          id: Date.now(),
          employeeId: parseInt(newShift.employeeId),
          date: newShift.date,
          startTime: newShift.startTime,
          endTime: newShift.endTime,
          role: newShift.role,
          repeatType: 'none'
        }]);
      }
      setNewShift({ employeeId: '', date: '', startTime: '', endTime: '', role: '', repeatType: 'none', repeatUntil: '' });
      setShowAddShift(false);
    }
  };

  const handleUpdateShift = () => {
    if (editingShift && !hasConflict(editingShift.employeeId, editingShift.date, editingShift.startTime, editingShift.endTime, editingShift.id)) {
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
      setUploadMessage('✅ משמרות הוחלפו בהצלחה!');
      setTimeout(() => setUploadMessage(''), 3000);
    }
  };

  const handleAddCustomStatus = () => {
    if (newStatusName.trim() && !customStatuses.includes(newStatusName.trim())) {
      setCustomStatuses([...customStatuses, newStatusName.trim()]);
      setNewStatusName('');
      setShowAddStatus(false);
    }
  };

  const handleDeleteStatus = (s) => {
    if (customStatuses.length > 1) setCustomStatuses(customStatuses.filter(x => x !== s));
  };

  const getEmployee = (employeeId) => employees.find(e => e.id === employeeId);
  const getShiftsByDate = (date) => filteredShifts.filter(s => s.date === date).sort((a, b) => a.startTime.localeCompare(b.startTime));
  const getDates = () => Array.from(new Set(filteredShifts.map(s => s.date))).sort();

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
    headerRight: { display: 'flex', gap: '12px' },
    iconBtn: { padding: '8px', background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '8px', cursor: 'pointer', color: 'white', transition: 'all 0.2s' },
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
      background: darkMode ? (color === 'green' ? '#065f46' : color === 'indigo' ? '#4338ca' : color === 'blue' ? '#1e40af' : color === 'purple' ? '#6b21a8' : color === 'red' ? '#991b1b' : '#374151') : 
                            (color === 'green' ? '#16a34a' : color === 'indigo' ? '#6366f1' : color === 'blue' ? '#3b82f6' : color === 'purple' ? '#9333ea' : color === 'red' ? '#dc2626' : '#6b7280'),
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
      background: type === 'green' ? '#dcfce7' : type === 'red' ? '#fee2e2' : type === 'yellow' ? '#fef3c7' : type === 'blue' ? '#dbeafe' : '#e5e7eb',
      color: type === 'green' ? '#166534' : type === 'red' ? '#991b1b' : type === 'yellow' ? '#854d0e' : type === 'blue' ? '#1e40af' : '#374151'
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
          <div style={{...styles.message, borderColor: uploadMessage.includes('⚠️') ? '#f59e0b' : '#10b981', background: uploadMessage.includes('⚠️') ? '#fef3c7' : '#d1fae5', color: uploadMessage.includes('⚠️') ? '#92400e' : '#065f46'}}>
            {uploadMessage}
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
                      <div style={{display: 'flex', gap: '8px'}}>
                        <select 
                          value={editingEmployee ? editingEmployee.status : newEmployee.status} 
                          onChange={(e) => editingEmployee 
                            ? setEditingEmployee({...editingEmployee, status: e.target.value})
                            : setNewEmployee({...newEmployee, status: e.target.value})} 
                          style={{...styles.select, flex: 1}}
                        >
                          <option value="">{t.select}</option>
                          {customStatuses.map(s => <option key={s} value={s}>{s}</option>)}
                        </select>
                        <button onClick={() => setShowAddStatus(true)} style={{...styles.btn('green'), padding: '8px'}}>
                          <Plus size={20} />
                        </button>
                      </div>
                      {showAddStatus && (
                        <div style={{marginTop: '8px', display: 'flex', gap: '8px'}}>
                          <input type="text" value={newStatusName} onChange={(e) => setNewStatusName(e.target.value)} placeholder="סטטוס חדש" style={{...styles.input, flex: 1}} onKeyPress={(e) => e.key === 'Enter' && handleAddCustomStatus()} />
                          <button onClick={handleAddCustomStatus} style={{...styles.btn('green'), padding: '8px'}}>
                            <Save size={16} />
                          </button>
                          <button onClick={() => { setShowAddStatus(false); setNewStatusName(''); }} style={{...styles.btn('gray'), padding: '8px'}}>
                            <X size={16} />
                          </button>
                        </div>
                      )}
                      {customStatuses.length > 0 && (
                        <div style={{marginTop: '8px', display: 'flex', flexWrap: 'wrap', gap: '4px'}}>
                          {customStatuses.map(status => (
                            <span key={status} style={{display: 'inline-flex', alignItems: 'center', gap: '4px', padding: '4px 8px', borderRadius: '4px', fontSize: '12px', background: darkMode ? '#4b5563' : '#e5e7eb', color: darkMode ? '#d1d5db' : '#374151'}}>
                              {status}
                              {customStatuses.length > 1 && (
                                <button onClick={() => handleDeleteStatus(status)} style={{background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: '#ef4444'}}>
                                  <X size={12} />
                                </button>
                              )}
                            </span>
                          ))}
                        </div>
                      )}
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
                    {filteredEmployees.map(emp => (
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
                          <span style={styles.badge('green')}>{emp.status || 'N/A'}</span>
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
                    ))}
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
                    {employees.length === 0 ? t.addEmployeesFirst : 'הוסף משמרות כדי לראות דוחות'}
                  </p>
                </div>
              ) : (
                <>
                  <div style={{marginBottom: '24px'}}>
                    <h3 style={{fontSize: '18px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '16px'}}>
                      טבלת סוגי משמרות וניקוד צדק
                    </h3>
                    <div style={styles.statsGrid}>
                      {Object.values(SHIFT_TYPES).map(type => (
                        <div key={type.name} style={{...styles.statBox, borderLeft: `4px solid ${type.color}`}}>
                          <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                            <div>
                              <h4 style={{margin: 0, fontSize: '16px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{type.name}</h4>
                              <p style={{margin: '4px 0 0 0', fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{type.description}</p>
                            </div>
                            <div style={{textAlign: 'center'}}>
                              <Award size={24} color={type.color} />
                              <p style={{margin: '4px 0 0 0', fontSize: '14px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{type.points} נקודות</p>
                            </div>
                          </div>
                        </div>
                      ))}
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
                          <th style={{...styles.th, textAlign: 'center'}}>{t.totalShifts}</th>
                          {Object.keys(SHIFT_TYPES).map(type => (
                            <th key={type} style={{...styles.th, textAlign: 'center'}}>{type}</th>
                          ))}
                          <th style={{...styles.th, textAlign: 'center'}}>{t.justicePoints}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {employeeStats.sort((a, b) => b.justicePoints - a.justicePoints).map(emp => (
                          <tr key={emp.id} style={styles.tr} onMouseEnter={(e) => e.currentTarget.style.background = darkMode ? '#374151' : '#f9fafb'} onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}>
                            <td style={styles.td}>
                              <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                                <div style={styles.avatar}>{emp.name.charAt(0)}</div>
                                <span style={{fontWeight: 500, color: darkMode ? 'white' : '#1f2937'}}>{emp.name}</span>
                              </div>
                            </td>
                            <td style={{...styles.td, textAlign: 'center', fontWeight: 600}}>{emp.totalShifts}</td>
                            {Object.keys(SHIFT_TYPES).map(type => (
                              <td key={type} style={{...styles.td, textAlign: 'center'}}>
                                {emp.shiftCounts[type] || 0}
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
                  <button onClick={handleExportShifts} style={styles.btn('green')} disabled={shifts.length === 0}>
                    <Download size={20} />
                    {t.export}
                  </button>
                  <button onClick={() => window.print()} style={styles.btn('purple')}>
                    <Printer size={20} />
                    {t.print}
                  </button>
                  <button onClick={() => setShowAddShift(true)} style={styles.btn('blue')} disabled={employees.length === 0}>
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
                  {Object.keys(SHIFT_TYPES).map(type => <option key={type} value={type}>{type}</option>)}
                </select>
                <select value={filterDepartment} onChange={(e) => setFilterDepartment(e.target.value)} style={styles.select}>
                  <option value="">{t.allDepartments}</option>
                  {departments.map(d => <option key={d} value={d}>{d}</option>)}
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

              {employees.length === 0 && (
                <div style={styles.emptyState}>
                  <Users size={64} color="#9ca3af" style={{margin: '0 auto 16px'}} />
                  <p style={{color: darkMode ? '#d1d5db' : '#6b7280', fontSize: '18px', margin: '8px 0'}}>{t.addEmployeesFirst}</p>
                  <p style={{color: darkMode ? '#9ca3af' : '#9ca3af', fontSize: '14px'}}>{t.needEmployees}</p>
                </div>
              )}

              {showAddShift && employees.length > 0 && (
                <div style={styles.modal}>
                  <div style={styles.modalHeader}>
                    <h3 style={styles.modalTitle}>{t.newShift}</h3>
                    <button onClick={() => setShowAddShift(false)} style={{background: 'none', border: 'none', cursor: 'pointer'}}>
                      <X size={20} color={darkMode ? 'white' : 'black'} />
                    </button>
                  </div>
                  <div style={styles.grid2}>
                    <div style={{gridColumn: '1 / -1'}}>
                      <label style={styles.label}>{t.employee} *</label>
                      <select value={newShift.employeeId} onChange={(e) => setNewShift({...newShift, employeeId: e.target.value})} style={styles.select}>
                        <option value="">{t.selectEmployee}</option>
                        {employees.map(emp => (
                          <option key={emp.id} value={emp.id}>{emp.name}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label style={styles.label}>{t.date} *</label>
                      <input type="date" value={newShift.date} onChange={(e) => setNewShift({...newShift, date: e.target.value})} style={styles.input} />
                    </div>
                    <div>
                      <label style={styles.label}>{t.shiftType} *</label>
                      <select value={newShift.role} onChange={(e) => setNewShift({...newShift, role: e.target.value})} style={styles.select}>
                        <option value="">{t.shiftTypePlaceholder}</option>
                        {Object.values(SHIFT_TYPES).map(type => (
                          <option key={type.name} value={type.name}>
                            {type.name} - {type.description} ({type.points} נקודות)
                          </option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label style={styles.label}>{t.startTime} *</label>
                      <input type="datetime-local" value={newShift.startTime} onChange={(e) => setNewShift({...newShift, startTime: e.target.value})} style={styles.input} />
                    </div>
                    <div>
                      <label style={styles.label}>{t.endTime} *</label>
                      <input type="datetime-local" value={newShift.endTime} onChange={(e) => setNewShift({...newShift, endTime: e.target.value})} style={styles.input} />
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
                      const shiftType = SHIFT_TYPES[shift.role] || SHIFT_TYPES['חול'];
                      return (
                        <div 
                          key={shift.id} 
                          onClick={() => handleSwapShifts(swappingShift.id, shift.id)}
                          style={{
                            padding: '12px',
                            marginBottom: '8px',
                            borderRadius: '8px',
                            border: `1px solid ${shiftType.color}`,
                            background: shiftType.color + '10',
                            cursor: 'pointer',
                            transition: 'all 0.2s'
                          }}
                          onMouseEnter={(e) => e.currentTarget.style.background = shiftType.color + '30'}
                          onMouseLeave={(e) => e.currentTarget.style.background = shiftType.color + '10'}
                        >
                          <div style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{emp?.name}</div>
                          <div style={{fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280', marginTop: '4px'}}>
                            {new Date(shift.startTime).toLocaleTimeString(language === 'he' ? 'he-IL' : 'en-US', { hour: '2-digit', minute: '2-digit' })} - 
                            {new Date(shift.endTime).toLocaleTimeString(language === 'he' ? 'he-IL' : 'en-US', { hour: '2-digit', minute: '2-digit' })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {calendarView === 'list' && filteredShifts.length === 0 && !showAddShift && employees.length > 0 && (
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
                        const shiftType = SHIFT_TYPES[shift.role] || SHIFT_TYPES['חול'];
                        return editingShift?.id === shift.id ? (
                          <div key={shift.id} style={{...styles.shiftRow(false), background: darkMode ? '#374151' : '#eff6ff'}}>
                            <div style={{flex: 1, display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: '8px'}}>
                              <select value={editingShift.employeeId} onChange={(e) => setEditingShift({...editingShift, employeeId: parseInt(e.target.value)})} style={{...styles.select, padding: '6px'}}>
                                {employees.map(emp => <option key={emp.id} value={emp.id}>{emp.name}</option>)}
                              </select>
                              <input type="date" value={editingShift.date} onChange={(e) => setEditingShift({...editingShift, date: e.target.value})} style={{...styles.input, padding: '6px'}} />
                              <input type="datetime-local" value={editingShift.startTime} onChange={(e) => setEditingShift({...editingShift, startTime: e.target.value})} style={{...styles.input, padding: '6px'}} />
                              <input type="datetime-local" value={editingShift.endTime} onChange={(e) => setEditingShift({...editingShift, endTime: e.target.value})} style={{...styles.input, padding: '6px'}} />
                              <select value={editingShift.role} onChange={(e) => setEditingShift({...editingShift, role: e.target.value})} style={{...styles.select, padding: '6px'}}>
                                {Object.values(SHIFT_TYPES).map(type => (
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
                                <div style={styles.avatar}>{emp?.name.charAt(0)}</div>
                                <div>
                                  <p style={{margin: 0, fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{emp?.name}</p>
                                  <p style={{margin: 0, fontSize: '12px', color: darkMode ? '#9ca3af' : '#6b7280'}}>{emp?.department || t.noDept}</p>
                                </div>
                              </div>
                              <div style={{display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center'}}>
                                <span style={{...styles.badge('blue'), background: shiftType.color + '20', color: shiftType.color, border: `1px solid ${shiftType.color}`, fontWeight: 600}}>
                                  {shift.role}
                                </span>
                                <span style={{fontSize: '14px', color: darkMode ? '#d1d5db' : '#6b7280'}}>
                                  {new Date(shift.startTime).toLocaleTimeString(language === 'he' ? 'he-IL' : 'en-US', { hour: '2-digit', minute: '2-digit' })} - 
                                  {new Date(shift.endTime).toLocaleTimeString(language === 'he' ? 'he-IL' : 'en-US', { hour: '2-digit', minute: '2-digit' })}
                                </span>
                                <span style={{fontSize: '12px', color: darkMode ? '#9ca3af' : '#9ca3af'}}>
                                  {shiftType.points} נקודות
                                </span>
                                {shift.repeatType && shift.repeatType !== 'none' && (
                                  <span style={{...styles.badge('yellow'), fontSize: '11px'}}>
                                    <Repeat size={12} style={{display: 'inline', marginLeft: '2px'}} />
                                    {t[shift.repeatType]}
                                  </span>
                                )}
                                {isConflict && (
                                  <span style={{...styles.badge('red'), fontSize: '12px'}}>
                                    <AlertCircle size={14} style={{display: 'inline', marginLeft: '4px'}} />
                                    התנגשות
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
                              const shiftType = SHIFT_TYPES[shift.role] || SHIFT_TYPES['חול'];
                              return (
                                <div key={shift.id} style={{marginBottom: '8px', padding: '8px', borderRadius: '6px', background: shiftType.color + '15', border: `1px solid ${shiftType.color}40`}}>
                                  <div style={{fontSize: '12px', fontWeight: 600, color: darkMode ? 'white' : '#1f2937', marginBottom: '4px'}}>{emp?.name}</div>
                                  <div style={{fontSize: '10px', color: darkMode ? '#9ca3af' : '#6b7280'}}>
                                    {new Date(shift.startTime).toLocaleTimeString(language === 'he' ? 'he-IL' : 'en-US', { hour: '2-digit', minute: '2-digit' })}
                                  </div>
                                  <div style={{fontSize: '10px', fontWeight: 600, color: shiftType.color, marginTop: '4px'}}>{shift.role}</div>
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
                            const shiftType = SHIFT_TYPES[shift.role] || SHIFT_TYPES['חול'];
                            return (
                              <div key={shift.id} style={{marginBottom: '4px', padding: '4px', borderRadius: '4px', background: shiftType.color + '15', border: `1px solid ${shiftType.color}40`}}>
                                <div style={{fontWeight: 600, color: darkMode ? 'white' : '#1f2937'}}>{emp?.name}</div>
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
    </div>
  );
}