import React, { useState, useMemo } from 'react';
import { Calendar, Clock, Users, Plus, Edit2, Trash2, Save, X, Upload, Download, Search, Filter, Printer, AlertCircle, Moon, Sun, Globe } from 'lucide-react';
import * as XLSX from 'xlsx';

const translations = {
  en: {
    appTitle: 'Shift Manager',
    subtitle: 'Company Scheduling System',
    employees: 'Employees',
    schedule: 'Schedule',
    employeeManagement: 'Employee Management',
    template: 'Template',
    import: 'Import',
    add: 'Add',
    searchPlaceholder: 'Search by name or ID...',
    allDepartments: 'All Departments',
    allStatuses: 'All Statuses',
    newEmployee: 'New Employee',
    name: 'Name',
    personalNumber: 'Personal Number',
    sex: 'Sex',
    status: 'Status',
    department: 'Department',
    select: 'Select',
    male: 'Male',
    female: 'Female',
    other: 'Other',
    active: 'Active',
    inactive: 'Inactive',
    onLeave: 'On Leave',
    addEmployee: 'Add Employee',
    noEmployees: 'No employees yet',
    importOrAdd: 'Import from Excel or add manually to get started',
    noMatching: 'No matching employees',
    adjustFilters: 'Try adjusting your filters',
    shifts: 'Shifts',
    actions: 'Actions',
    shiftSchedule: 'Shift Schedule',
    list: 'List',
    week: 'Week',
    month: 'Month',
    export: 'Export',
    print: 'Print',
    addShift: 'Add Shift',
    conflictDetected: 'shift conflict(s) detected! Some employees are scheduled for overlapping shifts.',
    addEmployeesFirst: 'Add employees first',
    needEmployees: 'You need to add employees before creating shifts',
    newShift: 'New Shift',
    employee: 'Employee',
    date: 'Date',
    startTime: 'Start Time',
    endTime: 'End Time',
    shiftType: 'Shift Type',
    shiftTypePlaceholder: 'e.g., Morning, Evening, Night',
    noShifts: 'No shifts scheduled',
    clickToAdd: 'Click "Add Shift" to create your first shift',
    previousWeek: '← Previous Week',
    nextWeek: 'Next Week →',
    weekOf: 'Week of',
    sun: 'Sun',
    mon: 'Mon',
    tue: 'Tue',
    wed: 'Wed',
    thu: 'Thu',
    fri: 'Fri',
    sat: 'Sat',
    imported: 'Successfully imported',
    employeesText: 'employees!',
    errorReading: 'Error reading file. Please check the format.',
    conflictWarning: '⚠️ Conflict detected! Employee already has a shift at this time.',
    noDept: 'No Dept',
    noDepartment: 'No Department',
    selectEmployee: 'Select Employee',
    deptPlaceholder: 'e.g., Sales, Marketing, IT'
  },
  he: {
    appTitle: 'מנהל משמרות',
    subtitle: 'מערכת ניהול משמרות',
    employees: 'עובדים',
    schedule: 'לוח משמרות',
    employeeManagement: 'ניהול עובדים',
    template: 'תבנית',
    import: 'ייבוא',
    add: 'הוסף',
    searchPlaceholder: 'חיפוש לפי שם או מספר אישי...',
    allDepartments: 'כל המחלקות',
    allStatuses: 'כל הסטטוסים',
    newEmployee: 'עובד חדש',
    name: 'שם',
    personalNumber: 'מספר אישי',
    sex: 'מין',
    status: 'סטטוס',
    department: 'מחלקה',
    select: 'בחר',
    male: 'זכר',
    female: 'נקבה',
    other: 'אחר',
    active: 'פעיל',
    inactive: 'לא פעיל',
    onLeave: 'בחופשה',
    addEmployee: 'הוסף עובד',
    noEmployees: 'אין עובדים עדיין',
    importOrAdd: 'ייבא מאקסל או הוסף ידנית כדי להתחיל',
    noMatching: 'אין עובדים תואמים',
    adjustFilters: 'נסה להתאים את המסננים',
    shifts: 'משמרות',
    actions: 'פעולות',
    shiftSchedule: 'לוח משמרות',
    list: 'רשימה',
    week: 'שבוע',
    month: 'חודש',
    export: 'ייצוא',
    print: 'הדפסה',
    addShift: 'הוסף משמרת',
    conflictDetected: 'התנגשויות משמרת זוהו! חלק מהעובדים מתוזמנים למשמרות חופפות.',
    addEmployeesFirst: 'הוסף עובדים תחילה',
    needEmployees: 'עליך להוסיף עובדים לפני יצירת משמרות',
    newShift: 'משמרת חדשה',
    employee: 'עובד',
    date: 'תאריך',
    startTime: 'שעת התחלה',
    endTime: 'שעת סיום',
    shiftType: 'סוג משמרת',
    shiftTypePlaceholder: 'לדוגמה: בוקר, ערב, לילה',
    noShifts: 'אין משמרות מתוזמנות',
    clickToAdd: 'לחץ על "הוסף משמרת" כדי ליצור את המשמרת הראשונה שלך',
    previousWeek: '→ שבוע קודם',
    nextWeek: 'שבוע הבא ←',
    weekOf: 'שבוע של',
    sun: 'א\'',
    mon: 'ב\'',
    tue: 'ג\'',
    wed: 'ד\'',
    thu: 'ה\'',
    fri: 'ו\'',
    sat: 'ש\'',
    imported: 'יובאו בהצלחה',
    employeesText: 'עובדים!',
    errorReading: 'שגיאה בקריאת הקובץ. אנא בדוק את הפורמט.',
    conflictWarning: '⚠️ התנגשות זוהתה! לעובד כבר יש משמרת בזמן זה.',
    noDept: 'ללא מחלקה',
    noDepartment: 'ללא מחלקה',
    selectEmployee: 'בחר עובד',
    deptPlaceholder: 'לדוגמה: מכירות, שיווק, IT'
  }
};

export default function ShiftManager() {
  const [employees, setEmployees] = useState([]);
  const [shifts, setShifts] = useState([]);
  const [activeTab, setActiveTab] = useState('employees');
  const [showAddShift, setShowAddShift] = useState(false);
  const [showAddEmployee, setShowAddEmployee] = useState(false);
  const [editingShift, setEditingShift] = useState(null);
  const [uploadMessage, setUploadMessage] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [filterDepartment, setFilterDepartment] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [calendarView, setCalendarView] = useState('list');
  const [currentWeekStart, setCurrentWeekStart] = useState(new Date());
  const [language, setLanguage] = useState('en');
  const [darkMode, setDarkMode] = useState(false);
  const [customStatuses, setCustomStatuses] = useState(['Active', 'Inactive', 'On Leave', 'Vacation', 'Sick Leave']);
  const [showAddStatus, setShowAddStatus] = useState(false);
  const [newStatusName, setNewStatusName] = useState('');

  const t = translations[language];

  // Load data from persistent storage on mount
  React.useEffect(() => {
    const loadData = async () => {
      try {
        const employeesData = await window.storage.get('employees');
        const shiftsData = await window.storage.get('shifts');
        const langData = await window.storage.get('language');
        const darkModeData = await window.storage.get('darkMode');
        const statusesData = await window.storage.get('customStatuses');
        
        if (employeesData?.value) {
          setEmployees(JSON.parse(employeesData.value));
        }
        if (shiftsData?.value) {
          setShifts(JSON.parse(shiftsData.value));
        }
        if (langData?.value) {
          setLanguage(langData.value);
        }
        if (darkModeData?.value) {
          setDarkMode(darkModeData.value === 'true');
        }
        if (statusesData?.value) {
          setCustomStatuses(JSON.parse(statusesData.value));
        }
      } catch (error) {
        console.log('No saved data found, starting fresh');
      }
    };
    loadData();
  }, []);

  // Save employees to storage whenever they change
  React.useEffect(() => {
    const saveEmployees = async () => {
      try {
        await window.storage.set('employees', JSON.stringify(employees));
      } catch (error) {
        console.error('Error saving employees:', error);
      }
    };
    saveEmployees();
  }, [employees]);

  // Save shifts to storage whenever they change
  React.useEffect(() => {
    const saveShifts = async () => {
      try {
        await window.storage.set('shifts', JSON.stringify(shifts));
      } catch (error) {
        console.error('Error saving shifts:', error);
      }
    };
    saveShifts();
  }, [shifts]);

  // Save language preference
  React.useEffect(() => {
    const saveLang = async () => {
      try {
        await window.storage.set('language', language);
      } catch (error) {
        console.error('Error saving language:', error);
      }
    };
    saveLang();
  }, [language]);

  // Save dark mode preference
  React.useEffect(() => {
    const saveDarkMode = async () => {
      try {
        await window.storage.set('darkMode', String(darkMode));
      } catch (error) {
        console.error('Error saving dark mode:', error);
      }
    };
    saveDarkMode();
  }, [darkMode]);

  // Save custom statuses
  React.useEffect(() => {
    const saveStatuses = async () => {
      try {
        await window.storage.set('customStatuses', JSON.stringify(customStatuses));
      } catch (error) {
        console.error('Error saving statuses:', error);
      }
    };
    saveStatuses();
  }, [customStatuses]);

  const handleAddCustomStatus = () => {
    if (newStatusName.trim() && !customStatuses.includes(newStatusName.trim())) {
      setCustomStatuses([...customStatuses, newStatusName.trim()]);
      setNewStatusName('');
      setShowAddStatus(false);
    }
  };

  const handleDeleteStatus = (statusToDelete) => {
    if (customStatuses.length > 1) {
      setCustomStatuses(customStatuses.filter(s => s !== statusToDelete));
    }
  };

  const [newShift, setNewShift] = useState({
    employeeId: '',
    date: '',
    startTime: '',
    endTime: '',
    role: ''
  });

  const [newEmployee, setNewEmployee] = useState({
    name: '',
    personalNumber: '',
    sex: '',
    status: '',
    department: ''
  });

  const departments = useMemo(() => {
    return [...new Set(employees.map(e => e.department).filter(Boolean))];
  }, [employees]);

  const statuses = useMemo(() => {
    return [...new Set(employees.map(e => e.status).filter(Boolean))];
  }, [employees]);

  const filteredEmployees = useMemo(() => {
    return employees.filter(emp => {
      const matchesSearch = emp.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           emp.personalNumber.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesDept = !filterDepartment || emp.department === filterDepartment;
      const matchesStatus = !filterStatus || emp.status === filterStatus;
      return matchesSearch && matchesDept && matchesStatus;
    });
  }, [employees, searchTerm, filterDepartment, filterStatus]);

  const hasConflict = (employeeId, date, startTime, endTime, excludeShiftId = null) => {
    const employeeShifts = shifts.filter(s => 
      s.employeeId === employeeId && 
      s.date === date && 
      s.id !== excludeShiftId
    );

    return employeeShifts.some(shift => {
      const newStart = startTime;
      const newEnd = endTime;
      const existingStart = shift.startTime;
      const existingEnd = shift.endTime;

      return (newStart < existingEnd && newEnd > existingStart);
    });
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        const importedEmployees = jsonData.map((row, index) => ({
          id: Date.now() + index,
          name: row.Name || row.name || '',
          personalNumber: row['Personal Number'] || row.personalNumber || row['Personal_Number'] || '',
          sex: row.Sex || row.sex || row.Gender || row.gender || '',
          status: row.Status || row.status || '',
          department: row.Department || row.department || ''
        }));

        setEmployees([...employees, ...importedEmployees]);
        setUploadMessage(`${t.imported} ${importedEmployees.length} ${t.employeesText}`);
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
    const template = [
      {
        Name: 'John Doe',
        'Personal Number': '12345',
        Sex: 'Male',
        Status: 'Active',
        Department: 'Sales'
      },
      {
        Name: 'Jane Smith',
        'Personal Number': '67890',
        Sex: 'Female',
        Status: 'Active',
        Department: 'Marketing'
      }
    ];

    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Employees');
    XLSX.writeFile(wb, 'employee_template.xlsx');
  };

  const handleExportShifts = () => {
    const exportData = shifts.map(shift => {
      const employee = employees.find(e => e.id === shift.employeeId);
      return {
        'Employee Name': employee?.name || 'Unknown',
        'Personal Number': employee?.personalNumber || '',
        'Department': employee?.department || '',
        'Date': shift.date,
        'Start Time': shift.startTime,
        'End Time': shift.endTime,
        'Shift Type': shift.role
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shifts');
    XLSX.writeFile(wb, `shifts_export_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handlePrintSchedule = () => {
    window.print();
  };

  const handleAddShift = () => {
    if (newShift.employeeId && newShift.date && newShift.startTime && newShift.endTime) {
      if (hasConflict(parseInt(newShift.employeeId), newShift.date, newShift.startTime, newShift.endTime)) {
        setUploadMessage(t.conflictWarning);
        setTimeout(() => setUploadMessage(''), 3000);
        return;
      }

      const shift = {
        id: Date.now(),
        employeeId: parseInt(newShift.employeeId),
        date: newShift.date,
        startTime: newShift.startTime,
        endTime: newShift.endTime,
        role: newShift.role || 'Regular'
      };
      setShifts([...shifts, shift]);
      setNewShift({ employeeId: '', date: '', startTime: '', endTime: '', role: '' });
      setShowAddShift(false);
    }
  };

  const handleUpdateShift = () => {
    if (editingShift) {
      if (hasConflict(editingShift.employeeId, editingShift.date, editingShift.startTime, editingShift.endTime, editingShift.id)) {
        setUploadMessage(t.conflictWarning);
        setTimeout(() => setUploadMessage(''), 3000);
        return;
      }
      setShifts(shifts.map(s => s.id === editingShift.id ? editingShift : s));
      setEditingShift(null);
    }
  };

  const handleDeleteShift = (id) => {
    setShifts(shifts.filter(s => s.id !== id));
  };

  const handleAddEmployee = () => {
    if (newEmployee.name && newEmployee.personalNumber) {
      const employee = {
        id: Date.now(),
        name: newEmployee.name,
        personalNumber: newEmployee.personalNumber,
        sex: newEmployee.sex,
        status: newEmployee.status,
        department: newEmployee.department
      };
      setEmployees([...employees, employee]);
      setNewEmployee({ name: '', personalNumber: '', sex: '', status: '', department: '' });
      setShowAddEmployee(false);
    }
  };

  const handleDeleteEmployee = (id) => {
    setEmployees(employees.filter(e => e.id !== id));
    setShifts(shifts.filter(s => s.employeeId !== id));
  };

  const getEmployee = (employeeId) => {
    return employees.find(e => e.id === employeeId);
  };

  const getShiftsByDate = (date) => {
    return shifts.filter(s => s.date === date).sort((a, b) => a.startTime.localeCompare(b.startTime));
  };

  const getDates = () => {
    const dates = new Set(shifts.map(s => s.date));
    const today = new Date().toISOString().split('T')[0];
    if (shifts.length > 0) dates.add(today);
    return Array.from(dates).sort();
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
    const year = now.getFullYear();
    const month = now.getMonth();
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const dates = [];
    
    for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
      dates.push(new Date(d).toISOString().split('T')[0]);
    }
    return dates;
  };

  const navigateWeek = (direction) => {
    const newDate = new Date(currentWeekStart);
    newDate.setDate(newDate.getDate() + (direction * 7));
    setCurrentWeekStart(newDate);
  };

  const conflictingShifts = useMemo(() => {
    const conflicts = new Set();
    shifts.forEach(shift => {
      if (hasConflict(shift.employeeId, shift.date, shift.startTime, shift.endTime, shift.id)) {
        conflicts.add(shift.id);
      }
    });
    return conflicts;
  }, [shifts]);

  const toggleLanguage = () => {
    setLanguage(language === 'en' ? 'he' : 'en');
  };

  const toggleDarkMode = () => {
    setDarkMode(!darkMode);
  };

  return (
    <div className={`min-h-screen p-4 ${darkMode ? 'bg-gray-900' : 'bg-gradient-to-br from-blue-50 to-indigo-100'}`} dir={language === 'he' ? 'rtl' : 'ltr'}>
      <div className="max-w-7xl mx-auto">
        <div className={`rounded-lg shadow-lg overflow-hidden ${darkMode ? 'bg-gray-800' : 'bg-white'}`}>
          <div className={`p-6 ${darkMode ? 'bg-gradient-to-r from-indigo-900 to-purple-900' : 'bg-gradient-to-r from-blue-600 to-indigo-600'} text-white`}>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <Calendar className="w-8 h-8" />
                <div>
                  <h1 className="text-3xl font-bold">{t.appTitle}</h1>
                  <p className="text-sm opacity-90">{t.subtitle}</p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <button
                  onClick={toggleLanguage}
                  className="p-2 bg-white bg-opacity-20 hover:bg-opacity-30 rounded-lg transition-all"
                  title={language === 'en' ? 'Switch to Hebrew' : 'עבור לאנגלית'}
                >
                  <Globe className="w-5 h-5" />
                </button>
                <button
                  onClick={toggleDarkMode}
                  className="p-2 bg-white bg-opacity-20 hover:bg-opacity-30 rounded-lg transition-all"
                  title={darkMode ? 'Light Mode' : 'Dark Mode'}
                >
                  {darkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
                </button>
              </div>
            </div>
          </div>

          {uploadMessage && (
            <div className={`${uploadMessage.includes('⚠️') ? 'bg-yellow-100 border-yellow-500 text-yellow-700' : 'bg-green-100 border-green-500 text-green-700'} ${language === 'he' ? 'border-r-4' : 'border-l-4'} p-4`}>
              <p>{uploadMessage}</p>
            </div>
          )}

          <div className={`flex border-b ${darkMode ? 'border-gray-700' : ''}`}>
            <button
              onClick={() => setActiveTab('employees')}
              className={`flex-1 py-4 px-6 font-medium transition-colors ${
                activeTab === 'employees'
                  ? `${darkMode ? 'bg-gray-700 text-blue-400 border-b-2 border-blue-400' : 'bg-blue-50 text-blue-600 border-b-2 border-blue-600'}`
                  : `${darkMode ? 'text-gray-300 hover:bg-gray-700' : 'text-gray-600 hover:bg-gray-50'}`
              }`}
            >
              <Users className="w-5 h-5 inline mr-2" />
              {t.employees}
            </button>
            <button
              onClick={() => setActiveTab('schedule')}
              className={`flex-1 py-4 px-6 font-medium transition-colors ${
                activeTab === 'schedule'
                  ? `${darkMode ? 'bg-gray-700 text-blue-400 border-b-2 border-blue-400' : 'bg-blue-50 text-blue-600 border-b-2 border-blue-600'}`
                  : `${darkMode ? 'text-gray-300 hover:bg-gray-700' : 'text-gray-600 hover:bg-gray-50'}`
              }`}
            >
              <Clock className="w-5 h-5 inline mr-2" />
              {t.schedule}
            </button>
          </div>

          <div className="p-6">
            {activeTab === 'employees' && (
              <div>
                <div className="flex flex-wrap justify-between items-center gap-4 mb-6">
                  <h2 className={`text-2xl font-bold ${darkMode ? 'text-white' : 'text-gray-800'}`}>{t.employeeManagement}</h2>
                  <div className="flex gap-3">
                    <button
                      onClick={handleDownloadTemplate}
                      className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors ${darkMode ? 'bg-green-700 hover:bg-green-600' : 'bg-green-600 hover:bg-green-700'} text-white`}
                    >
                      <Download className="w-5 h-5" />
                      {t.template}
                    </button>
                    <label className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors cursor-pointer ${darkMode ? 'bg-indigo-700 hover:bg-indigo-600' : 'bg-indigo-600 hover:bg-indigo-700'} text-white`}>
                      <Upload className="w-5 h-5" />
                      {t.import}
                      <input
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={handleFileUpload}
                        className="hidden"
                      />
                    </label>
                    <button
                      onClick={() => setShowAddEmployee(true)}
                      className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors ${darkMode ? 'bg-blue-700 hover:bg-blue-600' : 'bg-blue-600 hover:bg-blue-700'} text-white`}
                    >
                      <Plus className="w-5 h-5" />
                      {t.add}
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                  <div className="relative">
                    <Search className={`absolute ${language === 'he' ? 'right-3' : 'left-3'} top-1/2 transform -translate-y-1/2 w-5 h-5 ${darkMode ? 'text-gray-400' : 'text-gray-400'}`} />
                    <input
                      type="text"
                      placeholder={t.searchPlaceholder}
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className={`w-full ${language === 'he' ? 'pr-10 pl-4' : 'pl-10 pr-4'} py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                    />
                  </div>
                  <div>
                    <select
                      value={filterDepartment}
                      onChange={(e) => setFilterDepartment(e.target.value)}
                      className={`w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                    >
                      <option value="">{t.allDepartments}</option>
                      {departments.map(dept => (
                        <option key={dept} value={dept}>{dept}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <select
                      value={filterStatus}
                      onChange={(e) => setFilterStatus(e.target.value)}
                      className={`w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                    >
                      <option value="">{t.allStatuses}</option>
                      {statuses.map(status => (
                        <option key={status} value={status}>{status}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {showAddEmployee && (
                  <div className={`mb-6 p-4 rounded-lg border-2 ${darkMode ? 'bg-gray-700 border-blue-500' : 'bg-blue-50 border-blue-200'}`}>
                    <div className="flex justify-between items-center mb-4">
                      <h3 className={`text-lg font-semibold ${darkMode ? 'text-white' : 'text-gray-800'}`}>{t.newEmployee}</h3>
                      <button onClick={() => setShowAddEmployee(false)} className={`${darkMode ? 'text-gray-400 hover:text-gray-200' : 'text-gray-500 hover:text-gray-700'}`}>
                        <X className="w-5 h-5" />
                      </button>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.name} *</label>
                        <input
                          type="text"
                          value={newEmployee.name}
                          onChange={(e) => setNewEmployee({ ...newEmployee, name: e.target.value })}
                          className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                          placeholder={t.name}
                        />
                      </div>
                      <div>
                        <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.personalNumber} *</label>
                        <input
                          type="text"
                          value={newEmployee.personalNumber}
                          onChange={(e) => setNewEmployee({ ...newEmployee, personalNumber: e.target.value })}
                          className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                          placeholder={t.personalNumber}
                        />
                      </div>
                      <div>
                        <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.sex}</label>
                        <select
                          value={newEmployee.sex}
                          onChange={(e) => setNewEmployee({ ...newEmployee, sex: e.target.value })}
                          className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                        >
                          <option value="">{t.select}</option>
                          <option value="Male">{t.male}</option>
                          <option value="Female">{t.female}</option>
                          <option value="Other">{t.other}</option>
                        </select>
                      </div>
                      <div>
                        <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.status}</label>
                        <div className="flex gap-2">
                          <select
                            value={newEmployee.status}
                            onChange={(e) => setNewEmployee({ ...newEmployee, status: e.target.value })}
                            className={`flex-1 p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                          >
                            <option value="">{t.select}</option>
                            {customStatuses.map(status => (
                              <option key={status} value={status}>{status}</option>
                            ))}
                          </select>
                          <button
                            onClick={() => setShowAddStatus(true)}
                            className={`p-2 rounded-lg ${darkMode ? 'bg-green-700 hover:bg-green-600' : 'bg-green-600 hover:bg-green-700'} text-white`}
                            title={language === 'he' ? 'הוסף סטטוס חדש' : 'Add new status'}
                          >
                            <Plus className="w-5 h-5" />
                          </button>
                        </div>
                        {showAddStatus && (
                          <div className="mt-2 flex gap-2">
                            <input
                              type="text"
                              value={newStatusName}
                              onChange={(e) => setNewStatusName(e.target.value)}
                              placeholder={language === 'he' ? 'שם סטטוס חדש' : 'New status name'}
                              className={`flex-1 p-2 border rounded-lg text-sm ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                              onKeyPress={(e) => e.key === 'Enter' && handleAddCustomStatus()}
                            />
                            <button
                              onClick={handleAddCustomStatus}
                              className={`p-2 rounded-lg text-white ${darkMode ? 'bg-green-700 hover:bg-green-600' : 'bg-green-600 hover:bg-green-700'}`}
                            >
                              <Save className="w-4 h-4" />
                            </button>
                            <button
                              onClick={() => {
                                setShowAddStatus(false);
                                setNewStatusName('');
                              }}
                              className={`p-2 rounded-lg text-white ${darkMode ? 'bg-gray-700 hover:bg-gray-600' : 'bg-gray-600 hover:bg-gray-700'}`}
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>
                        )}
                        {customStatuses.length > 0 && (
                          <div className="mt-2 flex flex-wrap gap-1">
                            {customStatuses.map(status => (
                              <span key={status} className={`inline-flex items-center gap-1 px-2 py-1 rounded text-xs ${darkMode ? 'bg-gray-600 text-gray-200' : 'bg-gray-200 text-gray-700'}`}>
                                {status}
                                {customStatuses.length > 1 && (
                                  <button
                                    onClick={() => handleDeleteStatus(status)}
                                    className="hover:text-red-600"
                                  >
                                    <X className="w-3 h-3" />
                                  </button>
                                )}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                      <div className="col-span-2">
                        <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.department}</label>
                        <input
                          type="text"
                          value={newEmployee.department}
                          onChange={(e) => setNewEmployee({ ...newEmployee, department: e.target.value })}
                          className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                          placeholder={t.deptPlaceholder}
                        />
                      </div>
                    </div>
                    <button
                      onClick={handleAddEmployee}
                      className={`mt-4 px-4 py-2 rounded-lg w-full text-white ${darkMode ? 'bg-blue-700 hover:bg-blue-600' : 'bg-blue-600 hover:bg-blue-700'}`}
                    >
                      {t.addEmployee}
                    </button>
                  </div>
                )}

                {filteredEmployees.length === 0 ? (
                  <div className={`text-center py-12 rounded-lg border-2 border-dashed ${darkMode ? 'bg-gray-700 border-gray-600' : 'bg-gray-50 border-gray-300'}`}>
                    <Users className={`w-16 h-16 mx-auto mb-4 ${darkMode ? 'text-gray-500' : 'text-gray-400'}`} />
                    <p className={`mb-2 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{employees.length === 0 ? t.noEmployees : t.noMatching}</p>
                    <p className={`text-sm ${darkMode ? 'text-gray-400' : 'text-gray-500'}`}>{employees.length === 0 ? t.importOrAdd : t.adjustFilters}</p>
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full">
                      <thead className={`border-b-2 ${darkMode ? 'bg-gray-700 border-gray-600' : 'bg-gray-100 border-gray-200'}`}>
                        <tr>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.name}</th>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.personalNumber}</th>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.sex}</th>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.status}</th>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.department}</th>
                          <th className={`px-4 py-3 text-${language === 'he' ? 'right' : 'left'} text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.shifts}</th>
                          <th className={`px-4 py-3 text-center text-sm font-semibold ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.actions}</th>
                        </tr>
                      </thead>
                      <tbody className={`divide-y ${darkMode ? 'divide-gray-700' : 'divide-gray-200'}`}>
                        {filteredEmployees.map(employee => {
                          const empShifts = shifts.filter(s => s.employeeId === employee.id);
                          return (
                            <tr key={employee.id} className={`transition-colors ${darkMode ? 'hover:bg-gray-700' : 'hover:bg-gray-50'}`}>
                              <td className="px-4 py-3">
                                <div className="flex items-center gap-3">
                                  <div className={`w-10 h-10 rounded-full flex items-center justify-center text-white font-bold ${darkMode ? 'bg-gradient-to-br from-indigo-600 to-purple-600' : 'bg-gradient-to-br from-blue-500 to-indigo-500'}`}>
                                    {employee.name.charAt(0)}
                                  </div>
                                  <span className={`font-medium ${darkMode ? 'text-gray-200' : 'text-gray-800'}`}>{employee.name}</span>
                                </div>
                              </td>
                              <td className={`px-4 py-3 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{employee.personalNumber}</td>
                              <td className={`px-4 py-3 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{employee.sex || '-'}</td>
                              <td className="px-4 py-3">
                                <span className={`px-2 py-1 rounded-full text-xs font-medium ${
                                  employee.status === 'Active' ? 'bg-green-100 text-green-700' :
                                  employee.status === 'Inactive' ? 'bg-red-100 text-red-700' :
                                  employee.status === 'On Leave' ? 'bg-yellow-100 text-yellow-700' :
                                  employee.status === 'Vacation' ? 'bg-blue-100 text-blue-700' :
                                  employee.status === 'Sick Leave' ? 'bg-orange-100 text-orange-700' :
                                  'bg-gray-100 text-gray-700'
                                }`}>
                                  {employee.status || 'N/A'}
                                </span>
                              </td>
                              <td className={`px-4 py-3 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{employee.department || '-'}</td>
                              <td className={`px-4 py-3 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{empShifts.length}</td>
                              <td className="px-4 py-3 text-center">
                                <button
                                  onClick={() => handleDeleteEmployee(employee.id)}
                                  className={`p-2 rounded transition-colors ${darkMode ? 'text-red-400 hover:bg-gray-600' : 'text-red-600 hover:bg-red-50'}`}
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'schedule' && (
              <div>
                <div className="flex flex-wrap justify-between items-center gap-4 mb-6">
                  <h2 className={`text-2xl font-bold ${darkMode ? 'text-white' : 'text-gray-800'}`}>{t.shiftSchedule}</h2>
                  <div className="flex gap-3">
                    <div className={`flex rounded-lg p-1 ${darkMode ? 'bg-gray-700' : 'bg-gray-100'}`}>
                      <button
                        onClick={() => setCalendarView('list')}
                        className={`px-3 py-1 rounded text-sm ${calendarView === 'list' ? (darkMode ? 'bg-gray-600 shadow' : 'bg-white shadow') : ''}`}
                      >
                        {t.list}
                      </button>
                      <button
                        onClick={() => setCalendarView('week')}
                        className={`px-3 py-1 rounded text-sm ${calendarView === 'week' ? (darkMode ? 'bg-gray-600 shadow' : 'bg-white shadow') : ''}`}
                      >
                        {t.week}
                      </button>
                      <button
                        onClick={() => setCalendarView('month')}
                        className={`px-3 py-1 rounded text-sm ${calendarView === 'month' ? (darkMode ? 'bg-gray-600 shadow' : 'bg-white shadow') : ''}`}
                      >
                        {t.month}
                      </button>
                    </div>
                    {shifts.length > 0 && (
                      <>
                        <button
                          onClick={handleExportShifts}
                          className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors text-white ${darkMode ? 'bg-green-700 hover:bg-green-600' : 'bg-green-600 hover:bg-green-700'}`}
                        >
                          <Download className="w-5 h-5" />
                          {t.export}
                        </button>
                        <button
                          onClick={handlePrintSchedule}
                          className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors text-white ${darkMode ? 'bg-purple-700 hover:bg-purple-600' : 'bg-purple-600 hover:bg-purple-700'}`}
                        >
                          <Printer className="w-5 h-5" />
                          {t.print}
                        </button>
                      </>
                    )}
                    <button
                      onClick={() => setShowAddShift(true)}
                      disabled={employees.length === 0}
                      className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-colors ${
                        employees.length === 0 
                          ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                          : `text-white ${darkMode ? 'bg-blue-700 hover:bg-blue-600' : 'bg-blue-600 hover:bg-blue-700'}`
                      }`}
                    >
                      <Plus className="w-5 h-5" />
                      {t.addShift}
                    </button>
                  </div>
                </div>

                {conflictingShifts.size > 0 && (
                  <div className={`mb-6 p-4 ${language === 'he' ? 'border-r-4' : 'border-l-4'} ${darkMode ? 'bg-yellow-900 border-yellow-600' : 'bg-yellow-100 border-yellow-500'}`}>
                    <div className="flex items-center gap-2">
                      <AlertCircle className={`w-5 h-5 ${darkMode ? 'text-yellow-400' : 'text-yellow-700'}`} />
                      <p className={`font-medium ${darkMode ? 'text-yellow-300' : 'text-yellow-700'}`}>
                        {conflictingShifts.size} {t.conflictDetected}
                      </p>
                    </div>
                  </div>
                )}

                {employees.length === 0 ? (
                  <div className={`text-center py-12 rounded-lg border-2 border-dashed ${darkMode ? 'bg-gray-700 border-gray-600' : 'bg-gray-50 border-gray-300'}`}>
                    <Clock className={`w-16 h-16 mx-auto mb-4 ${darkMode ? 'text-gray-500' : 'text-gray-400'}`} />
                    <p className={`mb-2 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{t.addEmployeesFirst}</p>
                    <p className={`text-sm ${darkMode ? 'text-gray-400' : 'text-gray-500'}`}>{t.needEmployees}</p>
                  </div>
                ) : (
                  <>
                    {showAddShift && (
                      <div className={`mb-6 p-4 rounded-lg border-2 ${darkMode ? 'bg-gray-700 border-blue-500' : 'bg-blue-50 border-blue-200'}`}>
                        <div className="flex justify-between items-center mb-4">
                          <h3 className={`text-lg font-semibold ${darkMode ? 'text-white' : 'text-gray-800'}`}>{t.newShift}</h3>
                          <button onClick={() => setShowAddShift(false)} className={`${darkMode ? 'text-gray-400 hover:text-gray-200' : 'text-gray-500 hover:text-gray-700'}`}>
                            <X className="w-5 h-5" />
                          </button>
                        </div>
                        <div className="grid grid-cols-2 gap-4">
                          <div>
                            <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.employee}</label>
                            <select
                              value={newShift.employeeId}
                              onChange={(e) => setNewShift({ ...newShift, employeeId: e.target.value })}
                              className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                            >
                              <option value="">{t.selectEmployee}</option>
                              {employees.map(emp => (
                                <option key={emp.id} value={emp.id}>
                                  {emp.name} - {emp.department || t.noDept}
                                </option>
                              ))}
                            </select>
                          </div>
                          <div>
                            <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.date}</label>
                            <input
                              type="date"
                              value={newShift.date}
                              onChange={(e) => setNewShift({ ...newShift, date: e.target.value })}
                              className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                            />
                          </div>
                          <div>
                            <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.startTime}</label>
                            <input
                              type="time"
                              value={newShift.startTime}
                              onChange={(e) => setNewShift({ ...newShift, startTime: e.target.value })}
                              className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                            />
                          </div>
                          <div>
                            <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.endTime}</label>
                            <input
                              type="time"
                              value={newShift.endTime}
                              onChange={(e) => setNewShift({ ...newShift, endTime: e.target.value })}
                              className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                            />
                          </div>
                          <div className="col-span-2">
                            <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>{t.shiftType}</label>
                            <input
                              type="text"
                              value={newShift.role}
                              onChange={(e) => setNewShift({ ...newShift, role: e.target.value })}
                              placeholder={t.shiftTypePlaceholder}
                              className={`w-full p-2 border rounded-lg ${darkMode ? 'bg-gray-600 border-gray-500 text-white' : ''}`}
                            />
                          </div>
                        </div>
                        <button
                          onClick={handleAddShift}
                          className={`mt-4 px-4 py-2 rounded-lg w-full text-white ${darkMode ? 'bg-blue-700 hover:bg-blue-600' : 'bg-blue-600 hover:bg-blue-700'}`}
                        >
                          {t.addShift}
                        </button>
                      </div>
                    )}

                    {calendarView === 'list' && (
                      <>
                        {shifts.length === 0 ? (
                          <div className={`text-center py-12 rounded-lg border-2 border-dashed ${darkMode ? 'bg-gray-700 border-gray-600' : 'bg-gray-50 border-gray-300'}`}>
                            <Calendar className={`w-16 h-16 mx-auto mb-4 ${darkMode ? 'text-gray-500' : 'text-gray-400'}`} />
                            <p className={`mb-2 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>{t.noShifts}</p>
                            <p className={`text-sm ${darkMode ? 'text-gray-400' : 'text-gray-500'}`}>{t.clickToAdd}</p>
                          </div>
                        ) : (
                          <div className="space-y-6">
                            {getDates().map(date => {
                              const dayShifts = getShiftsByDate(date);
                              if (dayShifts.length === 0) return null;
                              
                              return (
                                <div key={date} className={`border rounded-lg overflow-hidden ${darkMode ? 'border-gray-700' : ''}`}>
                                  <div className={`px-4 py-3 border-b ${darkMode ? 'bg-gray-700 border-gray-600' : 'bg-gray-100'}`}>
                                    <h3 className={`font-semibold ${darkMode ? 'text-gray-200' : 'text-gray-800'}`}>
                                      {new Date(date + 'T00:00:00').toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { 
                                        weekday: 'long', 
                                        year: 'numeric', 
                                        month: 'long', 
                                        day: 'numeric' 
                                      })}
                                    </h3>
                                  </div>
                                  <div className={`divide-y ${darkMode ? 'divide-gray-700' : ''}`}>
                                    {dayShifts.map(shift => {
                                      const employee = getEmployee(shift.employeeId);
                                      const hasConflictFlag = conflictingShifts.has(shift.id);
                                      return (
                                        <div key={shift.id}>
                                          {editingShift?.id === shift.id ? (
                                            <div className={`p-4 ${darkMode ? 'bg-gray-600' : 'bg-yellow-50'}`}>
                                              <div className="grid grid-cols-4 gap-4">
                                                <select
                                                  value={editingShift.employeeId}
                                                  onChange={(e) => setEditingShift({ ...editingShift, employeeId: parseInt(e.target.value) })}
                                                  className={`p-2 border rounded ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                                                >
                                                  {employees.map(emp => (
                                                    <option key={emp.id} value={emp.id}>{emp.name}</option>
                                                  ))}
                                                </select>
                                                <input
                                                  type="time"
                                                  value={editingShift.startTime}
                                                  onChange={(e) => setEditingShift({ ...editingShift, startTime: e.target.value })}
                                                  className={`p-2 border rounded ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                                                />
                                                <input
                                                  type="time"
                                                  value={editingShift.endTime}
                                                  onChange={(e) => setEditingShift({ ...editingShift, endTime: e.target.value })}
                                                  className={`p-2 border rounded ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : ''}`}
                                                />
                                                <div className="flex gap-2">
                                                  <button
                                                    onClick={handleUpdateShift}
                                                    className={`flex-1 px-3 py-2 rounded flex items-center justify-center gap-1 text-white ${darkMode ? 'bg-green-700 hover:bg-green-600' : 'bg-green-600 hover:bg-green-700'}`}
                                                  >
                                                    <Save className="w-4 h-4" />
                                                  </button>
                                                  <button
                                                    onClick={() => setEditingShift(null)}
                                                    className={`flex-1 px-3 py-2 rounded flex items-center justify-center gap-1 text-white ${darkMode ? 'bg-gray-700 hover:bg-gray-600' : 'bg-gray-600 hover:bg-gray-700'}`}
                                                  >
                                                    <X className="w-4 h-4" />
                                                  </button>
                                                </div>
                                              </div>
                                            </div>
                                          ) : (
                                            <div className={`p-4 flex items-center justify-between transition-colors ${hasConflictFlag ? (darkMode ? 'bg-red-900' : 'bg-red-50') : (darkMode ? 'hover:bg-gray-700' : 'hover:bg-gray-50')}`}>
                                              <div className="flex items-center gap-4 flex-1">
                                                {hasConflictFlag && (
                                                  <AlertCircle className={`w-5 h-5 ${darkMode ? 'text-red-400' : 'text-red-600'}`} />
                                                )}
                                                <div className="w-40">
                                                  <div className={`font-semibold ${darkMode ? 'text-gray-200' : 'text-gray-800'}`}>{employee?.name || 'Unknown'}</div>
                                                  <div className={`text-xs ${darkMode ? 'text-gray-400' : 'text-gray-500'}`}>{employee?.department || t.noDepartment}</div>
                                                </div>
                                                <div className={`flex items-center gap-2 ${darkMode ? 'text-gray-300' : 'text-gray-600'}`}>
                                                  <Clock className="w-4 h-4" />
                                                  <span>{shift.startTime} - {shift.endTime}</span>
                                                </div>
                                                <div className={`px-3 py-1 rounded-full text-sm ${darkMode ? 'bg-blue-900 text-blue-300' : 'bg-blue-100 text-blue-700'}`}>
                                                  {shift.role}
                                                </div>
                                              </div>
                                              <div className="flex gap-2">
                                                <button
                                                  onClick={() => setEditingShift(shift)}
                                                  className={`p-2 rounded transition-colors ${darkMode ? 'text-blue-400 hover:bg-gray-600' : 'text-blue-600 hover:bg-blue-50'}`}
                                                >
                                                  <Edit2 className="w-4 h-4" />
                                                </button>
                                                <button
                                                  onClick={() => handleDeleteShift(shift.id)}
                                                  className={`p-2 rounded transition-colors ${darkMode ? 'text-red-400 hover:bg-gray-600' : 'text-red-600 hover:bg-red-50'}`}
                                                >
                                                  <Trash2 className="w-4 h-4" />
                                                </button>
                                              </div>
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
                        )}
                      </>
                    )}

                    {calendarView === 'week' && (
                      <div>
                        <div className="flex items-center justify-between mb-4">
                          <button
                            onClick={() => navigateWeek(-1)}
                            className={`px-4 py-2 rounded-lg ${darkMode ? 'bg-gray-700 hover:bg-gray-600 text-white' : 'bg-gray-200 hover:bg-gray-300'}`}
                          >
                            {t.previousWeek}
                          </button>
                          <h3 className={`font-semibold text-lg ${darkMode ? 'text-white' : ''}`}>
                            {t.weekOf} {new Date(currentWeekStart).toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'long', day: 'numeric', year: 'numeric' })}
                          </h3>
                          <button
                            onClick={() => navigateWeek(1)}
                            className={`px-4 py-2 rounded-lg ${darkMode ? 'bg-gray-700 hover:bg-gray-600 text-white' : 'bg-gray-200 hover:bg-gray-300'}`}
                          >
                            {t.nextWeek}
                          </button>
                        </div>
                        <div className="grid grid-cols-7 gap-2">
                          {getWeekDates(currentWeekStart).map(date => {
                            const dayShifts = getShiftsByDate(date);
                            const dayName = new Date(date + 'T00:00:00').toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { weekday: 'short' });
                            const dayNum = new Date(date + 'T00:00:00').getDate();
                            return (
                              <div key={date} className={`border rounded-lg overflow-hidden ${darkMode ? 'border-gray-700' : ''}`}>
                                <div className={`p-2 text-center ${darkMode ? 'bg-gray-700' : 'bg-gray-100'}`}>
                                  <div className={`font-semibold text-sm ${darkMode ? 'text-gray-300' : ''}`}>{dayName}</div>
                                  <div className={`text-xl font-bold ${darkMode ? 'text-white' : ''}`}>{dayNum}</div>
                                </div>
                                <div className="p-2 min-h-32 space-y-1">
                                  {dayShifts.map(shift => {
                                    const employee = getEmployee(shift.employeeId);
                                    const hasConflictFlag = conflictingShifts.has(shift.id);
                                    return (
                                      <div key={shift.id} className={`text-xs p-2 rounded ${hasConflictFlag ? 'bg-red-100 border border-red-300' : (darkMode ? 'bg-blue-900' : 'bg-blue-50')}`}>
                                        <div className={`font-semibold truncate ${darkMode && !hasConflictFlag ? 'text-blue-300' : ''}`}>{employee?.name}</div>
                                        <div className={`${darkMode && !hasConflictFlag ? 'text-blue-400' : 'text-gray-600'}`}>{shift.startTime}-{shift.endTime}</div>
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
                      <div>
                        <h3 className={`font-semibold text-lg mb-4 text-center ${darkMode ? 'text-white' : ''}`}>
                          {new Date().toLocaleDateString(language === 'he' ? 'he-IL' : 'en-US', { month: 'long', year: 'numeric' })}
                        </h3>
                        <div className="grid grid-cols-7 gap-2">
                          {[t.sun, t.mon, t.tue, t.wed, t.thu, t.fri, t.sat].map(day => (
                            <div key={day} className={`text-center font-semibold text-sm p-2 ${darkMode ? 'bg-gray-700 text-gray-300' : 'bg-gray-100'}`}>
                              {day}
                            </div>
                          ))}
                          {getMonthDates().map(date => {
                            const dayShifts = getShiftsByDate(date);
                            const dayNum = new Date(date + 'T00:00:00').getDate();
                            return (
                              <div key={date} className={`border rounded aspect-square overflow-hidden ${darkMode ? 'border-gray-700' : ''}`}>
                                <div className={`text-${language === 'he' ? 'left' : 'right'} p-1 text-sm font-semibold ${darkMode ? 'text-gray-300' : ''}`}>{dayNum}</div>
                                <div className="px-1 space-y-0.5 text-xs">
                                  {dayShifts.slice(0, 2).map(shift => {
                                    const employee = getEmployee(shift.employeeId);
                                    const hasConflictFlag = conflictingShifts.has(shift.id);
                                    return (
                                      <div key={shift.id} className={`p-1 rounded truncate ${hasConflictFlag ? 'bg-red-100' : (darkMode ? 'bg-blue-900 text-blue-300' : 'bg-blue-100')}`}>
                                        {employee?.name}
                                      </div>
                                    );
                                  })}
                                  {dayShifts.length > 2 && (
                                    <div className={`text-center ${darkMode ? 'text-gray-400' : 'text-gray-500'}`}>+{dayShifts.length - 2}</div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}