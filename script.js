/***************************************************
 *               Local Storage Logic
 ***************************************************/
function saveToLocalStorage() {
  localStorage.setItem('employees', JSON.stringify(employees));
  localStorage.setItem('meetings', JSON.stringify(meetings));
  localStorage.setItem('attendanceRecords', JSON.stringify(attendanceRecords));
  localStorage.setItem('tardinessRecords', JSON.stringify(tardinessRecords));
  localStorage.setItem('tardinessIdCounter', String(tardinessIdCounter));
}

function loadFromLocalStorage() {
  const lsEmployees = localStorage.getItem('employees');
  if(lsEmployees) {
    employees = JSON.parse(lsEmployees);
  } else {
    // נתוני דוגמה ראשוניים
    employees = [
      {
        id: 1,
        name: 'שרה',
        role: 'מפתחת',
        startDate: '2025-01-01',
        endDate: '',
        status: 'active',     // active/inactive
        percentPosition: 100
      },
      {
        id: 2,
        name: 'דוד',
        role: 'מעצב',
        startDate: '2025-02-01',
        endDate: '',
        status: 'active',
        percentPosition: 80
      },
    ];
  }

  const lsMeetings = localStorage.getItem('meetings');
  if(lsMeetings) {
    meetings = JSON.parse(lsMeetings);
  } else {
    meetings = [
      { id: 1, title: 'סיעור מוחות', date: '2025-02-01', employeeId: 1 },
      { id: 2, title: 'עדכון פרויקט', date: '2025-02-10', employeeId: 2 },
    ];
  }

  const lsAttendance = localStorage.getItem('attendanceRecords');
  if(lsAttendance) {
    attendanceRecords = JSON.parse(lsAttendance);
  }

  const lsTardiness = localStorage.getItem('tardinessRecords');
  if(lsTardiness) {
    tardinessRecords = JSON.parse(lsTardiness);
  }

  const lsTardinessId = localStorage.getItem('tardinessIdCounter');
  if(lsTardinessId) {
    tardinessIdCounter = parseInt(lsTardinessId, 10);
  } else {
    tardinessIdCounter = 1;
  }
}

/***************************************************
 *            Global State & Data
 ***************************************************/
let employees = [];
let meetings = [];
let attendanceRecords = {};  // { empId: { "YYYY-MM-DD": { present:false, reason } } }
let tardinessRecords = {};   // { empId: { "YYYY-MM-DD": [ { id, reason, shiftStart, arrivalTime, minutesLate }, ... ] } }
let tardinessIdCounter = 1;

loadFromLocalStorage();

let currentTab = 'dashboard';
let globalFilterMonth = getYearMonth(getToday());
let alerts = [];

// שדות עבור עובד חדש
let newEmployeeName = '';
let newEmployeeRole = '';
let newEmployeeStartDate = '';
let newEmployeeStatus = 'active'; // active/inactive
let newEmployeePercentPosition = 100;

// Meetings
let newMeetingTitle = '';
let newMeetingDate = getToday();
let newMeetingEmployee = '';

// Attendance
let attendanceSelectedEmployee = '';
let attendanceDate = getToday();
let attendanceReason = '';
let attendanceExpanded = {};

// Tardiness
let tardySelectedEmployee = '';
let tardyReason = '';
let tardyDate = getToday();
let tardinessExpanded = {};
let newTardyShiftStart = '08:00';  // שעת תחילת משמרת ברירת מחדל
let newTardyArrivalTime = '08:00'; // שעת הגעה בפועל

/***************************************************
 *         Basic Helper Functions
 ***************************************************/
function getYearMonth(dateStr) {
  return dateStr.slice(0, 7); // "YYYY-MM"
}
function getToday() {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// מחשב הבדל דקות בין "HH:MM" ל־"HH:MM"
function calcTimeDiffInMinutes(time1, time2) {
  // time1, time2 in "HH:MM"
  const [h1,m1] = time1.split(':').map(Number);
  const [h2,m2] = time2.split(':').map(Number);
  const total1 = h1*60 + m1;
  const total2 = h2*60 + m2;
  return total2 - total1;
}

/***************************************************
 *          Alerts Logic
 ***************************************************/
function updateAlerts() {
  // לא נשתמש יותר בראש הדף, אלא רק עבור סמל קריאה
  alerts = [];
  // נשמור כאן את רשימת employeeId שיש להם 2+ איחורים
  // אבל בפועל נציב את הסימן ליד השם.
  // לצורך הדגמה, נשים ב-alerts את ה-id:
  employees.forEach(emp => {
    if(emp.status==='active') {
      const tCount = getTardinessCountForMonth(emp.id);
      if(tCount >= 2) {
        alerts.push(emp.id); // שמירה לצורך זיהוי
      }
    }
  });
}

/***************************************************
 *         Attendance (Absences) Logic
 ***************************************************/
// הוספנו חסימה: לא לאפשר חיסור כפול באותו יום לאותו עובד
function markAbsence(employeeId, dateString, reason) {
  if(!attendanceRecords[employeeId]) {
    attendanceRecords[employeeId] = {};
  }
  if(attendanceRecords[employeeId][dateString]) {
    alert('כבר דווח חיסור לעובד זה באותו היום!');
    return;
  }
  attendanceRecords[employeeId][dateString] = {
    present: false,
    reason: reason || ''
  };
  saveToLocalStorage();
}

function getAbsencesCountForMonth(employeeId) {
  if(!attendanceRecords[employeeId]) return 0;
  let count = 0;
  for(const dateStr in attendanceRecords[employeeId]) {
    const rec = attendanceRecords[employeeId][dateStr];
    if(!rec.present && getYearMonth(dateStr)===globalFilterMonth) {
      count++;
    }
  }
  return count;
}
function getAbsencesDetailForMonth(employeeId) {
  if(!attendanceRecords[employeeId]) return [];
  const result = [];
  for(const dateStr in attendanceRecords[employeeId]) {
    if(getYearMonth(dateStr)===globalFilterMonth) {
      const rec = attendanceRecords[employeeId][dateStr];
      if(!rec.present) {
        result.push({ date: dateStr, reason: rec.reason });
      }
    }
  }
  result.sort((a,b)=> (a.date<b.date ? -1:1));
  return result;
}
function editAbsence(employeeId, dateString) {
  const rec = attendanceRecords[employeeId][dateString];
  if(!rec) return;
  const newDate = prompt('ערוך תאריך (YYYY-MM-DD):', dateString);
  if(newDate===null) return;
  const newReason = prompt('ערוך סיבה:', rec.reason);
  if(newReason===null) return;

  // בדיקה האם כבר קיים באותו יום
  if(attendanceRecords[employeeId][newDate] && newDate!==dateString) {
    alert('כבר דווח חיסור באותו יום - אי אפשר לעדכן לתאריך זה!');
    return;
  }
  delete attendanceRecords[employeeId][dateString];
  attendanceRecords[employeeId][newDate] = { present: false, reason: newReason };
  saveToLocalStorage();
  render();
}
function deleteAbsence(employeeId, dateString) {
  if(!confirm('למחוק חיסור זה?')) return;
  delete attendanceRecords[employeeId][dateString];
  saveToLocalStorage();
  render();
}

/***************************************************
 *         Tardiness (איחורים) Logic
 ***************************************************/
// לא מאפשרים יותר מרשומה אחת ביום לאותו עובד
function addTardiness(employeeId, dateString, reason, shiftStart, arrivalTime) {
  if(!reason) reason='';
  if(!tardinessRecords[employeeId]) {
    tardinessRecords[employeeId] = {};
  }
  if(!tardinessRecords[employeeId][dateString]) {
    tardinessRecords[employeeId][dateString] = [];
  }
  // בודקים האם כבר קיים איחור ליום זה
  if(tardinessRecords[employeeId][dateString].length>0) {
    alert('כבר דווח איחור לעובד זה באותו היום!');
    return;
  }
  const lateMinutes = calcTimeDiffInMinutes(shiftStart, arrivalTime);
  // אם lateMinutes <= 0 אז לא באמת איחור, אפשר לבחור איך להתייחס...
  let finalMinutes = lateMinutes>0 ? lateMinutes : 0;
  tardinessRecords[employeeId][dateString].push({
    id: tardinessIdCounter++,
    reason,
    shiftStart,
    arrivalTime,
    minutesLate: finalMinutes
  });
  saveToLocalStorage();
}
function getTardinessCountForMonth(employeeId) {
  if(!tardinessRecords[employeeId]) return 0;
  let total=0;
  for(const dateStr in tardinessRecords[employeeId]) {
    if(getYearMonth(dateStr)===globalFilterMonth) {
      total += tardinessRecords[employeeId][dateStr].length;
    }
  }
  return total;
}
function getTardinessDetailForMonth(employeeId) {
  if(!tardinessRecords[employeeId]) return [];
  const result = [];
  for(const dateStr in tardinessRecords[employeeId]) {
    if(getYearMonth(dateStr)===globalFilterMonth) {
      tardinessRecords[employeeId][dateStr].forEach(r=>{
        result.push({ date: dateStr, reasonObj: r });
      });
    }
  }
  result.sort((a,b)=>(a.reasonObj.id<b.reasonObj.id?-1:1));
  return result;
}
function editTardiness(employeeId, dateString, tardyId) {
  const arr = tardinessRecords[employeeId][dateString];
  if(!arr) return;
  const item = arr.find(i=>i.id===tardyId);
  if(!item) return;

  const newDate = prompt('ערוך תאריך (YYYY-MM-DD):', dateString);
  if(newDate===null) return;
  const newReason = prompt('ערוך סיבת איחור:', item.reason);
  if(newReason===null) return;

  // עריכת שעת תחילת משמרת
  const newShiftStart = prompt('ערוך שעת תחילת משמרת (HH:MM):', item.shiftStart);
  if(newShiftStart===null) return;
  const newArrivalTime = prompt('ערוך שעת הגעה (HH:MM):', item.arrivalTime);
  if(newArrivalTime===null) return;

  // בדיקה אם כבר יש איחור ביום הזה
  if(!tardinessRecords[employeeId][newDate]) {
    tardinessRecords[employeeId][newDate] = [];
  }
  if(newDate!==dateString && tardinessRecords[employeeId][newDate].length>0) {
    alert('כבר דווח איחור לעובד זה בתאריך החדש. לא ניתן לעדכן.');
    return;
  }
  // הסרה מהישן
  const idx= arr.indexOf(item);
  arr.splice(idx,1);
  // הוספה לחדש
  const lateMinutes = calcTimeDiffInMinutes(newShiftStart, newArrivalTime);
  tardinessRecords[employeeId][newDate].push({
    id: tardinessIdCounter++,
    reason: newReason,
    shiftStart: newShiftStart,
    arrivalTime: newArrivalTime,
    minutesLate: (lateMinutes>0? lateMinutes: 0)
  });
  saveToLocalStorage();
  render();
}
function deleteTardiness(employeeId, dateString, tardyId) {
  if(!confirm('למחוק איחור זה?')) return;
  const arr = tardinessRecords[employeeId][dateString];
  if(!arr) return;
  const idx = arr.findIndex(i=>i.id===tardyId);
  if(idx!==-1) arr.splice(idx,1);
  saveToLocalStorage();
  render();
}

/***************************************************
 *            Meetings (פגישות) Logic
 ***************************************************/
function getMeetingsCountForMonth(employeeId) {
  return meetings.filter(m=>
    m.employeeId===employeeId &&
    getYearMonth(m.date)===globalFilterMonth
  ).length;
}

/***************************************************
 *       Export to "Excel" (CSV) + Send Email
 ***************************************************/
function downloadCSV(csvContent, filename) {
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function exportToExcel(section) {
  let csv = '';
  switch(section) {
    case 'Dashboard': {
      csv += 'שם עובד,אחוז משרה,איחורים,חיסורים,פגישות,אחוז עמידה במשרה\n';
      employees.forEach(emp=>{
        if(emp.status==='active'){
          const tardies = getTardinessCountForMonth(emp.id);
          const absences= getAbsencesCountForMonth(emp.id);
          const meets   = getMeetingsCountForMonth(emp.id);
          const posRate = calculatePositionRate(emp.id);
          csv += `${emp.name},${emp.percentPosition}%,${tardies},${absences},${meets},${posRate}%\n`;
        }
      });
      downloadCSV(csv, 'dashboard.csv');
      break;
    }
    case 'Tardiness': {
      csv += 'שם עובד,תאריך איחור,סיבה,שעת תחילת משמרת,שעת הגעה,איחור(דקות)\n';
      employees.forEach(emp=>{
        const details = getTardinessDetailForMonth(emp.id);
        details.forEach(item=>{
          csv += `${emp.name},${item.date},${item.reasonObj.reason},${item.reasonObj.shiftStart},${item.reasonObj.arrivalTime},${item.reasonObj.minutesLate}\n`;
        });
      });
      downloadCSV(csv, 'tardiness.csv');
      break;
    }
    case 'Attendance': {
      csv += 'שם עובד,תאריך חיסור,סיבה\n';
      employees.forEach(emp=>{
        const absList= getAbsencesDetailForMonth(emp.id);
        absList.forEach(a=>{
          csv += `${emp.name},${a.date},${a.reason}\n`;
        });
      });
      downloadCSV(csv, 'attendance.csv');
      break;
    }
    case 'Meetings': {
      csv += 'שם עובד,תאריך פגישה,כותרת\n';
      employees.forEach(emp=>{
        const arr = meetings.filter(m=>
          m.employeeId===emp.id &&
          getYearMonth(m.date)===globalFilterMonth
        );
        arr.forEach(m=>{
          csv += `${emp.name},${m.date},${m.title}\n`;
        });
      });
      downloadCSV(csv, 'meetings.csv');
      break;
    }
    default:
      alert('ייצוא לא מוגדר עבור: ' + section);
  }
}

// שליחת מייל עם "קובץ אקסל" (CSV). לצורך הדגמה נבצע יצירה של CSV טקסט, ונצרף ב-"mailto"...
function sendDataByEmail(section) {
  let csv = '';
  let subject = `דו"ח ${section}`;
  let body = '';
  // ניצור CSV בקובץ
  switch(section) {
    case 'Dashboard':
      csv += 'שם עובד,אחוז משרה,איחורים,חיסורים,פגישות\n';
      employees.forEach(emp=>{
        if(emp.status==='active') {
          const tardies = getTardinessCountForMonth(emp.id);
          const absences= getAbsencesCountForMonth(emp.id);
          const meets= getMeetingsCountForMonth(emp.id);
          csv += `${emp.name},${emp.percentPosition}%,${tardies},${absences},${meets}\n`;
        }
      });
      break;
    case 'Tardiness':
      csv += 'שם עובד,תאריך איחור,סיבה,שעת תחילת משמרת,שעת הגעה,איחור(דקות)\n';
      employees.forEach(emp=>{
        const details = getTardinessDetailForMonth(emp.id);
        details.forEach(item=>{
          csv += `${emp.name},${item.date},${item.reasonObj.reason},${item.reasonObj.shiftStart},${item.reasonObj.arrivalTime},${item.reasonObj.minutesLate}\n`;
        });
      });
      break;
    case 'Attendance':
      csv += 'שם עובד,תאריך חיסור,סיבה\n';
      employees.forEach(emp=>{
        const absList = getAbsencesDetailForMonth(emp.id);
        absList.forEach(a=>{
          csv += `${emp.name},${a.date},${a.reason}\n`;
        });
      });
      break;
    case 'Meetings':
      csv += 'שם עובד,תאריך פגישה,כותרת\n';
      employees.forEach(emp=>{
        const arr = meetings.filter(m=>
          m.employeeId===emp.id && getYearMonth(m.date)===globalFilterMonth
        );
        arr.forEach(m=>{
          csv += `${emp.name},${m.date},${m.title}\n`;
        });
      });
      break;
    default:
      alert('לא קיים ייצוא/מייל עבור: ' + section);
      return;
  }
  // למייל, נדרוש כתובת
  const email = prompt('הזן כתובת מייל לשליחה:');
  if(!email) return;

  // נצטרך להפוך את ה-CSV ל-URI מוכר. (פשוט נכניס לגוף הטקסט, לא באמת צרופה...)
  // במערכת אמיתית נשתמש ב-API לשליחת מייל.
  body = encodeURIComponent(csv);
  subject = encodeURIComponent(subject);

  // mailto
  const mailtoLink = `mailto:${email}?subject=${subject}&body=${body}`;
  window.location.href = mailtoLink;
}

/***************************************************
 *          Month & Tab Handling
 ***************************************************/
const globalMonthInput = document.getElementById('globalMonth');
globalMonthInput.value = globalFilterMonth;
globalMonthInput.onchange=(e)=>{
  globalFilterMonth = e.target.value; // "YYYY-MM"
  render();
};

function render() {
  renderNavigation();
  updateAlerts();

  const contentDiv = document.getElementById('mainContent');
  contentDiv.innerHTML = '';

  switch(currentTab) {
    case 'dashboard': 
      renderDashboard(contentDiv);
      break;
    case 'employees':
      renderEmployees(contentDiv);
      break;
    case 'attendance':
      renderMonthlyAttendance(contentDiv);
      break;
    case 'tardiness':
      renderTardiness(contentDiv);
      break;
    case 'meetings':
      renderMeetings(contentDiv);
      break;
    default:
      break;
  }
}

/***************************************************
 *          Navigation (Tabs)
 ***************************************************/
function renderNavigation(){
  const nav = document.getElementById('navButtons');
  nav.innerHTML='';
  const tabs=[
    {key:'dashboard', label:'מסך ראשי'},
    {key:'employees', label:'עובדים'},
    {key:'attendance', label:'נוכחות חודשית'},
    {key:'tardiness', label:'איחורים'},
    {key:'meetings', label:'פגישות'},
  ];
  tabs.forEach(tab=>{
    const btn = document.createElement('button');
    btn.className = `px-4 py-2 border rounded-lg transition ${
      currentTab===tab.key
        ? 'bg-blue-500 text-white border-blue-500'
        : 'bg-white text-blue-500 border-blue-500 hover:bg-blue-50'
    }`;
    btn.textContent= tab.label;
    btn.onclick=()=>{
      currentTab=tab.key;
      render();
    };
    nav.appendChild(btn);
  });
}

/***************************************************
 *          Dashboard
 ***************************************************/
// חישוב "אחוז עמידה במשרה"
function calculatePositionRate(empId) {
  const emp = employees.find(e=> e.id===empId);
  if(!emp) return 0;
  const baseDays=20;
  const potentialDays = baseDays*(emp.percentPosition/100);
  const absCount = getAbsencesCountForMonth(empId);
  const actualDays = potentialDays-absCount;
  if(potentialDays<=0) return 0;
  return Math.round((actualDays/potentialDays)*100);
}

// סיכומים כלליים לכלל העובדים
function getAllActiveEmployeesCount(){
  return employees.filter(e=> e.status==='active').length;
}
function getAllTardiesThisMonth(){
  let sum=0;
  employees.forEach(emp=>{
    if(emp.status==='active') sum+= getTardinessCountForMonth(emp.id);
  });
  return sum;
}
function getAllAbsencesThisMonth(){
  let sum=0;
  employees.forEach(emp=>{
    if(emp.status==='active') sum+= getAbsencesCountForMonth(emp.id);
  });
  return sum;
}
function getAllMeetingsThisMonth(){
  let sum=0;
  employees.forEach(emp=>{
    if(emp.status==='active') {
      sum += getMeetingsCountForMonth(emp.id);
    }
  });
  return sum;
}

function renderDashboard(container){
  // כותרת: ריבועי סיכום (כמות עובדים, איחורים, חיסורים, פגישות)
  const statsContainer = document.createElement('div');
  statsContainer.className = 'grid grid-cols-2 md:grid-cols-4 gap-4 mb-4';

  // כמות עובדים פעילים
  {
    const box = document.createElement('div');
    box.className = 'bg-white shadow p-4 rounded flex flex-col items-center justify-center';
    const label = document.createElement('div');
    label.className='text-gray-600 font-semibold mb-2';
    label.textContent='עובדים פעילים';
    const val = document.createElement('div');
    val.className='text-2xl font-bold text-blue-600';
    val.textContent= getAllActiveEmployeesCount();
    box.appendChild(label);
    box.appendChild(val);
    statsContainer.appendChild(box);
  }
  // כמות איחורים
  {
    const box = document.createElement('div');
    box.className = 'bg-white shadow p-4 rounded flex flex-col items-center justify-center';
    const label = document.createElement('div');
    label.className='text-gray-600 font-semibold mb-2';
    label.textContent='איחורים החודש';
    const val = document.createElement('div');
    val.className='text-2xl font-bold text-blue-600';
    val.textContent= getAllTardiesThisMonth();
    box.appendChild(label);
    box.appendChild(val);
    statsContainer.appendChild(box);
  }
  // כמות חיסורים
  {
    const box = document.createElement('div');
    box.className = 'bg-white shadow p-4 rounded flex flex-col items-center justify-center';
    const label = document.createElement('div');
    label.className='text-gray-600 font-semibold mb-2';
    label.textContent='חיסורים החודש';
    const val = document.createElement('div');
    val.className='text-2xl font-bold text-blue-600';
    val.textContent= getAllAbsencesThisMonth();
    box.appendChild(label);
    box.appendChild(val);
    statsContainer.appendChild(box);
  }
  // כמות פגישות
  {
    const box = document.createElement('div');
    box.className = 'bg-white shadow p-4 rounded flex flex-col items-center justify-center';
    const label = document.createElement('div');
    label.className='text-gray-600 font-semibold mb-2';
    label.textContent='פגישות החודש';
    const val = document.createElement('div');
    val.className='text-2xl font-bold text-blue-600';
    val.textContent= getAllMeetingsThisMonth();
    box.appendChild(label);
    box.appendChild(val);
    statsContainer.appendChild(box);
  }

  container.appendChild(statsContainer);

  // טבלה מפורטת
  const card = document.createElement('div');
  card.className='p-4 shadow-lg rounded-2xl bg-white';

  const title= document.createElement('h2');
  title.className='text-2xl mb-2 font-bold text-gray-700';
  title.textContent='מסך ראשי - סיכום חודשי';
  card.appendChild(title);

  const tableContainer= document.createElement('div');
  tableContainer.className='overflow-x-auto mt-4';

  const table = document.createElement('table');
  table.className='min-w-full text-right border-collapse';

  const thead = document.createElement('thead');
  thead.innerHTML=`
    <tr class="border-b">
      <th class="px-4 py-2">עובד</th>
      <th class="px-4 py-2">אחוז משרה</th>
      <th class="px-4 py-2">איחורים</th>
      <th class="px-4 py-2">חיסורים</th>
      <th class="px-4 py-2">אחוז עמידה במשרה</th>
      <th class="px-4 py-2">פגישות</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody= document.createElement('tbody');
  employees.forEach(e=>{
    if(e.status!=='active') return; // רק פעילים
    const tardies= getTardinessCountForMonth(e.id);
    const absences= getAbsencesCountForMonth(e.id);
    const meets= getMeetingsCountForMonth(e.id);
    const posRate= calculatePositionRate(e.id);

    const row= document.createElement('tr');
    row.className='border-b hover:bg-gray-50';

    // שם עובד + סימן קריאה אדום אם tCount>=2
    const nameTd= document.createElement('td');
    nameTd.className='px-4 py-2';
    // נבנה "span" עם שם העובד, ואם יש התרעה (alerts.includes(e.id)) נוסיף אייקון
    const spanName= document.createElement('span');
    spanName.className='text-blue-700 underline cursor-pointer';
    spanName.textContent= e.name;
    spanName.onclick=()=> showEmployeeDetails(e.id);
    nameTd.appendChild(spanName);

    if(alerts.includes(e.id)) {
      // סימן קריאה
      const iconContainer= document.createElement('span');
      iconContainer.className='tooltip-container ml-2 text-red-500';
      iconContainer.innerHTML=`
        &#9888;
        <span class="tooltip-text">
          לעובד יש 2+ איחורים החודש
        </span>
      `;
      nameTd.appendChild(iconContainer);
    }
    row.appendChild(nameTd);

    // אחוז משרה
    const percentTd= document.createElement('td');
    percentTd.className='px-4 py-2';
    percentTd.textContent= e.percentPosition+'%';
    row.appendChild(percentTd);

    // איחורים (קליק => מעבר לחוצץ איחורים)
    const tardyTd= document.createElement('td');
    tardyTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    tardyTd.textContent= tardies;
    tardyTd.onclick=()=>{
      currentTab='tardiness';
      render();
    };
    row.appendChild(tardyTd);

    // חיסורים (קליק => מעבר לחוצץ נוכחות)
    const absTd= document.createElement('td');
    absTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    absTd.textContent= absences;
    absTd.onclick=()=>{
      currentTab='attendance';
      render();
    };
    row.appendChild(absTd);

    // אחוז עמידה במשרה
    const posTd= document.createElement('td');
    posTd.className='px-4 py-2';
    posTd.textContent= posRate+'%';
    row.appendChild(posTd);

    // פגישות (קליק => מעבר לחוצץ פגישות)
    const meetTd= document.createElement('td');
    meetTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    meetTd.textContent= meets;
    meetTd.onclick=()=>{
      currentTab='meetings';
      render();
    };
    row.appendChild(meetTd);

    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  card.appendChild(tableContainer);

  // כפתורי יצוא/שליחה
  const btnContainer= document.createElement('div');
  btnContainer.className='flex gap-2 mt-4';

  const exportBtn= document.createElement('button');
  exportBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  exportBtn.textContent='יצוא לאקסל';
  exportBtn.onclick=()=> exportToExcel('Dashboard');
  btnContainer.appendChild(exportBtn);

  const mailBtn= document.createElement('button');
  mailBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  mailBtn.textContent='שלח במייל';
  mailBtn.onclick=()=> sendDataByEmail('Dashboard');
  btnContainer.appendChild(mailBtn);

  card.appendChild(btnContainer);

  container.appendChild(card);
}

/***************************************************
 *          Employees
 ***************************************************/
// כעת נוסיף אפשרות עריכת סטטוס (פעיל/לא פעיל) בתוך הטבלה ישירות.
// אם העובד לא פעיל - מציגים שדה endDate
function renderEmployees(container){
  const card= document.createElement('div');
  card.className='p-4 shadow-lg rounded-2xl bg-white mb-4';

  const title= document.createElement('h2');
  title.className='text-2xl mb-2 font-bold text-gray-700';
  title.textContent='עובדים';
  card.appendChild(title);

  // Form להוספת עובד חדש
  const formDiv= document.createElement('div');
  formDiv.className='grid gap-2 mb-4 md:grid-cols-2 lg:grid-cols-3';

  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='שם העובד';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value= newEmployeeName;
    inp.oninput=(e)=>{ newEmployeeName=e.target.value; };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='תפקיד העובד';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value= newEmployeeRole;
    inp.oninput=(e)=>{ newEmployeeRole=e.target.value; };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='תאריך תחילת עבודה';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='date';
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value=newEmployeeStartDate;
    inp.onchange=(e)=>{ newEmployeeStartDate=e.target.value; };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='אחוז משרה';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='number';
    inp.min='0';
    inp.max='100';
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value=newEmployeePercentPosition;
    inp.onchange=(e)=>{
      newEmployeePercentPosition=parseInt(e.target.value,10);
      if(isNaN(newEmployeePercentPosition)) newEmployeePercentPosition=100;
    };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='סטטוס עובד';
    formDiv.appendChild(lbl);

    const sel= document.createElement('select');
    sel.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    sel.innerHTML=`
      <option value="active">פעיל</option>
      <option value="inactive">לא פעיל</option>
    `;
    sel.value=newEmployeeStatus;
    sel.onchange=(e)=>{ newEmployeeStatus=e.target.value; };
    formDiv.appendChild(sel);
  }

  // כפתור הוספה
  {
    const btn= document.createElement('button');
    btn.className='mt-2 bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600';
    btn.textContent='הוסף עובד';
    btn.onclick=()=>{
      if(!newEmployeeName || !newEmployeeRole) return;
      const newId = employees.length? Math.max(...employees.map(e=>e.id))+1 :1;
      const emp = {
        id:newId,
        name:newEmployeeName.trim(),
        role:newEmployeeRole.trim(),
        startDate:newEmployeeStartDate || '',
        endDate:'',
        status:newEmployeeStatus,
        percentPosition:newEmployeePercentPosition
      };
      employees.push(emp);
      // איפוס
      newEmployeeName='';
      newEmployeeRole='';
      newEmployeeStartDate='';
      newEmployeeStatus='active';
      newEmployeePercentPosition=100;
      saveToLocalStorage();
      render();
    };
    formDiv.appendChild(btn);
  }

  card.appendChild(formDiv);

  // טבלת עובדים
  const hr= document.createElement('hr');
  hr.className='my-4';
  card.appendChild(hr);

  const tableContainer= document.createElement('div');
  tableContainer.className='overflow-x-auto';

  const table= document.createElement('table');
  table.className='min-w-full text-right border-collapse';

  const thead= document.createElement('thead');
  thead.innerHTML=`
    <tr class="border-b bg-gray-100">
      <th class="px-4 py-2">שם</th>
      <th class="px-4 py-2">תפקיד</th>
      <th class="px-4 py-2">תאריך תחילה</th>
      <th class="px-4 py-2">תאריך סיום</th>
      <th class="px-4 py-2">סטטוס</th>
      <th class="px-4 py-2">אחוז משרה</th>
      <th class="px-4 py-2">מחק</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody= document.createElement('tbody');
  employees.forEach(e=>{
    const row= document.createElement('tr');
    row.className='border-b hover:bg-gray-50';

    // שם
    const nameTd= document.createElement('td');
    nameTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    nameTd.textContent=e.name;
    nameTd.onclick=()=> showEmployeeDetails(e.id);
    row.appendChild(nameTd);

    // תפקיד
    const roleTd= document.createElement('td');
    roleTd.className='px-4 py-2';
    roleTd.textContent=e.role;
    row.appendChild(roleTd);

    // תאריך תחילה
    const startTd= document.createElement('td');
    startTd.className='px-4 py-2';
    startTd.textContent=e.startDate;
    row.appendChild(startTd);

    // תאריך סיום (אם status inactive)
    const endTd= document.createElement('td');
    endTd.className='px-4 py-2';

    // אם סטטוס לא פעיל, נציג input type=date במקום טקסט
    if(e.status==='inactive'){
      const endInput= document.createElement('input');
      endInput.type='date';
      endInput.value=e.endDate;
      endInput.className='border rounded px-2 py-1 focus:outline-none focus:ring';
      endInput.onchange=(ev)=>{
        e.endDate= ev.target.value;
        saveToLocalStorage();
      };
      endTd.appendChild(endInput);
    } else {
      endTd.textContent= e.endDate || '';
    }
    row.appendChild(endTd);

    // סטטוס (select inline)
    const statusTd= document.createElement('td');
    statusTd.className='px-4 py-2';
    const statusSelect= document.createElement('select');
    statusSelect.className='border rounded px-2 py-1 focus:outline-none focus:ring';

    const optActive= document.createElement('option');
    optActive.value='active'; 
    optActive.textContent='פעיל';
    statusSelect.appendChild(optActive);

    const optInactive= document.createElement('option');
    optInactive.value='inactive';
    optInactive.textContent='לא פעיל';
    statusSelect.appendChild(optInactive);

    statusSelect.value= e.status;
    statusSelect.onchange=(ev)=>{
      e.status= ev.target.value;
      // אם הפך ללא פעיל => נבקש תאריך סיום
      if(e.status==='inactive') {
        // נראה את ה-endDate input
        e.endDate= e.endDate||''; 
      } else {
        // חזר לפעיל, ננקה תאריך סיום
        e.endDate='';
      }
      saveToLocalStorage();
      render(); // rerender כדי להציג את השינוי
    };
    statusTd.appendChild(statusSelect);
    row.appendChild(statusTd);

    // אחוז משרה
    const percentTd= document.createElement('td');
    percentTd.className='px-4 py-2';
    percentTd.textContent= e.percentPosition+'%';
    row.appendChild(percentTd);

    // מחיקה
    const delTd= document.createElement('td');
    delTd.className='px-4 py-2';
    const delBtn= document.createElement('button');
    delBtn.className='border border-red-300 text-red-500 rounded px-2 py-1 text-sm hover:bg-red-50';
    delBtn.textContent='מחק';
    delBtn.onclick=()=>{
      if(!confirm('האם למחוק עובד זה?')) return;
      employees= employees.filter(emp=> emp.id!== e.id);
      saveToLocalStorage();
      render();
    };
    delTd.appendChild(delBtn);
    row.appendChild(delTd);

    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  card.appendChild(tableContainer);

  container.appendChild(card);
}

/***************************************************
 *          Meetings
 ***************************************************/
function renderMeetings(container){
  const card= document.createElement('div');
  card.className='p-4 shadow-lg rounded-2xl bg-white mb-4';

  const title= document.createElement('h2');
  title.className='text-2xl mb-2 font-bold text-gray-700';
  title.textContent='פגישות - חודשי';
  card.appendChild(title);

  // Form להוספת פגישה
  const formDiv= document.createElement('div');
  formDiv.className='grid gap-2 mb-4';

  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='כותרת פגישה';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value= newMeetingTitle;
    inp.oninput=(e)=>{ newMeetingTitle=e.target.value; };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='תאריך';
    formDiv.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='date';
    inp.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    inp.value= newMeetingDate;
    inp.onchange=(e)=>{ newMeetingDate=e.target.value; };
    formDiv.appendChild(inp);
  }
  {
    const lbl= document.createElement('label');
    lbl.className='font-semibold text-gray-600';
    lbl.textContent='בחר עובד';
    formDiv.appendChild(lbl);

    const sel= document.createElement('select');
    sel.className='border rounded px-2 py-1 focus:outline-none focus:ring';
    sel.innerHTML='<option value="">בחר עובד...</option>';
    employees.forEach(emp=>{
      const opt= document.createElement('option');
      opt.value= emp.id;
      opt.textContent= emp.name;
      sel.appendChild(opt);
    });
    sel.value= newMeetingEmployee;
    sel.onchange=(e)=>{ newMeetingEmployee=e.target.value; };
    formDiv.appendChild(sel);
  }

  const addBtn= document.createElement('button');
  addBtn.className='mt-2 bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600';
  addBtn.textContent='הוסף פגישה';
  addBtn.onclick=()=>{
    if(!newMeetingTitle || !newMeetingDate || !newMeetingEmployee) return;
    const newId= meetings.length? Math.max(...meetings.map(m=>m.id))+1 :1;
    meetings.push({
      id:newId,
      title:newMeetingTitle.trim(),
      date:newMeetingDate,
      employeeId:Number(newMeetingEmployee)
    });
    newMeetingTitle='';
    newMeetingDate=getToday();
    newMeetingEmployee='';
    saveToLocalStorage();
    render();
  };
  formDiv.appendChild(addBtn);

  card.appendChild(formDiv);

  const hr= document.createElement('hr');
  hr.className='my-4';
  card.appendChild(hr);

  // טבלת פגישות לחודש
  const table= document.createElement('table');
  table.className='min-w-full text-right border-collapse';
  const thead= document.createElement('thead');
  thead.innerHTML=`
    <tr class="border-b bg-gray-100">
      <th class="px-4 py-2">כותרת</th>
      <th class="px-4 py-2">תאריך</th>
      <th class="px-4 py-2">עובד</th>
      <th class="px-4 py-2">ערוך</th>
      <th class="px-4 py-2">מחק</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody= document.createElement('tbody');
  const sortedMeetings= [...meetings].sort((a,b)=>(a.date< b.date? -1:1));
  sortedMeetings.forEach(m=>{
    if(getYearMonth(m.date)!==globalFilterMonth) return;
    const emp= employees.find(e=> e.id===m.employeeId);

    const row= document.createElement('tr');
    row.className='border-b hover:bg-gray-50';

    // כותרת
    {
      const td= document.createElement('td');
      td.className='px-4 py-2';
      td.textContent=m.title;
      row.appendChild(td);
    }
    // תאריך
    {
      const td= document.createElement('td');
      td.className='px-4 py-2';
      td.textContent=m.date;
      row.appendChild(td);
    }
    // עובד
    {
      const td= document.createElement('td');
      td.className='px-4 py-2 text-blue-700 underline cursor-pointer';
      td.textContent= emp? emp.name:'לא קיים';
      if(emp) td.onclick=()=> showEmployeeDetails(emp.id);
      row.appendChild(td);
    }
    // ערוך
    {
      const td= document.createElement('td');
      td.className='px-4 py-2';
      const editBtn= document.createElement('button');
      editBtn.className='border border-gray-300 rounded px-2 py-1 text-sm hover:bg-gray-100';
      editBtn.textContent='ערוך';
      editBtn.onclick=()=>{
        const newTitle= prompt('ערוך כותרת פגישה:', m.title);
        if(newTitle===null) return;
        const newDate= prompt('ערוך תאריך (YYYY-MM-DD):', m.date);
        if(newDate===null) return;
        const newEmpId= prompt('ערוך מזהה עובד:', m.employeeId);
        if(newEmpId===null) return;
        m.title=newTitle.trim();
        m.date=newDate;
        m.employeeId=Number(newEmpId);
        saveToLocalStorage();
        render();
      };
      td.appendChild(editBtn);
      row.appendChild(td);
    }
    // מחק
    {
      const td= document.createElement('td');
      td.className='px-4 py-2';
      const delBtn= document.createElement('button');
      delBtn.className='border border-red-300 text-red-500 rounded px-2 py-1 text-sm hover:bg-red-50';
      delBtn.textContent='מחק';
      delBtn.onclick=()=>{
        if(!confirm('למחוק פגישה זו?')) return;
        meetings= meetings.filter(mm=> mm.id!== m.id);
        saveToLocalStorage();
        render();
      };
      td.appendChild(delBtn);
      row.appendChild(td);
    }

    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  card.appendChild(table);

  // כפתורי ייצוא/שליחה
  const btnContainer= document.createElement('div');
  btnContainer.className='flex gap-2 mt-4';

  const exportBtn= document.createElement('button');
  exportBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  exportBtn.textContent='יצוא לאקסל';
  exportBtn.onclick=()=> exportToExcel('Meetings');
  btnContainer.appendChild(exportBtn);

  const mailBtn= document.createElement('button');
  mailBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  mailBtn.textContent='שלח במייל';
  mailBtn.onclick=()=> sendDataByEmail('Meetings');
  btnContainer.appendChild(mailBtn);

  card.appendChild(btnContainer);

  container.appendChild(card);
}

/***************************************************
 *       Tardiness
 ***************************************************/
function toggleTardinessDetail(empId){
  tardinessExpanded[empId] = !tardinessExpanded[empId];
  render();
}

function TardinessDetail(empId){
  const detail= getTardinessDetailForMonth(empId);
  if(detail.length===0) {
    return '<div class="text-sm text-gray-500">אין איחורים עבור עובד זה בחודש הנוכחי.</div>';
  }
  let rows='';
  detail.forEach(item=>{
    rows+=`
      <tr class="border-b hover:bg-gray-50">
        <td class="px-2 py-1">${item.date}</td>
        <td class="px-2 py-1">${item.reasonObj.reason}</td>
        <td class="px-2 py-1">${item.reasonObj.shiftStart}</td>
        <td class="px-2 py-1">${item.reasonObj.arrivalTime}</td>
        <td class="px-2 py-1">${item.reasonObj.minutesLate}</td>
        <td class="px-2 py-1">
          <button class="border border-gray-300 px-2 py-1 rounded text-sm text-blue-600 hover:bg-gray-100"
            onclick="editTardiness(${empId}, '${item.date}', ${item.reasonObj.id})">ערוך</button>
          <button class="border border-red-300 ml-2 px-2 py-1 rounded text-sm text-red-500 hover:bg-red-50"
            onclick="deleteTardiness(${empId}, '${item.date}', ${item.reasonObj.id})">מחק</button>
        </td>
      </tr>
    `;
  });
  return `
    <div class="overflow-x-auto mt-2">
      <table class="min-w-full text-right text-sm border-collapse">
        <thead class="bg-gray-100">
          <tr class="border-b">
            <th class="px-2 py-1">תאריך</th>
            <th class="px-2 py-1">סיבה</th>
            <th class="px-2 py-1">תחילת משמרת</th>
            <th class="px-2 py-1">שעת הגעה</th>
            <th class="px-2 py-1">איחור (דק')</th>
            <th class="px-2 py-1">פעולות</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderTardiness(container){
  const card= document.createElement('div');
  card.className='p-4 shadow-lg rounded-2xl bg-white mb-4';

  const title= document.createElement('h2');
  title.className='text-2xl mb-2 font-bold text-gray-700';
  title.textContent='איחורים - חודשי';
  card.appendChild(title);

  // Form הוספת איחור
  const formDiv= document.createElement('div');
  formDiv.className='flex gap-4 items-end flex-wrap mb-4';

  // בחירת עובד
  {
    const div= document.createElement('div');
    div.className='w-40';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='בחר עובד:';
    div.appendChild(lbl);

    const sel= document.createElement('select');
    sel.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    sel.innerHTML='<option value="">בחר עובד...</option>';
    employees.forEach(e=>{
      const opt= document.createElement('option');
      opt.value=e.id;
      opt.textContent=e.name;
      sel.appendChild(opt);
    });
    sel.value=tardySelectedEmployee;
    sel.onchange=(ev)=>{ tardySelectedEmployee= ev.target.value; };
    div.appendChild(sel);
    formDiv.appendChild(div);
  }
  // תאריך
  {
    const div= document.createElement('div');
    div.className='w-40';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='תאריך איחור:';
    div.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='date';
    inp.value=tardyDate;
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.onchange=(e)=>{ tardyDate=e.target.value; };
    div.appendChild(inp);
    formDiv.appendChild(div);
  }
  // שעת תחילת משמרת
  {
    const div= document.createElement('div');
    div.className='w-32';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='תחילת משמרת:';
    div.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='time';
    inp.value=newTardyShiftStart;
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.onchange=(e)=>{ newTardyShiftStart=e.target.value; };
    div.appendChild(inp);
    formDiv.appendChild(div);
  }
  // שעת הגעה בפועל
  {
    const div= document.createElement('div');
    div.className='w-32';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='שעת הגעה:';
    div.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='time';
    inp.value=newTardyArrivalTime;
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.onchange=(e)=>{ newTardyArrivalTime=e.target.value; };
    div.appendChild(inp);
    formDiv.appendChild(div);
  }

  // סיבה
  {
    const div= document.createElement('div');
    div.className='w-60';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='סיבת איחור:';
    div.appendChild(lbl);

    const inp= document.createElement('input');
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.value=tardyReason;
    inp.oninput=(e)=>{ tardyReason=e.target.value; };
    div.appendChild(inp);
    formDiv.appendChild(div);
  }

  // כפתור הוספה
  {
    const btn= document.createElement('button');
    btn.className='h-9 bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600';
    btn.textContent='הוסף איחור';
    btn.onclick=()=>{
      if(!tardySelectedEmployee) return;
      addTardiness(
        Number(tardySelectedEmployee),
        tardyDate,
        tardyReason,
        newTardyShiftStart,
        newTardyArrivalTime
      );
      // reset
      tardyReason='';
      tardyDate=getToday();
      render();
    };
    formDiv.appendChild(btn);
  }

  card.appendChild(formDiv);
  container.appendChild(card);

  // טבלת סיכום
  const summaryCard= document.createElement('div');
  summaryCard.className='p-4 shadow-lg rounded-2xl bg-white';

  const summaryTitle= document.createElement('h2');
  summaryTitle.className='text-xl mb-2 font-bold text-gray-700';
  summaryTitle.textContent='סיכום איחורים לחודש הנוכחי';
  summaryCard.appendChild(summaryTitle);

  const table= document.createElement('table');
  table.className='min-w-full text-right mb-4 border-collapse';
  const thead= document.createElement('thead');
  thead.innerHTML=`
    <tr class="border-b bg-gray-100">
      <th class="px-4 py-2">עובד</th>
      <th class="px-4 py-2">כמות איחורים</th>
      <th class="px-4 py-2">+</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody= document.createElement('tbody');
  employees.forEach(e=>{
    // אפשר להסתיר לא פעילים, אבל נציג כאן לכולם
    const row= document.createElement('tr');
    row.className='border-b hover:bg-gray-50';

    const tardies= getTardinessCountForMonth(e.id);

    const nameTd= document.createElement('td');
    nameTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    nameTd.textContent= e.name + (e.status==='inactive'?' (לא פעיל)':'');
    nameTd.onclick=()=> showEmployeeDetails(e.id);
    row.appendChild(nameTd);

    const countTd= document.createElement('td');
    countTd.className='px-4 py-2';
    countTd.textContent= tardies;
    row.appendChild(countTd);

    const expandTd= document.createElement('td');
    expandTd.className='px-4 py-2';
    const expandBtn= document.createElement('button');
    expandBtn.className='border border-gray-300 px-2 py-1 rounded text-sm hover:bg-gray-100';
    expandBtn.textContent= tardinessExpanded[e.id]? '-' : '+';
    expandBtn.onclick=()=> toggleTardinessDetail(e.id);
    expandTd.appendChild(expandBtn);
    row.appendChild(expandTd);

    tbody.appendChild(row);

    if(tardinessExpanded[e.id]){
      const detailRow= document.createElement('tr');
      detailRow.innerHTML=`
        <td colspan="3" class="bg-gray-50 p-2">${ TardinessDetail(e.id) }</td>
      `;
      tbody.appendChild(detailRow);
    }
  });
  table.appendChild(tbody);
  summaryCard.appendChild(table);

  const btnContainer= document.createElement('div');
  btnContainer.className='flex gap-2';

  const exportBtn= document.createElement('button');
  exportBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  exportBtn.textContent='יצוא לאקסל';
  exportBtn.onclick=()=> exportToExcel('Tardiness');
  btnContainer.appendChild(exportBtn);

  const mailBtn= document.createElement('button');
  mailBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  mailBtn.textContent='שלח במייל';
  mailBtn.onclick=()=> sendDataByEmail('Tardiness');
  btnContainer.appendChild(mailBtn);

  summaryCard.appendChild(btnContainer);

  container.appendChild(summaryCard);
}

/***************************************************
 *        Monthly Attendance (Absences)
 ***************************************************/
function toggleAttendanceDetail(empId){
  attendanceExpanded[empId]= !attendanceExpanded[empId];
  render();
}
function AbsencesDetailForMonth(empId){
  const arr= getAbsencesDetailForMonth(empId);
  if(arr.length===0) {
    return '<div class="text-sm text-gray-500">אין חיסורים לעובד זה בחודש הנוכחי.</div>';
  }
  let rows='';
  arr.forEach(item=>{
    rows+=`
      <tr class="border-b hover:bg-gray-50">
        <td class="px-2 py-1">${item.date}</td>
        <td class="px-2 py-1">${item.reason}</td>
        <td class="px-2 py-1">
          <button class="border border-gray-300 px-2 py-1 rounded text-sm text-blue-600 hover:bg-gray-100"
            onclick="editAbsence(${empId}, '${item.date}')">ערוך</button>
          <button class="border border-red-300 ml-2 px-2 py-1 rounded text-sm text-red-500 hover:bg-red-50"
            onclick="deleteAbsence(${empId}, '${item.date}')">מחק</button>
        </td>
      </tr>
    `;
  });
  return `
    <div class="overflow-x-auto mt-2">
      <table class="min-w-full text-right text-sm border-collapse">
        <thead class="bg-gray-100">
          <tr class="border-b">
            <th class="px-2 py-1">תאריך</th>
            <th class="px-2 py-1">סיבה</th>
            <th class="px-2 py-1">פעולות</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderMonthlyAttendance(container){
  const card= document.createElement('div');
  card.className='p-4 shadow-lg rounded-2xl bg-white mb-4';

  const title= document.createElement('h2');
  title.className='text-2xl mb-2 font-bold text-gray-700';
  title.textContent='נוכחות חודשית (הוספת חיסור)';
  card.appendChild(title);

  // form להוספת חיסור
  const formDiv= document.createElement('div');
  formDiv.className='flex gap-4 items-end flex-wrap mb-4';

  {
    const d= document.createElement('div');
    d.className='w-40';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='בחר עובד:';
    d.appendChild(lbl);

    const sel= document.createElement('select');
    sel.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    sel.innerHTML='<option value="">בחר עובד...</option>';
    employees.forEach(e=>{
      const opt= document.createElement('option');
      opt.value=e.id;
      opt.textContent=e.name;
      sel.appendChild(opt);
    });
    sel.value= attendanceSelectedEmployee;
    sel.onchange=(e)=>{ attendanceSelectedEmployee=e.target.value; };
    d.appendChild(sel);
    formDiv.appendChild(d);
  }
  {
    const d= document.createElement('div');
    d.className='w-40';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='תאריך חיסור:';
    d.appendChild(lbl);

    const inp= document.createElement('input');
    inp.type='date';
    inp.value= attendanceDate;
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.onchange=(e)=>{ attendanceDate=e.target.value; };
    d.appendChild(inp);
    formDiv.appendChild(d);
  }
  {
    const d= document.createElement('div');
    d.className='w-60';
    const lbl= document.createElement('label');
    lbl.className='block font-semibold text-gray-600';
    lbl.textContent='סיבת חיסור:';
    d.appendChild(lbl);

    const inp= document.createElement('input');
    inp.className='border rounded px-2 py-1 w-full focus:outline-none focus:ring';
    inp.value= attendanceReason;
    inp.oninput=(e)=>{ attendanceReason=e.target.value; };
    d.appendChild(inp);
    formDiv.appendChild(d);
  }

  const addBtn= document.createElement('button');
  addBtn.className='h-9 bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600';
  addBtn.textContent='הוסף חיסור';
  addBtn.onclick=()=>{
    if(!attendanceSelectedEmployee) return;
    markAbsence(Number(attendanceSelectedEmployee), attendanceDate, attendanceReason);
    attendanceReason='';
    attendanceDate= getToday();
    render();
  };
  formDiv.appendChild(addBtn);

  card.appendChild(formDiv);
  container.appendChild(card);

  // טבלת סיכום
  const summaryCard= document.createElement('div');
  summaryCard.className='p-4 shadow-lg rounded-2xl bg-white';

  const summaryTitle= document.createElement('h2');
  summaryTitle.className='text-xl mb-2 font-bold text-gray-700';
  summaryTitle.textContent='סיכום חיסורים לחודש הנוכחי';
  summaryCard.appendChild(summaryTitle);

  const table= document.createElement('table');
  table.className='min-w-full text-right mb-4 border-collapse';
  const thead= document.createElement('thead');
  thead.innerHTML=`
    <tr class="border-b bg-gray-100">
      <th class="px-4 py-2">עובד</th>
      <th class="px-4 py-2">כמות חיסורים</th>
      <th class="px-4 py-2">+</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody= document.createElement('tbody');
  employees.forEach(e=>{
    // מציגים גם לא פעילים לפי דרישה...
    const row= document.createElement('tr');
    row.className='border-b hover:bg-gray-50';

    const count= getAbsencesCountForMonth(e.id);

    const nameTd= document.createElement('td');
    nameTd.className='px-4 py-2 text-blue-700 underline cursor-pointer';
    nameTd.textContent= e.name + (e.status==='inactive'?' (לא פעיל)':'');
    nameTd.onclick=()=> showEmployeeDetails(e.id);
    row.appendChild(nameTd);

    const countTd= document.createElement('td');
    countTd.className='px-4 py-2';
    countTd.textContent= count;
    row.appendChild(countTd);

    const expandTd= document.createElement('td');
    expandTd.className='px-4 py-2';
    const expandBtn= document.createElement('button');
    expandBtn.className='border border-gray-300 px-2 py-1 rounded text-sm hover:bg-gray-100';
    expandBtn.textContent= attendanceExpanded[e.id]? '-' : '+';
    expandBtn.onclick=()=> toggleAttendanceDetail(e.id);
    expandTd.appendChild(expandBtn);
    row.appendChild(expandTd);

    tbody.appendChild(row);

    if(attendanceExpanded[e.id]) {
      const detailRow= document.createElement('tr');
      detailRow.innerHTML=`
        <td colspan="3" class="bg-gray-50 p-2">
          ${ AbsencesDetailForMonth(e.id) }
        </td>
      `;
      tbody.appendChild(detailRow);
    }
  });
  table.appendChild(tbody);
  summaryCard.appendChild(table);

  // כפתורי יצוא/שליחה
  const btnContainer= document.createElement('div');
  btnContainer.className='flex gap-2';

  const exportBtn= document.createElement('button');
  exportBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  exportBtn.textContent='יצוא לאקסל';
  exportBtn.onclick=()=> exportToExcel('Attendance');
  btnContainer.appendChild(exportBtn);

  const mailBtn= document.createElement('button');
  mailBtn.className='border border-gray-300 rounded px-4 py-2 hover:bg-gray-100';
  mailBtn.textContent='שלח במייל';
  mailBtn.onclick=()=> sendDataByEmail('Attendance');
  btnContainer.appendChild(mailBtn);

  summaryCard.appendChild(btnContainer);

  container.appendChild(summaryCard);
}

/***************************************************
 *         Show Employee Details (Popup)
 ***************************************************/
function showEmployeeDetails(empId){
  const emp= employees.find(e=> e.id===empId);
  if(!emp) return;
  const tardies= getTardinessCountForMonth(empId);
  const absences= getAbsencesCountForMonth(empId);
  const meets= getMeetingsCountForMonth(empId);
  const posRate= calculatePositionRate(empId);

  alert(
    `פרטי עובד:\n`+
    `שם: ${emp.name}\n`+
    `תפקיד: ${emp.role}\n`+
    `סטטוס: ${emp.status}\n`+
    `אחוז משרה: ${emp.percentPosition}%\n`+
    `תאריך תחילה: ${emp.startDate}\n`+
    (emp.status==='inactive'? `תאריך סיום: ${emp.endDate}\n`: '')+
    `\n---\n`+
    `איחורים בחודש: ${tardies}\n`+
    `חיסורים בחודש: ${absences}\n`+
    `פגישות בחודש: ${meets}\n`+
    `אחוז עמידה במשרה: ${posRate}%`
  );
}

/***************************************************
 *      Start
 ***************************************************/
render();