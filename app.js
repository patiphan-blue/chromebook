const API_URL = "https://script.google.com/macros/s/AKfycbzZgri1saLnzCrQ2WPyCIeqWMA_r9L0h9Faodi8WTczJjqp4BwTVK9-AKnM0yMFvhrJ/exec";
const IMPORT_CHUNK_SIZE = 400;

const state = {
  admin: JSON.parse(localStorage.getItem('chromebook_admin') || 'null'),
  studentRows: [],
  borrowRows: [],
  availableDevices: [],
  availableDeviceReport: [],
  borrowRequests: [],
  assignClassStudents: [],
  assignStudentSearch: '',
  dashboardTables: { students: [], teachers: [], returned: [], available: [] },
  activeDashboardTable: 'students',
  dashboardSearch: '',
  dashboardClassFilter: '',
  dashboardPage: 1,
  dashboardPageSize: 50,
};

const STUDENT_ID_KEYS = ['เลขประจำตัวนักเรียน', 'รหัสนักเรียน', 'student_id', 'เลขประจำตัว'];
const CITIZEN_ID_KEYS = ['เลขบัตรประชาชนนักเรียน', 'เลขบัตรประชาชน', 'citizen_id', 'national_id', 'เลขประจำตัวประชาชน'];
const PHONE_KEYS = ['เบอร์โทรศัพท์', 'เบอร์โทรศัพท์มือถือ', 'phone', 'โทรศัพท์', 'เบอร์โทร', 'หมายเลขโทรศัพท์', 'โทรศัพท์มือถือ'];

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  setTodayDefaults();
  renderAuthState();
  loadPublicDashboard();
});

function bindEvents() {
  document.getElementById('publicRefreshBtn').addEventListener('click', loadPublicDashboard);
  document.getElementById('mobileRefreshBtn').addEventListener('click', loadPublicDashboard);
  document.getElementById('mobileLoginBtn').addEventListener('click', openLoginModal);
  document.getElementById('mobileLogoutBtn').addEventListener('click', logout);
  document.querySelectorAll('[data-scroll-target]').forEach((button) => {
    button.addEventListener('click', () => {
      if (button.classList.contains('side-link')) {
        document.querySelectorAll('.side-link').forEach((link) => link.classList.toggle('active', link === button));
      }
      if (button.dataset.scrollTarget === 'adminPanel' && !state.admin) {
        openLoginModal();
        return;
      }
      if (button.dataset.scrollTarget === 'publicRequestSection') {
        setBorrowRequestOpen(true);
      }
      document.getElementById(button.dataset.scrollTarget).scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
  document.getElementById('dashboardSearchInput').addEventListener('input', (event) => {
    state.dashboardSearch = event.target.value;
    state.dashboardPage = 1;
    renderActiveDashboardTable();
  });
  document.getElementById('dashboardClassFilter').addEventListener('change', (event) => {
    state.dashboardClassFilter = event.target.value;
    state.dashboardPage = 1;
    renderActiveDashboardTable();
  });
  document.getElementById('dashboardPageSize').addEventListener('change', (event) => {
    state.dashboardPageSize = Number(event.target.value || 50);
    state.dashboardPage = 1;
    renderActiveDashboardTable();
  });
  document.getElementById('dashboardPrevPage').addEventListener('click', () => {
    if (state.dashboardPage > 1) {
      state.dashboardPage--;
      renderActiveDashboardTable();
    }
  });
  document.getElementById('dashboardNextPage').addEventListener('click', () => {
    state.dashboardPage++;
    renderActiveDashboardTable();
  });
  document.querySelectorAll('.table-tab').forEach((button) => {
    button.addEventListener('click', () => {
      state.activeDashboardTable = button.dataset.table;
      state.dashboardPage = 1;
      document.querySelectorAll('.table-tab').forEach((tab) => tab.classList.toggle('active', tab === button));
      renderActiveDashboardTable();
    });
  });
  document.getElementById('openLoginBtn').addEventListener('click', openLoginModal);
  document.getElementById('closeLoginBtn').addEventListener('click', closeLoginModal);
  document.getElementById('loginModal').addEventListener('click', (event) => {
    if (event.target.id === 'loginModal') closeLoginModal();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !document.getElementById('loginModal').classList.contains('hidden')) {
      closeLoginModal();
    }
  });
  document.getElementById('loginForm').addEventListener('submit', onLogin);
  document.getElementById('logoutBtn').addEventListener('click', logout);
  document.getElementById('assignBorrowerType').addEventListener('change', updateAssignBorrowerFields);
  document.getElementById('assignForm').addEventListener('submit', onAssignDevice);
  document.getElementById('assignStudentSearchInput').addEventListener('input', (event) => {
    state.assignStudentSearch = event.target.value;
    renderAssignClassTable();
  });
  document.getElementById('loadClassStudentsBtn').addEventListener('click', loadAssignClassStudents);
  document.getElementById('bulkAssignBtn').addEventListener('click', onBulkAssignDevices);
  document.getElementById('loadBorrowersBtn').addEventListener('click', loadBorrowers);
  document.getElementById('bulkReturnBtn').addEventListener('click', onBulkReturn);
  document.getElementById('selectAllReturn').addEventListener('change', toggleAllReturns);
  document.getElementById('studentMasterFile').addEventListener('change', (event) => onExcelSelected(event, 'students'));
  document.getElementById('borrowHistoryFile').addEventListener('change', (event) => onExcelSelected(event, 'borrow'));
  document.getElementById('importStudentsBtn').addEventListener('click', onImportStudents);
  document.getElementById('importBorrowBtn').addEventListener('click', onImportBorrowHistory);
  document.getElementById('toggleBorrowRequestBtn').addEventListener('click', () => {
    const form = document.getElementById('borrowRequestForm');
    setBorrowRequestOpen(form.classList.contains('hidden'));
  });
  document.getElementById('borrowRequestForm').addEventListener('submit', onCreateBorrowRequest);
  document.getElementById('loadBorrowRequestsBtn').addEventListener('click', loadBorrowRequests);
  document.getElementById('borrowRequestStatusFilter').addEventListener('change', loadBorrowRequests);
  document.getElementById('borrowRequestSearch').addEventListener('input', debounce(loadBorrowRequests, 350));
  document.getElementById('loadDeviceReportBtn').addEventListener('click', loadDeviceReport);
  document.getElementById('printDeviceReportBtn').addEventListener('click', printDeviceReport);
  document.getElementById('downloadDeviceReportBtn').addEventListener('click', downloadDeviceReportCsv);
  document.querySelectorAll('.tab-btn').forEach((button) => {
    button.addEventListener('click', () => showAdminPage(button.dataset.page));
  });
  document.querySelectorAll('.assign-mode-btn').forEach((button) => {
    button.addEventListener('click', () => showAssignMode(button.dataset.assignMode));
  });
}

function setBorrowRequestOpen(isOpen) {
  const form = document.getElementById('borrowRequestForm');
  const button = document.getElementById('toggleBorrowRequestBtn');
  if (!form || !button) return;
  form.classList.toggle('hidden', !isOpen);
  form.setAttribute('aria-hidden', String(!isOpen));
  button.setAttribute('aria-expanded', String(isOpen));
  button.textContent = isOpen ? 'ซ่อนแบบฟอร์ม' : 'เริ่มกรอกคำขอ';
  if (isOpen) {
    window.setTimeout(() => {
      const firstInput = document.getElementById('requestCitizenId');
      if (firstInput) firstInput.focus({ preventScroll: true });
    }, 180);
  }
}

function setTodayDefaults() {
  const today = new Date().toISOString().slice(0, 10);
  document.getElementById('assignBorrowDate').value = today;
  document.getElementById('returnDate').value = today;
  updateAssignBorrowerFields();
}

function renderAuthState() {
  const loggedIn = Boolean(state.admin);
  document.getElementById('adminPanel').classList.toggle('hidden', !loggedIn);
  document.getElementById('openLoginBtn').classList.toggle('hidden', loggedIn);
  document.getElementById('logoutBtn').classList.toggle('hidden', !loggedIn);
  document.getElementById('mobileLoginBtn').classList.toggle('hidden', loggedIn);
  document.getElementById('mobileLogoutBtn').classList.toggle('hidden', !loggedIn);

  if (loggedIn) {
    document.getElementById('adminName').textContent = `เข้าสู่ระบบโดย ${state.admin.full_name || state.admin.username}`;
    showAdminPage('assign');
    loadClasses();
    loadAvailableDeviceOptions();
  }
}

function openLoginModal() {
  document.getElementById('loginModal').classList.remove('hidden');
  document.getElementById('loginModal').classList.add('flex');
  document.getElementById('username').focus();
}

function closeLoginModal() {
  document.getElementById('loginModal').classList.add('hidden');
  document.getElementById('loginModal').classList.remove('flex');
}

async function onLogin(event) {
  event.preventDefault();
  const button = event.submitter || document.querySelector('#loginForm button[type="submit"]');
  setButtonBusy(button, true, 'กำลังเข้าสู่ระบบ...');
  try {
    const res = await api('login', {
      username: document.getElementById('username').value,
      password: document.getElementById('password').value,
    });
    state.admin = res;
    localStorage.setItem('chromebook_admin', JSON.stringify(res));
    closeLoginModal();
    renderAuthState();
    document.getElementById('adminPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
    toast('เข้าสู่ระบบสำเร็จ');
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

function logout() {
  state.admin = null;
  localStorage.removeItem('chromebook_admin');
  renderAuthState();
  toast('ออกจากระบบแล้ว');
}

function showAdminPage(page) {
  document.querySelectorAll('.admin-page').forEach((el) => el.classList.add('hidden'));
  document.getElementById(`${page}Page`).classList.remove('hidden');
  document.querySelectorAll('.tab-btn').forEach((button) => {
    button.classList.toggle('active', button.dataset.page === page);
  });
  if (page === 'return') loadClasses();
  if (page === 'assign') {
    loadClasses();
    loadAvailableDeviceOptions();
  }
  if (page === 'requests') loadBorrowRequests();
  if (page === 'deviceReport') loadDeviceReport();
}

async function loadPublicDashboard() {
  renderDashboardLoading();
  try {
    const [dashboard, tables, availableDevices] = await Promise.all([
      api('dashboard'),
      api('dashboardTables'),
      api('listAvailableDevices'),
    ]);

    document.getElementById('totalDevices').textContent = dashboard.total_devices || 0;
    document.getElementById('availableDevices').textContent = dashboard.available || 0;
    document.getElementById('borrowedDevices').textContent = dashboard.borrowed || 0;
    document.getElementById('repairDevices').textContent = dashboard.repairing || 0;
    document.getElementById('totalStudents').textContent = dashboard.total_students || 0;

    state.dashboardTables = {
      students: tables.students || [],
      teachers: tables.teachers || [],
      returned: tables.returned || [],
      available: availableDevices || [],
    };
    renderCharts(dashboard);
    renderDashboardClassFilter();
    renderActiveDashboardTable();
  } catch (error) {
    toast(error.message, true);
  }
}

function renderActiveDashboardTable() {
  const type = state.activeDashboardTable;
  const filteredRows = filterDashboardRows(state.dashboardTables[type] || [], type);
  const total = filteredRows.length;
  const pageSize = state.dashboardPageSize || 50;
  const totalPages = Math.max(Math.ceil(total / pageSize), 1);
  state.dashboardPage = Math.min(Math.max(state.dashboardPage || 1, 1), totalPages);
  const startIndex = total ? (state.dashboardPage - 1) * pageSize : 0;
  const rows = filteredRows.slice(startIndex, startIndex + pageSize);
  const head = document.getElementById('dashboardTableHead');
  const tbody = document.getElementById('dashboardTableRows');
  const meta = document.getElementById('dashboardTableMeta');
  const pageInfo = document.getElementById('dashboardPageInfo');
  const prevButton = document.getElementById('dashboardPrevPage');
  const nextButton = document.getElementById('dashboardNextPage');
  const classFilter = document.getElementById('dashboardClassFilter');

  head.innerHTML = type === 'available'
    ? '<tr><th>เลขเครื่อง</th><th>เลขที่ทรัพย์สิน</th><th>สถานะ</th></tr>'
    : '<tr><th>รหัส</th><th>ชื่อ-สกุล</th><th>ชั้น+ห้อง</th><th>วันยืม</th><th>วันคืน</th><th>สถานะ</th><th>เลขเครื่อง</th></tr>';
  tbody.innerHTML = '';
  meta.textContent = `แสดง ${rows.length} จาก ${total} รายการ`;
  if (pageInfo) {
    const from = total ? startIndex + 1 : 0;
    const to = total ? startIndex + rows.length : 0;
    pageInfo.textContent = `${from}-${to} จาก ${total}`;
  }
  if (prevButton) prevButton.disabled = state.dashboardPage <= 1;
  if (nextButton) nextButton.disabled = state.dashboardPage >= totalPages;
  if (classFilter) {
    classFilter.classList.toggle('hidden', type === 'teachers' || type === 'available');
  }

  rows.forEach((row) => {
    if (type === 'available') {
      tbody.insertAdjacentHTML('beforeend', `
        <tr class="row-ready">
          <td>${escapeHtml(row.device_key || '-')}</td>
          <td>${escapeHtml(row.asset_no || '-')}</td>
          <td>${statusBadge('ว่าง')}</td>
        </tr>
      `);
      return;
    }

    tbody.insertAdjacentHTML('beforeend', `
      <tr class="${dashboardRowClass(type, row)}">
        <td>${escapeHtml(row.borrower_id || '-')}</td>
        <td>${escapeHtml(row.full_name || '-')}</td>
        <td>${escapeHtml(type === 'teachers' ? 'ครู' : row.grade_level || '-')}</td>
        <td>${escapeHtml(row.borrow_date || '-')}</td>
        <td>${escapeHtml(row.return_date || '-')}</td>
        <td>${statusBadge(row.status)}</td>
        <td>${escapeHtml(row.device_key || '-')}</td>
      </tr>
    `);
  });

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="${type === 'available' ? 3 : 7}" class="text-center text-slate-500">ไม่พบข้อมูล</td></tr>`;
  }
}

function renderDashboardClassFilter() {
  const select = document.getElementById('dashboardClassFilter');
  if (!select) return;
  const current = select.value;
  const classes = [
    ...state.dashboardTables.students.map((row) => row.grade_level),
    ...state.dashboardTables.returned.map((row) => row.borrower_type === 'student' ? row.grade_level : ''),
  ]
    .filter(Boolean)
    .filter((value, index, arr) => arr.indexOf(value) === index)
    .sort((a, b) => String(a).localeCompare(String(b), 'th', { numeric: true, sensitivity: 'base' }));

  select.innerHTML = '<option value="">ทุกห้อง</option>';
  classes.forEach((className) => {
    select.insertAdjacentHTML('beforeend', `<option value="${escapeAttr(className)}">${escapeHtml(className)}</option>`);
  });
  if (current && classes.includes(current)) {
    select.value = current;
  } else {
    select.value = '';
    state.dashboardClassFilter = '';
  }
}

function dashboardRowClass(type, row) {
  if (type === 'returned' || row.status === 'คืนแล้ว') return 'row-returned';
  if (row.status === 'ส่งซ่อม') return 'row-repair';
  return 'row-borrowing';
}

function filterDashboardRows(rows, type) {
  const q = normalizeSearch(state.dashboardSearch);
  const classFilter = state.dashboardClassFilter;
  return rows.filter((row) => {
    if (classFilter && type !== 'teachers' && type !== 'available' && row.grade_level !== classFilter) return false;
    if (!q) return true;
    const haystack = type === 'available'
      ? [row.device_key, row.asset_no, 'ว่าง']
      : [row.borrower_id, row.full_name, row.grade_level, row.device_key, row.status];
    return normalizeSearch(haystack.join(' ')).includes(q);
  });
}

function renderCharts(dashboard) {
  const available = Number(dashboard.available || 0);
  const borrowed = Number(dashboard.borrowed || 0);
  const repair = Number(dashboard.repairing || 0);
  const total = Math.max(available + borrowed + repair, 1);
  const availableDeg = (available / total) * 360;
  const borrowedDeg = availableDeg + (borrowed / total) * 360;
  document.getElementById('statusDonut').style.background = `conic-gradient(#15803d 0deg ${availableDeg}deg, #b45309 ${availableDeg}deg ${borrowedDeg}deg, #b91c1c ${borrowedDeg}deg 360deg)`;
  document.getElementById('legendAvailable').textContent = available;
  document.getElementById('legendBorrowed').textContent = borrowed;
  document.getElementById('legendRepair').textContent = repair;

  const studentCount = state.dashboardTables.students.length;
  const teacherCount = state.dashboardTables.teachers.length;
  const activeTotal = Math.max(studentCount + teacherCount, 1);
  document.getElementById('studentBar').style.width = `${(studentCount / activeTotal) * 100}%`;
  document.getElementById('teacherBar').style.width = `${(teacherCount / activeTotal) * 100}%`;
  document.getElementById('studentBarText').textContent = studentCount;
  document.getElementById('teacherBarText').textContent = teacherCount;
}

function normalizeSearch(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

async function onAssignDevice(event) {
  event.preventDefault();
  try {
    const borrowerType = document.getElementById('assignBorrowerType').value;
    const res = await api('assignDevice', {
      borrower_type: borrowerType,
      student_id: document.getElementById('assignStudentId').value,
      teacher_name: document.getElementById('assignTeacherName').value,
      device_key: document.getElementById('assignDeviceKey').value,
      borrow_date: document.getElementById('assignBorrowDate').value,
      note: document.getElementById('assignNote').value,
    });
    toast(res.message || 'บันทึกสำเร็จ');
    event.target.reset();
    setTodayDefaults();
    loadAvailableDeviceOptions();
    loadPublicDashboard();
    loadClasses();
  } catch (error) {
    toast(error.message, true);
  }
}

async function loadAvailableDeviceOptions() {
  renderAvailableDeviceLoading();
  try {
    const devices = await api('listAvailableDevices');
    state.availableDevices = devices;
    const select = document.getElementById('assignDeviceKey');
    const current = select.value;
    select.innerHTML = '<option value="">เลือกเลขเครื่องที่ว่าง</option>';

    devices.forEach((device) => {
      const label = device.asset_no
        ? `${device.device_key} - ${device.asset_no}`
        : device.device_key;
      select.insertAdjacentHTML('beforeend', `<option value="${escapeAttr(device.device_key)}">${escapeHtml(label)}</option>`);
    });

    if (current && devices.some((device) => String(device.device_key) === String(current))) {
      select.value = current;
    }

    if (!devices.length) {
      select.innerHTML = '<option value="">ไม่มีเครื่องว่าง</option>';
    }
    renderAvailableDevicePanel();
  } catch (error) {
    toast(error.message, true);
  }
}

function showAssignMode(mode) {
  document.querySelectorAll('.assign-mode-btn').forEach((button) => {
    button.classList.toggle('active', button.dataset.assignMode === mode);
  });
  document.getElementById('assignClassMode').classList.toggle('hidden', mode !== 'class');
  document.getElementById('assignForm').classList.toggle('hidden', mode !== 'single');
  if (mode === 'class') {
    loadClasses();
    loadAvailableDeviceOptions();
  }
}

function updateAssignBorrowerFields() {
  const type = document.getElementById('assignBorrowerType').value;
  const isTeacher = type === 'teacher';
  const studentInput = document.getElementById('assignStudentId');
  document.getElementById('assignStudentLabel').textContent = isTeacher ? 'รหัสครู (ใส่ 0 ได้)' : 'รหัสนักเรียน';
  studentInput.required = !isTeacher;
  if (isTeacher) studentInput.value = '0';
  document.getElementById('assignTeacherNameWrap').classList.toggle('hidden', !isTeacher);
  document.getElementById('assignTeacherName').required = isTeacher;
}

async function loadClasses() {
  try {
    const classes = await api('listClasses');
    const select = document.getElementById('classSelect');
    const assignSelect = document.getElementById('assignClassSelect');
    const current = select.value;
    const assignCurrent = assignSelect.value;
    select.innerHTML = '<option value="">เลือกห้องเรียน / ครู</option>';
    assignSelect.innerHTML = '<option value="">เลือกห้องเรียน</option>';
    classes.forEach((className) => {
      select.insertAdjacentHTML('beforeend', `<option value="${escapeAttr(className)}">${escapeHtml(className)}</option>`);
      if (className !== 'ครู') {
        assignSelect.insertAdjacentHTML('beforeend', `<option value="${escapeAttr(className)}">${escapeHtml(className)}</option>`);
      }
    });
    if (current) select.value = current;
    if (assignCurrent) assignSelect.value = assignCurrent;
  } catch (error) {
    toast(error.message, true);
  }
}

async function loadAssignClassStudents() {
  const gradeLevel = document.getElementById('assignClassSelect').value;
  if (!gradeLevel) {
    toast('กรุณาเลือกห้องเรียน', true);
    return;
  }

  const button = document.getElementById('loadClassStudentsBtn');
  setButtonBusy(button, true, 'กำลังดึงรายชื่อ...');
  setTableLoading('assignClassTable', 5, 'กำลังโหลดรายชื่อนักเรียน...');
  document.getElementById('assignClassMeta').textContent = 'กำลังดึงรายชื่อจากฐานข้อมูล กรุณารอสักครู่...';
  try {
    await loadAvailableDeviceOptions();
    const students = await api('getStudentsByClass', { grade_level: gradeLevel });
    state.assignClassStudents = students;
    state.assignStudentSearch = '';
    document.getElementById('assignStudentSearchInput').value = '';
    renderAssignClassTable();
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

function renderAssignClassTable(students = state.assignClassStudents) {
  const tbody = document.getElementById('assignClassTable');
  const meta = document.getElementById('assignClassMeta');
  const selectedByStudent = {};
  document.querySelectorAll('.class-device-select').forEach((select) => {
    if (select.value) selectedByStudent[select.dataset.studentId] = select.value;
  });
  const rows = filterAssignStudents(students);
  const readyCount = students.filter((student) => !student.already_borrowing).length;
  const selectedCount = Object.keys(selectedByStudent).length;

  tbody.innerHTML = '';
  meta.textContent = students.length
    ? `พบ ${students.length} คน | พร้อมยืม ${readyCount} คน | เลือกเครื่องแล้ว ${selectedCount} คน | แสดง ${rows.length} คน`
    : 'เลือกห้องแล้วกดดึงรายชื่อ เพื่อเริ่มเลือกเลขเครื่องที่ยังว่าง';

  rows.forEach((student) => {
    tbody.insertAdjacentHTML('beforeend', `
      <tr class="${student.already_borrowing ? 'row-borrowing' : 'row-ready'}">
        <td>${escapeHtml(student.grade_level || '-')}</td>
        <td>${escapeHtml(student.student_id)}</td>
        <td>${escapeHtml(student.full_name)}</td>
        <td>${student.already_borrowing ? statusBadge('กำลังยืม') : statusBadge('พร้อมยืม')}</td>
        <td>
          <select class="input class-device-select" data-student-id="${escapeAttr(student.student_id)}" ${student.already_borrowing ? 'disabled' : ''}>
            <option value="">เลือกเครื่อง</option>
            ${state.availableDevices.map((device) => {
              const label = device.asset_no ? `${device.device_key} - ${device.asset_no}` : device.device_key;
              const selected = String(selectedByStudent[student.student_id] || '') === String(device.device_key) ? ' selected' : '';
              return `<option value="${escapeAttr(device.device_key)}"${selected}>${escapeHtml(label)}</option>`;
            }).join('')}
          </select>
        </td>
      </tr>
    `);
  });
  document.querySelectorAll('.class-device-select').forEach((select) => {
    select.addEventListener('change', () => renderAssignClassMeta());
  });
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-slate-500">ไม่พบนักเรียนตามเงื่อนไขนี้</td></tr>';
  }
}

function renderAssignClassMeta() {
  const meta = document.getElementById('assignClassMeta');
  const students = state.assignClassStudents;
  const rows = filterAssignStudents(students);
  const readyCount = students.filter((student) => !student.already_borrowing).length;
  const selectedCount = Array.from(document.querySelectorAll('.class-device-select')).filter((select) => select.value).length;
  meta.textContent = students.length
    ? `พบ ${students.length} คน | พร้อมยืม ${readyCount} คน | เลือกเครื่องแล้ว ${selectedCount} คน | แสดง ${rows.length} คน`
    : 'เลือกห้องแล้วกดดึงรายชื่อ เพื่อเริ่มเลือกเลขเครื่องที่ยังว่าง';
}

function filterAssignStudents(students) {
  const q = normalizeSearch(state.assignStudentSearch);
  if (!q) return students;
  return students.filter((student) => normalizeSearch([
    student.student_id,
    student.full_name,
    student.grade_level,
  ].join(' ')).includes(q));
}

function renderAvailableDevicePanel() {
  const list = document.getElementById('assignAvailableList');
  if (!list) return;
  const devices = state.availableDevices || [];
  if (!devices.length) {
    list.innerHTML = '<p class="empty-state">ยังไม่มีเครื่องว่างในระบบ</p>';
    return;
  }
  list.innerHTML = devices.slice(0, 80).map((device) => `
    <div class="device-chip" title="${escapeAttr([device.device_key, device.asset_no].filter(Boolean).join(' - '))}">
      <strong>${escapeHtml(device.device_key || '-')}</strong>
      <span>${escapeHtml(device.asset_no || 'ไม่มีเลขทรัพย์สิน')}</span>
    </div>
  `).join('');
}

async function onBulkAssignDevices() {
  const selects = Array.from(document.querySelectorAll('.class-device-select'));
  const assignments = selects
    .filter((select) => select.value)
    .map((select) => ({ student_id: select.dataset.studentId, device_key: select.value }));
  const duplicate = findDuplicate(assignments.map((item) => item.device_key));
  if (duplicate) {
    toast(`เลือกเครื่อง ${duplicate} ซ้ำ`, true);
    return;
  }
  if (!assignments.length) {
    toast('กรุณาเลือกเครื่องอย่างน้อย 1 คน', true);
    return;
  }

  const button = document.getElementById('bulkAssignBtn');
  setButtonBusy(button, true, 'กำลังบันทึก...');
  try {
    const res = await api('bulkAssignDevices', {
      assignments,
      borrow_date: document.getElementById('assignBorrowDate').value,
      note: 'ยืมรายห้อง',
    });
    toast(`${res.message}: สำเร็จ ${res.assigned_count}, ข้าม ${res.skipped_count}`);
    loadAssignClassStudents();
    loadPublicDashboard();
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

function findDuplicate(values) {
  const seen = new Set();
  for (const value of values) {
    if (seen.has(value)) return value;
    seen.add(value);
  }
  return '';
}

async function loadBorrowers() {
  const gradeLevel = document.getElementById('classSelect').value;
  if (!gradeLevel) {
    toast('กรุณาเลือกห้องเรียน / ครู', true);
    return;
  }

  const button = document.getElementById('loadBorrowersBtn');
  setButtonBusy(button, true, 'กำลังดึงรายชื่อ...');
  setTableLoading('borrowersTable', 6, 'กำลังโหลดรายการที่ยืมอยู่...');
  try {
    const rows = await api('getBorrowersByClass', { grade_level: gradeLevel });
    const tbody = document.getElementById('borrowersTable');
    tbody.innerHTML = '';
    rows.forEach((row) => {
      tbody.insertAdjacentHTML('beforeend', `
        <tr>
          <td><input class="return-check" type="checkbox" value="${escapeAttr(row.transaction_id)}"></td>
          <td>${escapeHtml(row.grade_level || '-')}</td>
          <td>${escapeHtml(row.student_id || '-')}</td>
          <td>${escapeHtml(row.full_name || '-')}</td>
          <td>${escapeHtml(row.device_key || '-')}</td>
          <td>${escapeHtml(row.borrow_date || '-')}</td>
        </tr>
      `);
    });

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="6" class="text-center text-slate-500">ไม่พบรายการกำลังยืม</td></tr>';
    }
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

function renderDashboardLoading() {
  const tbody = document.getElementById('dashboardTableRows');
  const meta = document.getElementById('dashboardTableMeta');
  if (meta) meta.textContent = 'กำลังโหลดข้อมูลล่าสุด...';
  if (tbody) setTableLoading('dashboardTableRows', state.activeDashboardTable === 'available' ? 3 : 7, 'กำลังโหลดตารางติดตาม...');
}

function renderAvailableDeviceLoading() {
  const list = document.getElementById('assignAvailableList');
  if (!list) return;
  list.innerHTML = `
    <div class="loading-card"></div>
    <div class="loading-card"></div>
    <div class="loading-card"></div>
  `;
}

function setTableLoading(tbodyId, colspan, message) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  tbody.innerHTML = `
    <tr class="loading-row">
      <td colspan="${colspan}">
        <span class="loading-dot" aria-hidden="true"></span>
        ${escapeHtml(message)}
      </td>
    </tr>
  `;
}

function toggleAllReturns(event) {
  document.querySelectorAll('.return-check').forEach((checkbox) => {
    checkbox.checked = event.target.checked;
  });
}

async function onBulkReturn() {
  const transactionIds = Array.from(document.querySelectorAll('.return-check:checked')).map((el) => el.value);
  if (!transactionIds.length) {
    toast('กรุณาเลือกรายการที่ต้องการคืน', true);
    return;
  }

  const button = document.getElementById('bulkReturnBtn');
  setButtonBusy(button, true, 'กำลังคืนเครื่อง...');
  try {
    const res = await api('bulkReturn', {
      transaction_ids: transactionIds,
      return_date: document.getElementById('returnDate').value,
    });
    toast(res.message || 'คืนเครื่องสำเร็จ');
    document.getElementById('selectAllReturn').checked = false;
    loadBorrowers();
    loadAvailableDeviceOptions();
    loadPublicDashboard();
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

async function onCreateBorrowRequest(event) {
  event.preventDefault();
  const button = event.submitter || document.querySelector('#borrowRequestForm button[type="submit"]');
  setButtonBusy(button, true, 'กำลังบันทึกคำขอ...');
  try {
    const res = await api('createBorrowRequest', {
      citizen_id: document.getElementById('requestCitizenId').value,
      student_id: document.getElementById('requestStudentId').value,
      full_name: document.getElementById('requestFullName').value,
      parent_name: document.getElementById('requestParentName').value,
      grade_level: document.getElementById('requestGradeLevel').value,
      phone: document.getElementById('requestPhone').value,
      house_no: document.getElementById('requestHouseNo').value,
      village_no: document.getElementById('requestVillageNo').value,
      subdistrict: document.getElementById('requestSubdistrict').value,
      district: document.getElementById('requestDistrict').value,
      province: document.getElementById('requestProvince').value,
      note: document.getElementById('requestNote').value,
    });
    toast(res.message || 'บันทึกคำขอสำเร็จ');
    event.target.reset();
    setBorrowRequestOpen(false);
    document.getElementById('publicRequestSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

async function loadBorrowRequests() {
  const tbody = document.getElementById('borrowRequestTable');
  const meta = document.getElementById('borrowRequestMeta');
  if (!tbody) return;
  setTableLoading('borrowRequestTable', 8, 'กำลังโหลดรายการคำขอยืม...');
  if (meta) meta.textContent = 'กำลังโหลดคำขอล่าสุด...';
  try {
    const rows = await api('listBorrowRequests', {
      status: document.getElementById('borrowRequestStatusFilter').value,
      query: document.getElementById('borrowRequestSearch').value,
      limit: 200,
    });
    state.borrowRequests = rows || [];
    renderBorrowRequests();
  } catch (error) {
    toast(error.message, true);
    tbody.innerHTML = '<tr><td colspan="9" class="text-center text-slate-500">โหลดรายการคำขอไม่สำเร็จ</td></tr>';
  }
}

function renderBorrowRequests() {
  const tbody = document.getElementById('borrowRequestTable');
  const meta = document.getElementById('borrowRequestMeta');
  const rows = state.borrowRequests || [];
  tbody.innerHTML = '';
  if (meta) meta.textContent = `แสดง ${rows.length} รายการล่าสุด`;

  rows.forEach((row) => {
    tbody.insertAdjacentHTML('beforeend', `
      <tr>
        <td>${requestStatusBadge(row.request_status)}</td>
        <td>${escapeHtml(row.student_id || '-')}</td>
        <td>${escapeHtml(row.full_name || '-')}</td>
        <td>${escapeHtml(row.parent_name || '-')}</td>
        <td>${escapeHtml(row.grade_level || '-')}</td>
        <td>${escapeHtml(row.phone || '-')}</td>
        <td class="request-address">${escapeHtml(row.address || '-')}</td>
        <td>${escapeHtml(row.created_at || '-')}</td>
        <td>
          <select class="input request-status-select" data-request-id="${escapeAttr(row.request_id)}" data-previous-value="${escapeAttr(row.request_status || '')}">
            ${['รอตรวจสอบ', 'อนุมัติแล้ว', 'ปฏิเสธ', 'ยกเลิก'].map((status) => `
              <option value="${escapeAttr(status)}"${String(row.request_status || '') === status ? ' selected' : ''}>${escapeHtml(status)}</option>
            `).join('')}
          </select>
        </td>
      </tr>
    `);
  });

  document.querySelectorAll('.request-status-select').forEach((select) => {
    select.addEventListener('change', () => updateBorrowRequestStatus(select));
  });

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="9" class="text-center text-slate-500">ยังไม่มีรายการคำขอตามเงื่อนไขนี้</td></tr>';
  }
}

async function updateBorrowRequestStatus(select) {
  const previousText = select.dataset.previousValue || '';
  select.disabled = true;
  try {
    const res = await api('updateBorrowRequestStatus', {
      request_id: select.dataset.requestId,
      request_status: select.value,
    });
    toast(res.message || 'อัปเดตสถานะคำขอสำเร็จ');
    select.dataset.previousValue = select.value;
    loadBorrowRequests();
  } catch (error) {
    if (previousText) select.value = previousText;
    toast(error.message, true);
  } finally {
    select.disabled = false;
  }
}

async function loadDeviceReport() {
  const button = document.getElementById('loadDeviceReportBtn');
  setButtonBusy(button, true, 'กำลังโหลด...');
  setTableLoading('deviceReportTable', 6, 'กำลังโหลดเครื่องที่ยังว่าง...');
  try {
    const report = await api('listAvailableDeviceReport');
    state.availableDeviceReport = report.devices || [];
    renderDeviceReport(report);
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

function renderDeviceReport(report = {}) {
  const rows = report.devices || state.availableDeviceReport || [];
  const tbody = document.getElementById('deviceReportTable');
  const meta = document.getElementById('deviceReportMeta');
  const generated = document.getElementById('deviceReportGeneratedAt');
  tbody.innerHTML = '';
  if (meta) meta.textContent = `พบเครื่องว่าง ${rows.length} เครื่อง`;
  if (generated) generated.textContent = `อัปเดต ${report.generated_at || '-'}`;

  rows.forEach((row, index) => {
    tbody.insertAdjacentHTML('beforeend', `
      <tr class="row-ready">
        <td>${index + 1}</td>
        <td>${escapeHtml(row.device_key || '-')}</td>
        <td>${escapeHtml(row.asset_no || '-')}</td>
        <td>${statusBadge(row.device_status || 'ว่าง')}</td>
        <td></td>
        <td></td>
      </tr>
    `);
  });

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-slate-500">ไม่มีเครื่องว่างในระบบ</td></tr>';
  }
}

function printDeviceReport() {
  if (!state.availableDeviceReport.length) {
    toast('กรุณาโหลดเครื่องว่างก่อนพิมพ์', true);
    return;
  }
  window.print();
}

function downloadDeviceReportCsv() {
  const rows = state.availableDeviceReport || [];
  if (!rows.length) {
    toast('กรุณาโหลดเครื่องว่างก่อนดาวน์โหลด', true);
    return;
  }
  const csvRows = [
    ['ลำดับ', 'เลขเครื่อง', 'เลขที่ทรัพย์สิน', 'สถานะ', 'จัดให้', 'หมายเหตุ'],
    ...rows.map((row, index) => [
      index + 1,
      row.device_key || '',
      row.asset_no || '',
      row.device_status || 'ว่าง',
      '',
      '',
    ]),
  ];
  const csv = csvRows.map((row) => row.map(csvCell).join(',')).join('\r\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `available-chromebooks-${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvCell(value) {
  return `"${String(value ?? '').replaceAll('"', '""')}"`;
}

async function onExcelSelected(event, type) {
  const file = event.target.files[0];
  if (!file) return;

  const rows = type === 'students' ? await readStudentWorkbook(file) : await readFirstSheet(file);
  if (type === 'students') {
    state.studentRows = rows;
    document.getElementById('studentCount').textContent = `${rows.length} แถว`;
    renderPreview(rows.slice(0, 10), 'studentPreviewHead', 'studentPreviewBody');
  } else {
    state.borrowRows = rows;
    document.getElementById('borrowCount').textContent = `${rows.length} แถว`;
    renderPreview(rows.slice(0, 10), 'borrowPreviewHead', 'borrowPreviewBody');
  }

  toast(`อ่านไฟล์ได้ ${rows.length} แถว`);
}

async function readFirstSheet(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  return sheetToObjects(workbook.Sheets[workbook.SheetNames[0]]);
}

async function readStudentWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheetNames = workbook.SheetNames;
  const mainRows = sheetToStudentRows(workbook.Sheets[sheetNames[0]]);
  const phoneRows = sheetNames.slice(1).flatMap((sheetName) => sheetToObjects(workbook.Sheets[sheetName]));
  return mergePhoneRows(mainRows, phoneRows);
}

function sheetToObjects(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const headerIndex = detectHeaderRow(matrix);
  if (headerIndex < 0) return [];

  const headers = matrix[headerIndex].map((value) => String(value || '').trim());
  return matrix.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell || '').trim() !== ''))
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        if (!header) return;
        const key = obj[header] === undefined ? header : `${header}_${index}`;
        obj[key] = row[index] ?? '';
      });
      return obj;
    });
}

function sheetToStudentRows(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const headerIndex = detectStudentHeaderRow(matrix);
  if (headerIndex < 0) return [];

  const headers = matrix[headerIndex].map((value) => String(value || '').trim());
  const normalizedHeaders = headers.map(normalizeHeader);
  const citizenIndex = findCitizenIdIndex(normalizedHeaders);
  const studentIdIndex = findStudentIdIndex(normalizedHeaders);
  const gradeLevelIndex = findHeaderIndex(normalizedHeaders, ['ชั้น+ห้อง', 'ระดับชั้น+ห้อง', 'ชั้น/ห้อง', 'grade_level']);
  const gradeIndex = findHeaderIndex(normalizedHeaders, ['ชั้น', 'ระดับชั้น']);
  const roomIndex = findHeaderIndex(normalizedHeaders, ['ห้อง']);
  const fullNameIndex = findHeaderIndex(normalizedHeaders, ['คำนำหน้า+ชื่อ+สกุล', 'ชื่อสกุล', 'ชื่อ-นามสกุล', 'ชื่อ นามสกุล', 'full_name']);
  const prefixIndex = findHeaderIndex(normalizedHeaders, ['คำนำหน้าชื่อ', 'คำนำหน้า']);
  const firstNameIndex = findHeaderIndex(normalizedHeaders, ['ชื่อ']);
  const lastNameIndex = findHeaderIndex(normalizedHeaders, ['นามสกุล', 'สกุล']);
  const phoneIndex = findHeaderIndex(normalizedHeaders, PHONE_KEYS);

  return matrix.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell || '').trim() !== ''))
    .map((row) => {
      const grade = getCell(row, gradeIndex);
      const room = getCell(row, roomIndex);
      const prefix = getCell(row, prefixIndex);
      const firstName = getCell(row, firstNameIndex);
      const lastName = getCell(row, lastNameIndex);
      return {
        citizen_id: getCell(row, citizenIndex),
        student_id: getCell(row, studentIdIndex),
        grade_level: getCell(row, gradeLevelIndex) || (grade && room ? `${grade}/${room}` : grade || room),
        prefix,
        first_name: firstName,
        last_name: lastName,
        full_name: getCell(row, fullNameIndex) || [prefix, firstName, lastName].filter(Boolean).join(' '),
        birth_date: getCell(row, findHeaderIndex(normalizedHeaders, ['วันเกิด', 'วัน/เดือน/ปีเกิด'])),
        house_no: getCell(row, findHeaderIndex(normalizedHeaders, ['บ้านเลขที่', 'เลขที่บ้าน'])),
        village_no: getCell(row, findHeaderIndex(normalizedHeaders, ['หมู่', 'หมู่บ้าน'])),
        subdistrict: getCell(row, findHeaderIndex(normalizedHeaders, ['ตำบล'])),
        district: getCell(row, findHeaderIndex(normalizedHeaders, ['อำเภอ'])),
        province: getCell(row, findHeaderIndex(normalizedHeaders, ['จังหวัด'])),
        phone: getCell(row, phoneIndex),
      };
    });
}

function detectHeaderRow(matrix) {
  return matrix.findIndex((row) => {
    const text = row.map(normalizeHeader).join('|');
    return text.includes('เลขประจำตัวนักเรียน') || text.includes('รหัสนักเรียน') || text.includes('เลขเครื่องนิยม');
  });
}

function detectStudentHeaderRow(matrix) {
  return matrix.findIndex((row) => {
    const text = row.map(normalizeHeader).join('|');
    const hasStudentId = text.includes('เลขประจำตัวนักเรียน') || text.includes('รหัสนักเรียน');
    const hasName = text.includes('คำนำหน้า+ชื่อ+สกุล') || text.includes('ชื่อสกุล') || text.includes('ชื่อ-นามสกุล') || (text.includes('ชื่อ') && text.includes('นามสกุล'));
    return hasStudentId && hasName;
  });
}

function findCitizenIdIndex(headers) {
  const citizen = findHeaderIndex(headers, ['เลขบัตรประชาชนนักเรียน', 'เลขบัตรประชาชน', 'citizen_id', 'national_id', 'เลขประจำตัวประชาชน']);
  if (citizen >= 0) return citizen;
  const duplicateStudentIds = findAllHeaderIndexes(headers, 'เลขประจำตัวนักเรียน');
  return duplicateStudentIds[0] ?? -1;
}

function findStudentIdIndex(headers) {
  const direct = findHeaderIndex(headers, ['รหัสนักเรียน', 'student_id']);
  if (direct >= 0) return direct;
  const duplicateStudentIds = findAllHeaderIndexes(headers, 'เลขประจำตัวนักเรียน');
  return duplicateStudentIds.length > 1 ? duplicateStudentIds[1] : duplicateStudentIds[0] ?? -1;
}

function findHeaderIndex(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  return headers.findIndex((header) => normalizedCandidates.includes(header));
}

function findAllHeaderIndexes(headers, candidate) {
  const normalized = normalizeHeader(candidate);
  return headers.reduce((indexes, header, index) => {
    if (header === normalized) indexes.push(index);
    return indexes;
  }, []);
}

function getCell(row, index) {
  return index >= 0 ? String(row[index] ?? '').trim() : '';
}

function mergePhoneRows(mainRows, phoneRows) {
  if (!phoneRows.length) return mainRows;
  const phoneByStudentId = new Map();
  const phoneByCitizenId = new Map();

  phoneRows.forEach((row) => {
    const phone = pickValue(row, PHONE_KEYS);
    if (!phone) return;
    const studentId = normalizeKey(pickValue(row, STUDENT_ID_KEYS));
    const citizenId = normalizeKey(pickValue(row, CITIZEN_ID_KEYS));
    if (studentId) phoneByStudentId.set(studentId, phone);
    if (citizenId) phoneByCitizenId.set(citizenId, phone);
  });

  return mainRows.map((row) => {
    if (row.phone) return row;
    const studentId = normalizeKey(row.student_id);
    const citizenId = normalizeKey(row.citizen_id);
    return Object.assign({}, row, {
      phone: phoneByStudentId.get(studentId) || phoneByCitizenId.get(citizenId) || '',
    });
  });
}

function pickValue(row, keys) {
  for (const key of keys) {
    const direct = row[key];
    if (direct !== undefined && direct !== null && String(direct).trim() !== '') return String(direct).trim();
    const foundKey = Object.keys(row).find((candidate) => normalizeHeader(candidate) === normalizeHeader(key));
    if (foundKey && String(row[foundKey] || '').trim() !== '') return String(row[foundKey]).trim();
  }
  return '';
}

function normalizeHeader(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function normalizeKey(value) {
  return String(value || '').replace(/\D/g, '') || String(value || '').trim();
}

function renderPreview(rows, headId, bodyId) {
  const head = document.getElementById(headId);
  const body = document.getElementById(bodyId);
  head.innerHTML = '';
  body.innerHTML = '';
  if (!rows.length) {
    body.innerHTML = '<tr><td class="text-center text-slate-500">ยังไม่มีข้อมูลตัวอย่าง</td></tr>';
    return;
  }
  const columns = Object.keys(rows[0]).slice(0, 8);
  head.innerHTML = `<tr>${columns.map((key) => `<th>${escapeHtml(key)}</th>`).join('')}</tr>`;
  rows.forEach((row) => {
    body.insertAdjacentHTML('beforeend', `<tr>${columns.map((key) => `<td>${escapeHtml(formatCell(row[key]))}</td>`).join('')}</tr>`);
  });
}

async function onImportStudents() {
  if (!state.studentRows.length) {
    toast('กรุณาเลือกไฟล์ข้อมูลนักเรียนก่อน', true);
    return;
  }
  const button = document.getElementById('importStudentsBtn');
  setButtonBusy(button, true, 'กำลังนำเข้า...');
  try {
    const summary = await importInChunks('importStudentMaster', state.studentRows, 'กำลังนำเข้าฐานนักเรียน');
    toast(`นำเข้านักเรียนสำเร็จ: เพิ่ม ${summary.imported_students || 0}, อัปเดต ${summary.updated_students || 0}, ข้าม ${summary.skipped_rows || 0}`);
    loadPublicDashboard();
    loadClasses();
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

async function onImportBorrowHistory() {
  if (!state.borrowRows.length) {
    toast('กรุณาเลือกไฟล์ฐานข้อมูลยืมก่อน', true);
    return;
  }
  const button = document.getElementById('importBorrowBtn');
  setButtonBusy(button, true, 'กำลังนำเข้า...');
  try {
    const summary = await importInChunks('importBorrowHistory', state.borrowRows, 'กำลังนำเข้ายอดยืม');
    toast(`นำเข้ายอดยืมสำเร็จ: เครื่องใหม่ ${summary.imported_devices || 0}, ครูใหม่ ${summary.imported_teachers || 0}, รายการ ${summary.imported_transactions || 0}, ไม่พบนักเรียน ${summary.missing_students || 0}, ข้าม ${summary.skipped_rows || 0}`);
    loadPublicDashboard();
    loadClasses();
  } catch (error) {
    toast(error.message, true);
  } finally {
    setButtonBusy(button, false);
  }
}

async function importInChunks(action, rows, label) {
  const total = rows.length;
  const summary = {};
  for (let start = 0; start < total; start += IMPORT_CHUNK_SIZE) {
    const chunk = rows.slice(start, start + IMPORT_CHUNK_SIZE);
    toast(`${label} ${Math.min(start + chunk.length, total)} / ${total}`);
    mergeSummary(summary, await api(action, { rows: chunk }));
  }
  return summary;
}

function mergeSummary(target, source) {
  Object.keys(source || {}).forEach((key) => {
    if (typeof source[key] === 'number') target[key] = (target[key] || 0) + source[key];
    else target[key] = source[key];
  });
}

async function api(action, data = {}) {
  if (!API_URL || API_URL === 'รอใส่ URL ของ GAS') {
    throw new Error('กรุณาใส่ Web App URL ของ Google Apps Script ในไฟล์ app.js ก่อน');
  }
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, data }),
  });
  const result = await response.json();
  if (!result.success) throw new Error(result.message || 'เกิดข้อผิดพลาด');
  return result.data;
}

function statusBadge(status) {
  const text = status || '-';
  const className = text === 'คืนแล้ว' || text === 'ว่าง'
    ? 'badge-green'
    : text === 'ส่งซ่อม'
      ? 'badge-red'
      : text === 'กำลังยืม'
        ? 'badge-orange'
        : 'badge-blue';
  return `<span class="badge ${className}">${escapeHtml(text)}</span>`;
}

function requestStatusBadge(status) {
  const text = status || 'รอตรวจสอบ';
  const className = text === 'อนุมัติแล้ว'
    ? 'badge-green'
    : text === 'ปฏิเสธ' || text === 'ยกเลิก'
      ? 'badge-red'
      : 'badge-orange';
  return `<span class="badge ${className}">${escapeHtml(text)}</span>`;
}

function debounce(fn, wait = 250) {
  let timer = 0;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function toast(message, isError = false) {
  const el = document.getElementById('toast');
  el.textContent = message;
  el.style.background = isError ? '#b91c1c' : '#0f172a';
  el.classList.remove('hidden');
  clearTimeout(window.toastTimer);
  window.toastTimer = setTimeout(() => el.classList.add('hidden'), 3600);
}

function setButtonBusy(button, isBusy, busyText = 'กำลังทำงาน...') {
  if (!button) return;
  if (isBusy) {
    button.dataset.originalText = button.textContent;
    button.textContent = busyText;
    button.disabled = true;
    button.setAttribute('aria-busy', 'true');
    return;
  }
  button.textContent = button.dataset.originalText || button.textContent;
  button.disabled = false;
  button.removeAttribute('aria-busy');
  delete button.dataset.originalText;
}

function formatCell(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return value ?? '';
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('`', '&#096;');
}
