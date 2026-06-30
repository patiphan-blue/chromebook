const SHEETS = {
  CONFIG: 'Config',
  ADMINS: 'Admins',
  STUDENTS: 'Students',
  TEACHERS: 'Teachers',
  CHROMEBOOKS: 'Chromebooks',
  TRANSACTIONS: 'Transactions',
  BORROW_REQUESTS: 'BorrowRequests',
};

const STATUS = {
  AVAILABLE: 'ว่าง',
  BORROWED_DEVICE: 'ถูกยืม',
  BORROWING: 'กำลังยืม',
  RETURNED: 'คืนแล้ว',
  REPAIR: 'ส่งซ่อม',
  REQUEST_PENDING: 'รอตรวจสอบ',
  REQUEST_APPROVED: 'อนุมัติแล้ว',
  REQUEST_REJECTED: 'ปฏิเสธ',
  REQUEST_CANCELLED: 'ยกเลิก',
};

const HEADERS = {
  Config: ['key', 'value', 'updated_at'],
  Admins: ['admin_id', 'username', 'password', 'full_name', 'role', 'is_active', 'created_at'],
  Students: [
    'student_id',
    'citizen_id',
    'grade_level',
    'student_no',
    'prefix',
    'first_name',
    'last_name',
    'full_name',
    'birth_date',
    'house_no',
    'village_no',
    'subdistrict',
    'district',
    'province',
    'phone',
    'address',
    'created_at',
    'updated_at',
  ],
  Teachers: ['teacher_id', 'prefix', 'full_name', 'phone', 'created_at', 'updated_at'],
  Chromebooks: ['device_key', 'asset_no', 'device_status', 'current_student_id', 'updated_at'],
  Transactions: ['transaction_id', 'borrower_type', 'borrower_id', 'student_id', 'teacher_id', 'borrower_name', 'device_key', 'borrow_date', 'return_date', 'status', 'note', 'created_at', 'updated_at'],
  BorrowRequests: ['request_id', 'citizen_id', 'student_id', 'full_name', 'parent_name', 'grade_level', 'phone', 'house_no', 'village_no', 'subdistrict', 'district', 'province', 'address', 'request_status', 'note', 'created_at', 'updated_at'],
};

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const payload = parsePayload(e);
    const action = String(payload.action || '').trim();
    const data = payload.data || {};

    const routes = {
      ping: () => ({ ok: true, message: 'Chromebook API is ready' }),
      login: () => login(data),
      dashboard: () => getDashboard(),
      dashboardTables: () => getDashboardTables(),
      assignDevice: () => assignDevice(data),
      bulkAssignDevices: () => bulkAssignDevices(data),
      validateBulkLoans: () => validateBulkLoans(data),
      importBulkLoans: () => importBulkLoans(data),
      getStudentsByClass: () => getStudentsByClass(data),
      getBorrowersByClass: () => getBorrowersByClass(data),
      bulkReturn: () => bulkReturn(data),
      importStudentMaster: () => importStudentMaster(data),
      importBorrowHistory: () => importBorrowHistory(data),
      importData: () => importBorrowHistory(data),
      listClasses: () => listClasses(),
      listGradeGroups: () => listGradeGroups(),
      listUnborrowedStudentsByGrade: () => listUnborrowedStudentsByGrade(data),
      listAvailableDevices: () => listAvailableDevices(),
      listAvailableDeviceReport: () => listAvailableDeviceReport(),
      createBorrowRequest: () => createBorrowRequest(data),
      listBorrowRequests: () => listBorrowRequests(data),
      updateBorrowRequestStatus: () => updateBorrowRequestStatus(data),
    };

    if (!routes[action]) {
      return jsonResponse({ success: false, message: 'Invalid action: ' + action });
    }

    return jsonResponse({ success: true, data: routes[action]() });
  } catch (err) {
    return jsonResponse({ success: false, message: err.message, stack: err.stack });
  }
}

function parsePayload(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }

  const params = (e && e.parameter) || {};
  if (params.payload) return JSON.parse(params.payload);
  return params;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Google Sheet ที่ผูกกับ Apps Script กรุณาเปิดจาก Google Sheets > Extensions > Apps Script');
  }

  var sheetNames = Object.keys(HEADERS);
  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var headers = HEADERS[sheetName];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#dbeafe');
    sheet.autoResizeColumns(1, headers.length);
  }

  var now = nowText();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  configSheet.getRange(2, 1, 1, 3).setValues([[
    'school_name',
    'โรงเรียนของคุณ',
    now,
  ]]);

  var adminSheet = ss.getSheetByName(SHEETS.ADMINS);
  adminSheet.getRange(2, 1, 1, 7).setValues([[
    Utilities.getUuid(),
    'admin',
    'admin123',
    'ผู้ดูแลระบบ',
    'admin',
    'TRUE',
    now,
  ]]);

  SpreadsheetApp.flush();
  return 'Database setup completed. Default login: admin / admin123';
}

function upgradeDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Google Sheet ที่ผูกกับ Apps Script กรุณาเปิดจาก Google Sheets > Extensions > Apps Script');
  }

  Object.keys(HEADERS).forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]);
      sheet.setFrozenRows(1);
    } else {
      ensureHeaders(sheet, sheetName);
    }
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#dbeafe');
    sheet.autoResizeColumns(1, Math.max(sheet.getLastColumn(), HEADERS[sheetName].length));
  });

  SpreadsheetApp.flush();
  return 'Database upgraded safely. Existing data was not cleared.';
}

function login(data) {
  const username = String(data.username || '').trim();
  const password = String(data.password || '').trim();
  if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');

  const admins = getRows(SHEETS.ADMINS);
  const admin = admins.find((row) =>
    String(row.username) === username &&
    String(row.password) === password &&
    String(row.is_active).toUpperCase() !== 'FALSE'
  );

  if (!admin) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');

  return {
    admin_id: admin.admin_id,
    username: admin.username,
    full_name: admin.full_name,
    role: admin.role,
  };
}

function getDashboard() {
  const inventory = getSynchronizedInventory(10000);
  const chromebooks = inventory.chromebooks;
  const transactions = inventory.transactions;
  const students = getRows(SHEETS.STUDENTS);
  const active = transactions.filter((row) => isBorrowingStatus(row.status));

  const statusCounts = chromebooks.reduce((acc, row) => {
    const status = normalizeDeviceStatus(row.device_status) || 'ไม่ระบุ';
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {});

  return {
    total_students: students.length,
    total_devices: chromebooks.length,
    available: statusCounts[STATUS.AVAILABLE] || 0,
    borrowed: (statusCounts[STATUS.BORROWED_DEVICE] || 0) + (statusCounts[STATUS.BORROWING] || 0),
    repairing: statusCounts[STATUS.REPAIR] || 0,
    active_transactions: active.length,
    status_counts: statusCounts,
    synced_devices: inventory.synced_devices,
    available_devices: chromebooks
      .filter((row) => normalizeDeviceStatus(row.device_status) === STATUS.AVAILABLE)
      .map((row) => ({ device_key: row.device_key, asset_no: row.asset_no })),
    recent_transactions: transactions.slice(-10).reverse(),
  };
}

function getSynchronizedInventory(waitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs || 10000)) {
    throw new Error('ระบบกำลังอัปเดตสถานะเครื่อง กรุณาลองใหม่อีกครั้ง');
  }

  try {
    const chromebooks = getRows(SHEETS.CHROMEBOOKS);
    const transactions = getRows(SHEETS.TRANSACTIONS);
    const syncedDevices = reconcileDeviceRows(chromebooks, transactions);
    if (syncedDevices > 0) {
      rewriteObjects(SHEETS.CHROMEBOOKS, chromebooks);
      SpreadsheetApp.flush();
    }

    return {
      chromebooks,
      transactions,
      synced_devices: syncedDevices,
    };
  } finally {
    lock.releaseLock();
  }
}

function reconcileDeviceRows(chromebooks, transactions) {
  const activeByDevice = {};
  transactions.forEach((row) => {
    const deviceKey = String(row.device_key || '').trim();
    if (deviceKey && isBorrowingStatus(row.status)) activeByDevice[deviceKey] = row;
  });

  let changed = 0;
  chromebooks.forEach((device) => {
    const deviceKey = String(device.device_key || '').trim();
    const active = activeByDevice[deviceKey];
    const rawStatus = String(device.device_status || '').trim();
    const currentStatus = normalizeDeviceStatus(rawStatus);
    let nextStatus = currentStatus;
    let nextBorrowerId = device.current_student_id || '';

    if (active && currentStatus !== STATUS.REPAIR) {
      nextStatus = STATUS.BORROWED_DEVICE;
      nextBorrowerId = getTransactionBorrowerType(active) === 'teacher'
        ? active.teacher_id || active.borrower_id || active.student_id || ''
        : active.student_id || active.borrower_id || '';
    } else if (!active && (currentStatus === STATUS.BORROWED_DEVICE || currentStatus === STATUS.BORROWING)) {
      nextStatus = STATUS.AVAILABLE;
      nextBorrowerId = '';
    } else if (!active && currentStatus === STATUS.AVAILABLE) {
      nextBorrowerId = '';
    }

    if (nextStatus !== rawStatus || String(nextBorrowerId) !== String(device.current_student_id || '')) {
      device.device_status = nextStatus;
      device.current_student_id = nextBorrowerId;
      device.updated_at = nowText();
      changed++;
    }
  });

  return changed;
}

function getDashboardTables() {
  const studentsById = indexBy(getRows(SHEETS.STUDENTS), 'student_id');
  const teachersById = indexBy(getRows(SHEETS.TEACHERS), 'teacher_id');
  const rows = getRows(SHEETS.TRANSACTIONS)
    .slice()
    .reverse()
    .map((tx) => formatDashboardTransaction(tx, studentsById, teachersById));

  return {
    students: rows.filter((row) => row.borrower_type === 'student' && row.status !== STATUS.RETURNED),
    teachers: rows.filter((row) => row.borrower_type === 'teacher' && row.status !== STATUS.RETURNED),
    returned: rows.filter((row) => row.status === STATUS.RETURNED),
  };
}

function formatDashboardTransaction(tx, studentsById, teachersById) {
  const borrowerType = getTransactionBorrowerType(tx);
  if (borrowerType === 'teacher') {
    const teacherId = tx.teacher_id || tx.borrower_id || tx.student_id || '';
    const teacher = teachersById[teacherId] || {};
    return {
      transaction_id: tx.transaction_id,
      borrower_type: 'teacher',
      borrower_id: teacherId || '0',
      full_name: tx.borrower_name || teacher.full_name || 'ครู',
      grade_level: 'ครู',
      borrow_date: tx.borrow_date || '',
      return_date: tx.return_date || '',
      status: tx.status || '',
      device_key: tx.device_key || '',
    };
  }

  const studentId = tx.student_id || tx.borrower_id || '';
  const student = studentsById[studentId] || {};
  return {
    transaction_id: tx.transaction_id,
    borrower_type: 'student',
    borrower_id: studentId,
    full_name: student.full_name || tx.borrower_name || '',
    grade_level: student.grade_level || '',
    borrow_date: tx.borrow_date || '',
    return_date: tx.return_date || '',
    status: tx.status || '',
    device_key: tx.device_key || '',
  };
}

function assignDevice(data) {
  const borrowerType = String(data.borrower_type || 'student').trim();
  const studentId = String(data.student_id || '').trim();
  const teacherName = String(data.teacher_name || '').trim();
  const deviceKey = required(data.device_key, 'เลขเครื่องนิยม');
  const borrowDate = data.borrow_date || todayText();
  const note = data.note || '';
  let borrower;

  if (borrowerType === 'teacher') {
    if (!teacherName) throw new Error('กรุณาระบุชื่อครู');
    borrower = getOrCreateTeacher(teacherName, data.phone || '');
  } else {
    const requiredStudentId = required(studentId, 'รหัสนักเรียน');
    const student = findRowByValue(SHEETS.STUDENTS, 'student_id', requiredStudentId);
    if (!student) throw new Error('ไม่พบนักเรียนรหัส ' + requiredStudentId);
    borrower = {
      borrower_type: 'student',
      borrower_id: requiredStudentId,
      student_id: requiredStudentId,
      teacher_id: '',
      borrower_name: student.row.full_name,
    };
  }

  const device = findRowByValue(SHEETS.CHROMEBOOKS, 'device_key', deviceKey);
  if (!device) throw new Error('ไม่พบ Chromebook เลขเครื่อง ' + deviceKey);
  const deviceStatus = normalizeDeviceStatus(device.row.device_status);
  if (deviceStatus === STATUS.REPAIR) throw new Error('เครื่องนี้อยู่ระหว่างส่งซ่อม');
  if (deviceStatus === STATUS.BORROWED_DEVICE || deviceStatus === STATUS.BORROWING) {
    throw new Error('เครื่องนี้ถูกยืมอยู่แล้ว');
  }

  const devicesByKey = indexBy(getRows(SHEETS.CHROMEBOOKS), 'device_key');
  const existing = getRows(SHEETS.TRANSACTIONS).find((row) =>
    getTransactionBorrowerKey(row) === borrower.borrower_type + ':' + borrower.borrower_id &&
    isBorrowingStatus(row.status) &&
    isTransactionDeviceStillBorrowed(row, devicesByKey)
  );
  if (existing) throw new Error('ผู้ยืมคนนี้ยังมีรายการยืมค้างอยู่');

  const now = nowText();
  appendObjects(SHEETS.TRANSACTIONS, [{
    transaction_id: Utilities.getUuid(),
    borrower_type: borrower.borrower_type,
    borrower_id: borrower.borrower_id,
    student_id: borrower.student_id,
    teacher_id: borrower.teacher_id,
    borrower_name: borrower.borrower_name,
    device_key: deviceKey,
    borrow_date: borrowDate,
    return_date: '',
    status: STATUS.BORROWING,
    note,
    created_at: now,
    updated_at: now,
  }]);

  updateRowByIndex(SHEETS.CHROMEBOOKS, device.index, {
    device_status: STATUS.BORROWED_DEVICE,
    current_student_id: borrower.borrower_id,
    updated_at: now,
  });

  return { message: 'บันทึกการยืมสำเร็จ', borrower_id: borrower.borrower_id, device_key: deviceKey };
}

function getStudentsByClass(data) {
  const gradeLevel = required(data.grade_level, 'ห้องเรียน');
  const devicesByKey = indexBy(getRows(SHEETS.CHROMEBOOKS), 'device_key');
  const activeBorrowerKeys = new Set(getRows(SHEETS.TRANSACTIONS)
    .filter((row) => isBorrowingStatus(row.status) && isTransactionDeviceStillBorrowed(row, devicesByKey))
    .map((row) => getTransactionBorrowerKey(row)));

  return getRows(SHEETS.STUDENTS)
    .filter((row) => row.grade_level === gradeLevel)
    .sort((a, b) => Number(a.student_no || 9999) - Number(b.student_no || 9999))
    .map((row) => ({
      student_id: row.student_id,
      full_name: row.full_name,
      grade_level: row.grade_level,
      student_no: row.student_no,
      phone: row.phone,
      already_borrowing: activeBorrowerKeys.has('student:' + row.student_id),
    }));
}

function bulkAssignDevices(data) {
  const assignments = data.assignments || [];
  const borrowDate = data.borrow_date || todayText();
  const note = data.note || '';
  if (!Array.isArray(assignments) || assignments.length === 0) {
    throw new Error('กรุณาเลือกรายการยืมอย่างน้อย 1 รายการ');
  }

  const now = nowText();
  const studentsById = indexBy(getRows(SHEETS.STUDENTS), 'student_id');
  const deviceRows = getRows(SHEETS.CHROMEBOOKS);
  const devicesByKey = indexBy(deviceRows, 'device_key');
  const txRows = getRows(SHEETS.TRANSACTIONS);
  const activeBorrowerKeys = new Set(txRows
    .filter((row) => isBorrowingStatus(row.status) && isTransactionDeviceStillBorrowed(row, devicesByKey))
    .map((row) => getTransactionBorrowerKey(row)));
  const usedDevices = new Set();
  let assigned = 0;
  let skipped = 0;
  const errors = [];

  assignments.forEach((item) => {
    const studentId = String(item.student_id || '').trim();
    const deviceKey = String(item.device_key || '').trim();
    if (!studentId || !deviceKey) {
      skipped++;
      return;
    }
    const student = studentsById[studentId];
    const device = devicesByKey[deviceKey];
    if (!student) {
      skipped++;
      errors.push(studentId + ': ไม่พบนักเรียน');
      return;
    }
    if (!device) {
      skipped++;
      errors.push(deviceKey + ': ไม่พบเครื่อง');
      return;
    }
    if (normalizeDeviceStatus(device.device_status) !== STATUS.AVAILABLE || usedDevices.has(deviceKey)) {
      skipped++;
      errors.push(deviceKey + ': เครื่องไม่ว่าง');
      return;
    }
    if (activeBorrowerKeys.has('student:' + studentId)) {
      skipped++;
      errors.push(studentId + ': มีรายการยืมค้างอยู่');
      return;
    }

    txRows.push({
      transaction_id: Utilities.getUuid(),
      borrower_type: 'student',
      borrower_id: studentId,
      student_id: studentId,
      teacher_id: '',
      borrower_name: student.full_name,
      device_key: deviceKey,
      borrow_date: borrowDate,
      return_date: '',
      status: STATUS.BORROWING,
      note,
      created_at: now,
      updated_at: now,
    });
    device.device_status = STATUS.BORROWED_DEVICE;
    device.current_student_id = studentId;
    device.updated_at = now;
    usedDevices.add(deviceKey);
    activeBorrowerKeys.add('student:' + studentId);
    assigned++;
  });

  rewriteObjects(SHEETS.CHROMEBOOKS, deviceRows);
  rewriteObjects(SHEETS.TRANSACTIONS, txRows);

  return {
    message: 'บันทึกยืมรายห้องสำเร็จ',
    assigned_count: assigned,
    skipped_count: skipped,
    errors: errors.slice(0, 20),
  };
}

function validateBulkLoans(data) {
  return processBulkLoans(data, false);
}

function importBulkLoans(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่อีกครั้ง');
  }

  try {
    return processBulkLoans(data, true);
  } finally {
    lock.releaseLock();
  }
}

function processBulkLoans(data, shouldWrite) {
  const rows = data.rows || [];
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error('ไม่พบรายการยืมสำหรับนำเข้า');
  }
  if (rows.length > 500) {
    throw new Error('นำเข้าได้ไม่เกิน 500 แถวต่อครั้ง');
  }

  const now = nowText();
  const studentRows = getRows(SHEETS.STUDENTS);
  const deviceRows = getRows(SHEETS.CHROMEBOOKS);
  const txRows = getRows(SHEETS.TRANSACTIONS);
  const studentsById = indexBy(studentRows, 'student_id');
  const devicesByKey = indexBy(deviceRows, 'device_key');
  const devicesByAssetNo = {};
  deviceRows.forEach((device) => {
    const assetKey = normalizeAssetNo(device.asset_no);
    if (!assetKey) return;
    if (!devicesByAssetNo[assetKey]) devicesByAssetNo[assetKey] = [];
    devicesByAssetNo[assetKey].push(device);
  });
  const activeBorrowerKeys = new Set(txRows
    .filter((row) => isBorrowingStatus(row.status) && isTransactionDeviceStillBorrowed(row, devicesByKey))
    .map((row) => getTransactionBorrowerKey(row)));
  const usedStudentIds = new Set();
  const usedDeviceKeys = new Set();
  const results = [];
  let assigned = 0;
  let skipped = 0;

  rows.forEach((raw, index) => {
    const mapped = mapBulkLoanRow(raw, index);
    const student = studentsById[mapped.student_id];
    const assetMatches = mapped.asset_no
      ? devicesByAssetNo[normalizeAssetNo(mapped.asset_no)] || []
      : [];
    const device = mapped.asset_no
      ? (assetMatches.length === 1 ? assetMatches[0] : null)
      : devicesByKey[mapped.device_key];
    const resolvedDeviceKey = device ? String(device.device_key || '').trim() : String(mapped.device_key || '').trim();
    let error = '';

    if (!mapped.student_id) error = 'ไม่พบรหัสนักเรียน';
    else if (!mapped.asset_no && !mapped.device_key) error = 'ไม่พบเลขที่ทรัพย์สิน';
    else if (!mapped.borrow_date) error = 'วันที่ยืมต้องเป็นรูปแบบ YYYY-MM-DD';
    else if (usedStudentIds.has(mapped.student_id)) error = 'รหัสนักเรียนซ้ำในไฟล์';
    else if (mapped.asset_no && assetMatches.length === 0) error = 'ไม่พบเลขที่ทรัพย์สินในทะเบียนเครื่อง';
    else if (mapped.asset_no && assetMatches.length > 1) error = 'เลขที่ทรัพย์สินซ้ำในทะเบียนเครื่อง';
    else if (!student) error = 'ไม่พบนักเรียนในฐานข้อมูล';
    else if (!device) error = 'ไม่พบเครื่องในฐานข้อมูล';
    else if (usedDeviceKeys.has(resolvedDeviceKey)) error = 'เลขที่ทรัพย์สินหรือเลขเครื่องซ้ำในไฟล์';
    else if (normalizeDeviceStatus(device.device_status) !== STATUS.AVAILABLE) error = 'เครื่องไม่ว่าง';
    else if (activeBorrowerKeys.has('student:' + mapped.student_id)) error = 'นักเรียนมีรายการยืมค้างอยู่';

    usedStudentIds.add(mapped.student_id);
    if (resolvedDeviceKey) usedDeviceKeys.add(resolvedDeviceKey);
    if (device) {
      mapped.device_key = resolvedDeviceKey;
      mapped.asset_no = String(device.asset_no || mapped.asset_no || '').trim();
    }

    if (error) {
      skipped++;
      results.push({
        source_sheet: mapped.source_sheet,
        source_row: mapped.source_row,
        student_id: mapped.student_id,
        asset_no: mapped.asset_no,
        device_key: mapped.device_key,
        borrow_date: mapped.borrow_date,
        success: false,
        message: error,
      });
      return;
    }

    activeBorrowerKeys.add('student:' + mapped.student_id);
    device.device_status = STATUS.BORROWED_DEVICE;
    device.current_student_id = mapped.student_id;
    device.updated_at = now;

    if (shouldWrite) {
      txRows.push({
        transaction_id: Utilities.getUuid(),
        borrower_type: 'student',
        borrower_id: mapped.student_id,
        student_id: mapped.student_id,
        teacher_id: '',
        borrower_name: student.full_name,
        device_key: mapped.device_key,
        borrow_date: mapped.borrow_date,
        return_date: '',
        status: STATUS.BORROWING,
        note: mapped.note || 'นำเข้ารายการยืมจาก Excel',
        created_at: now,
        updated_at: now,
      });
      assigned++;
    }

    results.push({
      source_sheet: mapped.source_sheet,
      source_row: mapped.source_row,
      student_id: mapped.student_id,
      full_name: student.full_name || '',
      grade_level: student.grade_level || '',
      asset_no: mapped.asset_no,
      device_key: mapped.device_key,
      borrow_date: mapped.borrow_date,
      success: true,
      message: shouldWrite ? 'บันทึกแล้ว' : 'พร้อมบันทึก',
    });
  });

  if (shouldWrite && assigned > 0) {
    rewriteObjects(SHEETS.CHROMEBOOKS, deviceRows);
    rewriteObjects(SHEETS.TRANSACTIONS, txRows);
    SpreadsheetApp.flush();
  }

  return {
    message: shouldWrite ? 'นำเข้ารายการยืมสำเร็จ' : 'ตรวจสอบไฟล์เรียบร้อย',
    total_count: rows.length,
    valid_count: results.filter((row) => row.success).length,
    assigned_count: assigned,
    skipped_count: skipped,
    results,
  };
}

function mapBulkLoanRow(row, index) {
  const get = makeGetter(row || {});
  const sourceRow = Number(row.source_row || row.__source_row || index + 2);
  return {
    source_sheet: String(row.source_sheet || row.__source_sheet || '').trim(),
    source_row: sourceRow,
    student_id: get('student_id', 'รหัสนักเรียน', 'เลขนักเรียน', 'เลขประจำตัวนักเรียน'),
    asset_no: get('asset_no', 'เลขที่ทรัพย์สิน', 'เลขทรัพย์สิน', 'เลขครุภัณฑ์'),
    device_key: get('device_key', 'เลขเครื่องนิยม', 'เลขเครื่อง', 'หมายเลขเครื่อง', 'Key', 'key'),
    borrow_date: normalizeBulkLoanDate(getRawValue(row, 'borrow_date', 'วันที่ยืม')),
    note: get('note', 'หมายเหตุ'),
  };
}

function normalizeAssetNo(value) {
  return String(value || '')
    .normalize('NFC')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function normalizeBulkLoanDate(value) {
  const text = String(normalizeExcelDate(value) || '').trim() || todayText();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return '';

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return '';
  return text;
}

function getBorrowersByClass(data) {
  const gradeLevel = required(data.grade_level, 'ห้องเรียน');
  if (gradeLevel === 'ครู' || gradeLevel === '__TEACHERS__') {
    return getTeacherBorrowers();
  }

  const studentsById = indexBy(getRows(SHEETS.STUDENTS), 'student_id');
  const active = getRows(SHEETS.TRANSACTIONS)
    .filter((row) => isBorrowingStatus(row.status))
    .map((row) => ({ transaction: row, student: studentsById[row.student_id] }))
    .filter((item) => item.student && item.student.grade_level === gradeLevel)
    .map((item) => ({
      transaction_id: item.transaction.transaction_id,
      student_id: item.student.student_id,
      borrower_type: 'student',
      full_name: item.student.full_name,
      grade_level: item.student.grade_level,
      student_no: item.student.student_no,
      phone: item.student.phone,
      device_key: item.transaction.device_key,
      borrow_date: item.transaction.borrow_date,
    }));

  return active;
}

function getTeacherBorrowers() {
  const teachersById = indexBy(getRows(SHEETS.TEACHERS), 'teacher_id');
  return getRows(SHEETS.TRANSACTIONS)
    .filter((row) => isBorrowingStatus(row.status) && getTransactionBorrowerType(row) === 'teacher')
    .map((row) => {
      const teacher = teachersById[row.teacher_id || row.borrower_id] || {};
      return {
        transaction_id: row.transaction_id,
        student_id: row.teacher_id || row.borrower_id || '0',
        borrower_type: 'teacher',
        full_name: row.borrower_name || teacher.full_name || 'ครู',
        grade_level: 'ครู',
        student_no: '',
        phone: teacher.phone || '',
        device_key: row.device_key,
        borrow_date: row.borrow_date,
      };
    });
}

function bulkReturn(data) {
  const transactionIds = data.transaction_ids || [];
  const returnDate = data.return_date || todayText();
  if (!Array.isArray(transactionIds) || transactionIds.length === 0) {
    throw new Error('กรุณาเลือกรายการที่ต้องการคืน');
  }

  const txSheet = getSheet(SHEETS.TRANSACTIONS);
  const txValues = txSheet.getDataRange().getValues();
  const txHeaders = txValues[0];
  const txIdCol = txHeaders.indexOf('transaction_id');
  const statusCol = txHeaders.indexOf('status');
  const returnCol = txHeaders.indexOf('return_date');
  const updatedCol = txHeaders.indexOf('updated_at');
  const deviceCol = txHeaders.indexOf('device_key');
  const idSet = new Set(transactionIds.map(String));
  const borrowerKeysToClose = new Set();
  const returnedDevices = [];
  const now = nowText();

  for (let i = 1; i < txValues.length; i++) {
    const row = txValues[i];
    if (idSet.has(String(row[txIdCol])) && isBorrowingStatus(row[statusCol])) {
      borrowerKeysToClose.add(getTransactionBorrowerKey(rowToObject(txHeaders, row)));
      row[statusCol] = STATUS.RETURNED;
      row[returnCol] = returnDate;
      row[updatedCol] = now;
      returnedDevices.push(row[deviceCol]);
    }
  }

  for (let i = 1; i < txValues.length; i++) {
    const row = txValues[i];
    const rowObj = rowToObject(txHeaders, row);
    if (borrowerKeysToClose.has(getTransactionBorrowerKey(rowObj)) && isBorrowingStatus(row[statusCol])) {
      row[statusCol] = STATUS.RETURNED;
      row[returnCol] = row[returnCol] || returnDate;
      row[updatedCol] = now;
      returnedDevices.push(row[deviceCol]);
    }
  }

  if (txValues.length > 1) {
    txSheet.getRange(2, 1, txValues.length - 1, txHeaders.length).setValues(txValues.slice(1));
  }

  const deviceRows = getRows(SHEETS.CHROMEBOOKS);
  const returnedSet = new Set(returnedDevices.map(String));
  deviceRows.forEach((row) => {
    if (returnedSet.has(String(row.device_key))) {
      row.device_status = STATUS.AVAILABLE;
      row.current_student_id = '';
      row.updated_at = now;
    }
  });
  rewriteObjects(SHEETS.CHROMEBOOKS, deviceRows);

  return { message: 'คืนเครื่องสำเร็จ', returned_count: returnedDevices.length };
}

function importStudentMaster(data) {
  const rows = data.rows || [];
  if (!Array.isArray(rows) || rows.length === 0) throw new Error('ไม่พบข้อมูลนักเรียนสำหรับนำเข้า');

  const now = nowText();
  const existingRows = getRows(SHEETS.STUDENTS);
  const studentMap = indexBy(existingRows, 'student_id');
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  rows.forEach((raw) => {
    const mapped = mapStudentRow(raw);
    if (!mapped.student_id || !mapped.full_name) {
      skipped++;
      return;
    }

    if (studentMap[mapped.student_id]) {
      Object.assign(studentMap[mapped.student_id], {
        citizen_id: mapped.citizen_id,
        full_name: mapped.full_name,
        grade_level: mapped.grade_level,
        student_no: mapped.student_no,
        prefix: mapped.prefix,
        first_name: mapped.first_name,
        last_name: mapped.last_name,
        birth_date: mapped.birth_date,
        house_no: mapped.house_no,
        village_no: mapped.village_no,
        subdistrict: mapped.subdistrict,
        district: mapped.district,
        province: mapped.province,
        phone: mapped.phone,
        address: mapped.address,
        updated_at: now,
      });
      updated++;
    } else {
      const row = {
        student_id: mapped.student_id,
        citizen_id: mapped.citizen_id,
        full_name: mapped.full_name,
        grade_level: mapped.grade_level,
        student_no: mapped.student_no,
        prefix: mapped.prefix,
        first_name: mapped.first_name,
        last_name: mapped.last_name,
        birth_date: mapped.birth_date,
        house_no: mapped.house_no,
        village_no: mapped.village_no,
        subdistrict: mapped.subdistrict,
        district: mapped.district,
        province: mapped.province,
        phone: mapped.phone,
        address: mapped.address,
        created_at: now,
        updated_at: now,
      };
      existingRows.push(row);
      studentMap[mapped.student_id] = row;
      inserted++;
    }
  });

  rewriteObjects(SHEETS.STUDENTS, existingRows);
  return {
    message: 'นำเข้าฐานข้อมูลนักเรียนสำเร็จ',
    imported_students: inserted,
    updated_students: updated,
    skipped_rows: skipped,
  };
}

function importBorrowHistory(data) {
  const rows = data.rows || [];
  if (!Array.isArray(rows) || rows.length === 0) throw new Error('ไม่พบข้อมูลยอดยืมสำหรับนำเข้า');

  const now = nowText();
  const studentRows = getRows(SHEETS.STUDENTS);
  const teacherRows = getRows(SHEETS.TEACHERS);
  const deviceRows = getRows(SHEETS.CHROMEBOOKS);
  const txRows = getRows(SHEETS.TRANSACTIONS);
  const studentMap = indexBy(studentRows, 'student_id');
  const teacherMap = indexBy(teacherRows, 'teacher_id');
  const deviceMap = indexBy(deviceRows, 'device_key');
  const existingActive = new Set(txRows
    .filter((row) => isBorrowingStatus(row.status))
    .map((row) => getTransactionBorrowerKey(row) + '|' + String(row.device_key)));
  const activeTxByDevice = {};
  const activeTxListByDevice = {};
  txRows.forEach((row) => {
    if (isBorrowingStatus(row.status) && row.device_key) {
      const deviceKey = String(row.device_key);
      if (!activeTxByDevice[deviceKey]) activeTxByDevice[deviceKey] = row;
      if (!activeTxListByDevice[deviceKey]) activeTxListByDevice[deviceKey] = [];
      activeTxListByDevice[deviceKey].push(row);
    }
  });

  let deviceInserted = 0;
  let deviceUpdated = 0;
  let transactionInserted = 0;
  let transactionUpdated = 0;
  let duplicateTransactionsClosed = 0;
  let missingStudents = 0;
  let importedTeachers = 0;
  let skipped = 0;

  rows.forEach((raw) => {
    const mapped = mapBorrowRow(raw);
    if (!getMappedBorrowerId(mapped) || !mapped.device_key) {
      skipped++;
      return;
    }

    const borrower = resolveBorrowerFromImport(mapped, studentMap, teacherRows, teacherMap, now);
    if (borrower.borrower_type === 'student' && !studentMap[borrower.student_id]) {
      missingStudents++;
    }
    if (borrower.created_teacher) importedTeachers++;

    if (deviceMap[mapped.device_key]) {
      Object.assign(deviceMap[mapped.device_key], {
        asset_no: mapped.asset_no || deviceMap[mapped.device_key].asset_no,
        device_status: STATUS.BORROWED_DEVICE,
        current_student_id: borrower.borrower_id,
        updated_at: now,
      });
      deviceUpdated++;
    } else {
      const device = {
        device_key: mapped.device_key,
        asset_no: mapped.asset_no,
        device_status: STATUS.BORROWED_DEVICE,
        current_student_id: borrower.borrower_id,
        updated_at: now,
      };
      deviceRows.push(device);
      deviceMap[mapped.device_key] = device;
      deviceInserted++;
    }

    const txKey = borrower.borrower_type + ':' + borrower.borrower_id + '|' + String(mapped.device_key);
    const deviceTxKey = String(mapped.device_key);
    const activeTx = activeTxByDevice[deviceTxKey];
    if (activeTx) {
      activeTx.borrower_type = borrower.borrower_type;
      activeTx.borrower_id = borrower.borrower_id;
      activeTx.student_id = borrower.student_id;
      activeTx.teacher_id = borrower.teacher_id;
      activeTx.borrower_name = borrower.borrower_name;
      activeTx.borrow_date = mapped.borrow_date || activeTx.borrow_date || todayText();
      activeTx.return_date = '';
      activeTx.status = STATUS.BORROWING;
      activeTx.note = 'Updated from Excel import';
      activeTx.updated_at = now;
      existingActive.add(txKey);
      transactionUpdated++;

      const duplicateTxRows = (activeTxListByDevice[deviceTxKey] || []).filter((row) => row !== activeTx);
      duplicateTxRows.forEach((row) => {
        row.status = STATUS.RETURNED;
        row.return_date = row.return_date || todayText();
        row.note = row.note || 'Closed duplicate active transaction during Excel import';
        row.updated_at = now;
        duplicateTransactionsClosed++;
      });
    } else if (!existingActive.has(txKey)) {
      txRows.push({
        transaction_id: Utilities.getUuid(),
        borrower_type: borrower.borrower_type,
        borrower_id: borrower.borrower_id,
        student_id: borrower.student_id,
        teacher_id: borrower.teacher_id,
        borrower_name: borrower.borrower_name,
        device_key: mapped.device_key,
        borrow_date: mapped.borrow_date || todayText(),
        return_date: '',
        status: STATUS.BORROWING,
        note: 'ยอดยกมาจาก Excel',
        created_at: now,
        updated_at: now,
      });
      existingActive.add(txKey);
      activeTxByDevice[deviceTxKey] = txRows[txRows.length - 1];
      activeTxListByDevice[deviceTxKey] = [txRows[txRows.length - 1]];
      transactionInserted++;
    }
  });

  rewriteObjects(SHEETS.TEACHERS, teacherRows);
  rewriteObjects(SHEETS.CHROMEBOOKS, deviceRows);
  rewriteObjects(SHEETS.TRANSACTIONS, txRows);

  return {
    message: 'นำเข้ายอดยืมเก่าสำเร็จ',
    imported_devices: deviceInserted,
    updated_devices: deviceUpdated,
    imported_transactions: transactionInserted,
    updated_transactions: transactionUpdated,
    closed_duplicate_transactions: duplicateTransactionsClosed,
    imported_teachers: importedTeachers,
    missing_students: missingStudents,
    skipped_rows: skipped,
  };
}

function mapStudentRow(row) {
  const get = makeGetter(row);
  const prefix = get('คำนำหน้า', 'prefix', 'คำนำหน้าชื่อ');
  const firstName = get('ชื่อ', 'first_name');
  const lastName = get('สกุล', 'นามสกุล', 'last_name');
  const fullName = get('คำนำหน้า+ชื่อ+สกุล', 'ชื่อสกุล', 'ชื่อ-นามสกุล', 'full_name', 'ชื่อ - นามสกุล', 'ชื่อ นามสกุล') ||
    [prefix + firstName, lastName].filter(Boolean).join(' ');
  const houseNo = get('บ้านเลขที่', 'house_no', 'เลขที่บ้าน');
  const villageNo = get('หมู่', 'หมู่บ้าน', 'village_no', 'village');
  const subdistrict = get('ตำบล', 'subdistrict');
  const district = get('อำเภอ', 'district');
  const province = get('จังหวัด', 'province');
  const addressParts = [
    houseNo ? 'บ้านเลขที่ ' + houseNo : '',
    villageNo ? 'หมู่ ' + villageNo : '',
    subdistrict,
    district,
    province,
  ].filter(Boolean);

  return {
    student_id: get('รหัสนักเรียน', 'student_id', 'เลขประจำตัว', 'เลขประจำตัวนักเรียน'),
    citizen_id: get('เลขบัตรประชาชนนักเรียน', 'เลขบัตรประชาชน', 'citizen_id', 'national_id', 'เลขประจำตัวประชาชน'),
    grade_level: get('ชั้น+ห้อง', 'ระดับชั้น+ห้อง', 'grade_level', 'ห้อง', 'ชั้น', 'ระดับชั้น', 'ชั้น/ห้อง'),
    student_no: get('เลขที่', 'student_no'),
    prefix,
    first_name: firstName,
    last_name: lastName,
    full_name: fullName,
    birth_date: normalizeExcelDate(getRawValue(row, 'วันเกิด', 'birth_date', 'วัน/เดือน/ปีเกิด')),
    house_no: houseNo,
    village_no: villageNo,
    subdistrict,
    district,
    province,
    phone: get('เบอร์โทรศัพท์', 'เบอร์โทรศัพท์มือถือ', 'เบอร์โทร', 'phone', 'โทรศัพท์', 'โทรศัพท์มือถือ'),
    address: get('address', 'ที่อยู่') || addressParts.join(' '),
  };
}

function mapBorrowRow(row) {
  const student = mapStudentRow(row);
  const get = makeGetter(row);
  const borrowerType = normalizeBorrowerType(get('borrower_type', 'ประเภทผู้ยืม', 'ประเภท'));
  const borrowerId = get('borrower_id', 'รหัสผู้ยืม', 'รหัส', 'รหัสครู', 'teacher_id', 'เลขประจำตัวครู', 'เลขครู');
  const teacherId = get('teacher_id', 'รหัสครู', 'เลขประจำตัวครู', 'เลขครู');
  return Object.assign(student, {
    borrower_type: borrowerType,
    borrower_id: borrowerId,
    teacher_id: teacherId || (looksLikeTeacherId(borrowerId) ? borrowerId : ''),
    device_key: get('เลขเครื่องนิยม', 'Key', 'key', 'device_key', 'เลขเครื่อง', 'หมายเลขเครื่อง'),
    asset_no: get('เลขที่ทรัพย์สิน', 'asset_no', 'เลขครุภัณฑ์'),
    borrow_date: normalizeExcelDate(getRawValue(row, 'วันที่ตรวจสอบ', 'borrow_date', 'วันที่ยืม')),
  });
}

function resolveBorrowerFromImport(mapped, studentMap, teacherRows, teacherMap, now) {
  if (isTeacherBorrower(mapped)) {
    const teacher = getOrCreateTeacherInRows(
      mapped.full_name || 'ครู',
      mapped.phone || '',
      teacherRows,
      teacherMap,
      now,
      mapped.teacher_id || mapped.borrower_id
    );
    return {
      borrower_type: 'teacher',
      borrower_id: teacher.teacher_id,
      student_id: '0',
      teacher_id: teacher.teacher_id,
      borrower_name: teacher.full_name,
      created_teacher: teacher.created_teacher,
    };
  }

  const studentId = mapped.student_id || mapped.borrower_id;
  const student = studentMap[studentId] || {};
  return {
    borrower_type: 'student',
    borrower_id: studentId,
    student_id: studentId,
    teacher_id: '',
    borrower_name: student.full_name || mapped.full_name || '',
    created_teacher: false,
  };
}

function isTeacherBorrower(mapped) {
  const id = String(mapped.student_id || '').trim();
  const borrowerId = String(mapped.borrower_id || '').trim();
  const teacherId = String(mapped.teacher_id || '').trim();
  const type = String(mapped.borrower_type || '').trim().toLowerCase();
  const name = String(mapped.full_name || '').trim();
  const prefix = String(mapped.prefix || '').trim();
  return type === 'teacher' ||
    type === 'ครู' ||
    looksLikeTeacherId(teacherId) ||
    looksLikeTeacherId(borrowerId) ||
    id === '0' ||
    prefix === 'ครู' ||
    name.indexOf('ครู') === 0;
}

function getOrCreateTeacher(fullName, phone) {
  const rows = getRows(SHEETS.TEACHERS);
  const map = indexBy(rows, 'teacher_id');
  const teacher = getOrCreateTeacherInRows(fullName, phone, rows, map, nowText());
  if (teacher.created_teacher) rewriteObjects(SHEETS.TEACHERS, rows);
  return {
    borrower_type: 'teacher',
    borrower_id: teacher.teacher_id,
    student_id: '0',
    teacher_id: teacher.teacher_id,
    borrower_name: teacher.full_name,
  };
}

function getOrCreateTeacherInRows(fullName, phone, rows, map, now, preferredTeacherId) {
  const cleanName = normalizeTeacherName(fullName);
  const cleanNameKey = teacherNameKey(cleanName);
  const existingByName = rows.find((row) => teacherNameKey(row.full_name) === cleanNameKey);
  const teacherId = normalizeTeacherId(preferredTeacherId) || (existingByName && existingByName.teacher_id) || makeTeacherId(cleanName);
  if (map[teacherId]) {
    if (phone && !map[teacherId].phone) map[teacherId].phone = phone;
    if (cleanName && (!map[teacherId].full_name || map[teacherId].full_name === 'ครูไม่ระบุชื่อ')) map[teacherId].full_name = cleanName;
    map[teacherId].updated_at = now;
    map[teacherId].created_teacher = false;
    return map[teacherId];
  }
  if (existingByName) {
    if (phone && !existingByName.phone) existingByName.phone = phone;
    existingByName.updated_at = now;
    existingByName.created_teacher = false;
    map[existingByName.teacher_id] = existingByName;
    return existingByName;
  }

  const teacher = {
    teacher_id: teacherId,
    prefix: 'ครู',
    full_name: cleanName,
    phone: phone || '',
    created_at: now,
    updated_at: now,
    created_teacher: true,
  };
  rows.push(teacher);
  map[teacherId] = teacher;
  return teacher;
}

function normalizeTeacherName(fullName) {
  const text = String(fullName || '')
    .replace(/^ครูผู้สอน\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return 'ครูไม่ระบุชื่อ';
  return text.indexOf('ครู') === 0 ? text : 'ครู ' + text;
}

function teacherNameKey(fullName) {
  return String(fullName || '')
    .replace(/^ครูผู้สอน\s*/i, '')
    .replace(/^ครู\s*/i, '')
    .replace(/^(นาย|นาง|นางสาว|เด็กชาย|เด็กหญิง)\s*/i, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function makeTeacherId(fullName) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, fullName);
  const hex = digest.map((byte) => {
    const value = (byte + 256) % 256;
    return ('0' + value.toString(16)).slice(-2);
  }).join('').slice(0, 10).toUpperCase();
  return 'T' + hex;
}

function getMappedBorrowerId(mapped) {
  return String(mapped.teacher_id || mapped.borrower_id || mapped.student_id || '').trim();
}

function normalizeBorrowerType(value) {
  const text = String(value || '').trim().toLowerCase();
  if (!text) return '';
  if (text === 'teacher' || text === 'ครู') return 'teacher';
  if (text === 'student' || text === 'นักเรียน') return 'student';
  return text;
}

function normalizeTeacherId(value) {
  const text = String(value || '').trim().toUpperCase();
  if (!text) return '';
  if (text === '0') return '';
  return text;
}

function looksLikeTeacherId(value) {
  return /^T[0-9A-Z_-]+$/i.test(String(value || '').trim());
}

function getTransactionBorrowerType(row) {
  if (row.borrower_type) return String(row.borrower_type);
  return String(row.student_id) === '0' || row.teacher_id ? 'teacher' : 'student';
}

function getTransactionBorrowerKey(row) {
  const type = getTransactionBorrowerType(row);
  if (type === 'teacher') return 'teacher:' + String(row.teacher_id || row.borrower_id || row.student_id || '');
  return 'student:' + String(row.student_id || row.borrower_id || '');
}

function isBorrowingStatus(status) {
  return String(status || '').trim() === STATUS.BORROWING;
}

function normalizeDeviceStatus(status) {
  const text = String(status || '')
    .normalize('NFC')
    .trim()
    .replace(/\s+/g, '')
    .replace(/([\u0E47-\u0E4E])\1+/g, '$1');
  const knownStatuses = [
    STATUS.AVAILABLE,
    STATUS.BORROWED_DEVICE,
    STATUS.BORROWING,
    STATUS.REPAIR,
  ];
  const matched = knownStatuses.find((value) => text === String(value).replace(/\s+/g, ''));
  return matched || text;
}

function isTransactionDeviceStillBorrowed(row, devicesByKey) {
  const device = devicesByKey[row.device_key];
  if (!device) return false;
  const status = normalizeDeviceStatus(device.device_status);
  return status === STATUS.BORROWED_DEVICE || status === STATUS.BORROWING;
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index];
  });
  return obj;
}

function makeGetter(row) {
  return (...keys) => {
    for (let i = 0; i < keys.length; i++) {
      const value = getRawValue(row, keys[i]);
      if (value !== undefined && value !== null && String(value).trim() !== '') return String(value).trim();
    }
    return '';
  };
}

function getRawValue(row, ...keys) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
    const foundKey = Object.keys(row).find((candidate) => normalizeHeader(candidate) === normalizeHeader(key));
    if (foundKey) return row[foundKey];
  }
  return '';
}

function normalizeHeader(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function normalizeSearchText(value) {
  return String(value || '').replace(/\s+/g, '').trim().toLowerCase();
}

function listClasses() {
  const classes = getRows(SHEETS.STUDENTS)
    .map((row) => row.grade_level)
    .filter(Boolean);
  const hasTeachers = getRows(SHEETS.TEACHERS).length > 0 ||
    getRows(SHEETS.TRANSACTIONS).some((row) => getTransactionBorrowerType(row) === 'teacher');
  const result = Array.from(new Set(classes)).sort(naturalClassSort);
  if (hasTeachers) result.push('ครู');
  return result;
}

function listGradeGroups() {
  const groups = {};
  getRows(SHEETS.STUDENTS).forEach((student) => {
    const gradeLevel = String(student.grade_level || '').trim();
    const gradePrefix = getGradePrefix(gradeLevel);
    if (!gradePrefix) return;
    if (!groups[gradePrefix]) {
      groups[gradePrefix] = {
        grade_prefix: gradePrefix,
        student_count: 0,
        rooms: new Set(),
      };
    }
    groups[gradePrefix].student_count++;
    if (gradeLevel) groups[gradePrefix].rooms.add(gradeLevel);
  });

  return Object.keys(groups)
    .sort(naturalClassSort)
    .map((key) => ({
      grade_prefix: groups[key].grade_prefix,
      student_count: groups[key].student_count,
      room_count: groups[key].rooms.size,
    }));
}

function listUnborrowedStudentsByGrade(data) {
  const gradePrefix = required(data.grade_prefix, 'ระดับชั้น');
  const inventory = getSynchronizedInventory(10000);
  const activeStudentIds = new Set(inventory.transactions
    .filter((row) => isBorrowingStatus(row.status) && getTransactionBorrowerType(row) === 'student')
    .map((row) => String(row.student_id || row.borrower_id || '').trim())
    .filter(Boolean));
  const seenStudentIds = new Set();
  const grouped = {};

  getRows(SHEETS.STUDENTS)
    .filter((student) => getGradePrefix(student.grade_level) === gradePrefix)
    .filter((student) => {
      const studentId = String(student.student_id || '').trim();
      if (!studentId || activeStudentIds.has(studentId) || seenStudentIds.has(studentId)) return false;
      seenStudentIds.add(studentId);
      return true;
    })
    .sort((a, b) => {
      const classCompare = naturalClassSort(a.grade_level, b.grade_level);
      if (classCompare !== 0) return classCompare;
      const numberCompare = Number(a.student_no || 9999) - Number(b.student_no || 9999);
      return numberCompare || String(a.student_id).localeCompare(String(b.student_id), 'th', { numeric: true });
    })
    .forEach((student) => {
      const gradeLevel = String(student.grade_level || gradePrefix).trim();
      if (!grouped[gradeLevel]) grouped[gradeLevel] = [];
      grouped[gradeLevel].push({
        student_id: student.student_id,
        full_name: student.full_name,
        grade_level: gradeLevel,
        student_no: student.student_no || '',
      });
    });

  const rooms = Object.keys(grouped)
    .sort(naturalClassSort)
    .map((gradeLevel) => ({
      grade_level: gradeLevel,
      student_count: grouped[gradeLevel].length,
      students: grouped[gradeLevel],
    }));

  return {
    grade_prefix: gradePrefix,
    generated_at: nowText(),
    total_students: rooms.reduce((sum, room) => sum + room.student_count, 0),
    room_count: rooms.length,
    rooms,
  };
}

function getGradePrefix(value) {
  const text = String(value || '').replace(/\s+/g, '').trim();
  if (!text) return '';
  const match = text.match(/^(.+?\d+)(?:[/\\-]|ห้อง)/i);
  return match ? match[1] : text;
}

function listAvailableDevices() {
  return getSynchronizedInventory(10000).chromebooks
    .filter((row) => normalizeDeviceStatus(row.device_status) === STATUS.AVAILABLE)
    .map((row) => ({ device_key: row.device_key, asset_no: row.asset_no }));
}

function listAvailableDeviceReport() {
  const devices = listAvailableDevices()
    .sort((a, b) => String(a.device_key).localeCompare(String(b.device_key), 'th', { numeric: true, sensitivity: 'base' }))
    .map((row, index) => ({
      no: index + 1,
      device_key: row.device_key,
      asset_no: row.asset_no,
      device_status: STATUS.AVAILABLE,
      assign_to: '',
      note: '',
    }));

  return {
    generated_at: nowText(),
    total_available: devices.length,
    devices,
  };
}

function createBorrowRequest(data) {
  const mapped = mapBorrowRequest(data);
  if (!mapped.student_id) throw new Error('กรุณาระบุเลขนักเรียน');

  const now = nowText();
  const student = findRowByValue(SHEETS.STUDENTS, 'student_id', mapped.student_id);
  const source = student ? student.row : {};
  if (!mapped.full_name && !source.full_name) throw new Error('กรุณาระบุชื่อนักเรียน');
  const row = {
    request_id: Utilities.getUuid(),
    citizen_id: mapped.citizen_id || source.citizen_id || '',
    student_id: mapped.student_id,
    full_name: mapped.full_name || source.full_name || '',
    parent_name: mapped.parent_name || '',
    grade_level: mapped.grade_level || source.grade_level || '',
    phone: mapped.phone || source.phone || '',
    house_no: mapped.house_no || source.house_no || '',
    village_no: mapped.village_no || source.village_no || '',
    subdistrict: mapped.subdistrict || source.subdistrict || '',
    district: mapped.district || source.district || '',
    province: mapped.province || source.province || '',
    address: mapped.address || source.address || buildAddress(mapped),
    request_status: STATUS.REQUEST_PENDING,
    note: mapped.note || '',
    created_at: now,
    updated_at: now,
  };

  appendObjects(SHEETS.BORROW_REQUESTS, [row]);
  return { message: 'บันทึกคำขอยืมสำเร็จ', request_id: row.request_id };
}

function listBorrowRequests(data) {
  const status = String(data.status || '').trim();
  const q = normalizeSearchText(data.query || '');
  const limit = Math.min(Number(data.limit || 200), 500);
  let rows = getRows(SHEETS.BORROW_REQUESTS)
    .slice()
    .reverse();

  if (status) rows = rows.filter((row) => String(row.request_status || '') === status);
  if (q) {
    rows = rows.filter((row) => normalizeSearchText([
      row.citizen_id,
      row.student_id,
      row.full_name,
      row.parent_name,
      row.grade_level,
      row.phone,
      row.address,
      row.request_status,
      row.note,
    ].join(' ')).indexOf(q) !== -1);
  }

  return rows.slice(0, limit);
}

function updateBorrowRequestStatus(data) {
  const requestId = required(data.request_id, 'รหัสคำขอ');
  const status = required(data.request_status || data.status, 'สถานะคำขอ');
  const allowed = [
    STATUS.REQUEST_PENDING,
    STATUS.REQUEST_APPROVED,
    STATUS.REQUEST_REJECTED,
    STATUS.REQUEST_CANCELLED,
  ];
  if (allowed.indexOf(status) === -1) throw new Error('สถานะคำขอไม่ถูกต้อง');

  const found = findRowByValue(SHEETS.BORROW_REQUESTS, 'request_id', requestId);
  if (!found) throw new Error('ไม่พบคำขอที่เลือก');

  updateRowByIndex(SHEETS.BORROW_REQUESTS, found.index, {
    request_status: status,
    note: data.note !== undefined ? data.note : found.row.note,
    updated_at: nowText(),
  });

  return { message: 'อัปเดตสถานะคำขอสำเร็จ', request_id: requestId, request_status: status };
}

function mapBorrowRequest(data) {
  const get = makeGetter(data || {});
  const houseNo = get('house_no', 'บ้านเลขที่', 'เลขที่บ้าน');
  const villageNo = get('village_no', 'หมู่', 'หมู่บ้าน');
  const subdistrict = get('subdistrict', 'ตำบล');
  const district = get('district', 'อำเภอ');
  const province = get('province', 'จังหวัด');
  const mapped = {
    citizen_id: get('citizen_id', 'เลขบัตรประชาชน', 'เลขบัตรประชาชนนักเรียน'),
    student_id: get('student_id', 'เลขนักเรียน', 'รหัสนักเรียน', 'เลขประจำตัวนักเรียน'),
    full_name: get('full_name', 'ชื่อนักเรียน', 'ชื่อ-สกุล', 'ชื่อสกุล'),
    parent_name: get('parent_name', 'ชื่อผู้ปกครอง', 'ผู้ปกครอง', 'parent'),
    grade_level: get('grade_level', 'ชั้นห้อง', 'ชั้น+ห้อง', 'ระดับชั้น+ห้อง'),
    phone: get('phone', 'เบอร์โทร', 'เบอร์โทรศัพท์'),
    house_no: houseNo,
    village_no: villageNo,
    subdistrict,
    district,
    province,
    address: get('address', 'ที่อยู่'),
    note: get('note', 'หมายเหตุ'),
  };
  if (!mapped.address) mapped.address = buildAddress(mapped);
  return mapped;
}

function buildAddress(row) {
  return [
    row.house_no ? 'บ้านเลขที่ ' + row.house_no : '',
    row.village_no ? 'หมู่ ' + row.village_no : '',
    row.subdistrict,
    row.district,
    row.province,
  ].filter(Boolean).join(' ');
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && HEADERS[sheetName]) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS[sheetName].length)
      .setFontWeight('bold')
      .setBackground('#dbeafe');
    sheet.autoResizeColumns(1, HEADERS[sheetName].length);
  }
  if (!sheet) throw new Error('ไม่พบชีต ' + sheetName + ' กรุณารัน setupDatabase() ก่อน');
  ensureHeaders(sheet, sheetName);
  return sheet;
}

function ensureHeaders(sheet, sheetName) {
  const expected = HEADERS[sheetName];
  if (!expected || sheet.getLastRow() === 0) return;

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String);
  const missing = expected.filter((header) => current.indexOf(header) === -1);
  if (!missing.length) return;

  sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  sheet.getRange(1, 1, 1, current.length + missing.length)
    .setFontWeight('bold')
    .setBackground('#dbeafe');
}

function getRows(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter((row) => row.some((cell) => cell !== '')).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] instanceof Date ? formatDate(row[index]) : row[index];
    });
    return obj;
  });
}

function appendObjects(sheetName, objects) {
  if (!objects.length) return;
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = objects.map((obj) => headers.map((header) => obj[header] !== undefined ? obj[header] : ''));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function rewriteObjects(sheetName, objects) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (!objects.length) return;
  const values = objects.map((obj) => headers.map((header) => obj[header] !== undefined ? obj[header] : ''));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function findRowByValue(sheetName, key, value) {
  const rows = getRows(sheetName);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][key]) === String(value)) {
      return { row: rows[i], index: i + 2 };
    }
  }
  return null;
}

function updateRowByIndex(sheetName, rowIndex, patch) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Object.keys(patch).forEach((key) => {
    const col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowIndex, col + 1).setValue(patch[key]);
  });
}

function indexBy(rows, key) {
  return rows.reduce((acc, row) => {
    if (row[key] !== undefined && row[key] !== '') acc[String(row[key])] = row;
    return acc;
  }, {});
}

function compactObject(obj) {
  Object.keys(obj).forEach((key) => {
    if (obj[key] === undefined || obj[key] === null || obj[key] === '') delete obj[key];
  });
  return obj;
}

function required(value, label) {
  const text = String(value || '').trim();
  if (!text) throw new Error('กรุณาระบุ' + label);
  return text;
}

function nowText() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function todayText() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeExcelDate(value) {
  if (!value) return '';
  if (value instanceof Date) return formatDate(value);
  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).trim();
}

function naturalClassSort(a, b) {
  return String(a).localeCompare(String(b), 'th', { numeric: true, sensitivity: 'base' });
}
