const TERM_AUDIT_COLUMNS = [
  ['academic_year', 'ปีการศึกษา'], ['term', 'เทอม'], ['student_id', 'รหัสนักเรียน'],
  ['full_name', 'ชื่อ-สกุล'], ['grade_level', 'ชั้น+ห้อง'], ['transaction_id', 'รหัสรายการยืม'],
  ['device_key', 'รหัสเครื่อง'], ['asset_no', 'เลขที่ทรัพย์สิน'], ['checked_on', 'วันที่ตรวจ'],
  ['inspector', 'ผู้ตรวจ'], ['device_result', 'ผลตรวจเครื่อง'], ['pen', 'ปากกา'],
  ['pen_charger', 'ที่ชาร์จปากกา'], ['charger', 'สายชาร์จ / ที่ชาร์จเครื่อง'], ['note', 'หมายเหตุ'],
];
let termAuditRows = [];
let termAuditGeneration = 0;
const auditElement = (id) => document.getElementById('termAudit' + id);
const hasTermAuditEntry = (row) => ['checked_on', 'inspector', 'device_result', 'pen', 'pen_charger', 'charger', 'note'].some((key) => String(row[key] || '').trim());

async function loadTermAuditGrades() {
  const select = auditElement('Grade');
  if (select.options.length > 1) return;
  try {
    const groups = await api('listGradeGroups');
    select.innerHTML = '<option value="">เลือกระดับชั้น</option>' + groups.map((row) => `<option value="${escapeAttr(row.grade_prefix)}">${escapeHtml(row.grade_prefix)}</option>`).join('');
  } catch (error) { auditElement('Meta').textContent = error.message; }
}

function makeTermAuditWorkbook(report, year, term) {
  const workbook = XLSX.utils.book_new();
  const used = new Set();
  report.rooms.forEach((room) => {
    const rows = room.students.map((student) => TERM_AUDIT_COLUMNS.map(([key]) => {
      if (key === 'academic_year') return year;
      if (key === 'term') return term;
      return String(student[key] || '');
    }).concat(student.audit_status || 'ยังไม่ได้ตรวจ'));
    const sheet = XLSX.utils.aoa_to_sheet([TERM_AUDIT_COLUMNS.map(([, label]) => label).concat('สถานะการตรวจ'), ...rows]);
    sheet['!cols'] = TERM_AUDIT_COLUMNS.map(([key]) => ({ wch: ['full_name', 'device_key', 'transaction_id', 'note'].includes(key) ? 34 : 22 }));
    sheet['!cols'].push({ wch: 22 });
    sheet['!autofilter'] = { ref: 'A1:P' + (rows.length + 1) };
    XLSX.utils.book_append_sheet(workbook, sheet, makeUniqueSheetName(room.grade_level, used));
  });
  const instructions = XLSX.utils.aoa_to_sheet([
    ['คำแนะนำการตรวจประจำเทอม'],
    ['ผลตรวจเดิมของปีและเทอมที่เลือกเติมให้แล้ว อัปโหลดข้อมูลเดิมซ้ำจะถูกข้าม'],
    ['สถานะ ต้องตรวจใหม่ หมายถึงรายการยืมเปลี่ยน ช่องผลตรวจจะเว้นว่างเพื่อให้ตรวจเครื่องปัจจุบัน'],
    ['กรอกเฉพาะวันที่ตรวจ ผู้ตรวจ ผลตรวจเครื่อง อุปกรณ์ และหมายเหตุ เก็บรหัสเดิมไว้'],
    ['วันที่ตรวจใช้ YYYY-MM-DD ค.ศ. เช่น 2026-10-06'],
    ['ผลตรวจเครื่อง: พบเครื่อง-ปกติ / พบเครื่อง-ชำรุด / ไม่พบเครื่อง / ไม่ได้ยืม'],
    ['อุปกรณ์แต่ละชิ้น: ครบ / ขาด / ชำรุด / ยังไม่ได้ตรวจ'],
    ['คนที่ไม่ได้ยืม: กรอกวันที่ ผู้ตรวจ และผลตรวจเครื่องเป็น ไม่ได้ยืม เว้นอุปกรณ์ว่าง'],
    ['แถวที่ยังไม่ได้กรอกผลตรวจจะถูกข้าม รวมหลายห้องในไฟล์เดียวได้ ไม่เกิน 2000 แถว'],
    ['อัปโหลดที่เมนูตรวจประจำเทอม ไม่ใช่นำเข้ารายการยืม'],
    ['ผลตรวจไม่เปลี่ยนสถานะยืม คืน หรือซ่อม อุปกรณ์ขาดหรือชำรุดจะแสดงในเมนูอุปกรณ์ค้างคืน'],
  ]);
  instructions['!cols'] = [{ wch: 110 }];
  XLSX.utils.book_append_sheet(workbook, instructions, 'คำแนะนำ');
  return workbook;
}

auditElement('Download').addEventListener('click', async () => {
  const year = auditElement('Year').value;
  const term = auditElement('Term').value;
  const grade = auditElement('Grade').value;
  if (!/^25\d{2}$/.test(year) || !grade) { auditElement('Meta').textContent = 'กรุณาเลือกปีการศึกษาและระดับชั้น'; return; }
  const button = auditElement('Download');
  setButtonBusy(button, true, 'กำลังสร้างไฟล์...');
  try {
    const report = await api('listAnnualStudentDeviceAudit', { grade_prefix: grade, academic_year: year, term, repair_token: state.admin && state.admin.repair_token });
    if (String(report.academic_year) !== year || String(report.term) !== term) throw new Error('กรุณาอัปเดต Code.gs และ Deploy เวอร์ชันใหม่เพื่อดาวน์โหลดผลตรวจเดิม');
    if (!report.total_students) throw new Error('ไม่พบรายชื่อนักเรียน');
    // Reject an old deployment that omits the loan identifier required for matching.
    if (report.rooms.some((room) => room.students.some((row) => row.device_key && !row.transaction_id))) throw new Error('กรุณาอัปเดต Code.gs และ Deploy เวอร์ชันใหม่');
    XLSX.writeFile(makeTermAuditWorkbook(report, year, term), `ตรวจเทอม-${year}-${term}-${grade.replace(/[\\/:*?"<>|]/g, '-')}.xlsx`);
    auditElement('Meta').textContent = `ดาวน์โหลด ${report.total_students} คน แยก ${report.room_count} ห้องแล้ว`;
  } catch (error) { auditElement('Meta').textContent = error.message; }
  finally { setButtonBusy(button, false); }
});

function parseTermAuditWorkbook(workbook) {
  const rows = [];
  workbook.SheetNames.forEach((name) => {
    if (name === 'คำแนะนำ') return;
    const grid = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', blankrows: true });
    if (!grid.length || !grid.some((row) => row.some((cell) => String(cell).trim()))) return;
    const headers = grid[0].map((cell) => String(cell).trim());
    if (TERM_AUDIT_COLUMNS.some(([, label]) => headers.filter((header) => header === label).length !== 1)) throw new Error(`ชีต ${name}: หัวตารางไม่ตรง กรุณาใช้ไฟล์ตรวจประจำเทอมใหม่`);
    grid.slice(1).forEach((values, index) => {
      if (!values.some((cell) => String(cell).trim())) return;
      const row = { source_sheet: name, source_row: index + 2 };
      TERM_AUDIT_COLUMNS.forEach(([key, label]) => {
        let value = values[headers.indexOf(label)] ?? '';
        if (key === 'checked_on' && typeof value === 'number') {
          const date = XLSX.SSF.parse_date_code(value, { date1904: Boolean(workbook.Workbook?.WBProps?.date1904) });
          value = date ? `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}` : String(value);
        }
        row[key] = String(value).trim();
      });
      rows.push(row);
    });
  });
  if (!rows.length || rows.length > 2000) throw new Error('รองรับไฟล์ 1–2000 แถวต่อครั้ง');
  return rows;
}

auditElement('File').addEventListener('change', async (event) => {
  const generation = ++termAuditGeneration;
  termAuditRows = [];
  auditElement('Save').disabled = true;
  auditElement('Validate').disabled = true;
  auditElement('Preview').innerHTML = '';
  auditElement('Meta').textContent = '';
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (file.size > 10 * 1024 * 1024) throw new Error('ไฟล์ต้องมีขนาดไม่เกิน 10 MB');
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const rows = parseTermAuditWorkbook(workbook);
    if (generation !== termAuditGeneration) return;
    termAuditRows = rows;
    auditElement('Validate').disabled = false;
    auditElement('Meta').textContent = `อ่าน ${rows.length} แถว · กรอกผลตรวจ ${rows.filter(hasTermAuditEntry).length} คน`;
  } catch (error) { if (generation === termAuditGeneration) auditElement('Meta').textContent = error.message; }
});

async function submitTermAudit(write) {
  if (!termAuditRows.length) return;
  const generation = termAuditGeneration;
  auditElement('File').disabled = true;
  auditElement('Save').disabled = true;
  auditElement('Validate').disabled = true;
  auditElement('Meta').textContent = write ? 'กำลังบันทึกผลตรวจ...' : 'กำลังตรวจสอบไฟล์...';
  try {
    if (!write) {
      const file = auditElement('File').files[0];
      if (!file) throw new Error('กรุณาเลือกไฟล์ผลตรวจอีกครั้ง');
      termAuditRows = parseTermAuditWorkbook(XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false }));
    }
    const entered = termAuditRows.filter(hasTermAuditEntry);
    if (!entered.length) throw new Error('ไม่พบผลตรวจในไฟล์ที่เลือก กรุณาบันทึกไฟล์ Excel แล้วเลือกไฟล์ที่บันทึกล่าสุดอีกครั้ง');
    const result = await api(write ? 'importTermAudit' : 'validateTermAudit', { rows: entered, repair_token: state.admin && state.admin.repair_token });
    if (generation !== termAuditGeneration) return;
    if (!Array.isArray(result.results) || result.results.length !== entered.length) throw new Error('คำตอบจาก Apps Script ไม่ตรงกับไฟล์ กรุณาตรวจว่าเว็บใช้ Deployment ของ Code.gs เวอร์ชันล่าสุด');
    const ready = result.results.filter((row) => row.status === 'ready').length;
    const errors = result.results.filter((row) => row.status === 'error').length;
    const blank = result.results.filter((row) => row.message === 'ยังไม่ได้กรอกผลตรวจ').length;
    auditElement('Preview').innerHTML = result.results.slice().sort((a, b) => Number(b.status === 'error') - Number(a.status === 'error')).map((row) => `<tr><td>${escapeHtml(row.source_sheet)} / ${row.source_row}</td><td>${escapeHtml(row.student_id)}</td><td>${escapeHtml(write && !errors && row.status === 'ready' ? 'บันทึกแล้ว' : row.message)}</td></tr>`).join('');
    if (blank) throw new Error(`หน้าเว็บอ่านผลตรวจได้ ${entered.length} คน แต่ Apps Script แจ้งว่าว่าง กรุณาอัปเดต Code.gs และ Deploy เวอร์ชันใหม่`);
    auditElement('Meta').textContent = errors ? `กรอก ${entered.length} คน · พบ ${errors} แถวที่ต้องแก้ไข ยังไม่มีการบันทึก` : write ? `บันทึก ${result.saved_count} รายการแล้ว` : ready ? `พร้อมบันทึก ${ready} รายการ · ข้ามแถวที่ยังไม่ได้กรอก ${termAuditRows.length - entered.length} แถว` : `ผลตรวจ ${entered.length} คนบันทึกไว้แล้ว ไม่มีข้อมูลใหม่ให้บันทึก`;
    auditElement('Save').disabled = write || errors > 0 || !ready;
    if (write && result.saved_count > 0) {
      localStorage.removeItem(DASHBOARD_CACHE_KEY);
      await loadPublicDashboard();
      await loadTermAuditProgress();
    }
  } catch (error) { auditElement('Meta').textContent = error.message; }
  finally { auditElement('File').disabled = false; auditElement('Validate').disabled = false; }
}
auditElement('Validate').addEventListener('click', () => submitTermAudit(false));
auditElement('Save').addEventListener('click', () => submitTermAudit(true));

let termProgressReport = null;
let termProgressRequest = 0;
const progressElement = (id) => document.getElementById('termProgress' + id);

async function loadTermAuditProgress() {
  const request = ++termProgressRequest;
  termProgressReport = null;
  progressElement('Rows').innerHTML = '';
  progressElement('Meta').textContent = 'กำลังโหลดผลตรวจ...';
  try {
    const report = await api('listTermAuditProgress', {
      academic_year: auditElement('Year').value, term: auditElement('Term').value,
      grade_prefix: auditElement('Grade').value, repair_token: state.admin && state.admin.repair_token,
    });
    if (request !== termProgressRequest || !state.admin) return;
    termProgressReport = report;
    const selected = progressElement('Room').value;
    const rooms = Array.from(new Set(report.rows.map((row) => row.grade_level).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'th', { numeric: true }));
    progressElement('Room').innerHTML = '<option value="">ทุกห้อง</option>' + rooms.map((room) => `<option value="${escapeAttr(room)}">${escapeHtml(room)}</option>`).join('');
    if (rooms.includes(selected)) progressElement('Room').value = selected;
    renderTermAuditProgress();
  } catch (error) { if (request === termProgressRequest) progressElement('Meta').textContent = error.message; }
}

function renderTermAuditProgress() {
  if (!termProgressReport) return;
  const room = progressElement('Room').value;
  const status = progressElement('Status').value;
  const query = progressElement('Query').value.trim().toLowerCase();
  const scoped = termProgressReport.rows.filter((row) => !room || row.grade_level === room);
  const rows = scoped.filter((row) => (!status || row.audit_status === status) && [row.student_id, row.full_name, row.asset_no, row.device_key].join(' ').toLowerCase().includes(query));
  const count = (value) => scoped.filter((row) => row.audit_status === value).length;
  progressElement('Meta').textContent = `ปี ${termProgressReport.academic_year} เทอม ${termProgressReport.term} · ${room || 'ทุกห้องในระดับชั้นที่เลือก'} ${scoped.length} คน · ตรวจแล้ว ${count('ตรวจแล้ว')} · ยังไม่ได้ตรวจ ${count('ยังไม่ได้ตรวจ')} · ต้องตรวจใหม่ ${count('ต้องตรวจใหม่')} · แสดง ${rows.length} คน`;
  progressElement('Rows').innerHTML = rows.length ? rows.map((row) => `<tr>
    <td><strong>${escapeHtml(row.audit_status)}</strong>${row.audit_status === 'ต้องตรวจใหม่' ? '<br>รายการยืมเปลี่ยนหลังตรวจ' : ''}</td>
    <td>${escapeHtml(row.student_id)}<br>${escapeHtml(row.full_name)}</td><td>${escapeHtml(row.grade_level)}</td>
    <td>${escapeHtml(row.borrow_status)}<br>${escapeHtml(row.asset_no || '-')}<br>${escapeHtml(row.device_key || '')}</td>
    <td>${escapeHtml(row.checked_on || '-')}<br>${escapeHtml(row.inspector || '')}</td>
    <td>${escapeHtml(row.device_result || '-')}<br>${Object.entries(ACCESSORY_LABELS).map(([key, label]) => `${label}: ${escapeHtml(row[key] || '-')}`).join('<br>')}${row.note ? `<br>${escapeHtml(row.note)}` : ''}</td>
  </tr>`).join('') : '<tr><td colspan="6">ไม่พบรายชื่อตามเงื่อนไข</td></tr>';
}
['Year', 'Term', 'Grade'].forEach((id) => auditElement(id).addEventListener('change', loadTermAuditProgress));
['Room', 'Status'].forEach((id) => progressElement(id).addEventListener('change', renderTermAuditProgress));
progressElement('Query').addEventListener('input', renderTermAuditProgress);
progressElement('Reload').addEventListener('click', loadTermAuditProgress);
