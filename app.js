'use strict';
const $ = id => document.getElementById(id);
let selected = null;
let importedNames = [];
let importedStudents = [];
let selectedDesks = [];
let randomAssignments = [];
let animationToken = 0;
const seats = [];
const desks = [];

for (let d = 0; d < 24; d++) {
  const desk = document.createElement('div');
  desk.className = 'desk';
  desks.push(desk);
  for (let s = 0; s < 2; s++) {
    const i = d * 2 + s;
    const el = document.createElement('textarea');
    el.className = 'seat';
    el.maxLength = 70;
    el.setAttribute('aria-label', `Hàng ${Math.floor(d / 4) + 1}, tổ ${d % 4 + 1}, chỗ ${s + 1}`);
    el.addEventListener('input', updateCount);
    el.addEventListener('click', () => {
      if (!$('swap').checked) return;
      if (selected === null) { selected = i; el.classList.add('selected') }
      else {
        [seats[selected].value, el.value] = [el.value, seats[selected].value];
        resetSelection(); updateCount(); syncCurrentResults();
      }
    });
    desk.append(el); seats.push(el);
  }
  $('desks').append(desk);
}

function resetSelection() { selected = null; seats.forEach(s => s.classList.remove('selected')) }
function updateCount() { $('occupied').textContent = `${seats.filter(s => s.value.trim()).length} / 48 chỗ` }
function updateHeader() {
  $('title').textContent = `SƠ ĐỒ LỚP ${$('className').value.trim() || '…'} NĂM HỌC: ${$('year').value.trim() || '…'}`;
  $('subtitle').textContent = `GVCN: ${$('teacher').value.trim() || '…'}`;
}
['teacher', 'className', 'year'].forEach(id => $(id).addEventListener('input', updateHeader));
$('swap').onchange = () => { resetSelection(); seats.forEach(s => s.readOnly = $('swap').checked) };
$('clear').onclick = () => {
  if (seats.some(s => s.value) && confirm('Xóa toàn bộ tên trên sơ đồ?')) {
    seats.forEach(s => s.value = ''); resetArrangementState(); randomAssignments = []; $('export-random').disabled = true;
    resetSelection(); updateCount();
  }
};
$('print').onclick = () => window.print();

// Mẫu: cột A là Họ và tên, cột B (tùy chọn) là Số bàn muốn gán cố định khi random.
$('template').onclick = () => {
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; return; }
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Họ và tên', 'Số bàn cố định (tùy chọn)'], ...Array.from({ length: 48 }, () => ['', ''])]);
  sheet['!cols'] = [{ wch: 38 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(book, sheet, 'Danh sách');
  XLSX.writeFile(book, 'danh-sach-hoc-sinh.xlsx');
};
function extractFixedDesk(row) {
  const colA = String(row[0] ?? '').trim();
  // 1. Kiểm tra cột B (row[1]) nếu có số
  if (row[1] !== undefined && row[1] !== null && String(row[1]).trim() !== '') {
    const m = String(row[1]).trim().match(/\d+/);
    if (m) {
      const num = parseInt(m[0], 10);
      if (num > 0) return { name: colA, fixedDesk: num };
    }
  }
  // 2. Nếu cột B trống, kiểm tra xem ở đuôi cột A có ghi số không (ví dụ: "Nguyễn Văn A 5", "Nguyễn Văn A (bàn 5)", "Nguyễn Văn A - 5")
  const match = colA.match(/^(.*?)\s*[-–(,]?\s*(?:bàn\s*)?(\d+)\s*[)]?$/i);
  if (match && match[1].trim() && match[2]) {
    const num = parseInt(match[2], 10);
    if (num > 0) return { name: match[1].trim(), fixedDesk: num };
  }
  return { name: colA, fixedDesk: null };
}
function readStudentData(book) {
  const sheet = book.Sheets[book.SheetNames[0]];
  const a1 = String(sheet?.A1?.v || '').normalize('NFC').trim().toLocaleLowerCase('vi');
  if (!sheet || (a1 !== 'họ và tên' && a1 !== 'họ tên'))
    throw Error('File chưa đúng mẫu. Hãy tải Excel mẫu và điền tên vào cột Họ và tên.');
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, range: 0 });
  const students = [];
  const deskCountMap = {};
  for (const row of rows.slice(1)) {
    const { name, fixedDesk } = extractFixedDesk(row);
    if (!name) continue;
    if (name.length > 70) throw Error(`Tên "${name}" dài hơn 70 ký tự. Vui lòng kiểm tra lại.`);
    if (fixedDesk) {
      if (fixedDesk < 1 || fixedDesk > 24)
        throw Error(`Học sinh "${name}" có số bàn ${fixedDesk} không hợp lệ (sơ đồ có bàn 1 đến 24).`);
      deskCountMap[fixedDesk] = (deskCountMap[fixedDesk] || 0) + 1;
      if (deskCountMap[fixedDesk] > 2)
        throw Error(`Bàn ${fixedDesk} được gán cho nhiều hơn 2 học sinh. Mỗi bàn chỉ có tối đa 2 chỗ.`);
    }
    students.push({ name, fixedDesk });
  }
  if (students.length > 48) throw Error(`File có ${students.length} học sinh. Sơ đồ chỉ có tối đa 48 chỗ.`);
  if (!students.length) throw Error('File chưa có tên học sinh.');
  return students;
}
$('file').onchange = async () => {
  const input = $('file'), file = input.files[0]; if (!file) return;
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; input.value = ''; return; }
  input.disabled = true; $('status').textContent = 'Đang đọc file…';
  try {
    if (file.size > 10 * 1024 * 1024) throw Error('Vui lòng chọn file nhỏ hơn 10 MB.');
    let book;
    try { book = XLSX.read(await file.arrayBuffer(), { type: 'array' }); }
    catch { throw Error('Không đọc được file. Vui lòng dùng Excel mẫu.'); }
    importedStudents = readStudentData(book);
    importedNames = importedStudents.map(s => s.name);
    randomAssignments = []; selectedDesks = [];
    seats.forEach((seat, i) => seat.value = importedNames[i] || '');
    resetArrangementState();
    $('random').disabled = false; $('export-random').disabled = true;
    resetSelection(); updateCount();
    const fixedCount = importedStudents.filter(s => s.fixedDesk).length;
    $('status').textContent = `Đã nhận ${importedStudents.length} học sinh${fixedCount ? ` (có ${fixedCount} học sinh đã được nhận diện)` : ''}.`;
  } catch (error) { $('status').textContent = error.message; }
  finally { input.disabled = false; input.value = ''; }
};

// Chọn bàn theo thứ tự bấm. Thứ tự này chính là số bàn hiển thị khi quay random.
const pickerButtons = [];
for (let d = 0; d < 24; d++) {
  const button = document.createElement('button');
  button.type = 'button'; button.className = 'picker-desk';
  button.dataset.desk = String(d);
  button.setAttribute('aria-label', `Hàng ${Math.floor(d / 4) + 1}, Tổ ${d % 4 + 1}`);
  button.onclick = () => toggleDesk(d);
  pickerButtons.push(button); $('desk-picker').append(button);
}
// Sơ đồ thu nhỏ để nhìn tên học sinh bay vào đúng bàn trong lúc random.
const liveSection = document.createElement('section');
liveSection.className = 'random-live-section';
const liveTitle = document.createElement('h3'); liveTitle.textContent = 'Sơ đồ đang xếp chỗ';
const liveHeadings = document.createElement('div'); liveHeadings.className = 'live-headings';
['Tổ 1', 'Tổ 2', 'Tổ 3', 'Tổ 4'].forEach(label => { const span = document.createElement('span'); span.textContent = label; liveHeadings.append(span) });
const liveLayout = document.createElement('div'); liveLayout.className = 'random-live-layout';
const liveDesks = [];
for (let d = 0; d < 24; d++) {
  const desk = document.createElement('div'); desk.className = 'live-desk';
  const badge = document.createElement('b'); badge.className = 'live-desk-number';
  const left = document.createElement('span'), right = document.createElement('span'); left.className = right.className = 'live-seat';
  desk.append(badge, left, right); liveDesks.push(desk); liveLayout.append(desk);
}
liveSection.append(liveTitle, liveHeadings, liveLayout);
$('random-run-view').insertBefore(liveSection, $('assignment-list'));

function physicalDeskName(d) { return `Tổ ${d % 4 + 1} · hàng ${Math.floor(d / 4) + 1}` }
function refreshPicker() {
  pickerButtons.forEach((button, d) => {
    const order = selectedDesks.indexOf(d);
    button.classList.toggle('chosen', order >= 0);
    button.innerHTML = order >= 0 ? `<small>Bàn</small><strong>${order + 1}</strong>` : `<small>${Math.floor(d / 4) + 1}</small><span>Chọn</span>`;
    button.title = physicalDeskName(d);
  });
  const places = selectedDesks.length * 2;
  $('selected-count').textContent = selectedDesks.length ? `Đã chọn ${selectedDesks.length} bàn · ${places} chỗ` : 'Chưa chọn bàn';
  $('required-count').textContent = `Cần ${importedNames.length} chỗ`;
  $('random-error').textContent = '';
}
function toggleDesk(d) {
  const i = selectedDesks.indexOf(d);
  if (i >= 0) selectedDesks.splice(i, 1); else selectedDesks.push(d);
  refreshPicker();
}
function openRandom() {
  if (!importedNames.length) { $('status').textContent = 'Hãy tải danh sách Excel trước khi random.'; return; }
  selectedDesks = []; refreshPicker();
  $('random-select-view').hidden = false; $('random-run-view').hidden = true;
  $('random-modal').hidden = false; document.body.classList.add('modal-open');
}
function closeRandom() {
  animationToken++; $('random-modal').hidden = true; document.body.classList.remove('modal-open');
}
$('random').onclick = openRandom;
$('random-close').onclick = closeRandom;
$('random-modal').onclick = e => { if (e.target === $('random-modal') && !$('random-select-view').hidden) closeRandom() };
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('random-modal').hidden && !$('random-select-view').hidden) closeRandom() });
$('select-all').onclick = () => { selectedDesks = Array.from({ length: 24 }, (_, i) => i); refreshPicker() };

function shuffle(values) {
  const out = values.slice();
  for (let i = out.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1));[out[i], out[j]] = [out[j], out[i]] }
  return out;
}
function makeAssignments(studentsList) {
  const students = (studentsList && studentsList.length)
    ? studentsList
    : ((importedStudents && importedStudents.length === importedNames.length)
      ? importedStudents
      : importedNames.map(name => ({ name, fixedDesk: null })));

  // Tạo kho 2 chỗ ngồi cho mỗi bàn được chọn
  const deskSlotsMap = new Map();
  selectedDesks.forEach((desk, deskOrder) => {
    const deskNumber = deskOrder + 1;
    deskSlotsMap.set(deskNumber, [
      { seat: desk * 2, deskNumber, physicalDesk: desk },
      { seat: desk * 2 + 1, deskNumber, physicalDesk: desk }
    ]);
  });

  const assignmentMap = new Map();
  const unassignedStudents = [];

  // 1. Ưu tiên xếp các học sinh có số bàn gán cố định trước
  students.forEach(student => {
    if (student.fixedDesk && deskSlotsMap.has(student.fixedDesk)) {
      const available = deskSlotsMap.get(student.fixedDesk);
      if (available && available.length > 0) {
        const idx = Math.floor(Math.random() * available.length);
        assignmentMap.set(student, available.splice(idx, 1)[0]);
        return;
      }
    }
    unassignedStudents.push(student);
  });

  // 2. Gom toàn bộ chỗ ngồi còn trống từ các bàn
  const remainingSlots = [];
  deskSlotsMap.forEach(slots => remainingSlots.push(...slots));

  // 3. Xáo trộn ngẫu nhiên các chỗ ngồi còn lại
  const shuffledSlots = shuffle(remainingSlots);

  // 4. Phân bổ các học sinh còn lại vào các chỗ ngẫu nhiên
  unassignedStudents.forEach((student, idx) => {
    assignmentMap.set(student, shuffledSlots[idx]);
  });

  // 5. Trả về kết quả theo đúng thứ tự ban đầu của danh sách học sinh
  return students.map(student => {
    const slot = assignmentMap.get(student);
    return { name: student.name, fixedDesk: student.fixedDesk, ...slot };
  });
}
function prepareLiveLayout() {
  liveDesks.forEach((desk, d) => {
    const order = selectedDesks.indexOf(d);
    desk.classList.toggle('active', order >= 0); desk.classList.remove('complete-celebration');
    desk.querySelector('.live-desk-number').textContent = order >= 0 ? `Bàn ${order + 1}` : '';
    desk.querySelectorAll('.live-seat').forEach(seat => { seat.textContent = ''; seat.classList.remove('filled') });
  });
}
function placeNameInLiveDesk(item) {
  const slot = liveDesks[item.physicalDesk].querySelectorAll('.live-seat')[item.seat % 2];
  slot.textContent = item.name; slot.title = item.name; slot.classList.add('filled');
}
async function flyNameToDesk(item, token) {
  const source = $('draw-name'), target = liveDesks[item.physicalDesk], targetSlot = target.querySelectorAll('.live-seat')[item.seat % 2];
  const from = source.getBoundingClientRect(), to = targetSlot.getBoundingClientRect();
  const flyer = document.createElement('div'); flyer.className = 'flying-student'; flyer.textContent = item.name;
  flyer.style.left = `${from.left}px`; flyer.style.top = `${from.top}px`; flyer.style.width = `${Math.min(Math.max(from.width, 170), 300)}px`;
  document.body.append(flyer);
  const dx = to.left + to.width / 2 - (from.left + Math.min(Math.max(from.width, 170), 300) / 2);
  const dy = to.top + to.height / 2 - (from.top + from.height / 2);
  const keyframes = [
    { transform: 'translate(0,0) scale(1)', opacity: 1, offset: 0 },
    { transform: `translate(${dx * .48}px,${dy * .48 - 85}px) scale(1.08)`, opacity: 1, offset: .52 },
    { transform: `translate(${dx}px,${dy}px) scale(.3)`, opacity: .15, offset: 1 }
  ];
  if (flyer.animate) {
    const flight = flyer.animate(keyframes, { duration: 900, easing: 'cubic-bezier(.2,.75,.3,1)', fill: 'forwards' });
    try { await flight.finished } catch { }
  } else await wait(900);
  flyer.remove();
  if (token !== animationToken) return false;
  placeNameInLiveDesk(item);
  const full = [...target.querySelectorAll('.live-seat')].every(seat => seat.textContent.trim());
  if (full) {
    target.classList.add('complete-celebration'); desks[item.physicalDesk].classList.add('desk-complete');
    await wait(1000);
    target.classList.remove('complete-celebration'); desks[item.physicalDesk].classList.remove('desk-complete');
  }
  return token === animationToken;
}
$('start-random').onclick = () => {
  if (selectedDesks.length * 2 < importedNames.length) {
    const missing = importedNames.length - selectedDesks.length * 2;
    $('random-error').textContent = `Thiếu ${missing} chỗ ngồi. Hãy chọn thêm ${Math.ceil(missing / 2)} bàn.`;
    return;
  }
  if (!selectedDesks.length) { $('random-error').textContent = 'Hãy chọn ít nhất một bàn.'; return; }

  const students = (importedStudents && importedStudents.length === importedNames.length)
    ? importedStudents
    : importedNames.map(name => ({ name, fixedDesk: null }));

  for (const s of students) {
    if (s.fixedDesk && s.fixedDesk > selectedDesks.length) {
      $('random-error').textContent = `Học sinh "${s.name}" được gán Bàn ${s.fixedDesk}, nhưng bạn chỉ mới chọn ${selectedDesks.length} bàn. Vui lòng chọn thêm bàn để có Bàn ${s.fixedDesk}.`;
      return;
    }
  }

  $('random-close').disabled = true;
  $('export-random').disabled = true;
  clearArrangeHistory();
  randomAssignments = makeAssignments(students);
  seats.forEach(seat => seat.value = ''); updateCount();
  prepareLiveLayout();
  $('random-select-view').hidden = true; $('random-run-view').hidden = false;
  $('assignment-list').replaceChildren(); $('finish-random').hidden = true; $('skip-animation').hidden = false;
  $('random-progress').style.width = '0%'; $('progress-text').textContent = '';
  runAnimation();
};
function wait(ms) { return new Promise(resolve => setTimeout(resolve, ms)) }
async function animateDeskNumber(finalNumber, token) {
  const number = $('draw-number'), card = document.querySelector('.draw-card');
  number.classList.remove('landed'); number.classList.add('rolling'); card.classList.add('suspense');
  document.querySelector('.draw-label').textContent = 'ĐANG CHỌN BÀN…';
  const started = performance.now(), duration = 1900;
  while (performance.now() - started < duration) {
    if (token !== animationToken) return false;
    const progress = (performance.now() - started) / duration;
    number.textContent = String(1 + Math.floor(Math.random() * selectedDesks.length));
    await wait(55 + Math.pow(progress, 2) * 190);
  }
  if (token !== animationToken) return false;
  number.classList.remove('rolling'); number.textContent = String(finalNumber);
  void number.offsetWidth;
  number.classList.add('landed'); card.classList.remove('suspense'); document.querySelector('.draw-label').textContent = 'KẾT QUẢ';
  await wait(700);
  return token === animationToken;
}
async function runAnimation() {
  const token = ++animationToken;
  for (let i = 0; i < randomAssignments.length; i++) {
    if (token !== animationToken) return;
    const item = randomAssignments[i];
    const name = $('draw-name'); name.classList.remove('revealing'); name.textContent = item.name; void name.offsetWidth; name.classList.add('revealing');
    if (!await animateDeskNumber(item.deskNumber, token)) return;
    if (!await flyNameToDesk(item, token)) return;
    addAssignmentRow(item, i); applyOneAssignment(item);
    const done = i + 1;
    $('random-progress').style.width = `${done / randomAssignments.length * 100}%`;
    $('progress-text').textContent = `${done} / ${randomAssignments.length} học sinh`;
    await wait(450);
  }
  if (token === animationToken) finishAnimation();
}
function addAssignmentRow(item, i) {
  const row = document.createElement('div');
  const left = document.createElement('span'), number = document.createElement('b'), name = document.createTextNode(` ${item.name}`);
  number.textContent = `${i + 1}.`; left.append(number, name);
  const right = document.createElement('strong'); right.textContent = `Bàn ${item.deskNumber}`;
  row.append(left, right); $('assignment-list').append(row); $('assignment-list').scrollTop = $('assignment-list').scrollHeight;
}
function applyOneAssignment(item) { seats[item.seat].value = item.name; updateCount() }
function applyAllAssignments() {
  seats.forEach(seat => seat.value = '');
  prepareLiveLayout(); randomAssignments.forEach(placeNameInLiveDesk);
  randomAssignments.forEach(applyOneAssignment);
  $('assignment-list').replaceChildren(); randomAssignments.forEach(addAssignmentRow);
}
function finishAnimation() {
  $('random-close').disabled = false;
  refreshDeskLabels();
  $('skip-animation').hidden = true; $('finish-random').hidden = false;
  $('draw-name').textContent = 'Đã random xong!'; $('draw-number').textContent = '✓'; $('draw-number').classList.remove('rolling', 'landed');
  document.querySelector('.draw-card').classList.remove('suspense'); document.querySelector('.draw-label').textContent = 'HOÀN TẤT';
  $('random-progress').style.width = '100%'; $('progress-text').textContent = `Đã xếp ${randomAssignments.length} học sinh`;
  $('export-random').disabled = false; $('status').textContent = `Đã random chỗ ngồi cho ${randomAssignments.length} học sinh.`;
}
$('skip-animation').onclick = () => { animationToken++; applyAllAssignments(); finishAnimation() };
$('finish-random').onclick = closeRandom;

$('export-random').onclick = () => {
  if (!randomAssignments.length) return;
  if (!window.XLSX) { $('status').textContent = 'Thiếu file xlsx.full.min.js. Hãy tải đầy đủ các file của web lên GitHub.'; return; }
  const rows = [['STT', 'Họ và tên', 'Bàn', 'Vị trí trên sơ đồ']];
  randomAssignments.forEach((item, i) => rows.push([i + 1, item.name, item.deskNumber, physicalDeskName(item.physicalDesk)]));
  const book = XLSX.utils.book_new(), sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [{ wch: 7 }, { wch: 36 }, { wch: 10 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(book, sheet, 'Kết quả random');
  XLSX.writeFile(book, 'ket-qua-random-' + ($('className').value.trim() || 'lop') + '.xlsx');
};

// Tải Word từ mẫu gốc. Không gửi dữ liệu lên máy chủ.
$('word').onclick = async () => {
  const button = $('word'); button.disabled = true; button.textContent = 'Đang tạo Word…';
  try {
    if (!window.JSZip) throw Error('Thiếu file jszip.min.js.');
    const data = { teacher: $('teacher').value.trim(), className: $('className').value.trim(), year: $('year').value.trim(), names: seats.map(s => s.value) };
    const blob = await makeClassDocx(data);
    const url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = 'so-do-lop-' + (data.className || 'moi').replace(/[^\p{L}\p{N}_-]/gu, '-') + '.docx';
    document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 30000);
    $('status').textContent = 'Đã tạo file Word theo sơ đồ hiện tại.';
  } catch (error) { $('status').textContent = 'Không tạo được Word. Hãy giải nén đầy đủ các file rồi mở lại index.html.'; console.error(error) }
  finally { button.disabled = false; button.textContent = 'Tải Word (.docx)' }
};

updateHeader(); updateCount();
