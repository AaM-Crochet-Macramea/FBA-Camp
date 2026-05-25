const ADMIN_USERNAME = 'admin';
const ADMIN_PASSWORD = 'ebdaa';
const STORAGE_KEY = 'ai-web-kpis-members';
const ADMIN_SESSION = 'admin-session-active';
const MEMBER_SESSION = 'member-session';

const columnMap = {
  'اسم العضو': 'fullName',
  'الكود': 'password',
  'الرقم السري': 'password',
  'الرقم': 'phone',
  'الرقم القومي': 'nationalId',
  'العنوان': 'address',
  'المحافظة': 'governorate',
  'رقم الغرفة': 'room',
  'حضور التدريب الاول': 'attendance1',
  'استلام المشروع الاول': 'project1',
  'حضور التدريب الثاني': 'attendance2',
  'حضور التدريب الثالث': 'attendance3',
  'المشروع الثاني': 'project2',
  'المشروع الثالث': 'project3',
  'استلام المشروع النهائي': 'finalProject'
};

function getMembers() {
  const stored = localStorage.getItem(STORAGE_KEY);
  return stored ? JSON.parse(stored) : [];
}

function saveMembers(members) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(members));
}

function setAdminSession() {
  localStorage.setItem(ADMIN_SESSION, 'true');
}

function clearAdminSession() {
  localStorage.removeItem(ADMIN_SESSION);
}

function isAdminAuthenticated() {
  return localStorage.getItem(ADMIN_SESSION) === 'true';
}

function setMemberSession(member) {
  localStorage.setItem(MEMBER_SESSION, JSON.stringify(member));
}

function clearMemberSession() {
  localStorage.removeItem(MEMBER_SESSION);
}

function getMemberSession() {
  const value = localStorage.getItem(MEMBER_SESSION);
  return value ? JSON.parse(value) : null;
}

function findMemberFromSession(session) {
  if (!session) return null;
  const members = getMembers();
  return (
    members.find((member) => {
      if (session.nationalId && member.nationalId && member.nationalId === session.nationalId) {
        return true;
      }
      if (
        session.fullName &&
        member.fullName === session.fullName &&
        session.password &&
        member.password === session.password
      ) {
        return true;
      }
      if (session.phone && member.phone === session.phone && session.password && member.password === session.password) {
        return true;
      }
      return false;
    }) || null
  );
}

function isMemberAuthenticated() {
  return !!findMemberFromSession(getMemberSession());
}

function showMemberDashboard() {
  const loginPage = document.getElementById('login-page');
  const mainPage = document.getElementById('main-page');
  if (loginPage) loginPage.style.display = 'none';
  if (mainPage) mainPage.style.display = 'block';
  document.body.classList.add('member-logged-in');
  document.body.classList.remove('member-logged-out');
}

function showMemberLogin() {
  const loginPage = document.getElementById('login-page');
  const mainPage = document.getElementById('main-page');
  if (loginPage) loginPage.style.display = 'flex';
  if (mainPage) mainPage.style.display = 'none';
  document.body.classList.add('member-logged-out');
  document.body.classList.remove('member-logged-in');
}

function showAdminDashboard() {
  const loginPage = document.getElementById('admin-login-page');
  const adminPage = document.getElementById('admin-page');
  if (loginPage) loginPage.style.display = 'none';
  if (adminPage) adminPage.style.display = 'block';
  document.body.classList.add('admin-logged-in');
  document.body.classList.remove('admin-logged-out');
}

function showAdminLogin() {
  const loginPage = document.getElementById('admin-login-page');
  const adminPage = document.getElementById('admin-page');
  if (loginPage) loginPage.style.display = 'flex';
  if (adminPage) adminPage.style.display = 'none';
  document.body.classList.add('admin-logged-out');
  document.body.classList.remove('admin-logged-in');
}

function parseExcelFile(file, callback) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    callback(rows);
  };
  reader.readAsArrayBuffer(file);
}

function normalizeRow(row) {
  const member = {
    fullName: '',
    password: '',
    phone: '',
    nationalId: '',
    address: '',
    governorate: '',
    room: '',
    attendance1: false,
    project1: false,
    attendance2: false,
    attendance3: false,
    project2: false,
    project3: false,
    finalProject: false,
    // files for projects (store as { name, dataUrl, uploadedAt })
    project1File: null,
    project2File: null,
    project3File: null,
    finalProjectFile: null
  };

  Object.keys(row).forEach((key) => {
    const normalizedKey = key.trim();
    if (columnMap[normalizedKey]) {
      const value = row[key];
      const field = columnMap[normalizedKey];
      if (field.startsWith('attendance') || field.startsWith('project')) {
        member[field] = String(value).trim().toLowerCase() === 'نعم' || String(value).trim() === '1' || String(value).trim().toLowerCase() === 'true';
      } else {
        member[field] = String(value || '').trim();
      }
    }
  });

  if (!member.fullName && !member.nationalId && !member.password) {
    return null;
  }

  return member;
}

function mergeMembers(rows, replace = false) {
  const stored = getMembers();
  const updated = replace ? [] : [...stored];

  rows.forEach((row) => {
    const normalized = normalizeRow(row);
    if (!normalized) return;
    const existingIndex = updated.findIndex((member) => member.nationalId === normalized.nationalId || member.phone === normalized.phone || member.fullName === normalized.fullName);
    if (existingIndex >= 0) {
      updated[existingIndex] = { ...updated[existingIndex], ...normalized };
    } else {
      updated.push(normalized);
    }
  });

  saveMembers(updated);
  return updated;
}

function exportMembersToExcel() {
  const members = getMembers();
  if (!members.length) {
    alert('لا توجد بيانات لتصديرها');
    return;
  }

  const exportData = members.map((member) => ({
    'اسم المستخدم': member.fullName,
    'الكود': member.password,
    'الرقم': member.phone,
    'الرقم القومي': member.nationalId,
    'العنوان': member.address,
    'المحافظة': member.governorate,
    'رقم الغرفة': member.room,
    'حضور التدريب الأول': member.attendance1 ? 'نعم' : 'لا',
    'المشروع الأول': member.project1 ? 'نعم' : 'لا',
    'حضور التدريب الثاني': member.attendance2 ? 'نعم' : 'لا',
    'المشروع الثاني': member.project2 ? 'نعم' : 'لا',
    'حضور التدريب الثالث': member.attendance3 ? 'نعم' : 'لا',
    'المشروع الثالث': member.project3 ? 'نعم' : 'لا',
    'المشروع النهائي': member.finalProject ? 'نعم' : 'لا'
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Members');
  XLSX.writeFile(workbook, 'members_export.xlsx');
}

function createButton(label, handler) {
  const button = document.createElement('button');
  button.textContent = label;
  button.type = 'button';
  button.addEventListener('click', handler);
  return button;
}

function createFileInput(onFileSelected) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls';
  input.addEventListener('change', () => {
    if (!input.files || input.files.length === 0) return;
    onFileSelected(input.files[0]);
    input.value = '';
  });
  input.className = 'hidden-file-input';
  return input;
}

function renderAdminTable(container, filter = '') {
  const members = getMembers();
  const query = filter.trim().toLowerCase();
  const visibleMembers = query
    ? members.filter((member) => {
        return [
          member.fullName,
          member.password,
          member.phone,
          member.nationalId,
          member.address,
          member.governorate,
          member.room
        ]
          .filter(Boolean)
          .some((value) => value.toLowerCase().includes(query));
      })
    : members;

  container.innerHTML = '';
  const tableWrapper = document.createElement('div');
  tableWrapper.className = 'table-wrapper';
  const table = document.createElement('table');

  const header = document.createElement('tr');
  [
    'اسم المستخدم',
    'الكود',
    'الرقم',
    'الرقم القومي',
    'العنوان',
    'المحافظة',
    'رقم الغرفة',
    'حضور التدريب الأول',
    'المشروع الأول',
    'حضور التدريب الثاني',
    'المشروع الثاني',
    'حضور التدريب الثالث',
    'المشروع الثالث',
    'المشروع النهائي',
    'حذف العضو'
  ].forEach((text) => {
    const th = document.createElement('th');
    th.textContent = text;
    header.appendChild(th);
  });
  table.appendChild(header);

  visibleMembers.forEach((member, index) => {
    const row = document.createElement('tr');
    const textFields = ['fullName', 'password', 'phone', 'nationalId', 'address', 'governorate', 'room'];
    textFields.forEach((field) => {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.textContent = member[field] || '';
      td.addEventListener('blur', () => {
        member[field] = td.textContent.trim();
        saveMembers(members);
      });
      row.appendChild(td);
    });

    // Build cells in header order: attendance1, project1, attendance2, project2, attendance3, project3, finalProject
    const cellsOrder = [
      { type: 'attendance', field: 'attendance1' },
      { type: 'project', field: 'project1', fileKey: 'project1File' },
      { type: 'attendance', field: 'attendance2' },
      { type: 'project', field: 'project2', fileKey: 'project2File' },
      { type: 'attendance', field: 'attendance3' },
      { type: 'project', field: 'project3', fileKey: 'project3File' },
      { type: 'project', field: 'finalProject', fileKey: 'finalProjectFile' }
    ];

    cellsOrder.forEach((cell) => {
      if (cell.type === 'attendance') {
        const td = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = !!member[cell.field];
        checkbox.addEventListener('change', () => {
          member[cell.field] = checkbox.checked;
          saveMembers(members);
        });
        td.appendChild(checkbox);
        row.appendChild(td);
      } else if (cell.type === 'project') {
        const td = document.createElement('td');
        const fileObj = member[cell.fileKey];
        if (fileObj && fileObj.dataUrl) {
          const link = document.createElement('a');
          link.textContent = fileObj.name || 'ملف PDF';
          link.href = fileObj.dataUrl;
          link.target = '_blank';
          link.rel = 'noopener noreferrer';
          link.className = 'admin-file-link';
          td.appendChild(link);

          const dl = document.createElement('button');
          dl.type = 'button';
          dl.className = 'download-file-button';
          dl.textContent = 'تنزيل';
          dl.addEventListener('click', () => {
            const a = document.createElement('a');
            a.href = fileObj.dataUrl;
            a.download = fileObj.name || 'document.pdf';
            document.body.appendChild(a);
            a.click();
            a.remove();
          });
          td.appendChild(document.createTextNode(' '));
          td.appendChild(dl);
        } else {
          td.textContent = 'لم يتم رفع ملف';
        }
        row.appendChild(td);
      }
    });

    const deleteCell = document.createElement('td');
    const deleteButton = document.createElement('button');
    deleteButton.className = 'delete-button';
    deleteButton.textContent = 'حذف';
    deleteButton.addEventListener('click', () => {
      if (confirm('هل تريد حذف هذا العضو؟')) {
        const originalIndex = members.findIndex((item) => item.nationalId === member.nationalId && item.phone === member.phone && item.fullName === member.fullName);
        if (originalIndex >= 0) {
          members.splice(originalIndex, 1);
          saveMembers(members);
          renderAdminContent();
        }
      }
    });
    deleteCell.appendChild(deleteButton);
    row.appendChild(deleteCell);

    table.appendChild(row);
  });

  tableWrapper.appendChild(table);
  container.appendChild(tableWrapper);
}

function renderAdminContent() {
  if (!isAdminAuthenticated()) {
    showAdminLogin();
    return;
  }

  const content = document.getElementById('admin-content');
  showAdminDashboard();
  
  content.innerHTML = '';

  const welcomeCard = document.createElement('section');
  welcomeCard.className = 'card';
  welcomeCard.innerHTML = `
    <div class="admin-top-row">
      <div>
        <h2>لوحة إدارة الأعضاء</h2>
        <p>يمكنك إضافة أو استيراد أو تعديل بيانات الأعضاء من ملف Excel.</p>
      </div>
      <button class="logout-button admin-logout-button" type="button">تسجيل الخروج</button>
    </div>
  `;

  const welcomeLogout = welcomeCard.querySelector('.admin-logout-button');
  welcomeLogout.addEventListener('click', handleAdminLogout);

  const actions = document.createElement('div');
  actions.className = 'admin-actions card-actions grid-3';

  const replaceCard = document.createElement('section');
  replaceCard.className = 'action-card';
  replaceCard.innerHTML = '<strong>ارفاق البيانات</strong>';
  const replaceInput = createFileInput((file) => {
    parseExcelFile(file, (rows) => {
      mergeMembers(rows, true);
      renderAdminContent();
      alert('تم إرفاق البيانات بنجاح');
    });
  });
  const replaceButton = createButton('ارفاق البيانات', () => replaceInput.click());
  replaceCard.appendChild(replaceButton);
  replaceCard.appendChild(replaceInput);

  const appendCard = document.createElement('section');
  appendCard.className = 'action-card';
  appendCard.innerHTML = '<strong>اضافة بيانات جديدة</strong>';
  const appendInput = createFileInput((file) => {
    parseExcelFile(file, (rows) => {
      mergeMembers(rows, false);
      renderAdminContent();
      alert('تم إضافة البيانات الجديدة بنجاح');
    });
  });
  const appendButton = createButton('اضافة بيانات جديدة', () => appendInput.click());
  appendCard.appendChild(appendButton);
  appendCard.appendChild(appendInput);

  const exportCard = document.createElement('section');
  exportCard.className = 'action-card';
  exportCard.innerHTML = '<strong>تصدير بيانات الأعضاء</strong>';
  const exportButton = createButton('تصدير البيانات', exportMembersToExcel);
  exportCard.appendChild(exportButton);

  actions.appendChild(replaceCard);
  actions.appendChild(appendCard);
  actions.appendChild(exportCard);

  const section = document.createElement('section');
  section.className = 'card admin-table-card';

  const searchRow = document.createElement('div');
  searchRow.className = 'admin-search-row';
  const searchInput = document.createElement('input');
  searchInput.type = 'search';
  searchInput.placeholder = 'ابحث بالاسم أو الرقم أو العنوان';
  searchInput.className = 'search-input admin-search-input';
  searchRow.appendChild(searchInput);

  const scrollHint = document.createElement('p');
  scrollHint.className = 'admin-table-scroll-hint';
  scrollHint.textContent = '↔ اسحب الجدول أفقياً لعرض جميع الأعمدة';

  const tableContainer = document.createElement('div');
  tableContainer.className = 'admin-table-container';
  searchInput.addEventListener('input', () => renderAdminTable(tableContainer, searchInput.value));

  section.innerHTML = '<h2>بيانات الأعضاء</h2>';
  section.appendChild(searchRow);
  section.appendChild(scrollHint);
  renderAdminTable(tableContainer);
  section.appendChild(tableContainer);

  content.appendChild(welcomeCard);
  content.appendChild(actions);
  content.appendChild(section);
}

function renderAdminLogin() {
  const content = document.getElementById('admin-content');
  content.innerHTML = '';
  const card = document.createElement('section');
  card.className = 'card';
  card.innerHTML = '<h2>تسجيل دخول الأدمن</h2>';

  const form = document.createElement('form');
  form.className = 'form-group';
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const username = event.target.elements.username.value.trim();
    const password = event.target.elements.password.value.trim();
    if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
      setAdminSession();
      renderAdminContent();
      return;
    }
    alert('اسم المستخدم أو كلمة المرور غير صحيحة');
  });

  const usernameInput = document.createElement('input');
  usernameInput.name = 'username';
  usernameInput.placeholder = 'اسم المستخدم';
  usernameInput.value = ADMIN_USERNAME;

  const passwordInput = document.createElement('input');
  passwordInput.name = 'password';
  passwordInput.type = 'password';
  passwordInput.placeholder = 'كلمة المرور';
  passwordInput.value = ADMIN_PASSWORD;

  const submitButton = createButton('تسجيل الدخول');
  submitButton.type = 'submit';

  form.appendChild(usernameInput);
  form.appendChild(passwordInput);
  form.appendChild(submitButton);
  card.appendChild(form);
  content.appendChild(card);
}

function initAdmin() {
  if (isAdminAuthenticated()) {
    renderAdminContent();
  } else {
    showAdminLogin();
  }

  window.addEventListener('storage', (e) => {
    if (e.key === STORAGE_KEY && isAdminAuthenticated()) {
      try {
        renderAdminContent();
      } catch (err) {
        /* ignore */
      }
    }
  });
}

function buildLoginForm(container) {
  const card = document.createElement('section');
  card.className = 'card';
  card.innerHTML = '<h2>تسجيل الدخول</h2><p>استخدم بياناتك المسجلة في صفحة الأدمن لتسجيل الدخول.</p>';

  const form = document.createElement('form');
  form.className = 'form-group';
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const username = event.target.elements.username.value.trim();
    const password = event.target.elements.password.value.trim();
    const member = getMembers().find((item) => (item.fullName === username || item.phone === username) && item.password === password);
    if (!member) {
      alert('بيانات الدخول غير صحيحة');
      return;
    }
    setMemberSession(member);
    renderSiteDashboard();
  });

  const usernameInput = document.createElement('input');
  usernameInput.name = 'username';
  usernameInput.placeholder = 'اسم العضو أو الرقم';
  const passwordInput = document.createElement('input');
  passwordInput.name = 'password';
  passwordInput.type = 'password';
  passwordInput.placeholder = 'الكود';

  const submitButton = createButton('دخول');
  submitButton.type = 'submit';

  form.appendChild(usernameInput);
  form.appendChild(passwordInput);
  form.appendChild(submitButton);
  card.appendChild(form);
  container.appendChild(card);
}

function escapeHtml(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function createProfileCard(member) {
  const card = document.createElement('section');
  card.className = 'card member-card';
  card.innerHTML = `
    <div class="member-card-inner">
      <div class="member-card-body">
        <article class="visa-card-face" aria-label="بطاقة العضو">
          <div class="visa-card-pattern" aria-hidden="true"></div>
          <div class="visa-card-content">
            <div class="visa-card-toolbar">
              <button class="visa-toolbar-btn member-card-collapse-handle" type="button" aria-expanded="true">
                إغلاق البطاقة
              </button>
              <button class="visa-toolbar-btn member-logout-button" type="button">تسجيل الخروج</button>
            </div>
            <div class="visa-card-collapsible">
              <div class="visa-card-top">
                <div class="visa-chip" aria-hidden="true"></div>
                <svg class="visa-contactless" viewBox="0 0 24 24" aria-hidden="true">
                  <path d="M7 12c2 0 3-1 4-2M11 12c2 0 3-1 4-2M15 12c2 0 3-1 4-2" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
                </svg>
                <span class="visa-brand" dir="ltr">FBA -CARD</span>
              </div>
              <p class="visa-card-idline" dir="ltr">EBDAA(360)-MEMBER INFO</p>
              <div class="visa-card-holder">
                <span class="visa-label">اسم العضو</span>
                <span class="visa-name">${escapeHtml(member.fullName || '—')}</span>
              </div>
              <div class="visa-card-details">
                <div class="visa-meta">
                  <span class="visa-label">الرقم القومي</span>
                  <span class="visa-value visa-mono" dir="ltr">${escapeHtml(member.nationalId || '—')}</span>
                </div>
                <div class="visa-meta">
                  <span class="visa-label">الكود</span>
                  <span class="visa-value visa-mono" dir="ltr">${escapeHtml(member.password || '—')}</span>
                </div>
                <div class="visa-meta">
                  <span class="visa-label">الهاتف</span>
                  <span class="visa-value visa-mono" dir="ltr">${escapeHtml(member.phone || '—')}</span>
                </div>
                <div class="visa-meta">
                  <span class="visa-label">المحافظة</span>
                  <span class="visa-value">${escapeHtml(member.governorate || '—')}</span>
                </div>
                <div class="visa-meta">
                  <span class="visa-label">رقم الغرفة</span>
                  <span class="visa-value visa-mono" dir="ltr">${escapeHtml(member.room || '—')}</span>
                </div>
              </div>
            </div>
          </div>
        </article>
      </div>
    </div>
  `;

  const logoutButton = card.querySelector('.member-logout-button');
  logoutButton.addEventListener('click', handleMemberLogout);

  const collapseHandle = card.querySelector('.member-card-collapse-handle');
  collapseHandle.addEventListener('click', () => {
    const collapsed = card.classList.toggle('is-collapsed');
    collapseHandle.setAttribute('aria-expanded', String(!collapsed));
    collapseHandle.textContent = collapsed ? 'فتح البطاقة' : 'إغلاق البطاقة';
  });

  return card;
}

function createStatusChip(text, ok = true) {
  const span = document.createElement('span');
  span.className = `status-chip ${ok ? 'status-ok' : 'status-warning'}`;
  span.textContent = text;
  return span;
}

function buildKpiCard(member) {
  const trainingsDone = [member.attendance1, member.attendance2, member.attendance3].filter(Boolean).length;
  const projectsDone = [member.project1, member.project2, member.project3, member.finalProject].filter(Boolean).length;
  const totalTrainings = 3;
  const totalProjects = 4;
  const trainingOk = trainingsDone === totalTrainings;
  const projectOk = projectsDone === totalProjects;
  const overallProgress = Math.round(((trainingsDone + projectsDone) / (totalTrainings + totalProjects)) * 100);

  const card = document.createElement('section');
  card.className = 'card kpi-card';
  card.innerHTML = `
    <h2>📊 مؤشرات الأداء</h2>
    
    <div class="kpi-overview">
      <div class="progress-overview">
        <div class="overall-progress">
          <div class="progress-circle">
            <svg viewBox="0 0 100 100" class="circular-progress">
              <circle cx="50" cy="50" r="45" fill="none" stroke="rgba(45, 111, 196, 0.2)" stroke-width="8"/>
              <circle cx="50" cy="50" r="45" fill="none" stroke="#3b82d9" stroke-width="8"
                      stroke-dasharray="${overallProgress * 2.83} 283" 
                      transform="rotate(-90 50 50)"
                      style="transition: stroke-dasharray 0.5s ease;"/>
            </svg>
            <div class="progress-text">
              <span class="percentage"><span class="pct-value">${overallProgress}</span><span class="pct-symbol" aria-hidden="true">%</span></span>
              <span class="label">المجموع</span>
            </div>
          </div>
        </div>

        <div class="stats-grid">
          <div class="stat-item">
            <div class="stat-header">
              <span class="stat-label">التدريبات</span>
              <span class="stat-badge ${trainingOk ? 'badge-ok' : 'badge-pending'}">${trainingOk ? '✓ مكتمل' : 'قيد الإنجاز'}</span>
            </div>
            <div class="stat-bar">
              <div class="bar-fill" style="width: ${(trainingsDone / totalTrainings) * 100}%;"></div>
            </div>
            <div class="stat-text">\u2066${trainingsDone} من ${totalTrainings}\u2069</div>
          </div>

          <div class="stat-item">
            <div class="stat-header">
              <span class="stat-label">المشاريع</span>
              <span class="stat-badge ${projectOk ? 'badge-ok' : 'badge-pending'}">${projectOk ? '✓ مكتمل' : 'قيد الإنجاز'}</span>
            </div>
            <div class="stat-bar">
              <div class="bar-fill" style="width: ${(projectsDone / totalProjects) * 100}%;"></div>
            </div>
            <div class="stat-text">\u2066${projectsDone} من ${totalProjects}\u2069</div>
          </div>
        </div>
      </div>
    </div>

    <div class="details-grid"></div>

    <div class="overall-status">
      <div class="status-label">الحالة العامة:</div>
      <div class="status-badge ${overallProgress === 100 ? 'status-excellent' : overallProgress >= 75 ? 'status-good' : overallProgress >= 50 ? 'status-fair' : 'status-low'}">
        ${overallProgress === 100 ? '🎉 ممتاز - مكتمل' : overallProgress >= 75 ? '👍 جيد جداً' : overallProgress >= 50 ? '⚠️ جيد' : '📈 قيد التطور'}
      </div>
    </div>
  `;

  // helper to persist member changes
  function updateMember(updated) {
    const members = getMembers();
    const idx = members.findIndex((m) => m.nationalId === updated.nationalId);
    if (idx >= 0) {
      members[idx] = { ...members[idx], ...updated };
      saveMembers(members);
      setMemberSession(members[idx]);
    }
  }

  const detailsGrid = card.querySelector('.details-grid');

  // Trainings section (read-only; admin marks attendance)
  const trainSection = document.createElement('div');
  trainSection.className = 'detail-section';
  trainSection.innerHTML = '<h3>🎯 تفاصيل التدريبات</h3>';
  const checklist = document.createElement('div');
  checklist.className = 'checklist';
  ['التدريب الأول', 'التدريب الثاني', 'التدريب الثالث'].forEach((label, i) => {
    const field = `attendance${i + 1}`;
    const attended = !!member[field];
    const item = document.createElement('div');
    item.className = 'checklist-item';
    const status = document.createElement('span');
    status.className = attended ? 'attendance-status attendance-done' : 'attendance-status attendance-pending';
    status.setAttribute('aria-label', attended ? 'تم تسجيل الحضور' : 'لم يُسجَّل الحضور');
    if (attended) {
      status.textContent = '✓';
    }
    const lab = document.createElement('span');
    lab.className = 'check-label';
    lab.textContent = label;
    item.appendChild(status);
    item.appendChild(lab);
    checklist.appendChild(item);
  });
  trainSection.appendChild(checklist);
  detailsGrid.appendChild(trainSection);

  // Projects section with PDF upload
  const projSection = document.createElement('div');
  projSection.className = 'detail-section';
  projSection.innerHTML = '<h3>🚀 تفاصيل المشاريع</h3>';
  const projList = document.createElement('div');
  projList.className = 'checklist';

  const projects = [
    { label: 'المشروع الأول', key: 'project1', fileKey: 'project1File' },
    { label: 'المشروع الثاني', key: 'project2', fileKey: 'project2File' },
    { label: 'المشروع الثالث', key: 'project3', fileKey: 'project3File' },
    { label: 'المشروع النهائي', key: 'finalProject', fileKey: 'finalProjectFile' }
  ];

  projects.forEach((p) => {
    const item = document.createElement('div');
    item.className = 'checklist-item project-item';

    const title = document.createElement('span');
    title.className = 'check-label';
    title.textContent = p.label;
    item.appendChild(title);

    const fileArea = document.createElement('div');
    fileArea.className = 'project-file-area';

    const fileObj = member[p.fileKey];
    const info = document.createElement('div');
    info.className = 'project-file-info';
    if (fileObj && fileObj.dataUrl) {
      const a = document.createElement('a');
      a.href = fileObj.dataUrl;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = fileObj.name || 'ملف PDF';
      fileArea.appendChild(a);

      const dl = document.createElement('button');
      dl.type = 'button';
      dl.className = 'download-file-button';
      dl.textContent = 'تنزيل';
      dl.addEventListener('click', () => {
        const a2 = document.createElement('a');
        a2.href = fileObj.dataUrl;
        a2.download = fileObj.name || 'document.pdf';
        document.body.appendChild(a2);
        a2.click();
        a2.remove();
      });
      fileArea.appendChild(dl);
    } else {
      info.textContent = 'لم يتم رفع ملف بعد';
      fileArea.appendChild(info);
    }

    const uploadBtn = document.createElement('button');
    uploadBtn.type = 'button';
    uploadBtn.textContent = 'ارفاق ملف PDF';
    uploadBtn.addEventListener('click', () => input.click());

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'application/pdf';
    input.style.display = 'none';
    input.addEventListener('change', () => {
      if (!input.files || !input.files[0]) return;
      const f = input.files[0];
      const reader = new FileReader();
      reader.onload = (e) => {
        const dataUrl = e.target.result;
        member[p.fileKey] = { name: f.name, dataUrl, uploadedAt: Date.now() };
        member[p.key] = true;
        updateMember(member);
        // update UI
        fileArea.innerHTML = '';
        const a = document.createElement('a');
        a.href = member[p.fileKey].dataUrl;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.textContent = member[p.fileKey].name;
        fileArea.appendChild(a);
        const dl = document.createElement('button');
        dl.type = 'button';
        dl.className = 'download-file-button';
        dl.textContent = 'تنزيل';
        dl.addEventListener('click', () => {
          const a2 = document.createElement('a');
          a2.href = member[p.fileKey].dataUrl;
          a2.download = member[p.fileKey].name || 'document.pdf';
          document.body.appendChild(a2);
          a2.click();
          a2.remove();
        });
        fileArea.appendChild(dl);
      };
      reader.readAsDataURL(f);
      input.value = '';
    });

    item.appendChild(fileArea);
    item.appendChild(uploadBtn);
    item.appendChild(input);
    projList.appendChild(item);
  });

  projSection.appendChild(projList);
  detailsGrid.appendChild(projSection);

  return card;
}

function renderSiteDashboard() {
  const main = document.getElementById('main-content');
  const sessionMember = getMemberSession();
  const currentMember = findMemberFromSession(sessionMember);

  if (!currentMember) {
    clearMemberSession();
    showMemberLogin();
    return;
  }

  setMemberSession(currentMember);
  showMemberDashboard();

  const memberNameEl = document.getElementById('member-name');
  if (memberNameEl) {
    memberNameEl.textContent = currentMember.fullName;
  }

  main.innerHTML = '';
  const profile = createProfileCard(currentMember);
  const kpi = buildKpiCard(currentMember);

  main.appendChild(profile);
  main.appendChild(kpi);
}

function initSite() {
  if (isMemberAuthenticated()) {
    renderSiteDashboard();
    return;
  }

  showMemberLogin();
}

// دوال معالجة صفحات الدخول الجديدة
function handleMemberLogin(event) {
  event.preventDefault();
  const name = document.getElementById('username').value.trim();
  const password = document.getElementById('member-password').value.trim();
  
  const member = getMembers().find((item) => item.fullName === name && item.password === password);
  if (!member) {
    alert('اسم العضو أو الكود غير صحيح');
    return;
  }
  
  setMemberSession(member);
  renderSiteDashboard();
}

function handleMemberLogout() {
  clearMemberSession();
  showMemberLogin();
  const main = document.getElementById('main-content');
  if (main) main.innerHTML = '';
}

function handleAdminLogin(event) {
  event.preventDefault();
  const username = document.getElementById('admin-username').value.trim();
  const password = document.getElementById('admin-password').value.trim();
  
  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
    setAdminSession();
    renderAdminContent();
    return;
  }
  
  alert('اسم المستخدم أو كلمة المرور غير صحيحة');
}

function handleAdminLogout() {
  clearAdminSession();
  showAdminLogin();
  const content = document.getElementById('admin-content');
  if (content) content.innerHTML = '';
}

function initDevToolsProtection() {
  document.addEventListener('contextmenu', (event) => {
    event.preventDefault();
  });

  document.addEventListener('keydown', (event) => {
    const key = event.key.toUpperCase();
    const blocked =
      key === 'F12' ||
      (event.ctrlKey && event.shiftKey && ['I', 'J', 'C'].includes(key)) ||
      (event.ctrlKey && key === 'U');

    if (blocked) {
      event.preventDefault();
      event.stopPropagation();
      return false;
    }
  });
}

function bootstrapApp() {
  initDevToolsProtection();

  if (document.getElementById('admin-content')) {
    initAdmin();
  }

  if (document.getElementById('main-content') && document.getElementById('login-page')) {
    initSite();
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', bootstrapApp);
} else {
  bootstrapApp();
}
