(() => {
  const data = window.DASHBOARD_DATA;
  const allowedIdentities = new Set([
    'anhvn1', 'anhvn1@sav.gov.vn',
    'quangnd', 'quangnd@sav.gov.vn',
    'dungbt1', 'dungbt1@sav.gov.vn',
    'tungnt1', 'tungnt1@sav.gov.vn'
  ]);
  const gate = document.querySelector('#accessGate');
  const accessForm = document.querySelector('#accessForm');
  const accessInput = document.querySelector('#accessIdentity');
  const gateError = document.querySelector('#gateError');
  if (sessionStorage.getItem('sav-dashboard-access') === 'granted') gate.classList.add('unlocked');
  accessForm.addEventListener('submit', event => {
    event.preventDefault();
    const identity = accessInput.value.trim().toLocaleLowerCase('vi');
    if (allowedIdentities.has(identity)) {
      sessionStorage.setItem('sav-dashboard-access', 'granted');
      gate.classList.add('unlocked');
      gateError.textContent = '';
    } else {
      gateError.textContent = 'Tài khoản hoặc email chưa được cấp quyền.';
      accessInput.select();
    }
  });
  const state = { view: 'overview', page: 1, pageSize: 18, query: '', unit: '', code: '', frequency: '', issueUnit: 'Tất cả' };
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const esc = (v = '') => String(v).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  const fmt = n => new Intl.NumberFormat('vi-VN').format(n);
  const fileSize = bytes => bytes > 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
  const link = path => encodeURI(path).replace(/#/g, '%23');

  function switchView(view) {
    state.view = view;
    $$('.view').forEach(el => el.classList.toggle('active', el.id === `view-${view}`));
    $$('.nav-item').forEach(el => el.classList.toggle('active', el.dataset.view === view));
    const titles = {overview:'Tổng quan phản hồi',catalog:'Danh mục chi tiết',units:'Đơn vị và hồ sơ',issues:'Điểm cần rà soát'};
    $('#pageTitle').textContent = titles[view];
    window.scrollTo({top:0, behavior:'smooth'});
  }

  function renderOverview() {
    $('#updatedAt').textContent = new Date(data.meta.updated).toLocaleString('vi-VN');
    const req = data.request.files[0];
    $('#requestDownload').innerHTML = req ? `<a class="button ghost" download href="${link(req.download)}">Tải yêu cầu gốc</a>` : '';
    const metrics = [
      [data.meta.unit_count, 'đơn vị đã phản hồi'],
      [data.meta.detail_count, 'dòng danh mục đã trích xuất'],
      [data.meta.priority_count, 'đề xuất ưu tiên đang có'],
      [data.issues.length, 'điểm cần rà soát'],
    ];
    $('#metrics').innerHTML = metrics.map(([n,label]) => `<div class="metric"><strong>${fmt(n)}</strong><span>${label}</span></div>`).join('');
    $('#unitProgress').innerHTML = data.units.map(u => `<div class="unit-row">
      <div class="unit-name"><strong>${esc(u.code)}</strong><small>${fmt(u.detail_count)} nguồn · ${fmt(u.priority_count)} ưu tiên</small></div>
      <div class="form-pill">Biểu 01</div><div class="form-pill">Biểu 02</div>
      <div class="form-pill ${u.form3 === 'Chưa có' ? 'missing' : u.form3 === 'Một phần' ? 'partial' : ''}">${u.form3 === 'Đầy đủ' ? 'Biểu 03' : u.form3}</div>
      <div class="unit-total">${fmt(u.files.length)} tệp</div></div>`).join('');
    const max = Math.max(...data.common_priorities.map(x => x.units));
    $('#commonPriorities').innerHTML = data.common_priorities.slice(0,8).map(x => `<div class="bar-item"><code>${esc(x.code)}</code><div class="bar-track"><div class="bar-fill" style="width:${x.units/max*100}%"></div></div><span>${x.units}/${data.meta.unit_count}</span></div>`).join('');
    const kv8 = data.priorities.filter(p => p.unit === 'KV8').slice(0,10);
    $('#priorityHighlights').innerHTML = kv8.map(p => `<div class="priority-item"><span class="rank">${String(p.rank).padStart(2,'0')}</span><code>${esc(p.code)}</code><span>${esc(p.name)}</span></div>`).join('');
    $('#issueCount').textContent = data.issues.length;
    $('#issuePreview').innerHTML = data.issues.slice(0,3).map(i => `<div class="issue-mini"><strong>${esc(i.title)}</strong><span>${esc(i.unit)} · ${esc(i.severity)}</span></div>`).join('');
  }

  function filteredDetails() {
    const q = state.query.toLocaleLowerCase('vi');
    return data.details.filter(r => {
      const hay = [r.unit,r.code,r.group,r.name,r.indicators,r.owner,r.purpose,r.frequency].join(' ').toLocaleLowerCase('vi');
      return (!q || hay.includes(q)) && (!state.unit || r.unit === state.unit) && (!state.code || r.code === state.code) && (!state.frequency || r.frequency.toLocaleLowerCase('vi').includes(state.frequency));
    });
  }

  function renderCatalog() {
    const rows = filteredDetails();
    const pages = Math.max(1, Math.ceil(rows.length / state.pageSize));
    state.page = Math.min(state.page, pages);
    const slice = rows.slice((state.page-1)*state.pageSize, state.page*state.pageSize);
    $('#catalogCount').textContent = `${fmt(rows.length)} kết quả`;
    $('#catalogBody').innerHTML = slice.length ? slice.map((r, idx) => `<tr>
      <td class="row-unit">${esc(r.unit)}</td><td><code>${esc(r.code)}</code></td><td>${esc(r.group)}</td><td>${esc(r.name)}</td><td>${esc(r.stages)}</td><td>${esc(r.frequency)}</td>
      <td><button class="row-more" data-detail-index="${data.details.indexOf(r)}">Chi tiết</button></td></tr>`).join('') : `<tr><td colspan="7" class="empty">Không có kết quả phù hợp.</td></tr>`;
    $('#pagination').innerHTML = Array.from({length:pages},(_,i)=>i+1).slice(Math.max(0,state.page-3),Math.min(pages,state.page+2)).map(p => `<button class="page-btn ${p===state.page?'active':''}" data-page="${p}">${p}</button>`).join('');
  }

  function renderFilters() {
    $('#unitFilter').innerHTML += data.units.map(u => `<option value="${u.code}">${u.code}</option>`).join('');
    const codes = [...new Set(data.details.map(r => r.code).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi'));
    $('#codeFilter').innerHTML += codes.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('');
  }

  function fileList(files) {
    return files.map(f => `<a class="file-link" download href="${link(f.download)}"><span class="file-type">${esc(f.type)}</span><span><strong>${esc(f.name)}</strong><small>${fileSize(f.size)} · ${new Date(f.updated).toLocaleDateString('vi-VN')}</small></span><span class="download-symbol">↓</span></a>`).join('');
  }

  function renderUnits() {
    $('#unitCards').innerHTML = data.units.map(u => `<article class="unit-card">
      <div class="unit-card-top"><div><p class="eyebrow">${esc(u.code)}</p><h3>${esc(u.name)}</h3></div><span class="coverage">Biểu 03: ${esc(u.form3)}</span></div>
      <p>${esc(u.note)}</p><div class="unit-stats"><div><strong>${fmt(u.detail_count)}</strong><span>nguồn dữ liệu</span></div><div><strong>${fmt(u.priority_count)}</strong><span>đề xuất ưu tiên</span></div><div><strong>${u.files.length}</strong><span>tệp gốc</span></div></div>
      <div class="file-list">${fileList(u.files)}</div><button class="button wide" data-unit-open="${u.code}">Xem ưu tiên và nhận xét</button></article>`).join('');
  }

  function renderIssues() {
    const units = ['Tất cả', ...data.units.map(u => u.code)];
    $('#issueFilters').innerHTML = units.map(u => `<button class="filter-chip ${u===state.issueUnit?'active':''}" data-issue-unit="${u}">${u}</button>`).join('');
    const items = data.issues.filter(i => state.issueUnit === 'Tất cả' || i.unit === state.issueUnit);
    $('#issuesList').innerHTML = items.map(i => `<article class="issue-card ${i.severity==='Cao'?'high':''}"><span class="severity">${esc(i.severity)}</span><span class="issue-unit">${esc(i.unit)}</span><div><h3>${esc(i.title)}</h3><p>${esc(i.detail)}</p></div></article>`).join('');
  }

  function openDetail(index) {
    const r = data.details[index];
    $('#dialogContent').innerHTML = `<div class="dialog-body"><h2>${esc(r.name || r.group)}</h2><span class="dialog-code">${esc(r.unit)} · ${esc(r.code)}</span><div class="detail-grid">
      ${[['Nhóm dữ liệu',r.group],['Lĩnh vực',r.field],['Giai đoạn sử dụng',r.stages],['Mức độ sử dụng',r.frequency],['Cơ quan quản lý',r.owner],['Hình thức hiện có',r.format],['Nội dung, chỉ tiêu',r.indicators],['Mục đích kiểm toán',r.purpose]].map(([k,v],i)=>`<div class="detail-block ${i>5?'wide':''}"><strong>${k}</strong><span>${esc(v||'Chưa nêu')}</span></div>`).join('')}</div></div>`;
    $('#detailDialog').showModal();
  }

  function openUnit(code) {
    const unit = data.units.find(u => u.code === code);
    const priorities = data.priorities.filter(p => p.unit === code);
    const issues = data.issues.filter(i => i.unit === code);
    $('#dialogContent').innerHTML = `<div class="dialog-body"><p class="eyebrow">${esc(code)}</p><h2>${esc(unit.name)}</h2><p>${esc(unit.note)}</p>
      <h3>Danh sách ưu tiên (${priorities.length})</h3><div class="file-list">${priorities.slice(0,12).map(p=>`<div class="file-link"><span class="file-type">${p.rank}</span><span><strong>${esc(p.code)} · ${esc(p.name)}</strong><small>${esc(p.reason)}</small></span></div>`).join('')}</div>
      ${priorities.length>12?`<p class="legend">Đang hiển thị 12/${priorities.length} ưu tiên. Xem đầy đủ trong file gốc.</p>`:''}
      ${issues.length?`<h3>Điểm cần rà soát</h3>${issues.map(i=>`<div class="issue-mini"><strong>${esc(i.title)}</strong><span>${esc(i.detail)}</span></div>`).join('')}`:''}
      <div class="unit-dialog-files"><h3>Tệp gốc</h3>${fileList(unit.files)}</div></div>`;
    $('#detailDialog').showModal();
  }

  $$('.nav-item').forEach(b => b.addEventListener('click', () => switchView(b.dataset.view)));
  $$('[data-jump]').forEach(b => b.addEventListener('click', () => switchView(b.dataset.jump)));
  document.addEventListener('click', e => {
    const detail = e.target.closest('[data-detail-index]'); if (detail) openDetail(+detail.dataset.detailIndex);
    const unit = e.target.closest('[data-unit-open]'); if (unit) openUnit(unit.dataset.unitOpen);
    const page = e.target.closest('[data-page]'); if (page) { state.page = +page.dataset.page; renderCatalog(); }
    const issue = e.target.closest('[data-issue-unit]'); if (issue) { state.issueUnit = issue.dataset.issueUnit; renderIssues(); }
  });
  $('#globalSearch').addEventListener('input', e => { state.query = e.target.value.trim(); state.page = 1; if (state.query && state.view !== 'catalog') switchView('catalog'); renderCatalog(); });
  $('#unitFilter').addEventListener('change', e => { state.unit=e.target.value; state.page=1; renderCatalog(); });
  $('#codeFilter').addEventListener('change', e => { state.code=e.target.value; state.page=1; renderCatalog(); });
  $('#frequencyFilter').addEventListener('change', e => { state.frequency=e.target.value; state.page=1; renderCatalog(); });
  $('#clearFilters').addEventListener('click', () => { state.unit=state.code=state.frequency=state.query=''; $('#unitFilter').value=$('#codeFilter').value=$('#frequencyFilter').value=$('#globalSearch').value=''; state.page=1; renderCatalog(); });
  $('#printButton').addEventListener('click', () => window.print());
  $('#detailDialog').addEventListener('click', e => { if (e.target === $('#detailDialog')) $('#detailDialog').close(); });

  renderOverview(); renderFilters(); renderCatalog(); renderUnits(); renderIssues();
})();
