/* ===== văn phòng — App Script ===== */
(function() {
  'use strict';

  const templateContainer = document.getElementById('templatesContainer');
  const searchInput = document.getElementById('searchInput');
  const searchClear = document.getElementById('searchClear');
  const categoryChips = document.querySelectorAll('.chip');
  const resultCount = document.getElementById('resultCount');

  let currentCategory = 'all';
  let currentQuery = '';

  // ===== Normalize text for search =====
  function normalize(str) {
    return str.toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9\s]/g, ' ');
  }

  // ===== Detail modal =====
  function showModal(t) {
    const existing = document.getElementById('templateModal');
    if (existing) existing.remove();

    const badgeClass = t.format === 'excel' ? 'badge-excel' : 'badge-word';
    const badgeLabel = t.format === 'excel' ? 'Excel' : 'Word';
    const tagsHTML = t.tags.map(tag => `<span class="tag">${tag}</span>`).join('');

    const overlay = document.createElement('div');
    overlay.id = 'templateModal';
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal">
        <button class="modal-close" id="modalClose">✕</button>
        <div class="modal-body">
          <div class="modal-icon">${t.icon}</div>
          <h2 class="modal-title">${t.name}</h2>
          <p class="modal-title-en">${t.nameEn}</p>
          <span class="template-badge ${badgeClass}" style="margin:8px 0 16px">${badgeLabel}</span>
          <p class="modal-desc">${t.description}</p>
          <div class="modal-tags">${tagsHTML}</div>
          <div class="modal-meta">
            <span>📂 ${t.subcategory} · ${t.source}</span>
          </div>
          <div class="modal-actions">
            <a href="${t.fileUrl}" class="btn-download" target="_blank" download>
              ⬇️ Tải file
            </a>
            <button class="btn-secondary" id="modalClose2">Đóng</button>
          </div>
        </div>
        <div class="modal-preview">
          <div class="preview-wrap">
            <div class="preview-badge ${badgeClass}">${badgeLabel}</div>
            <img src="/van-phong/screenshots/screen_${t.id}.png" alt="${t.name}" class="preview-img" loading="lazy" />
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    document.body.style.overflow = 'hidden';

    // Close handlers
    function closeModal() { overlay.remove(); document.body.style.overflow = ''; }
    document.getElementById('modalClose').addEventListener('click', closeModal);
    document.getElementById('modalClose2').addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeModal(); });
    document.addEventListener('keydown', function escHandler(e) {
      if (e.key === 'Escape') { closeModal(); document.removeEventListener('keydown', escHandler); }
    });
  }

  // ===== Render template card =====
  function createCard(t) {
    const badgeClass = t.format === 'excel' ? 'badge-excel' : 'badge-word';
    const badgeLabel = t.format === 'excel' ? 'Excel' : 'Word';
    const tagsHTML = t.tags.slice(0, 4).map(tag =>
      `<span class="tag">${tag}</span>`
    ).join('');

    const card = document.createElement('div');
    card.className = 'template-card';
    card.innerHTML = `
      <div class="template-card-top">
        <div class="template-icon">${t.icon}</div>
        <div class="template-text">
          <div class="template-name">
            ${t.name}
            <span class="template-name-en">${t.nameEn}</span>
          </div>
          <div class="template-desc">${t.description}</div>
          <div class="template-tags">${tagsHTML}</div>
        </div>
      </div>
      <div class="template-card-footer">
        <div style="display:flex;align-items:center;gap:10px">
          <span class="template-badge ${badgeClass}">${badgeLabel}</span>
          <span class="template-source">${t.source}</span>
        </div>
        <button class="btn-detail">📋 Chi tiết</button>
      </div>
    `;

    // Click handlers
    card.addEventListener('click', (e) => {
      if (!e.target.closest('.btn-detail') && !e.target.closest('.btn-download')) {
        showModal(t);
      }
    });
    card.querySelector('.btn-detail').addEventListener('click', (e) => {
      e.stopPropagation();
      showModal(t);
    });

    return card;
  }

  // ===== Filter & Render =====
  function filterAndRender() {
    const q = normalize(currentQuery);
    const qWords = q ? q.split(/\s+/).filter(w => w.length > 0) : [];

    const filtered = TEMPLATES.filter(t => {
      if (currentCategory !== 'all' && t.category !== currentCategory) return false;
      if (qWords.length > 0) {
        const haystack = normalize(`${t.name} ${t.nameEn} ${t.description} ${t.tags.join(' ')}`);
        return qWords.every(word => haystack.includes(word));
      }
      return true;
    });

    if (filtered.length === 0) {
      templateContainer.innerHTML = `
        <div class="no-results">
          <div class="no-results-icon">🔍</div>
          <div class="no-results-text">Không tìm thấy template nào</div>
          <div class="no-results-hint">Thử từ khóa khác hoặc bỏ lọc thể loại</div>
        </div>
      `;
    } else {
      const grid = document.createElement('div');
      grid.className = 'template-grid';
      filtered.forEach(t => grid.appendChild(createCard(t)));
      templateContainer.innerHTML = '';
      templateContainer.appendChild(grid);
    }
    resultCount.textContent = `${filtered.length} mẫu biểu`;
  }

  // ===== Category click =====
  categoryChips.forEach(chip => {
    chip.addEventListener('click', () => {
      categoryChips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      currentCategory = chip.dataset.category;
      filterAndRender();
    });
  });

  // ===== Search =====
  let searchTimer = null;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      currentQuery = searchInput.value;
      searchClear.classList.toggle('hidden', !currentQuery);
      filterAndRender();
    }, 200);
  });

  searchClear.addEventListener('click', () => {
    searchInput.value = '';
    currentQuery = '';
    searchClear.classList.add('hidden');
    filterAndRender();
    searchInput.focus();
  });

  // ===== Initial render =====
  filterAndRender();
})();
