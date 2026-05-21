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
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // bỏ dấu
      .replace(/[^a-z0-9\s]/g, ' ');
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
        <a href="${t.fileUrl}" class="btn-download" target="_blank" download>
          ⬇️ Tải
        </a>
      </div>
    `;
    return card;
  }

  // ===== Filter & Render =====
  function filterAndRender() {
    const q = normalize(currentQuery);
    const qWords = q ? q.split(/\s+/).filter(w => w.length > 0) : [];

    const filtered = TEMPLATES.filter(t => {
      // Category filter
      if (currentCategory !== 'all' && t.category !== currentCategory) return false;

      // Search filter
      if (qWords.length > 0) {
        const haystack = normalize(`${t.name} ${t.nameEn} ${t.description} ${t.tags.join(' ')}`);
        return qWords.every(word => haystack.includes(word));
      }

      return true;
    });

    // Render
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
