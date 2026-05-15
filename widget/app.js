// ─── Поля в Битрикс24 ────────────────────────────────────────────────────────
const FIELD_CITIES    = 'UF_CRM_CITIES_AI';
const FIELD_REGIONS   = 'UF_CRM_REGIONS_AI';
const FIELD_DISTRICTS = 'UF_CRM_DISTRICTS_AI';

// ─── Данные ──────────────────────────────────────────────────────────────────
let citiesDB    = [];
let regionToFO  = {};
let foToRegions = {};

const ALL_FOS = ['ЦФО','СЗФО','ЮФО','СКФО','ПФО','УФО','СФО','ДФО'];

// ─── Состояние выбора ────────────────────────────────────────────────────────
let selectedFOs     = new Set(); // все выбранные ФО (включая авто)
let explicitFOs     = new Set(); // только те ФО, что нажали вручную — для фильтрации поиска
let selectedRegions = new Set(); // все выбранные регионы (включая авто)
let explicitRegions = new Set(); // только те регионы, что выбрали вручную — для фильтрации поиска городов
let selectedCities  = new Set();
let currentDealId   = null;

// ─── DOM ─────────────────────────────────────────────────────────────────────
const foButtonsEl    = document.getElementById('foButtons');
const foTagsEl       = document.getElementById('foTags');
const regionInput    = document.getElementById('regionInput');
const regionDropdown = document.getElementById('regionDropdown');
const regionTagsEl   = document.getElementById('regionTags');
const cityInput      = document.getElementById('cityInput');
const cityDropdown   = document.getElementById('cityDropdown');
const cityTagsEl     = document.getElementById('cityTags');
const saveBtn        = document.getElementById('saveBtn');
const statusEl       = document.getElementById('status');

// ─── Вспомогательные элементы ─────────────────────────────────────────────
const loadingEl  = document.getElementById('loading');
const mainEl     = document.getElementById('mainContent');
const errorMsgEl = document.getElementById('errorMsg');

function showMain() {
  loadingEl.style.display  = 'none';
  mainEl.style.display     = 'block';
}

function showError(msg) {
  loadingEl.style.display  = 'none';
  errorMsgEl.style.display = 'block';
  errorMsgEl.textContent   = '❌ ' + msg;
  console.error(msg);
}

// ─── Инициализация через BX24 ─────────────────────────────────────────────
console.log('[GeoWidget] Старт. BX24:', typeof BX24);

if (typeof BX24 === 'undefined') {
  showError('BX24.js не загружен. Виджет должен открываться внутри Битрикс24.');
} else {

  const bx24Timeout = setTimeout(() => {
    showError('BX24.init() не ответил за 6 сек — проверьте регистрацию Local App в Битрикс24.');
  }, 6000);

BX24.init(function () {
  clearTimeout(bx24Timeout);
  console.log('[GeoWidget] BX24.init сработал');

  // Получаем ID текущей сделки
  const placementInfo = BX24.placement.info();
  currentDealId = placementInfo && placementInfo.options ? placementInfo.options.ID : null;
  console.log('[GeoWidget] placement info:', JSON.stringify(placementInfo));
  console.log('[GeoWidget] Deal ID:', currentDealId);

  // Загружаем базу городов
  fetch('../cities_russia.json')
    .then(r => r.json())
    .then(data => {
      citiesDB = data.sort((a, b) => a.city.localeCompare(b.city, 'ru'));

      data.forEach(({ city, region, district }) => {
        regionToFO[region] = district;
        if (!foToRegions[district]) foToRegions[district] = new Set();
        foToRegions[district].add(region);
      });

      Object.keys(foToRegions).forEach(fo => {
        foToRegions[fo] = [...foToRegions[fo]].sort((a, b) => a.localeCompare(b, 'ru'));
      });

      renderFOButtons();
      showMain();

      // Загружаем уже сохранённые значения из сделки
      if (currentDealId) {
        loadExistingValues();
      }
    })
    .catch(err => showError('Не удалось загрузить базу городов: ' + err.message));

}); // BX24.init

} // typeof BX24

// ─── Загрузка текущих значений из сделки ─────────────────────────────────
function loadExistingValues() {
  BX24.callMethod('crm.deal.get', { id: currentDealId }, function (result) {
    if (result.error()) {
      console.warn('Не удалось загрузить данные сделки:', result.error());
      return;
    }

    const deal = result.data();

    const cities    = (deal[FIELD_CITIES]    || '').split(',').map(s => s.trim()).filter(Boolean);
    const regions   = (deal[FIELD_REGIONS]   || '').split(',').map(s => s.trim()).filter(Boolean);
    const districts = (deal[FIELD_DISTRICTS] || '').split(',').map(s => s.trim()).filter(Boolean);

    cities.forEach(c => selectedCities.add(c));
    regions.forEach(r => {
      if (regionToFO[r]) {
        selectedRegions.add(r);
        explicitRegions.add(r);
      }
    });
    districts.forEach(fo => {
      if (ALL_FOS.includes(fo)) {
        selectedFOs.add(fo);
        explicitFOs.add(fo); // сохранённые ФО считаем явными
      }
    });

    if (cities.length || regions.length || districts.length) {
      renderAll();
    }
  });
}

// ─── Рендер кнопок ФО ────────────────────────────────────────────────────────
function renderFOButtons() {
  foButtonsEl.innerHTML = ALL_FOS.map(fo => `
    <button class="fo-btn ${selectedFOs.has(fo) ? 'active' : ''}" data-fo="${fo}">${fo}</button>
  `).join('');
}

foButtonsEl.addEventListener('click', e => {
  const btn = e.target.closest('.fo-btn');
  if (!btn) return;
  const fo = btn.dataset.fo;
  if (selectedFOs.has(fo)) removeFO(fo);
  else addFO(fo);
});

// ─── Добавление / удаление ФО ────────────────────────────────────────────────
function addFO(fo) {
  selectedFOs.add(fo);
  explicitFOs.add(fo);   // нажали вручную
  (foToRegions[fo] || []).forEach(r => selectedRegions.add(r));
  renderAll();
}

function removeFO(fo) {
  selectedFOs.delete(fo);
  explicitFOs.delete(fo);
  renderAll();
}

// ─── Добавление / удаление Региона ───────────────────────────────────────────
function addRegion(region) {
  selectedRegions.add(region);
  explicitRegions.add(region); // выбрали вручную
  const fo = regionToFO[region];
  if (fo) selectedFOs.add(fo); // авто — не попадает в explicitFOs
  renderAll();
}

function removeRegion(region) {
  selectedRegions.delete(region);
  explicitRegions.delete(region);
  renderAll();
}

// ─── Добавление / удаление Города ────────────────────────────────────────────
function addCity(city, region, fo) {
  selectedCities.add(city);
  if (region) selectedRegions.add(region);
  if (fo) selectedFOs.add(fo);
  renderAll();
}

function removeCity(city) {
  selectedCities.delete(city);
  renderAll();
}

// ─── Рендер всех тегов и кнопок ──────────────────────────────────────────────
function renderAll() {
  renderFOButtons();

  foTagsEl.innerHTML = [...selectedFOs].map(fo => `
    <span class="tag fo">
      ${fo}
      <span class="tag-remove" data-fo="${fo}">×</span>
    </span>
  `).join('');

  if (selectedRegions.size === 0) {
    regionTagsEl.innerHTML = '';
  } else {
    const groups = {};
    [...selectedRegions].sort((a, b) => a.localeCompare(b, 'ru')).forEach(r => {
      const fo = regionToFO[r] || 'Другие';
      if (!groups[fo]) groups[fo] = [];
      groups[fo].push(r);
    });
    const foOrder = [...ALL_FOS, 'Другие'];
    const sortedFOs = foOrder.filter(fo => groups[fo]);

    regionTagsEl.innerHTML = `<div class="region-groups">` +
      sortedFOs.map(fo => `
        <div class="region-group">
          <div class="region-group-header">
            <span class="region-group-fo">${fo}</span>
            <span class="region-group-line"></span>
          </div>
          <div class="region-group-tags">
            ${groups[fo].map(r => `
              <span class="tag reg">
                ${r}
                <span class="tag-remove" data-region="${r}">×</span>
              </span>
            `).join('')}
          </div>
        </div>
      `).join('') +
    `</div>`;
  }

  cityTagsEl.innerHTML = [...selectedCities].sort((a, b) => a.localeCompare(b, 'ru')).map(c => `
    <span class="tag city">
      ${c}
      <span class="tag-remove" data-city="${c}">×</span>
    </span>
  `).join('');

  const hasAny = selectedFOs.size > 0 || selectedRegions.size > 0 || selectedCities.size > 0;
  saveBtn.disabled = !hasAny;
}

// ─── Клики по тегам (удаление) ───────────────────────────────────────────────
foTagsEl.addEventListener('click', e => {
  const btn = e.target.closest('.tag-remove[data-fo]');
  if (btn) removeFO(btn.dataset.fo);
});

regionTagsEl.addEventListener('click', e => {
  const btn = e.target.closest('.tag-remove[data-region]');
  if (btn) removeRegion(btn.dataset.region);
});

cityTagsEl.addEventListener('click', e => {
  const btn = e.target.closest('.tag-remove[data-city]');
  if (btn) removeCity(btn.dataset.city);
});

// ─── Автокомплит: Регион ─────────────────────────────────────────────────────
regionInput.addEventListener('input', () => {
  const q = regionInput.value.trim().toLowerCase();
  if (q.length < 1) { regionDropdown.style.display = 'none'; return; }

  const allRegions = Object.keys(regionToFO).sort((a, b) => a.localeCompare(b, 'ru'));
  const filterRegion = r => {
    if (!r.toLowerCase().includes(q)) return false;
    if (selectedRegions.has(r)) return false;
    if (explicitFOs.size > 0 && !explicitFOs.has(regionToFO[r])) return false;
    return true;
  };
  // Сначала те, что начинаются с введённых букв, потом остальные
  const matches = [
    ...allRegions.filter(r => filterRegion(r) && r.toLowerCase().startsWith(q)),
    ...allRegions.filter(r => filterRegion(r) && !r.toLowerCase().startsWith(q)),
  ].slice(0, 25);

  regionDropdown.innerHTML = matches.length
    ? matches.map(r => `
        <div class="dropdown-item" data-region="${r}" data-fo="${regionToFO[r]}">
          <div class="item-main">${r}</div>
          <div class="item-sub">${regionToFO[r]}</div>
        </div>
      `).join('')
    : '<div class="no-results">Регион не найден</div>';

  regionDropdown.style.display = 'block';
});

regionDropdown.addEventListener('click', e => {
  const item = e.target.closest('.dropdown-item');
  if (!item) return;
  addRegion(item.dataset.region);
  regionInput.value = '';
  regionDropdown.style.display = 'none';
  regionInput.focus();
});

// ─── Автокомплит: Город ──────────────────────────────────────────────────────
cityInput.addEventListener('input', () => {
  const q = cityInput.value.trim().toLowerCase();
  if (q.length < 1) { cityDropdown.style.display = 'none'; return; }

  const matches = citiesDB
    .filter(c => c.city.toLowerCase().startsWith(q))
    .concat(citiesDB.filter(c =>
      !c.city.toLowerCase().startsWith(q) && c.city.toLowerCase().includes(q)
    ))
    .filter(c => {
      if (selectedCities.has(c.city)) return false;
      // если выбраны ФО — только города из этих ФО
      if (explicitFOs.size > 0 && !explicitFOs.has(c.district)) return false;
      // фильтруем только по регионам, выбранным вручную (не авто)
      if (explicitRegions.size > 0 && !explicitRegions.has(c.region)) return false;
      return true;
    })
    .slice(0, 30);

  cityDropdown.innerHTML = matches.length
    ? matches.map(c => `
        <div class="dropdown-item" data-city="${c.city}" data-region="${c.region}" data-fo="${c.district}">
          <div class="item-main">${c.city}</div>
          <div class="item-sub">${c.region} · ${c.district}</div>
        </div>
      `).join('')
    : '<div class="no-results">Город не найден</div>';

  cityDropdown.style.display = 'block';
});

cityDropdown.addEventListener('click', e => {
  const item = e.target.closest('.dropdown-item');
  if (!item) return;
  addCity(item.dataset.city, item.dataset.region, item.dataset.fo);
  cityInput.value = '';
  cityDropdown.style.display = 'none';
  cityInput.focus();
});

// ─── Закрытие дропдаунов по клику вне ────────────────────────────────────────
document.addEventListener('click', e => {
  if (!e.target.closest('#regionInput') && !e.target.closest('#regionDropdown'))
    regionDropdown.style.display = 'none';
  if (!e.target.closest('#cityInput') && !e.target.closest('#cityDropdown'))
    cityDropdown.style.display = 'none';
});

// ─── Сохранение в Битрикс24 через BX24.js ────────────────────────────────────
saveBtn.addEventListener('click', () => {
  if (!currentDealId) {
    showStatus('Не удалось определить ID сделки', true);
    return;
  }

  const fields = {
    [FIELD_CITIES]:    [...selectedCities].sort((a, b) => a.localeCompare(b, 'ru')).join(', '),
    [FIELD_REGIONS]:   [...selectedRegions].sort((a, b) => a.localeCompare(b, 'ru')).join(', '),
    [FIELD_DISTRICTS]: [...selectedFOs].join(', '),
  };

  saveBtn.disabled = true;
  saveBtn.textContent = 'Сохраняю...';

  BX24.callMethod('crm.deal.update', { id: currentDealId, fields }, function (result) {
    if (result.error()) {
      showStatus('Ошибка: ' + result.error(), true);
    } else {
      showStatus('✅ Успешно сохранено!');
      // Обновляем карточку сделки
      BX24.reloadWindow();
    }
    saveBtn.disabled = false;
    saveBtn.textContent = 'Сохранить в сделку';
  });
});

// ─── Вспомогательные ─────────────────────────────────────────────────────────
function showStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.className = 'status' + (isError ? ' error' : '');
  statusEl.style.display = 'block';
  setTimeout(() => { statusEl.style.display = 'none'; }, 4000);
}
