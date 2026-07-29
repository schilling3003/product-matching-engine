(function () {
  'use strict';

  // State
  const state = {
    mode: 'between',
    catalog: null,
    customer: null,
    catalogConfig: null,
    customerConfig: null,
    results: [],
    resultRows: [],
    processing: false
  };

  // Elements
  const els = {
    modeRadios: document.querySelectorAll('input[name="matchingMode"]'),
    catalogInput: document.getElementById('catalog-file'),
    customerInput: document.getElementById('customer-file'),
    customerUploadBox: document.getElementById('customer-upload-box'),
    catalogFileName: document.getElementById('catalog-file-name'),
    customerFileName: document.getElementById('customer-file-name'),
    columnsArea: document.getElementById('columns-area'),
    columnMappings: document.getElementById('column-mappings'),
    sizeOptions: document.getElementById('size-options'),
    enableSize: document.getElementById('enable-size'),
    sizeTolerance: document.getElementById('size-tolerance'),
    sizeToleranceValue: document.getElementById('size-tolerance-value'),
    runBtn: document.getElementById('run-button'),
    resetBtn: document.getElementById('reset-button'),
    progressArea: document.getElementById('progress-area'),
    progressFill: document.getElementById('progress-fill'),
    progressText: document.getElementById('progress-text'),
    resultsSection: document.getElementById('results-section'),
    summary: document.getElementById('summary'),
    resultsTable: document.getElementById('results-table'),
    noResults: document.getElementById('no-results'),
    downloadCsv: document.getElementById('download-csv'),
    downloadExcel: document.getElementById('download-excel')
  };

  // Init
  function init() {
    els.modeRadios.forEach(r => r.addEventListener('change', onModeChange));
    els.catalogInput.addEventListener('change', () => handleFile('catalog', els.catalogInput.files[0]));
    els.customerInput.addEventListener('change', () => handleFile('customer', els.customerInput.files[0]));
    els.enableSize.addEventListener('change', toggleSizeOptions);
    els.sizeTolerance.addEventListener('input', updateSizeToleranceLabel);
    els.runBtn.addEventListener('click', runMatch);
    els.resetBtn.addEventListener('click', reset);
    els.downloadCsv.addEventListener('click', () => download('csv'));
    els.downloadExcel.addEventListener('click', () => download('excel'));
    updateMode();
    updateSizeToleranceLabel();
  }

  function updateSizeToleranceLabel() {
    els.sizeToleranceValue.textContent = `${els.sizeTolerance.value}%`;
  }

  function toggleSizeOptions() {
    els.sizeOptions.style.display = els.enableSize.checked ? 'grid' : 'none';
  }

  function onModeChange() {
    state.mode = document.querySelector('input[name="matchingMode"]:checked').value;
    updateMode();
    state.results = [];
    hideResults();
  }

  function updateMode() {
    const isWithin = state.mode === 'within';
    els.customerUploadBox.style.display = isWithin ? 'none' : 'block';
    if (isWithin && els.customerInput.files[0]) {
      els.customerInput.value = '';
      state.customer = null;
      state.customerConfig = null;
      document.getElementById('customer-file-name').textContent = 'No file chosen';
    }
    renderColumnMappings();
    validateReady();
  }

  async function handleFile(type, file) {
    if (!file) return;
    const label = type === 'catalog' ? els.catalogFileName : els.customerFileName;
    label.textContent = file.name;
    try {
      const parsed = await ProductMatcher.parseFile(file);
      state[type] = parsed;
      if (state.mode === 'within' && type === 'catalog') {
        state.customer = parsed;
      }
      detectColumns(type);
      renderColumnMappings();
      validateReady();
    } catch (err) {
      alert(`Could not parse ${file.name}: ${err.message || err}`);
      console.error(err);
    }
  }

  function detectColumns(type) {
    const data = state[type];
    if (!data) return;
    const headers = data.headers;
    const rows = data.rows;

    const productCols = ProductMatcher.detectProductNameColumns(headers);
    const gtinCols = ProductMatcher.detectGtinColumns(headers, rows);
    const sizeCols = ProductMatcher.detectSizeColumns(headers);
    const restrictionCols = ProductMatcher.detectRestrictionColumns(headers, rows);

    const cfg = {
      productCols: productCols.length ? productCols : [headers[0]],
      sizeCol: sizeCols ? sizeCols.combined : null,
      sizeValueCol: sizeCols ? sizeCols.value : null,
      sizeUnitCol: sizeCols ? sizeCols.unit : null,
      gtinCols,
      restrictionCols,
      outputCols: []
    };

    if (type === 'catalog') state.catalogConfig = cfg;
    state.customerConfig = cfg;
  }

  function renderColumnMappings() {
    if (!state.catalog) {
      els.columnsArea.style.display = 'none';
      return;
    }
    els.columnsArea.style.display = 'block';

    const within = state.mode === 'within';
    const html = [];

    if (state.catalog) {
      const cfg = state.catalogConfig || detectColumnsReturn('catalog');
      html.push(buildMappingCard('Catalog file', 'catalog', cfg));
    }
    if (!within && state.customer) {
      const cfg = state.customerConfig || detectColumnsReturn('customer');
      html.push(buildMappingCard('Customer file', 'customer', cfg));
    }

    els.columnMappings.innerHTML = html.join('');

    // attach change handlers
    for (const input of els.columnMappings.querySelectorAll('input, select')) {
      input.addEventListener('change', updateConfigFromUI);
    }
    updateConfigFromUI();
  }

  function detectColumnsReturn(type) {
    if (!state[type]) return null;
    detectColumns(type);
    return type === 'catalog' ? state.catalogConfig : state.customerConfig;
  }

  function buildMappingCard(title, type, cfg) {
    const data = state[type === 'catalog' ? 'catalog' : 'customer'];
    const cols = data.headers;
    const within = state.mode === 'within';

    const productChecked = (col) => cfg.productCols.includes(col) ? 'checked' : '';
    const gtinChecked = (col) => cfg.gtinCols.includes(col) ? 'checked' : '';
    const restrictionChecked = (col) => cfg.restrictionCols.includes(col) ? 'checked' : '';
    const outputChecked = (col) => cfg.outputCols.includes(col) ? 'checked' : '';

    const sizeOptions = [`<option value="">None</option>`];
    for (const c of cols) {
      const selected = c === (cfg.sizeCol || '') ? 'selected' : '';
      sizeOptions.push(`<option value="${escapeHtml(c)}" ${selected}>${escapeHtml(c)}</option>`);
    }

    const valueOptions = [`<option value="">None</option>`];
    const unitOptions = [`<option value="">None</option>`];
    for (const c of cols) {
      valueOptions.push(`<option value="${escapeHtml(c)}" ${c === (cfg.sizeValueCol || '') ? 'selected' : ''}>${escapeHtml(c)}</option>`);
      unitOptions.push(`<option value="${escapeHtml(c)}" ${c === (cfg.sizeUnitCol || '') ? 'selected' : ''}>${escapeHtml(c)}</option>`);
    }

    let restrictionBlock = '';
    if (within) {
      restrictionBlock = `
        <div class="mapping-group">
          <span class="mapping-title">Restrict matches by columns</span>
          <div class="checkbox-list">
            ${cols.map(c => `<label><input type="checkbox" data-type="restriction" value="${escapeHtml(c)}" ${restrictionChecked(c)} /> ${escapeHtml(c)}</label>`).join('')}
          </div>
        </div>`;
    }

    return `
      <div class="mapping-card" data-card-type="${type}">
        <h4>${escapeHtml(title)}</h4>
        <div class="mapping-group">
          <span class="mapping-title">Product name / description columns <span class="required">*</span></span>
          <div class="checkbox-list">
            ${cols.map(c => `<label><input type="checkbox" data-type="product" value="${escapeHtml(c)}" ${productChecked(c)} /> ${escapeHtml(c)}</label>`).join('')}
          </div>
        </div>
        <div class="mapping-group">
          <span class="mapping-title">Size column (combined)</span>
          <select data-type="size">${sizeOptions.join('')}</select>
          <span class="field-hint">Or use separate value/unit columns below.</span>
          <div class="sub-selects">
            <label>Size value <select data-type="sizeValue">${valueOptions.join('')}</select></label>
            <label>Size unit <select data-type="sizeUnit">${unitOptions.join('')}</select></label>
          </div>
        </div>
        <div class="mapping-group">
          <span class="mapping-title">GTIN / UPC / Barcode columns</span>
          <div class="checkbox-list">
            ${cols.map(c => `<label><input type="checkbox" data-type="gtin" value="${escapeHtml(c)}" ${gtinChecked(c)} /> ${escapeHtml(c)}</label>`).join('')}
          </div>
        </div>
        ${restrictionBlock}
        <div class="mapping-group">
          <span class="mapping-title">Additional output columns</span>
          <div class="checkbox-list">
            ${cols.map(c => `<label><input type="checkbox" data-type="output" value="${escapeHtml(c)}" ${outputChecked(c)} /> ${escapeHtml(c)}</label>`).join('')}
          </div>
        </div>
      </div>
    `;
  }

  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function updateConfigFromUI() {
    const within = state.mode === 'within';
    state.catalogConfig = readMappingCard('catalog');
    if (!within) state.customerConfig = readMappingCard('customer');
    else state.customerConfig = state.catalogConfig;
    validateReady();
  }

  function readMappingCard(type) {
    const card = els.columnMappings.querySelector(`[data-card-type="${type}"]`);
    if (!card) return state[type === 'catalog' ? 'catalogConfig' : 'customerConfig'];
    const getChecked = (t) => Array.from(card.querySelectorAll(`[data-type="${t}"]:checked`)).map(i => i.value);
    return {
      productCols: getChecked('product'),
      sizeCol: card.querySelector('[data-type="size"]').value || null,
      sizeValueCol: card.querySelector('[data-type="sizeValue"]').value || null,
      sizeUnitCol: card.querySelector('[data-type="sizeUnit"]').value || null,
      gtinCols: getChecked('gtin'),
      restrictionCols: getChecked('restriction'),
      outputCols: getChecked('output')
    };
  }

  function validateReady() {
    const hasCatalog = !!state.catalog && state.catalogConfig && state.catalogConfig.productCols.length > 0;
    const hasCustomer = state.mode === 'within' || (!!state.customer && state.customerConfig && state.customerConfig.productCols.length > 0);
    els.runBtn.disabled = !(hasCatalog && hasCustomer) || state.processing;
  }

  function getSettings() {
    const focus = document.getElementById('text-focus').value;
    const focusMap = {
      exact: [0.1, 0.9],
      'mostly-spelling': [0.3, 0.7],
      balanced: [0.5, 0.5],
      'mostly-meaning': [0.7, 0.3],
      meaning: [0.9, 0.1]
    };
    const [baseTfidf, baseFuzzy] = focusMap[focus];
    return {
      matchingMode: state.mode,
      similarityThreshold: parseInt(document.getElementById('strictness').value, 10),
      maxMatchesPerProduct: parseInt(document.getElementById('max-matches').value, 10),
      textEnabled: document.getElementById('enable-text').checked,
      gtinEnabled: document.getElementById('enable-gtin').checked,
      sizeEnabled: document.getElementById('enable-size').checked,
      baseTfidf,
      baseFuzzy,
      sizeWeight: parseFloat(document.getElementById('size-importance').value),
      sizeTolerance: parseInt(document.getElementById('size-tolerance').value, 10),
      removeStopWords: document.getElementById('remove-stop-words').checked,
      caseSensitive: document.getElementById('case-sensitive').checked,
      includeSizeInText: false
    };
  }

  function runMatch() {
    if (state.processing) return;
    state.processing = true;
    els.runBtn.disabled = true;
    els.progressArea.style.display = 'block';
    els.resultsSection.style.display = 'none';
    setProgress(0, 'Preparing...');

    setTimeout(async () => {
      try {
        await doMatch();
      } catch (err) {
        console.error(err);
        alert('Error during matching: ' + (err.message || err));
      } finally {
        state.processing = false;
        validateReady();
        els.progressArea.style.display = 'none';
      }
    }, 50);
  }

  async function doMatch() {
    const settings = getSettings();
    const catalogItems = ProductMatcher.preprocess(state.catalog.rows, state.catalogConfig, settings);
    const customerItems = state.mode === 'within'
      ? catalogItems
      : ProductMatcher.preprocess(state.customer.rows, state.customerConfig, settings);

    if (!catalogItems.length || !customerItems.length) {
      alert('No rows to match.');
      return;
    }

    setProgress(0.05, 'Matching...');

    const matches = ProductMatcher.computeMatch(catalogItems, customerItems, settings, (p, cur, total) => {
      const pct = 5 + Math.round(p * 90);
      setProgress(pct / 100, `Matching product ${cur} of ${total}...`);
    });

    setProgress(0.95, 'Building results...');

    const resultRows = buildResultRows(matches, catalogItems, customerItems, settings);
    state.resultRows = resultRows;
    state.results = matches;

    renderResults(resultRows, settings);
    setProgress(1, 'Done');
    els.resultsSection.style.display = 'block';
    if (resultRows.length === 0) els.noResults.style.display = 'block';
    else els.noResults.style.display = 'none';
    setTimeout(() => els.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50);
  }

  function buildResultRows(matches, catalogItems, customerItems, settings) {
    const within = settings.matchingMode === 'within';
    const rows = [];
    const custCfg = state.mode === 'within' ? state.catalogConfig : state.customerConfig;
    const catCfg = state.catalogConfig;
    const custProductCol = custCfg.productCols[0];
    const catProductCol = catCfg.productCols[0];

    for (const m of matches) {
      const cust = customerItems[m.i].original;
      const cat = catalogItems[m.j].original;
      const row = {};
      if (within) {
        row['Product 1'] = cust[custProductCol];
        row['Product 2'] = cat[catProductCol];
      } else {
        row['Customer Product'] = cust[custProductCol];
        row['Catalog Product'] = cat[catProductCol];
      }
      row['Confidence Score'] = `${m.combined.toFixed(2)}%`;
      if (settings.textEnabled) {
        row['TF-IDF Score'] = `${m.tfidf.toFixed(2)}%`;
        row['Fuzzy Score'] = `${m.fuzzy.toFixed(2)}%`;
      }
      if (settings.gtinEnabled) {
        row['GTIN Score'] = `${m.gtinScore.toFixed(2)}%`;
        if (m.gtins && m.gtins.length) row['Matching GTINs'] = m.gtins.join(', ');
        if (m.gtinType) row['GTIN Match Type'] = m.gtinType;
      }
      if (settings.sizeEnabled) {
        row['Size Score'] = `${m.sizeSim.toFixed(2)}%`;
        row[within ? 'Product 1 Size' : 'Customer Size'] = customerItems[m.i].stdSize;
        row[within ? 'Product 2 Size' : 'Catalog Size'] = catalogItems[m.j].stdSize;
      }
      for (const c of custCfg.outputCols) row[(within ? 'Product 1 ' : 'Customer ') + c] = cust[c];
      for (const c of catCfg.outputCols) row[(within ? 'Product 2 ' : 'Catalog ') + c] = cat[c];
      rows.push(row);
    }
    return rows;
  }

  function renderResults(rows, settings) {
    els.summary.innerHTML = `
      <div class="metric"><span class="metric-value">${rows.length}</span><span class="metric-label">Matches</span></div>
      <div class="metric"><span class="metric-value">${settings.similarityThreshold}%</span><span class="metric-label">Threshold</span></div>
    `;

    const table = els.resultsTable;
    if (!rows.length) {
      table.querySelector('thead').innerHTML = '';
      table.querySelector('tbody').innerHTML = '';
      return;
    }
    const keys = Object.keys(rows[0]);
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = `<tr>${keys.map(k => `<th data-key="${escapeHtml(k)}">${escapeHtml(k)} <span class="sort-indicator">↕</span></th>`).join('')}</tr>`;
    tbody.innerHTML = rows.map(r => `<tr>${keys.map(k => `<td>${escapeHtml(String(r[k] ?? ''))}</td>`).join('')}</tr>`).join('');

    // Sorting
    let sortKey = null;
    let sortDir = 1;
    for (const th of thead.querySelectorAll('th')) {
      th.addEventListener('click', () => {
        const key = th.dataset.key;
        if (sortKey === key) sortDir *= -1;
        else { sortKey = key; sortDir = 1; }
        const sorted = [...state.resultRows].sort((a, b) => {
          const av = a[key], bv = b[key];
          if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
          return String(av).localeCompare(String(bv)) * sortDir;
        });
        renderTableBody(sorted, keys);
        thead.querySelectorAll('th').forEach(t => t.querySelector('.sort-indicator').textContent = '↕');
        th.querySelector('.sort-indicator').textContent = sortDir === 1 ? '↑' : '↓';
      });
    }
  }

  function renderTableBody(rows, keys) {
    const tbody = els.resultsTable.querySelector('tbody');
    tbody.innerHTML = rows.map(r => `<tr>${keys.map(k => `<td>${escapeHtml(String(r[k] ?? ''))}</td>`).join('')}</tr>`).join('');
  }

  function setProgress(p, text) {
    const pct = Math.max(0, Math.min(100, p * 100));
    els.progressFill.style.width = `${pct}%`;
    els.progressText.textContent = text;
  }

  function hideResults() {
    els.resultsSection.style.display = 'none';
  }

  function reset() {
    state.catalog = null;
    state.customer = null;
    state.catalogConfig = null;
    state.customerConfig = null;
    state.results = [];
    state.resultRows = [];
    els.catalogInput.value = '';
    els.customerInput.value = '';
    els.catalogFileName.textContent = 'No file chosen';
    els.customerFileName.textContent = 'No file chosen';
    els.columnsArea.style.display = 'none';
    hideResults();
    setProgress(0, '');
    els.progressArea.style.display = 'none';
    updateMode();
  }

  function download(format) {
    if (!state.resultRows.length) return;
    if (format === 'csv') {
      const csv = Papa.unparse(state.resultRows);
      downloadBlob(csv, 'product_matches.csv', 'text/csv');
    } else {
      const ws = XLSX.utils.json_to_sheet(state.resultRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Matches');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      downloadBlob(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'product_matches.xlsx');
    }
  }

  function downloadBlob(content, fileName, type) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: type || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  init();
})();
