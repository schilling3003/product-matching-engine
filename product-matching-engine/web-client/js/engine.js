const ProductMatcher = (function () {
  'use strict';

  const STOP_WORDS = new Set([
    'a','an','the','and','or','but','in','on','at','to','of','with','by','is','am','are','was','were','be','been','being','have','has','had','do','does','did','case','pack','brand','product','item','food','natural','premium','quality'
  ]);

  const UNIT_CONVERSION_MAP = {
    'fl oz': 29.5735, 'gallon': 3785.41, 'gallon': 3785.41, 'oz': 28.35,
    'lb': 453.592, 'kg': 1000, 'l': 1000, 'ml': 1, 'g': 1
  };
  const UNIT_KEYS = Object.keys(UNIT_CONVERSION_MAP).sort((a, b) => b.length - a.length);

  // ===== File parsing =====

  async function parseFile(file) {
    const name = file.name.toLowerCase();
    if (name.endsWith('.csv')) {
      return parseCSV(file);
    }
    return parseExcel(file);
  }

  function parseCSV(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const headers = results.meta.fields || [];
          const rows = results.data.map((row) => objectToStrings(row));
          resolve({ headers, rows });
        },
        error: (err) => reject(err)
      });
    });
  }

  async function parseExcel(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (json.length === 0) return { headers: [], rows: [] };
    const headers = json[0].map(String);
    const rows = json.slice(1).map((arr) => {
      const row = {};
      headers.forEach((h, i) => { row[h] = arr[i] ?? ''; });
      return objectToStrings(row);
    });
    return { headers, rows };
  }

  function objectToStrings(obj) {
    const out = {};
    for (const key of Object.keys(obj)) {
      const v = obj[key];
      out[key] = v == null ? '' : String(v);
    }
    return out;
  }

  // ===== Column detection =====

  const PRODUCT_NAME_HINTS = ['product_name', 'description', 'short_name', 'long_name', 'name', 'title', 'product', 'print name', 'item name'];
  const GTIN_HINTS = ['gtin','upc','ean','barcode','bar_code','product_code','item_code','sku','code','id'];
  const GTIN_EXCLUDE = ['type','description','name','category'];
  const SIZE_HINTS = ['size', 'weight', 'volume', 'package_size'];
  const SIZE_VALUE_HINTS = ['size_value', 'value', 'qty', 'quantity', 'net_weight'];
  const SIZE_UNIT_HINTS = ['size_unit', 'unit', 'uom', 'unit_of_measure'];
  const RESTRICTION_HINTS = ['category','commodity','department','type','class','group','segment','division','section','line','family','brand','supplier','vendor','product_type','product_category','item_type','item_category','category_name','commodity_code','dept','dept_name'];

  function normalizeHeader(h) {
    return h.toLowerCase().replace(/[_\s]+/g, ' ').trim();
  }

  function detectProductNameColumns(headers) {
    return headers.filter(h => {
      const n = normalizeHeader(h);
      return PRODUCT_NAME_HINTS.some(hint => n.includes(hint.replace(/_/g, ' '))) || n === 'name' || n === 'title';
    });
  }

  function detectSizeColumns(headers) {
    const value = headers.find(h => SIZE_VALUE_HINTS.some(hint => normalizeHeader(h) === hint.replace(/_/g, ' ')));
    const unit = headers.find(h => SIZE_UNIT_HINTS.some(hint => normalizeHeader(h) === hint.replace(/_/g, ' ')));
    if (value && unit) return { combined: null, value, unit };

    const byHeader = headers.filter(h => {
      const tokens = normalizeHeader(h).split(/\s+/);
      return SIZE_HINTS.some(hint => tokens.includes(hint.replace(/_/g, ' ')));
    });
    if (byHeader.length) return { combined: byHeader[0], value: null, unit: null };

    if (value) return { combined: null, value, unit };
    return null;
  }

  function detectGtinColumns(headers, rows) {
    const candidates = headers.filter(h => {
      const n = normalizeHeader(h);
      const matches = GTIN_HINTS.some(hint => n.includes(hint.replace(/_/g, ' ')));
      const excludes = GTIN_EXCLUDE.some(ex => n.includes(ex));
      return matches && !excludes;
    });
    const validated = candidates.filter(h => {
      const sample = rows.slice(0, 10).map(r => String(r[h] ?? ''));
      let good = 0;
      for (const v of sample) {
        const digits = v.replace(/\D/g, '');
        if (digits.length >= 7) good++;
      }
      return good >= sample.length * 0.5 || good > 0;
    });
    return validated.length ? validated : candidates;
  }

  function detectRestrictionColumns(headers, rows) {
    return headers.filter(h => {
      const n = normalizeHeader(h);
      if (RESTRICTION_HINTS.some(hint => n.includes(hint.replace(/_/g, ' ')))) return true;
      const unique = new Set();
      for (const r of rows) unique.add(r[h]);
      const uniq = unique.size;
      return uniq >= 2 && uniq <= 50 && uniq < rows.length * 0.5;
    });
  }

  // ===== Text & size preprocessing =====

  function cleanText(text, removeStopWords, caseSensitive) {
    let s = String(text ?? '');
    if (!caseSensitive) s = s.toLowerCase();
    s = s.replace(/[^\w\s]/g, ' ');
    s = s.replace(/\s+/g, ' ').trim();
    if (removeStopWords) {
      s = s.split(/\s+/).filter(w => w && !STOP_WORDS.has(w.toLowerCase())).join(' ');
    }
    return s;
  }

  function standardizeSize(text) {
    const s = String(text ?? '').toLowerCase().replace(/,/g, '');
    for (const unit of UNIT_KEYS) {
      const pattern = new RegExp('(\\d*\\.?\\d+)\\s*' + escapeRegex(unit) + '\\b', 'i');
      const m = s.match(pattern);
      if (m) {
        const val = parseFloat(m[1]);
        if (!isNaN(val)) return `${(val * UNIT_CONVERSION_MAP[unit]).toFixed(1)}g`;
      }
    }
    return '';
  }

  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function extractSizeNumber(stdSize) {
    const m = String(stdSize).match(/(\d*\.?\d+)/);
    return m ? parseFloat(m[1]) : NaN;
  }

  function calculateSizeSimilarity(size1, size2, tolerance) {
    const v1 = extractSizeNumber(size1);
    const v2 = extractSizeNumber(size2);
    if (isNaN(v1) || isNaN(v2)) return 0;
    if (v1 === 0 && v2 === 0) return 100;
    if (v1 === 0 || v2 === 0) return 0;
    const larger = Math.max(v1, v2);
    const smaller = Math.min(v1, v2);
    const pctDiff = ((larger - smaller) / larger) * 100;
    if (pctDiff >= tolerance) return 0;
    return 100 * (1 - pctDiff / tolerance);
  }

  // ===== GTIN processing =====

  function calculateGtinCheckDigit(base) {
    if (!base || !/^\d+$/.test(base)) return '';
    let padded = base;
    const l = base.length;
    if ([7,11,13].includes(l)) padded = base.padStart(l, '0');
    else return '';
    let total = 0;
    for (let i = 0; i < padded.length; i++) {
      const digit = parseInt(padded[padded.length - 1 - i], 10);
      const weight = i % 2 === 0 ? 3 : 1;
      total += digit * weight;
    }
    const check = (10 - (total % 10)) % 10;
    return padded + check;
  }

  function validateGtin(gtin) {
    if (!gtin || !/^\d+$/.test(gtin)) return false;
    if (![8,12,13,14].includes(gtin.length)) return false;
    const base = gtin.slice(0, -1);
    return calculateGtinCheckDigit(base) === gtin;
  }

  function correctGtinCheckDigit(gtin) {
    if (!gtin || !/^\d+$/.test(gtin) || ![8,12,13,14].includes(gtin.length)) return '';
    return calculateGtinCheckDigit(gtin.slice(0, -1));
  }

  function extractUnitGtinFromCase(caseGtin) {
    if (!caseGtin || caseGtin.length !== 14 || !/^\d+$/.test(caseGtin)) return '';
    const base = caseGtin.slice(1, 13);
    const unit = calculateGtinCheckDigit(base);
    return unit.length === 13 ? unit : '';
  }

  function normalizeAndGenerateVariants(value) {
    const variants = {};
    const cleaned = String(value ?? '').replace(/\D/g, '');
    if (!cleaned) return variants;
    const length = cleaned.length;

    // original padded to 14
    if (length <= 14) variants[cleaned.padStart(14, '0')] = 'original';

    if ([8,12,13,14].includes(length)) {
      const corrected = correctGtinCheckDigit(cleaned);
      if (corrected && corrected !== cleaned) variants[corrected.padStart(14, '0')] = 'corrected';
    }

    if ([7,11].includes(length)) {
      const complete = calculateGtinCheckDigit(cleaned);
      if (complete) variants[complete.padStart(14, '0')] = 'missing_check';
    }

    if (length === 12) {
      const ean13 = calculateGtinCheckDigit(cleaned);
      if (ean13 && ean13 !== cleaned) variants[ean13.padStart(14, '0')] = 'missing_check';
    }

    if (length === 14) {
      const unit = extractUnitGtinFromCase(cleaned);
      if (unit) {
        variants[unit.padStart(14, '0')] = 'case_to_unit';
        const unitCorrected = correctGtinCheckDigit(unit);
        if (unitCorrected && unitCorrected !== unit) variants[unitCorrected.padStart(14, '0')] = 'case_to_unit';
      }
    }

    for (let i = 0; i <= 14 - length; i++) {
      const alt = '0'.repeat(i) + cleaned + '0'.repeat(14 - length - i);
      if (alt.length === 14 && alt !== cleaned.padStart(14, '0') && !variants[alt]) variants[alt] = 'original';
    }

    const filtered = {};
    for (const k of Object.keys(variants)) {
      if (k.length === 14 && /^\d+$/.test(k)) filtered[k] = variants[k];
    }
    return filtered;
  }

  function consolidateGtins(row, columns) {
    const pool = {};
    for (const col of columns) {
      if (row[col] == null || row[col] === '') continue;
      const v = normalizeAndGenerateVariants(row[col]);
      Object.assign(pool, v);
    }
    return pool;
  }

  function calculateGtinMatchConfidence(cust, cat) {
    const custKeys = Object.keys(cust);
    const catKeys = Object.keys(cat);
    if (!custKeys.length || !catKeys.length) return { score: 0, type: 'No Match', gtins: [] };
    const matching = [];
    for (const k of custKeys) if (catKeys.includes(k)) matching.push(k);
    if (!matching.length) return { score: 0, type: 'No Match', gtins: [] };

    let bestScore = 0, bestType = 'No Match';
    for (const gtin of matching) {
      const cs = cust[gtin], cts = cat[gtin];
      let score = 0, type = 'No Match';
      if (cs === 'original' && cts === 'original') { score = 120; type = 'Exact Match'; }
      else if (cs === 'corrected' || cts === 'corrected') { score = 92; type = 'Corrected Match'; }
      else if (cs === 'case_to_unit' || cts === 'case_to_unit') { score = 90; type = 'Case/Unit Match'; }
      else if (cs === 'missing_check' || cts === 'missing_check') { score = 92; type = 'Corrected Match'; }
      else { score = 120; type = 'Exact Match'; }
      if (score > bestScore) { bestScore = score; bestType = type; }
    }
    return { score: bestScore, type: bestType, gtins: matching.slice(0, 3) };
  }

  // ===== TF-IDF =====

  function tokenize(text) {
    return text.toLowerCase().split(/\s+/).filter(t => t);
  }

  function buildTfidfModel(texts, removeStopWords) {
    const n = texts.length;
    const df = new Map();
    const termIndex = new Map();
    const docTokens = [];
    const docVectors = [];

    for (const text of texts) {
      const tokens = tokenize(text);
      const filtered = removeStopWords ? tokens.filter(t => !STOP_WORDS.has(t)) : tokens;
      docTokens.push(filtered);
      const seen = new Set(filtered);
      for (const t of seen) df.set(t, (df.get(t) || 0) + 1);
    }

    const terms = Array.from(df.keys());
    terms.forEach((t, i) => termIndex.set(t, i));

    for (let i = 0; i < n; i++) {
      const counts = new Map();
      for (const t of docTokens[i]) counts.set(t, (counts.get(t) || 0) + 1);
      const vec = new Map();
      let norm = 0;
      for (const [t, cnt] of counts) {
        const idf = Math.log(n / (df.get(t) || 1));
        const w = cnt * idf;
        vec.set(termIndex.get(t), w);
        norm += w * w;
      }
      const normSqrt = Math.sqrt(norm) || 1;
      for (const [idx, w] of vec) vec.set(idx, w / normSqrt);
      docVectors.push(vec);
    }

    // Build inverted index for fast dot products
    const inverted = new Map();
    for (let docIdx = 0; docIdx < n; docIdx++) {
      for (const [termIdx, w] of docVectors[docIdx]) {
        if (!inverted.has(termIdx)) inverted.set(termIdx, []);
        inverted.get(termIdx).push({ docIdx, w });
      }
    }

    return { termIndex, df, docVectors, inverted, n };
  }

  function cosineScores(queryTokens, model, removeStopWords) {
    const scores = new Float64Array(model.n).fill(0);
    const counts = new Map();
    for (const t of queryTokens) {
      const key = t.toLowerCase();
      if (removeStopWords && STOP_WORDS.has(key)) continue;
      if (model.termIndex.has(key)) counts.set(key, (counts.get(key) || 0) + 1);
    }
    if (counts.size === 0) return scores;

    let norm = 0;
    const weights = new Map();
    for (const [t, cnt] of counts) {
      const termIdx = model.termIndex.get(t);
      const idf = Math.log(model.n / (model.df.get(t) || 1));
      const w = cnt * idf;
      weights.set(termIdx, w);
      norm += w * w;
    }
    const normSqrt = Math.sqrt(norm) || 1;
    for (const [termIdx, w] of weights) {
      const q = w / normSqrt;
      const postings = model.inverted.get(termIdx);
      if (!postings) continue;
      for (const { docIdx, w: catW } of postings) {
        scores[docIdx] += q * catW;
      }
    }
    return scores;
  }

  // ===== Main matching =====

  function preprocess(rows, config, settings) {
    const { productCols, sizeCol, sizeValueCol, sizeUnitCol, gtinCols, includeSizeInText } = config;
    const out = [];
    for (const row of rows) {
      const parts = productCols.filter(c => row[c] != null).map(c => cleanText(row[c], settings.removeStopWords, settings.caseSensitive));
      let combined = parts.join(' ').trim();
      let stdSize = '';
      if (sizeValueCol && sizeUnitCol && row[sizeValueCol] != null && row[sizeUnitCol] != null) {
        stdSize = standardizeSize(`${row[sizeValueCol]} ${row[sizeUnitCol]}`);
      } else if (sizeCol && row[sizeCol] != null) {
        stdSize = standardizeSize(row[sizeCol]);
      }
      if (settings.includeSizeInText && stdSize) {
        combined = (combined + ' ' + stdSize).trim();
      }
      const gtinPool = gtinCols && gtinCols.length ? consolidateGtins(row, gtinCols) : {};
      out.push({ original: row, combined, stdSize, gtinPool });
    }
    return out;
  }

  function buildGtinIndex(items) {
    const index = {};
    for (let i = 0; i < items.length; i++) {
      for (const gtin of Object.keys(items[i].gtinPool)) {
        if (!index[gtin]) index[gtin] = [];
        index[gtin].push(i);
      }
    }
    return index;
  }

  function computeMatch(catalogItems, customerItems, settings, onProgress) {
    const nCatalog = catalogItems.length;
    const nCustomer = customerItems.length;
    const threshold = settings.similarityThreshold;
    const maxMatches = settings.maxMatchesPerProduct;
    const textEnabled = settings.textEnabled;
    const gtinEnabled = settings.gtinEnabled;
    const sizeEnabled = settings.sizeEnabled;

    let tfidfWeight = 0, fuzzyWeight = 0, gtinWeight = 0, sizeWeight = 0;
    if (textEnabled && gtinEnabled) {
      gtinWeight = 0.5;
      const textTotal = 0.5;
      tfidfWeight = settings.baseTfidf * textTotal;
      fuzzyWeight = settings.baseFuzzy * textTotal;
    } else if (textEnabled) {
      tfidfWeight = settings.baseTfidf;
      fuzzyWeight = settings.baseFuzzy;
    } else if (gtinEnabled) {
      gtinWeight = 1.0;
    }
    if (sizeEnabled) sizeWeight = settings.sizeWeight;

    const allTexts = catalogItems.map(i => i.combined).concat(customerItems.map(i => i.combined));
    let model = null;
    if (textEnabled) {
      model = buildTfidfModel(allTexts, settings.removeStopWords);
      // split vectors: first catalog, then customer
    }
    const catalogModel = model ? { ...model, n: nCatalog, docVectors: model.docVectors.slice(0, nCatalog) } : null;
    const customerModel = model ? { ...model, n: nCustomer, docVectors: model.docVectors.slice(nCatalog) } : null;

    // rebuild catalog-only inverted index for fast dot product
    if (catalogModel) {
      const inverted = new Map();
      for (let docIdx = 0; docIdx < nCatalog; docIdx++) {
        for (const [termIdx, w] of catalogModel.docVectors[docIdx]) {
          if (!inverted.has(termIdx)) inverted.set(termIdx, []);
          inverted.get(termIdx).push({ docIdx, w });
        }
      }
      catalogModel.inverted = inverted;
    }

    const gtinIndex = gtinEnabled ? buildGtinIndex(catalogItems) : {};

    const results = [];
    let lastProgress = 0;

    for (let i = 0; i < nCustomer; i++) {
      const cust = customerItems[i];
      const scores = [];

      // determine candidate catalog indices
      let candidates;
      if (textEnabled && catalogModel) {
        const tfidfScores = cosineScores(tokenize(cust.combined), catalogModel, settings.removeStopWords);
        const minTfidf = Math.max(5, threshold * 0.2);
        const topK = Math.min(1000, Math.max(50, Math.round(0.1 * nCatalog)));

        const withScore = [];
        for (let j = 0; j < nCatalog; j++) withScore.push({ j, score: tfidfScores[j] });

        const above = withScore.filter(x => x.score >= minTfidf);
        const topByScore = withScore
          .slice()
          .sort((a, b) => b.score - a.score)
          .slice(0, topK);
        const set = new Map();
        for (const x of above) set.set(x.j, x.score);
        for (const x of topByScore) if (!set.has(x.j)) set.set(x.j, x.score);
        candidates = Array.from(set, ([j, tfidf]) => ({ j, tfidf }));
      } else if (gtinEnabled) {
        // GTIN-only mode: candidates are catalog items sharing any variant
        const candidateSet = new Map();
        for (const gtin of Object.keys(cust.gtinPool)) {
          if (gtinIndex[gtin]) {
            for (const j of gtinIndex[gtin]) candidateSet.set(j, 0);
          }
        }
        candidates = Array.from(candidateSet, ([j]) => ({ j, tfidf: 0 }));
      } else {
        candidates = [];
      }

      const rowMatches = [];
      for (const { j, tfidf } of candidates) {
        if (settings.matchingMode === 'within' && i === j) continue;

        const cat = catalogItems[j];

        let fuzzy = 0;
        if (textEnabled && fuzzyWeight > 0) {
          fuzzy = fuzzball.token_set_ratio(cust.combined, cat.combined, { full_process: false });
        }

        let sizeSim = 0;
        if (sizeEnabled && sizeWeight > 0) {
          sizeSim = calculateSizeSimilarity(cust.stdSize, cat.stdSize, settings.sizeTolerance);
        }

        let gtinScore = 0;
        let gtinType = 'No Match';
        let gtins = [];
        if (gtinEnabled && gtinWeight > 0) {
          const g = calculateGtinMatchConfidence(cust.gtinPool, cat.gtinPool);
          gtinScore = g.score;
          gtinType = g.type;
          gtins = g.gtins;
        }

        const combined = computeCombinedScore(tfidf, fuzzy, gtinScore, sizeSim, tfidfWeight, fuzzyWeight, gtinWeight, sizeWeight);

        if (combined >= threshold) {
          rowMatches.push({ i, j, combined, tfidf, fuzzy, gtinScore, gtinType, gtins, sizeSim });
        }
      }

      rowMatches.sort((a, b) => b.combined - a.combined);
      const top = rowMatches.slice(0, maxMatches);
      for (const m of top) results.push(m);

      const progress = (i + 1) / nCustomer;
      if (onProgress && progress - lastProgress > 0.05) {
        lastProgress = progress;
        onProgress(progress, i + 1, nCustomer);
      }
    }

    if (onProgress) onProgress(1, nCustomer, nCustomer);
    return results;
  }

  function computeCombinedScore(tfidf, fuzzy, gtinScore, sizeSim, tfidfWeight, fuzzyWeight, gtinWeight, sizeWeight) {
    const textTotal = tfidfWeight + fuzzyWeight;
    let textScore = 0;
    if (textTotal > 0) {
      textScore = tfidf * (tfidfWeight / textTotal) + fuzzy * (fuzzyWeight / textTotal);
    }

    let combined = textScore;
    if (gtinWeight > 0 && gtinScore > 0) {
      if (textTotal > 0) {
        combined = textScore * 0.5 + gtinScore * 0.5;
      } else {
        combined = gtinScore;
      }
    }

    if (sizeWeight > 0 && sizeSim > 0) {
      combined = combined * (1 - sizeWeight) + sizeSim * sizeWeight;
    }

    return Math.min(combined, 100);
  }

  // ===== Grouping (within-file) =====

  class UnionFind {
    constructor(n) {
      this.parent = new Int32Array(n);
      this.rank = new Int32Array(n);
      for (let i = 0; i < n; i++) this.parent[i] = i;
    }
    find(x) {
      let root = x;
      while (this.parent[root] !== root) root = this.parent[root];
      while (this.parent[x] !== root) { const next = this.parent[x]; this.parent[x] = root; x = next; }
      return root;
    }
    union(a, b) {
      const ra = this.find(a), rb = this.find(b);
      if (ra === rb) return;
      if (this.rank[ra] < this.rank[rb]) this.parent[ra] = rb;
      else if (this.rank[ra] > this.rank[rb]) this.parent[rb] = ra;
      else { this.parent[rb] = ra; this.rank[ra]++; }
    }
  }

  function buildGroups(results, n) {
    const uf = new UnionFind(n);
    for (const r of results) uf.union(r.i, r.j);
    const groups = {};
    for (let i = 0; i < n; i++) {
      const root = uf.find(i);
      if (!groups[root]) groups[root] = [];
      groups[root].push(i);
    }
    return Object.values(groups).filter(g => g.length > 1);
  }

  function makeGroupRows(groups, items, threshold) {
    const rows = [];
    for (let g = 0; g < groups.length; g++) {
      const members = groups[g];
      const rep = members[0];
      for (const idx of members) {
        rows.push({
          'Group ID': `G${g + 1}`,
          'Group Summary': items[rep].original,
          'Product': items[idx].original,
          'Is Representative': idx === rep
        });
      }
    }
    return rows;
  }

  // ===== Utilities =====

  function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }

  return {
    parseFile,
    detectProductNameColumns,
    detectSizeColumns,
    detectGtinColumns,
    detectRestrictionColumns,
    cleanText,
    standardizeSize,
    calculateSizeSimilarity,
    normalizeAndGenerateVariants,
    consolidateGtins,
    calculateGtinMatchConfidence,
    preprocess,
    computeMatch,
    buildGroups,
    makeGroupRows,
    STOP_WORDS,
    objectToStrings
  };
})();
