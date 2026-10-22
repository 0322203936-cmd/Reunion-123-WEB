// ============================================================================
// EXCEL NUMBER FORMATS - renders a value with an Excel format code
// (sections ;  colors [Red]  dates d/m/y h:m:s  0 # ? , . %  "text"  \x  _x  *x  @)
// ============================================================================

(function (root) {
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const COLORS = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#00b050', blue: '#0000ff', yellow: '#ffff00', magenta: '#ff00ff', cyan: '#00ffff' };
  const cache = new Map();

  function splitSections(fmt) {
    const out = [];
    let cur = '';
    let inQ = false;
    let inB = false;
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '\\' && !inQ) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
      if (ch === '"' && !inB) inQ = !inQ;
      else if (!inQ && ch === '[') inB = true;
      else if (!inQ && ch === ']') inB = false;
      if (ch === ';' && !inQ && !inB) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  // tokenize a section into literals and codes
  function lex(section) {
    const toks = [];
    let color = null;
    let cond = null;
    for (let i = 0; i < section.length; i++) {
      const ch = section[i];
      if (ch === '"') {
        const j = section.indexOf('"', i + 1);
        toks.push({ lit: section.slice(i + 1, j < 0 ? undefined : j) });
        i = j < 0 ? section.length : j;
      } else if (ch === '\\') {
        toks.push({ lit: section[i + 1] || '' }); i++;
      } else if (ch === '_') {
        toks.push({ lit: ' ' }); i++;
      } else if (ch === '*') {
        i++; // fill character: ignored
      } else if (ch === '[') {
        const j = section.indexOf(']', i);
        const inner = section.slice(i + 1, j);
        i = j;
        const low = inner.toLowerCase();
        if (COLORS[low]) color = COLORS[low];
        else if (/^color\d+$/.test(low)) color = null;
        else if (/^\$/.test(inner)) { const sym = inner.slice(1).split('-')[0]; if (sym) toks.push({ lit: sym }); }
        else if (/^[<>=]/.test(inner)) cond = inner;
        else if (/^(h+|m+|s+)$/i.test(inner)) toks.push({ code: `[${low}]` });
      } else {
        toks.push({ code: ch });
      }
    }
    return { toks, color, cond };
  }

  function compile(fmt) {
    if (cache.has(fmt)) return cache.get(fmt);
    const sections = splitSections(fmt).map(lex);
    const compiled = { sections };
    cache.set(fmt, compiled);
    return compiled;
  }

  // ---------------------------------------------------------------- dates
  function serialToParts(serial) {
    const ms = Math.round((serial - 25569) * 86400000); // 25569 = 1970-01-01
    const d = new Date(ms);
    return {
      y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate(), w: d.getUTCDay(),
      H: d.getUTCHours(), M: d.getUTCMinutes(), S: d.getUTCSeconds(), total: serial
    };
  }

  function formatDate(serial, toks) {
    const p = serialToParts(serial);
    // group consecutive identical code letters
    const groups = [];
    for (const t of toks) {
      if (t.lit !== undefined) { groups.push({ lit: t.lit }); continue; }
      const c = t.code;
      const last = groups[groups.length - 1];
      if (last && last.code && last.code[0].toLowerCase() === c.toLowerCase() && /^[ymdhs]$/i.test(c)) last.code += c;
      else groups.push({ code: c });
    }
    // detect AM/PM
    let ampm = false;
    for (let i = 0; i < groups.length; i++) {
      const g = groups[i];
      if (g.code && /^a$/i.test(g.code)) {
        const rest = groups.slice(i, i + 5).map(x => x.code || x.lit || '').join('');
        if (/^am\/pm/i.test(rest) || /^a\/p/i.test(rest)) ampm = true;
      }
    }
    let out = '';
    for (let i = 0; i < groups.length; i++) {
      const g = groups[i];
      if (g.lit !== undefined) { out += g.lit; continue; }
      const c = g.code;
      const lc = c.toLowerCase();
      if (lc[0] === 'y') { out += lc.length <= 2 ? String(p.y).slice(-2) : String(p.y); continue; }
      if (lc[0] === 'd') {
        out += lc.length === 1 ? p.d : lc.length === 2 ? String(p.d).padStart(2, '0') : lc.length === 3 ? DAYS[p.w].slice(0, 3) : DAYS[p.w];
        continue;
      }
      if (lc[0] === 'm') {
        // minutes if preceded by h or followed by s
        const prev = groups.slice(0, i).reverse().find(x => x.code && /^[a-z]/i.test(x.code));
        const next = groups.slice(i + 1).find(x => x.code && /^[a-z]/i.test(x.code));
        const isMinute = (prev && /^h/i.test(prev.code)) || (next && /^s/i.test(next.code));
        if (isMinute && lc.length <= 2) { out += lc.length === 1 ? p.M : String(p.M).padStart(2, '0'); continue; }
        out += lc.length === 1 ? p.m + 1 : lc.length === 2 ? String(p.m + 1).padStart(2, '0') : lc.length === 3 ? MONTHS[p.m].slice(0, 3) : lc.length === 5 ? MONTHS[p.m][0] : MONTHS[p.m];
        continue;
      }
      if (lc[0] === 'h') {
        let h = p.H;
        if (ampm) { h = h % 12; if (h === 0) h = 12; }
        out += lc.length === 1 ? h : String(h).padStart(2, '0');
        continue;
      }
      if (lc[0] === 's') { out += lc.length === 1 ? p.S : String(p.S).padStart(2, '0'); continue; }
      if (/^\[h+\]$/.test(lc)) { out += Math.floor(p.total * 24); continue; }
      if (/^a$/i.test(c)) {
        const rest = groups.slice(i, i + 5).map(x => x.code || x.lit || '').join('');
        if (/^am\/pm/i.test(rest)) { out += p.H < 12 ? 'AM' : 'PM'; i += 4; continue; }
        if (/^a\/p/i.test(rest)) { out += p.H < 12 ? 'A' : 'P'; i += 2; continue; }
        out += c; continue;
      }
      out += c;
    }
    return out;
  }

  // ---------------------------------------------------------------- numbers
  function formatNumber(value, toks) {
    // the numeric placeholder span: first digit placeholder + contiguous 0 # ? . ,
    let first = toks.findIndex(t => t.code && /[0#?.]/.test(t.code));
    if (first < 0) {
      // no digits: literals only (e.g. "-" for zero)
      return toks.map(t => t.lit !== undefined ? t.lit : (t.code === '@' ? '' : t.code)).join('');
    }
    let last = first;
    while (last + 1 < toks.length && toks[last + 1].code && /[0#?.,]/.test(toks[last + 1].code)) last++;
    const pattern = toks.slice(first, last + 1).map(t => t.code).join('');
    const percentCount = toks.filter(t => t.code === '%').length;

    let v = value * Math.pow(100, percentCount);
    // trailing commas scale by 1000
    let pat = pattern;
    while (/,$/.test(pat)) { v /= 1000; pat = pat.slice(0, -1); }
    const [intPat, decPat = ''] = pat.split('.');
    const thousands = intPat.includes(',');
    const intDigitsMin = (intPat.replace(/,/g, '').match(/0/g) || []).length;
    const decMax = (decPat.match(/[0#?]/g) || []).length;
    const decMin = (decPat.match(/0/g) || []).length;

    let abs = Math.abs(v);
    let fixed = abs.toFixed(decMax);
    let [ip, dp = ''] = fixed.split('.');
    // trim optional decimals (#)
    while (dp.length > decMin && dp.endsWith('0')) dp = dp.slice(0, -1);
    if (ip === '0' && intDigitsMin === 0) ip = '';
    if (ip.length < intDigitsMin) ip = ip.padStart(intDigitsMin, '0');
    if (thousands) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const numStr = ip + (decPat.length ? (dp.length || decPat.includes('0') || decPat.includes('?') ? '.' + dp : (decPat ? '.' : '')) : '');

    const before = toks.slice(0, first).map(t => t.lit !== undefined ? t.lit : (t.code === '%' ? '%' : t.code === '@' ? '' : t.code)).join('');
    const after = toks.slice(last + 1).map(t => t.lit !== undefined ? t.lit : (t.code === '%' ? '%' : t.code === '@' ? '' : t.code)).join('');
    return `${before}${numStr.replace(/\.$/, decPat ? '.' : '')}${after}`;
  }

  function generalNumber(v) {
    if (!Number.isFinite(v)) return String(v);
    if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
    const abs = Math.abs(v);
    if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
    return String(Number(v.toPrecision(10)));
  }

  /**
   * @returns {{ text: string, color: string|null, align: 'right'|'left'|null }}
   */
  function formatValue(value, fmt) {
    if (value === null || value === undefined || value === '') return { text: '', color: null };
    if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', color: null, align: 'center' };
    const code = fmt && String(fmt).trim();
    if (typeof value === 'string') {
      if (!code || /^general$/i.test(code)) return { text: value, color: null };
      const { sections } = compile(code);
      const textSec = sections.length >= 4 ? sections[3] : (sections.length === 1 && sections[0].toks.some(t => t.code === '@') ? sections[0] : null);
      if (!textSec) return { text: value, color: null };
      const text = textSec.toks.map(t => t.lit !== undefined ? t.lit : (t.code === '@' ? value : t.code)).join('');
      return { text, color: textSec.color };
    }
    if (typeof value !== 'number') return { text: String(value), color: null };
    if (!code || /^general$/i.test(code)) return { text: generalNumber(value), color: null };

    const { sections } = compile(code);
    let sec = sections[0];
    let useAbs = false;
    if (sections.length >= 2 && value < 0) { sec = sections[1]; useAbs = true; }
    else if (sections.length >= 3 && value === 0) { sec = sections[2]; }
    const toks = sec.toks;
    if (/^general$/i.test(toks.map(t => t.code || t.lit).join(''))) {
      return { text: (useAbs ? generalNumber(Math.abs(value)) : generalNumber(value)), color: sec.color };
    }
    const hasDate = toks.some(t => t.code && (/^[ymdhs]$/i.test(t.code) || /^\[(h+|m+|s+)\]$/.test(t.code)));
    if (hasDate) {
      if (value < 0) return { text: '#'.repeat(8), color: null };
      return { text: formatDate(value, toks), color: sec.color };
    }
    const n = useAbs ? Math.abs(value) : value;
    let text = formatNumber(n, toks);
    if (!useAbs && n < 0 && !/^-/.test(text)) {
      // a single section shows the minus sign; keep it before the number
      const roundedZero = /^[^1-9]*$/.test(text);
      text = roundedZero ? text : `-${text}`;
    }
    return { text, color: sec.color };
  }

  function isDateFormat(fmt) {
    if (!fmt) return false;
    const sec = compile(String(fmt)).sections[0];
    return sec.toks.some(t => t.code && (/^[ymdhs]$/i.test(t.code)));
  }

  const api = { formatValue, isDateFormat, generalNumber };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ExcelNumberFormat = api;
})(typeof window !== 'undefined' ? window : globalThis);
