/* Excel.Flo – eng begrenzter Formel-Auswerter
 *
 * Bewusst KEINE General-Purpose-Spreadsheet-Engine (z. B. HyperFormula): die
 * unterrichteten Funktionen sind fest bekannt, daher reicht ein kleiner,
 * lizenzfreier eigener Parser/Evaluator. Deckt genau die Funktionen ab, die
 * in FUNCTION_SIGNATURES (engine.js) unterrichtet werden.
 *
 * Nutzung:
 *   ExcelFloFormula.evaluate(formulaText, getCellValue) -> Wert | Error-Objekt
 *   ExcelFloFormula.acceptedFormulaMatch(rawInput, acceptedFormulas) -> boolean
 */

(function () {
  "use strict";

  /* ---------------- Normalisierung für Formel-Text-Vergleich ---------------- */

  // Baut auf dem Tokenizer auf (statt eigener Regex-Heuristik), damit Dezimal-Komma
  // (z. B. "1,1") und Komma als Argumenttrenner (z. B. "D2,A2:B7") sicher unterschieden
  // werden – der Tokenizer erkennt Zellbezüge/Zahlen bereits strukturell korrekt.
  function normalizeForCompare(text) {
    const raw = (text || "").trim();
    if (raw === "") return "";
    const tokens = tokenize(raw);
    if (!tokens.length) return "";

    const parts = tokens.map((tok) => {
      switch (tok.type) {
        case "ref":
          return tok.value;
        case "number":
          return String(tok.value);
        case "string":
          return '"' + tok.value.replace(/"/g, '""') + '"';
        case "ident":
          if (tok.value === "WAHR" || tok.value === "TRUE") return "1";
          if (tok.value === "FALSCH" || tok.value === "FALSE") return "0";
          return tok.value;
        case "op":
          return tok.value === "," ? ";" : tok.value;
        default:
          return "";
      }
    });
    return "=" + parts.join("");
  }

  // Vergleicht die (normalisierte) Nutzereingabe gegen eine Liste erlaubter,
  // im Klartext geschriebener Formeln (kein Regex nötig).
  function acceptedFormulaMatch(rawInput, acceptedFormulas) {
    if (!acceptedFormulas || !acceptedFormulas.length) return false;
    const input = normalizeForCompare(rawInput);
    if (!input) return false;
    return acceptedFormulas.some((f) => normalizeForCompare(f) === input);
  }

  /* ---------------- Tokenizer ---------------- */

  // Zahlen akzeptieren sowohl Punkt als auch Komma als Dezimaltrennzeichen (deutsches Excel: Komma).
  const TOKEN_RE = /\s*(?:(\$?[A-Za-z]{1,3}\$?\d{1,7})|(\d+(?:[.,]\d+)?)|("(?:[^"]|"")*")|([A-Za-z_][A-Za-z0-9_.äöüÄÖÜ]*)|(<>|<=|>=|[+\-*/^&=<>(),;:%]))/y;

  function tokenize(text) {
    const tokens = [];
    let s = text.trim();
    if (s.startsWith("=")) s = s.slice(1);
    TOKEN_RE.lastIndex = 0;
    let m;
    while (TOKEN_RE.lastIndex < s.length) {
      m = TOKEN_RE.exec(s);
      if (!m || m[0] === "") break;
      if (m[1]) tokens.push({ type: "ref", value: m[1].replace(/\$/g, "").toUpperCase() });
      else if (m[2]) tokens.push({ type: "number", value: parseFloat(m[2].replace(",", ".")) });
      else if (m[3]) tokens.push({ type: "string", value: m[3].slice(1, -1).replace(/""/g, '"') });
      else if (m[4]) tokens.push({ type: "ident", value: m[4].toUpperCase() });
      else if (m[5]) tokens.push({ type: "op", value: m[5] });
    }
    return tokens;
  }

  /* ---------------- Parser (recursive descent) ---------------- */

  function FormulaError(message) {
    this.message = message;
  }
  FormulaError.prototype.isFormulaError = true;

  function parse(tokens) {
    let pos = 0;

    function peek() {
      return tokens[pos];
    }
    function next() {
      return tokens[pos++];
    }
    function expectOp(op) {
      const t = next();
      if (!t || t.type !== "op" || t.value !== op) {
        throw new FormulaError("Erwartet '" + op + "'");
      }
    }

    function parseExpr() {
      return parseComparison();
    }

    function parseComparison() {
      let left = parseConcat();
      while (peek() && peek().type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().value)) {
        const op = next().value;
        const right = parseConcat();
        left = { type: "compare", op, left, right };
      }
      return left;
    }

    function parseConcat() {
      let left = parseAdditive();
      while (peek() && peek().type === "op" && peek().value === "&") {
        next();
        const right = parseAdditive();
        left = { type: "concat", left, right };
      }
      return left;
    }

    function parseAdditive() {
      let left = parseTerm();
      while (peek() && peek().type === "op" && (peek().value === "+" || peek().value === "-")) {
        const op = next().value;
        const right = parseTerm();
        left = { type: "binop", op, left, right };
      }
      return left;
    }

    function parseTerm() {
      let left = parseUnary();
      while (peek() && peek().type === "op" && (peek().value === "*" || peek().value === "/")) {
        const op = next().value;
        const right = parseUnary();
        left = { type: "binop", op, left, right };
      }
      return left;
    }

    function parseUnary() {
      if (peek() && peek().type === "op" && peek().value === "-") {
        next();
        return { type: "neg", value: parseUnary() };
      }
      return parsePrimary();
    }

    function parsePrimary() {
      const t = peek();
      if (!t) throw new FormulaError("Unerwartetes Formelende");

      if (t.type === "number") {
        next();
        return { type: "number", value: t.value };
      }
      if (t.type === "string") {
        next();
        return { type: "string", value: t.value };
      }
      if (t.type === "ref") {
        next();
        if (peek() && peek().type === "op" && peek().value === ":") {
          next();
          const end = next();
          if (!end || end.type !== "ref") throw new FormulaError("Ungültiger Bereich");
          return { type: "range", start: t.value, end: end.value };
        }
        return { type: "ref", value: t.value };
      }
      if (t.type === "ident") {
        next();
        if (peek() && peek().type === "op" && peek().value === "(") {
          next();
          const args = [];
          if (!(peek() && peek().type === "op" && peek().value === ")")) {
            args.push(parseExpr());
            while (peek() && peek().type === "op" && peek().value === ";") {
              next();
              args.push(parseExpr());
            }
          }
          expectOp(")");
          return { type: "call", name: t.value, args };
        }
        if (t.value === "WAHR" || t.value === "TRUE") return { type: "bool", value: true };
        if (t.value === "FALSCH" || t.value === "FALSE") return { type: "bool", value: false };
        throw new FormulaError("Unbekanntes Symbol: " + t.value);
      }
      if (t.type === "op" && t.value === "(") {
        next();
        const inner = parseExpr();
        expectOp(")");
        return inner;
      }
      throw new FormulaError("Unerwartetes Token: " + JSON.stringify(t));
    }

    const result = parseExpr();
    if (pos < tokens.length) throw new FormulaError("Unerwartete Zeichen am Ende der Formel");
    return result;
  }

  /* ---------------- Auswertung ---------------- */

  function colIndexFromLetters(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function colLettersFromIndex(index) {
    let n = index + 1;
    let s = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function refParts(ref) {
    const m = ref.match(/^([A-Za-z]{1,3})(\d{1,7})$/);
    return { col: m[1], row: parseInt(m[2], 10) };
  }

  function rangeToMatrix(rangeNode, getCellValue) {
    const a = refParts(rangeNode.start);
    const b = refParts(rangeNode.end);
    const c1 = colIndexFromLetters(a.col);
    const c2 = colIndexFromLetters(b.col);
    const colLo = Math.min(c1, c2);
    const colHi = Math.max(c1, c2);
    const rowLo = Math.min(a.row, b.row);
    const rowHi = Math.max(a.row, b.row);

    const matrix = [];
    for (let r = rowLo; r <= rowHi; r++) {
      const row = [];
      for (let c = colLo; c <= colHi; c++) {
        row.push(getCellValue(colLettersFromIndex(c) + r));
      }
      matrix.push(row);
    }
    return matrix;
  }

  function flatten(matrix) {
    const out = [];
    matrix.forEach((row) => row.forEach((v) => out.push(v)));
    return out;
  }

  function transpose(matrix) {
    return matrix[0].map((_, c) => matrix.map((row) => row[c]));
  }

  function toText(v) {
    if (v === undefined || v === null) return "";
    return String(v);
  }

  function toNumber(v) {
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (typeof v === "string") {
      const n = parseFloat(v.replace(",", "."));
      return Number.isNaN(n) ? 0 : n;
    }
    return 0;
  }

  function looseEquals(a, b) {
    if (typeof a === "number" || typeof b === "number") return toNumber(a) === toNumber(b);
    return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
  }

  function isTruthy(v) {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (typeof v === "string") return ["1", "WAHR", "TRUE"].includes(v.trim().toUpperCase());
    return !!v;
  }

  // Wandelt ein Excel-Platzhalter-Muster (* = beliebig viele Zeichen, ? = ein Zeichen) in einen RegExp um.
  function wildcardToRegex(pattern) {
    const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&");
    return new RegExp("^" + escaped.replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
  }

  // Einfache Kriterien wie SUMMEWENN/ZÄHLENWENN sie nutzen: "10", ">5", "<=3", "Text", "*Text*", "<>0"
  function matchesCriteria(value, criteria) {
    if (typeof criteria === "number") return toNumber(value) === criteria;
    const c = String(criteria).trim();
    const m = c.match(/^(<>|<=|>=|<|>|=)?(.*)$/);
    const op = m[1] || "=";
    const rest = m[2];
    const numRest = parseFloat(rest.replace(",", "."));
    const compareNumeric = !Number.isNaN(numRest) && typeof value !== "string";

    if (compareNumeric) {
      const v = toNumber(value);
      switch (op) {
        case ">": return v > numRest;
        case "<": return v < numRest;
        case ">=": return v >= numRest;
        case "<=": return v <= numRest;
        case "<>": return v !== numRest;
        default: return v === numRest;
      }
    }
    const hasWildcard = /[*?]/.test(rest);
    switch (op) {
      case "<>":
        return hasWildcard ? !wildcardToRegex(rest).test(String(value).trim()) : !looseEquals(value, rest);
      default:
        return hasWildcard ? wildcardToRegex(rest).test(String(value).trim()) : looseEquals(value, rest);
    }
  }

  // Array-bewusste Mini-Auswertung nur für SUMMENPRODUKT/FILTER: Bereiche werden zu flachen
  // Arrays, +-*/ und Vergleiche werden elementweise (broadcast) statt skalar ausgewertet. Der
  // normale evalNode bleibt davon unberührt – nur diese beiden Funktionen nutzen das hier.
  function broadcastNumeric(l, r, fn) {
    const la = Array.isArray(l), ra = Array.isArray(r);
    if (la && ra) return l.map((v, i) => fn(toNumber(v), toNumber(r[i])));
    if (la) return l.map((v) => fn(toNumber(v), toNumber(r)));
    if (ra) return r.map((v) => fn(toNumber(l), toNumber(v)));
    return fn(toNumber(l), toNumber(r));
  }

  function broadcastCompare(l, r, op) {
    const cmp = (a, b) => {
      switch (op) {
        case "=": return looseEquals(a, b) ? 1 : 0;
        case "<>": return looseEquals(a, b) ? 0 : 1;
        case "<": return toNumber(a) < toNumber(b) ? 1 : 0;
        case ">": return toNumber(a) > toNumber(b) ? 1 : 0;
        case "<=": return toNumber(a) <= toNumber(b) ? 1 : 0;
        case ">=": return toNumber(a) >= toNumber(b) ? 1 : 0;
      }
    };
    const la = Array.isArray(l), ra = Array.isArray(r);
    if (la && ra) return l.map((v, i) => cmp(v, r[i]));
    if (la) return l.map((v) => cmp(v, r));
    if (ra) return r.map((v) => cmp(l, v));
    return cmp(l, r);
  }

  function evalArrayAware(node, getCellValue) {
    switch (node.type) {
      case "range":
        return flatten(rangeToMatrix(node, getCellValue));
      case "ref":
        return getCellValue(node.value);
      case "number":
      case "string":
      case "bool":
        return node.value;
      case "neg": {
        const v = evalArrayAware(node.value, getCellValue);
        return Array.isArray(v) ? v.map((x) => -toNumber(x)) : -toNumber(v);
      }
      case "binop": {
        const l = evalArrayAware(node.left, getCellValue);
        const r = evalArrayAware(node.right, getCellValue);
        return broadcastNumeric(l, r, (a, b) => {
          switch (node.op) {
            case "+": return a + b;
            case "-": return a - b;
            case "*": return a * b;
            case "/": return b === 0 ? NaN : a / b;
          }
        });
      }
      case "compare": {
        const l = evalArrayAware(node.left, getCellValue);
        const r = evalArrayAware(node.right, getCellValue);
        return broadcastCompare(l, r, node.op);
      }
      default:
        return evalNode(node, getCellValue);
    }
  }

  // Datumswerte werden wie in Excel als fortlaufende Seriennummer (Tage seit 1899-12-30) gespeichert.
  const MS_PER_DAY = 86400000;
  const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

  function serialToDate(serial) {
    return new Date(EXCEL_EPOCH_UTC + serial * MS_PER_DAY);
  }
  function dateToSerial(y, m, d) {
    return Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY);
  }

  function roundHalfUp(num, digits) {
    const factor = Math.pow(10, digits);
    const sign = num < 0 ? -1 : 1;
    return (sign * Math.round(Math.abs(num) * factor)) / factor;
  }

  const FUNCTIONS = {
    SVERWEIS(args) {
      const [search, tableArg, colIndexArg, rangeArg] = args;
      if (tableArg.kind !== "matrix") throw new FormulaError("SVERWEIS: Matrix erwartet");
      const table = tableArg.value;
      const colIndex = Math.round(toNumber(colIndexArg.value)) - 1;
      const wantExact = rangeArg !== undefined && !isTruthy(rangeArg.value);

      if (wantExact) {
        for (const row of table) {
          if (looseEquals(row[0], search.value)) return row[colIndex];
        }
        throw new FormulaError("SVERWEIS: #NV (kein Treffer)");
      }
      // Näherungssuche: sortiert aufsteigend angenommen, größten Wert <= Suchkriterium
      let best = null;
      for (const row of table) {
        if (toNumber(row[0]) <= toNumber(search.value)) best = row;
        else break;
      }
      if (!best) throw new FormulaError("SVERWEIS: #NV (kein Treffer)");
      return best[colIndex];
    },

    WVERWEIS(args) {
      const [search, tableArg, rowIndexArg] = args;
      const table = tableArg.value; // rows x cols
      const rowIndex = Math.round(toNumber(rowIndexArg.value)) - 1;
      const headerRow = table.map((row) => row[0]);
      for (let c = 0; c < headerRow.length; c++) {
        if (looseEquals(table[c][0], search.value)) return table[c][rowIndex];
      }
      throw new FormulaError("WVERWEIS: #NV (kein Treffer)");
    },

    "XVERWEIS": function (args) {
      const [search, searchRangeArg, returnRangeArg, notFoundArg] = args;
      const searchList = flatten(searchRangeArg.value);
      const idx = searchList.findIndex((v) => looseEquals(v, search.value));
      if (idx === -1) {
        if (notFoundArg !== undefined) return notFoundArg.value;
        throw new FormulaError("XVERWEIS: #NV (kein Treffer)");
      }
      // Rückgabematrix zeilenweise indizieren (nicht flach!), sonst verschiebt sich bei mehr
      // als einer Rückgabespalte alles um den Faktor Spaltenanzahl. Eine einzelne Spalte
      // liefert einen Skalar, mehrere Spalten liefern die ganze Zeile (für Spill geeignet).
      const returnMatrix = returnRangeArg.value;
      const row = returnMatrix[idx];
      return row.length === 1 ? row[0] : row;
    },

    "WENNFEHLER": function (args) {
      return args[0].error ? args[1].value : args[0].value;
    },

    "SUMMENPRODUKT": function (args, rawArgs) {
      const arrays = rawArgs.map((a) => {
        const v = evalArrayAware(a.node, a.getCellValue);
        return Array.isArray(v) ? v.map(toNumber) : [toNumber(v)];
      });
      const len = Math.max(...arrays.map((a) => a.length));
      let sum = 0;
      for (let i = 0; i < len; i++) {
        let product = 1;
        arrays.forEach((arr) => { product *= arr.length === 1 ? arr[0] : arr[i]; });
        sum += product;
      }
      return sum;
    },

    "BEREICH.VERSCHIEBEN": function (args, rawArgs) {
      const baseNode = rawArgs[0].node;
      const getCellValue = rawArgs[0].getCellValue;
      const baseRef = baseNode.type === "range" ? baseNode.start : baseNode.type === "ref" ? baseNode.value : null;
      if (!baseRef) throw new FormulaError("BEREICH.VERSCHIEBEN: Bezug erwartet");

      const base = refParts(baseRef);
      const rowOffset = Math.round(toNumber(args[1].value));
      const colOffset = Math.round(toNumber(args[2].value));
      const height = args[3] !== undefined ? Math.round(toNumber(args[3].value)) : 1;
      const width = args[4] !== undefined ? Math.round(toNumber(args[4].value)) : 1;

      const startCol = colIndexFromLetters(base.col) + colOffset;
      const startRow = base.row + rowOffset;

      const matrix = [];
      for (let r = 0; r < height; r++) {
        const row = [];
        for (let c = 0; c < width; c++) {
          row.push(getCellValue(colLettersFromIndex(startCol + c) + (startRow + r)));
        }
        matrix.push(row);
      }
      return matrix;
    },

    "EINDEUTIG": function (args) {
      const matrix = args[0].value;
      if (!Array.isArray(matrix) || !Array.isArray(matrix[0])) throw new FormulaError("EINDEUTIG: Bereich erwartet");
      const byCol = args[1] !== undefined && isTruthy(args[1].value);
      const exactlyOnce = args[2] !== undefined && isTruthy(args[2].value);
      const source = byCol ? transpose(matrix) : matrix;
      const groups = new Map();
      source.forEach((row) => {
        const key = row.map((v) => String(v).trim().toLowerCase()).join("");
        if (!groups.has(key)) groups.set(key, { row: row, count: 0 });
        groups.get(key).count++;
      });
      const out = [];
      groups.forEach(function (g) {
        if (!exactlyOnce || g.count === 1) out.push(g.row);
      });
      return byCol ? transpose(out) : out;
    },

    "SORTIEREN": function (args) {
      const matrix = args[0].value;
      if (!Array.isArray(matrix) || !Array.isArray(matrix[0])) throw new FormulaError("SORTIEREN: Bereich erwartet");
      const colIdx = args[1] !== undefined ? Math.round(toNumber(args[1].value)) - 1 : 0;
      const order = args[2] !== undefined ? Math.round(toNumber(args[2].value)) : 1;
      const copy = matrix.slice();
      copy.sort((a, b) => {
        const av = a[colIdx], bv = b[colIdx];
        const cmp = typeof av === "number" && typeof bv === "number" ? av - bv : String(av).localeCompare(String(bv), "de");
        return order < 0 ? -cmp : cmp;
      });
      return copy;
    },

    "FILTER": function (args, rawArgs) {
      const matrix = args[0].value;
      if (!Array.isArray(matrix) || !Array.isArray(matrix[0])) throw new FormulaError("FILTER: Bereich erwartet");
      const condRaw = evalArrayAware(rawArgs[1].node, rawArgs[1].getCellValue);
      const cond = Array.isArray(condRaw) ? condRaw : matrix.map(() => condRaw);
      const out = matrix.filter((row, i) => isTruthy(cond[i]));
      if (!out.length) throw new FormulaError("FILTER: #CALC! (keine Treffer)");
      return out;
    },

    WENN(args, rawArgs) {
      const cond = isTruthy(args[0].value);
      const branch = cond ? rawArgs[1] : rawArgs[2];
      if (branch === undefined) return cond;
      return evalNode(branch.node, branch.getCellValue);
    },

    UND(args) {
      return args.every((a) => isTruthy(a.value));
    },

    ODER(args) {
      return args.some((a) => isTruthy(a.value));
    },

    SUMME(args) {
      let sum = 0;
      args.forEach((a) => {
        const values = a.kind === "matrix" ? flatten(a.value) : [a.value];
        values.forEach((v) => {
          if (typeof v === "number") sum += v;
        });
      });
      return sum;
    },

    SUMMEWENN(args) {
      const [rangeArg, criteriaArg, sumRangeArg] = args;
      const range = flatten(rangeArg.value);
      const sumRange = sumRangeArg ? flatten(sumRangeArg.value) : range;
      let sum = 0;
      range.forEach((v, i) => {
        if (matchesCriteria(v, criteriaArg.value)) sum += toNumber(sumRange[i]);
      });
      return sum;
    },

    "ZÄHLENWENN": function (args) {
      const [rangeArg, criteriaArg] = args;
      const range = flatten(rangeArg.value);
      return range.filter((v) => matchesCriteria(v, criteriaArg.value)).length;
    },

    "ZÄHLENWENNS": function (args) {
      if (args.length < 2 || args.length % 2 !== 0) throw new FormulaError("ZÄHLENWENNS: Kriterienbereich/Kriterium-Paare erwartet");
      const pairs = [];
      for (let i = 0; i < args.length; i += 2) {
        pairs.push({ range: flatten(args[i].value), criteria: args[i + 1].value });
      }
      let count = 0;
      for (let r = 0; r < pairs[0].range.length; r++) {
        if (pairs.every((p) => matchesCriteria(p.range[r], p.criteria))) count++;
      }
      return count;
    },

    RUNDEN(args) {
      return roundHalfUp(toNumber(args[0].value), Math.round(toNumber(args[1].value)));
    },

    VERGLEICH(args) {
      const [search, rangeArg, typeArg] = args;
      const list = flatten(rangeArg.value);
      const matchType = typeArg ? Math.round(toNumber(typeArg.value)) : 1;

      if (matchType === 0) {
        const idx = list.findIndex((v) => looseEquals(v, search.value));
        if (idx === -1) throw new FormulaError("VERGLEICH: #NV (kein Treffer)");
        return idx + 1;
      }
      let bestIdx = -1;
      for (let i = 0; i < list.length; i++) {
        if (matchType > 0 && toNumber(list[i]) <= toNumber(search.value)) bestIdx = i;
        else if (matchType < 0 && toNumber(list[i]) >= toNumber(search.value)) bestIdx = i;
      }
      if (bestIdx === -1) throw new FormulaError("VERGLEICH: #NV (kein Treffer)");
      return bestIdx + 1;
    },

    INDEX(args) {
      const [rangeArg, rowArg, colArg] = args;
      const matrix = rangeArg.value;
      const row = Math.round(toNumber(rowArg.value));
      const col = colArg ? Math.round(toNumber(colArg.value)) : 1;
      if (matrix.length === 1) return matrix[0][row - 1] !== undefined && !colArg ? matrix[0][row - 1] : matrix[row - 1][col - 1];
      return matrix[row - 1][col - 1];
    },

    MIN(args) {
      const nums = [];
      args.forEach((a) => (a.kind === "matrix" ? flatten(a.value) : [a.value]).forEach((v) => typeof v === "number" && nums.push(v)));
      return nums.length ? Math.min(...nums) : 0;
    },

    MAX(args) {
      const nums = [];
      args.forEach((a) => (a.kind === "matrix" ? flatten(a.value) : [a.value]).forEach((v) => typeof v === "number" && nums.push(v)));
      return nums.length ? Math.max(...nums) : 0;
    },

    MITTELWERT(args) {
      const nums = [];
      args.forEach((a) => (a.kind === "matrix" ? flatten(a.value) : [a.value]).forEach((v) => typeof v === "number" && nums.push(v)));
      if (!nums.length) throw new FormulaError("MITTELWERT: #DIV/0!");
      return nums.reduce((s, v) => s + v, 0) / nums.length;
    },

    ANZAHL(args) {
      let count = 0;
      args.forEach((a) => (a.kind === "matrix" ? flatten(a.value) : [a.value]).forEach((v) => typeof v === "number" && count++));
      return count;
    },

    GROSS(args) {
      return toText(args[0].value).toUpperCase();
    },

    KLEIN(args) {
      return toText(args[0].value).toLowerCase();
    },

    GROSS2(args) {
      return toText(args[0].value).replace(/\p{L}+/gu, (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());
    },

    "LÄNGE": function (args) {
      return toText(args[0].value).length;
    },

    LINKS(args) {
      const text = toText(args[0].value);
      const n = args[1] !== undefined ? Math.round(toNumber(args[1].value)) : 1;
      return text.slice(0, Math.max(0, n));
    },

    RECHTS(args) {
      const text = toText(args[0].value);
      const n = args[1] !== undefined ? Math.round(toNumber(args[1].value)) : 1;
      return n <= 0 ? "" : text.slice(-n);
    },

    "FINDEN": function (args) {
      const search = toText(args[0].value);
      const text = toText(args[1].value);
      const start = args[2] !== undefined ? Math.round(toNumber(args[2].value)) : 1;
      const idx = text.indexOf(search, Math.max(0, start - 1));
      if (idx === -1) throw new FormulaError("FINDEN: #WERT! (Text nicht gefunden)");
      return idx + 1;
    },

    TEXTVOR(args) {
      const text = toText(args[0].value);
      const delim = toText(args[1].value);
      const idx = text.indexOf(delim);
      if (idx === -1) throw new FormulaError("TEXTVOR: #N/A (Trennzeichen nicht gefunden)");
      return text.slice(0, idx);
    },

    TEXTNACH(args) {
      const text = toText(args[0].value);
      const delim = toText(args[1].value);
      const idx = text.indexOf(delim);
      if (idx === -1) throw new FormulaError("TEXTNACH: #N/A (Trennzeichen nicht gefunden)");
      return text.slice(idx + delim.length);
    },

    TEXTVERKETTEN(args) {
      const delim = toText(args[0].value);
      const ignoreEmpty = isTruthy(args[1].value);
      const parts = [];
      for (let i = 2; i < args.length; i++) {
        const a = args[i];
        const values = a.kind === "matrix" ? flatten(a.value) : [a.value];
        values.forEach((v) => {
          const t = toText(v);
          if (!(ignoreEmpty && t === "")) parts.push(t);
        });
      }
      return parts.join(delim);
    },

    TEXTTEILEN(args) {
      const text = toText(args[0].value);
      const delim = toText(args[1].value);
      const parts = delim === "" ? [text] : text.split(delim);
      return [parts];
    },

    "DATEDIF": function (args) {
      const start = serialToDate(toNumber(args[0].value));
      const end = serialToDate(toNumber(args[1].value));
      const unit = toText(args[2].value).trim().toUpperCase();

      if (unit === "D") return Math.round(toNumber(args[1].value)) - Math.round(toNumber(args[0].value));

      if (unit === "Y") {
        let years = end.getUTCFullYear() - start.getUTCFullYear();
        if (end.getUTCMonth() < start.getUTCMonth() || (end.getUTCMonth() === start.getUTCMonth() && end.getUTCDate() < start.getUTCDate())) years--;
        return years;
      }
      if (unit === "M") {
        let months = (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + (end.getUTCMonth() - start.getUTCMonth());
        if (end.getUTCDate() < start.getUTCDate()) months--;
        return months;
      }
      throw new FormulaError("DATEDIF: unbekannte Einheit (nutze \"Y\", \"M\" oder \"D\")");
    },

    "EDATUM": function (args) {
      const start = serialToDate(toNumber(args[0].value));
      const monthsDelta = Math.round(toNumber(args[1].value));
      const y = start.getUTCFullYear();
      const m = start.getUTCMonth();
      const day = start.getUTCDate();

      const totalMonths = m + monthsDelta;
      const newYear = y + Math.floor(totalMonths / 12);
      const newMonth = ((totalMonths % 12) + 12) % 12;
      const lastDayOfNewMonth = new Date(Date.UTC(newYear, newMonth + 1, 0)).getUTCDate();
      const newDay = Math.min(day, lastDayOfNewMonth);
      return dateToSerial(newYear, newMonth + 1, newDay);
    },
  };

  function evalNode(node, getCellValue) {
    switch (node.type) {
      case "number":
        return node.value;
      case "string":
        return node.value;
      case "bool":
        return node.value;
      case "ref":
        return getCellValue(node.value);
      case "range":
        return rangeToMatrix(node, getCellValue);
      case "neg":
        return -toNumber(evalNode(node.value, getCellValue));
      case "concat":
        return String(evalNode(node.left, getCellValue)) + String(evalNode(node.right, getCellValue));
      case "binop": {
        const l = toNumber(evalNode(node.left, getCellValue));
        const r = toNumber(evalNode(node.right, getCellValue));
        switch (node.op) {
          case "+": return l + r;
          case "-": return l - r;
          case "*": return l * r;
          case "/":
            if (r === 0) throw new FormulaError("#DIV/0!");
            return l / r;
        }
        throw new FormulaError("Unbekannter Operator: " + node.op);
      }
      case "compare": {
        const l = evalNode(node.left, getCellValue);
        const r = evalNode(node.right, getCellValue);
        switch (node.op) {
          case "=": return looseEquals(l, r);
          case "<>": return !looseEquals(l, r);
          case "<": return toNumber(l) < toNumber(r);
          case ">": return toNumber(l) > toNumber(r);
          case "<=": return toNumber(l) <= toNumber(r);
          case ">=": return toNumber(l) >= toNumber(r);
        }
        throw new FormulaError("Unbekannter Vergleich: " + node.op);
      }
      case "call": {
        const fn = FUNCTIONS[node.name];
        if (!fn) throw new FormulaError("Unbekannte Funktion: " + node.name);
        // WENNFEHLER muss Fehler aus seinen Argumenten selbst abfangen können, statt dass sie
        // sofort durchschlagen – nur für sie werden Auswertungsfehler pro Argument eingefangen
        // statt weitergeworfen. Für alle anderen Funktionen bleibt das Verhalten unverändert.
        const isErrorAware = node.name === "WENNFEHLER";
        const evaluatedArgs = node.args.map((argNode) => {
          let value, error = null;
          try {
            value = evalNode(argNode, getCellValue);
          } catch (e) {
            error = e && e.isFormulaError ? e : new FormulaError(e.message || "Formelfehler");
            if (!isErrorAware) throw error;
          }
          const isMatrix = argNode.type === "range" || Array.isArray(value);
          return { kind: isMatrix ? "matrix" : "scalar", value, error, node: argNode, getCellValue };
        });
        const rawArgs = node.args.map((argNode) => ({ node: argNode, getCellValue }));
        return fn(evaluatedArgs, rawArgs);
      }
      default:
        throw new FormulaError("Unbekannter Knoten: " + node.type);
    }
  }

  function evaluate(formulaText, getCellValue) {
    try {
      const tokens = tokenize(formulaText);
      if (!tokens.length) return new FormulaError("Leere Formel");
      const ast = parse(tokens);
      return evalNode(ast, getCellValue);
    } catch (e) {
      if (e && e.isFormulaError) return e;
      return new FormulaError(e.message || "Formelfehler");
    }
  }

  window.ExcelFloFormula = {
    evaluate,
    acceptedFormulaMatch,
    normalizeForCompare,
    isFormulaError: (v) => v instanceof FormulaError,
  };
})();
