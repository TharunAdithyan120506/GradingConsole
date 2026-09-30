    /*!
     * BITS Pilani Digital · Grading Console — grading core
     * Pure, DOM-free logic: file checks, workbook parsing, data validation,
     * grade-band rules, statistics, grading and CSV export.
     * Loaded as a classic <script> in the browser (window.GradingCore) and via require() in tests.
     */
    (function (root, factory) {
      if (typeof module === 'object' && module.exports) module.exports = factory();
      else root.GradingCore = factory();
    })(typeof self !== 'undefined' ? self : this, function () {
      'use strict';

      /* ------------------------------------------------------------------ constants */

      const GRADES = ['A', 'A-', 'B', 'B-', 'C', 'C-', 'D', 'E'];
      const DEFAULT_BANDS = Object.freeze([
        [80, 100], [70, 79], [60, 69], [50, 59], [40, 49], [30, 39], [20, 29], [0, 19],
      ].map(([min, max], i) => Object.freeze({ grade: GRADES[i], min, max })));

      const ACCEPTED_EXTENSIONS = ['xlsx', 'xls', 'csv'];
      const MAX_FILE_BYTES = 15 * 1024 * 1024;
      const MAX_ROWS = 100000;
      const HEADER_SCAN_ROWS = 30;
      const BORDERLINE_WINDOW = 2;

      const ROUNDING = Object.freeze({
        nearest: { label: 'Round to nearest', verb: 'rounded to nearest', apply: (v) => Math.round(v) },
        up: { label: 'Round up', verb: 'rounded up', apply: (v) => Math.ceil(v) },
        none: { label: 'Keep decimals', verb: 'kept as recorded', apply: (v) => v },
      });

      const FIELD_LABELS = { id: 'BITS ID', course: 'Course', marks: 'Total Marks' };

      /* ------------------------------------------------------------------ helpers */

      const normHeader = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const collapse = (s) => String(s).replace(/\s+/g, ' ').trim();
      const courseKeyOf = (name) => collapse(name).toLowerCase();
      const idKeyOf = (id) => String(id).replace(/\s+/g, '').toUpperCase();
      /** Strips binary floating-point noise such as 80.00000000000001 from spreadsheet numbers. */
      const clean = (v) => Math.round(v * 1e6) / 1e6;

      function extensionOf(name) {
        const m = /\.([a-z0-9]+)$/i.exec(String(name || ''));
        return m ? m[1].toLowerCase() : '';
      }

      function plural(n, one, many) { return `${n} ${n === 1 ? one : (many || one + 's')}`; }

      /* ------------------------------------------------------------------ file checks */

      function fileError(code, title, detail, hint) {
        return { ok: false, error: { code, title, detail, hint: hint || '' } };
      }

      /** Checks name and size before any bytes are read. */
      function checkFileMeta(meta) {
        const ext = extensionOf(meta && meta.name);
        if (!ACCEPTED_EXTENSIONS.includes(ext)) {
          return fileError('unsupported-type',
            `“${meta && meta.name ? meta.name : 'This file'}” isn't a marks spreadsheet`,
            'The console reads Excel workbooks (.xlsx or .xls) and CSV files.',
            'Save the marks sheet from Excel as .xlsx and upload it again.');
        }
        if (!meta.size) {
          return fileError('empty-file', 'This file is empty',
            `“${meta.name}” contains 0 bytes, so there are no marks to read.`,
            'Check that the file saved correctly, then upload it again.');
        }
        if (meta.size > MAX_FILE_BYTES) {
          return fileError('too-large', 'This file is too large',
            `“${meta.name}” is ${(meta.size / 1048576).toFixed(1)} MB; the limit is 15 MB.`,
            'Remove unused sheets, images or formatting and try again.');
        }
        return { ok: true, ext };
      }

      /* ------------------------------------------------------------------ cells */

      /** Normalises a SheetJS cell into a small tagged value. */
      function readCell(cell) {
        if (!cell || cell.t === 'z' || cell.v === undefined || cell.v === null) return { kind: 'blank' };
        switch (cell.t) {
          case 'n': {
            if (!Number.isFinite(cell.v)) return { kind: 'error', text: String(cell.w || cell.v) };
            if (typeof cell.w === 'string' && /^\d{1,4}[\/.-]\d{1,2}[\/.-]\d{1,4}/.test(cell.w)) return { kind: 'date', text: cell.w };
            return { kind: 'number', value: cell.v, text: cell.w != null ? String(cell.w) : String(cell.v) };
          }
          case 's': {
            const text = String(cell.v).trim();
            return text === '' ? { kind: 'blank' } : { kind: 'text', text };
          }
          case 'b': return { kind: 'bool', text: cell.v ? 'TRUE' : 'FALSE' };
          case 'e': return { kind: 'error', text: String(cell.w || '#ERROR') };
          case 'd': return { kind: 'date', text: String(cell.w || cell.v) };
          default: return { kind: 'text', text: String(cell.v).trim() };
        }
      }

      function cellText(c) {
        if (c.kind === 'number') return Number.isInteger(c.value) ? String(c.value) : String(c.value);
        return c.text || '';
      }

      /** Interprets a marks cell. Returns { ok, value, decimal } or { ok:false, code, shown }. */
      function readMarks(c) {
        if (c.kind === 'blank') return { ok: false, code: 'missing-marks', shown: '' };
        let value;
        if (c.kind === 'number') value = c.value;
        else if (c.kind === 'text' && /^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(c.text)) value = Number(c.text);
        else if (c.kind === 'error') return { ok: false, code: 'excel-error', shown: c.text };
        else if (c.kind === 'date') return { ok: false, code: 'date-marks', shown: c.text };
        else return { ok: false, code: 'text-marks', shown: c.text };
        value = clean(value);
        if (value < 0 || value > 100) return { ok: false, code: 'out-of-range', shown: String(value) };
        return { ok: true, value, decimal: !Number.isInteger(value) };
      }

      /* ------------------------------------------------------------------ header detection */

      /** Scores how well a header cell matches each required field (0 = no match). */
      function headerScores(text) {
        const n = normHeader(text);
        const words = String(text == null ? '' : text).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
        const s = { id: 0, course: 0, marks: 0 };
        if (!n) return s;
        if (n.includes('bitsid')) s.id = 3;
        else if (/^(student)?(id|idno|idnumber|rollno|rollnumber)$/.test(n)) s.id = 2;
        if (words.some((w) => w.startsWith('course'))) s.course = /^course(name|title)?$/.test(n) ? 3 : 2;
        else if (n.includes('totalmark')) s.marks = 3;
        else if (words.includes('total')) s.marks = 2;
        else if (words.some((w) => ['mark', 'marks', 'score', 'scores'].includes(w))) s.marks = 1;
        return s;
      }

      /** Finds the best column for each field in one row. */
      function mapHeaderRow(cells) {
        const best = {};
        const candidates = { id: [], course: [], marks: [] };
        cells.forEach((cell, col) => {
          const c = readCell(cell);
          if (c.kind === 'blank') return;
          const label = cellText(c);
          const scores = headerScores(label);
          for (const f of Object.keys(candidates)) {
            if (scores[f] > 0) {
              candidates[f].push({ col, label, score: scores[f] });
              if (!best[f] || scores[f] > best[f].score) best[f] = { col, label, score: scores[f] };
            }
          }
        });
        const found = Object.keys(best);
        const cols = new Set(found.map((f) => best[f].col));
        const complete = found.length === 3 && cols.size === 3;
        return { best, candidates, complete, foundCount: found.length };
      }

      /** Builds a sparse row grid from a SheetJS worksheet using only real cells. */
      function sheetRows(ws, XLSX) {
        const rows = new Map();
        let minR = Infinity, maxR = -1;
        for (const key in ws) {
          if (key.charCodeAt(0) === 33) continue; // "!ref", "!merges", ...
          const { r, c } = XLSX.utils.decode_cell(key);
          if (!rows.has(r)) rows.set(r, []);
          rows.get(r)[c] = ws[key];
          if (r < minR) minR = r;
          if (r > maxR) maxR = r;
        }
        return { rows, minR, maxR, empty: maxR < 0 };
      }

      /* ------------------------------------------------------------------ workbook parsing */

      /**
       * Reads bytes into a validated dataset.
       * @returns {{ok:true, dataset:object} | {ok:false, error:object}}
       */
      function parseFile(bytes, meta, XLSX) {
        const pre = checkFileMeta(meta);
        if (!pre.ok) return pre;
        let wb;
        try {
          wb = XLSX.read(bytes, {
            type: 'array', dense: false, cellDates: false, cellFormula: false,
            cellHTML: false, cellStyles: false, raw: pre.ext === 'csv',
          });
        } catch (err) {
          const msg = String(err && err.message || err);
          if (/password/i.test(msg)) {
            return fileError('protected', 'This workbook is password-protected',
              'The console cannot open encrypted workbooks.',
              'Open it in Excel, remove the password (File → Info → Protect Workbook), save, and upload again.');
          }
          return fileError('unreadable', "We couldn't open this file",
            `“${meta.name}” looks damaged or isn't a real ${pre.ext.toUpperCase()} file.`,
            'Open it in Excel and use File → Save As → Excel Workbook (.xlsx), then upload the new copy.');
        }
        return parseWorkbook(wb, XLSX, meta);
      }

      function parseWorkbook(wb, XLSX, meta) {
        const names = (wb && wb.SheetNames) || [];
        let best = null; // best partial header match, for error reporting
        let anyData = false;

        for (const sheetName of names) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const grid = sheetRows(ws, XLSX);
          if (grid.empty) continue;
          anyData = true;
          if (grid.maxR - grid.minR > MAX_ROWS) {
            return fileError('too-many-rows', 'This sheet has too many rows',
              `“${sheetName}” spans more than ${MAX_ROWS.toLocaleString('en-IN')} rows.`,
              'Keep only the marks table on the first sheet and upload again.');
          }
          const last = Math.min(grid.maxR, grid.minR + HEADER_SCAN_ROWS);
          for (let r = grid.minR; r <= last; r++) {
            const cells = grid.rows.get(r);
            if (!cells) continue;
            const map = mapHeaderRow(cells);
            if (map.complete) return buildDataset(grid, r, map, sheetName, meta, names.length);
            if (!best || map.foundCount > best.map.foundCount) best = { map, cells, sheetName };
          }
        }

        if (!anyData) {
          return fileError('no-data', 'This file has no data',
            `Every sheet in “${meta.name}” is empty.`,
            'Add a header row (BITS ID, Course, Total Marks) and one row per student.');
        }
        const missing = ['id', 'course', 'marks'].filter((f) => !best || !best.map.best[f]).map((f) => FIELD_LABELS[f]);
        const seen = best ? best.cells.map((c) => cellText(readCell(c))).filter(Boolean) : [];
        const list = (a) => a.map((x) => `“${x}”`).join(', ');
        return fileError('missing-columns',
          missing.length === 1 ? `Missing the ${missing[0]} column` : `Missing ${missing.length} required columns`,
          `We couldn't find ${list(missing)}${seen.length ? `. Closest header row found: ${list(seen.slice(0, 8))}` : ''}.`,
          'The first row of the marks table needs three headings: BITS ID, Course and Total Marks.');
      }

      function buildDataset(grid, headerR, map, sheetName, meta, sheetCount) {
        const col = { id: map.best.id.col, course: map.best.course.col, marks: map.best.marks.col };
        const usedCols = new Set(Object.values(col));
        const headerCells = grid.rows.get(headerR);
        const ignoredColumns = [];
        headerCells.forEach((cell, c) => {
          if (usedCols.has(c)) return;
          const t = cellText(readCell(cell));
          if (t) ignoredColumns.push(t);
        });

        const notes = [];
        for (const f of ['id', 'course', 'marks']) {
          const cands = map.candidates[f];
          if (cands.length > 1) {
            notes.push(`Several columns look like ${FIELD_LABELS[f]} (${cands.map((c) => `“${c.label}”`).join(', ')}); using “${map.best[f].label}”.`);
          }
        }
        if (ignoredColumns.length) notes.push(`Ignored ${plural(ignoredColumns.length, 'extra column')}: ${ignoredColumns.map((x) => `“${x}”`).join(', ')}.`);
        if (grid.minR < headerR) notes.push(`Skipped ${plural(headerR - grid.minR, 'row')} above the header.`);
        if (sheetCount > 1) notes.push(`Read the sheet “${sheetName}”.`);

        const issues = [];
        const courses = new Map();
        let rowsRead = 0, blankRows = 0;

        const courseFor = (name) => {
          const key = courseKeyOf(name);
          if (!courses.has(key)) courses.set(key, { key, name: collapse(name), variants: new Map(), entries: [], issueCount: 0 });
          const c = courses.get(key);
          const shown = collapse(name);
          c.variants.set(shown, (c.variants.get(shown) || 0) + 1);
          return c;
        };

        for (let r = headerR + 1; r <= grid.maxR; r++) {
          const cells = grid.rows.get(r) || [];
          const idC = readCell(cells[col.id]);
          const courseC = readCell(cells[col.course]);
          const marksC = readCell(cells[col.marks]);
          if (idC.kind === 'blank' && courseC.kind === 'blank' && marksC.kind === 'blank') { blankRows++; continue; }
          rowsRead++;
          const row = r + 1; // Excel row number
          const id = idC.kind === 'blank' || idC.kind === 'error' ? '' : collapse(cellText(idC));
          const courseName = courseC.kind === 'blank' || courseC.kind === 'error' ? '' : cellText(courseC);
          const course = courseName ? courseFor(courseName) : null;
          const base = { row, id, course: course ? course.name : '', courseKey: course ? course.key : null };

          if (!course) { issues.push({ ...base, severity: 'error', code: 'missing-course' }); continue; }
          if (!id) { issues.push({ ...base, severity: 'error', code: 'missing-id' }); course.issueCount++; continue; }
          const marks = readMarks(marksC);
          if (!marks.ok) { issues.push({ ...base, severity: 'error', code: marks.code, value: marks.shown }); course.issueCount++; continue; }
          course.entries.push({ id, idKey: idKeyOf(id), row, raw: marks.value, decimal: marks.decimal });
        }

        if (rowsRead === 0) {
          return fileError('no-rows', 'No student rows found',
            `The header is on row ${headerR + 1} of “${sheetName}”, but there are no students beneath it.`,
            'Add one row per student under the header, then upload again.');
        }
        if (courses.size === 0) {
          return fileError('no-courses', 'No course names found',
            `None of the ${plural(rowsRead, 'row')} has a course name, so there is nothing to grade.`,
            'Fill in the Course column for every student and upload again.');
        }

        const courseList = [];
        for (const c of courses.values()) {
          // Duplicate BITS IDs inside one course.
          const byId = new Map();
          c.entries.forEach((e) => { if (!byId.has(e.idKey)) byId.set(e.idKey, []); byId.get(e.idKey).push(e); });
          const students = [];
          for (const group of byId.values()) {
            if (group.length === 1) { students.push(group[0]); continue; }
            const rows = group.map((g) => g.row);
            const same = group.every((g) => g.raw === group[0].raw);
            if (same) {
              students.push(group[0]);
              issues.push({ row: rows[0], rows, id: group[0].id, course: c.name, courseKey: c.key, severity: 'warning', code: 'duplicate-identical', value: group[0].raw });
            } else {
              issues.push({ row: rows[0], rows, id: group[0].id, course: c.name, courseKey: c.key, severity: 'error', code: 'duplicate-conflict', values: group.map((g) => g.raw) });
              c.issueCount += group.length;
            }
          }
          students.sort((a, b) => a.row - b.row);
          students.forEach((s) => {
            if (s.decimal) issues.push({ row: s.row, id: s.id, course: c.name, courseKey: c.key, severity: 'warning', code: 'decimal', value: s.raw });
          });
          if (c.variants.size > 1) {
            const others = [...c.variants.keys()].filter((v) => v !== c.name);
            issues.push({ row: null, id: '', course: c.name, courseKey: c.key, severity: 'warning', code: 'course-variant', value: others });
          }
          courseList.push({ key: c.key, name: c.name, students, excluded: c.issueCount });
        }
        courseList.sort((a, b) => a.name.localeCompare(b.name, 'en', { numeric: true, sensitivity: 'base' }));
        issues.sort((a, b) => (a.row || 0) - (b.row || 0));

        const gradable = courseList.reduce((n, c) => n + c.students.length, 0);
        return {
          ok: true,
          dataset: {
            fileName: meta.name,
            fileSize: meta.size,
            isSample: !!meta.isSample,
            sheetName,
            headerRow: headerR + 1,
            columns: { id: map.best.id.label, course: map.best.course.label, marks: map.best.marks.label },
            notes,
            rowsRead,
            blankRows,
            gradable,
            excluded: issues.filter((i) => i.severity === 'error').reduce((n, i) => n + (i.code === 'duplicate-conflict' ? i.rows.length : 1), 0),
            courses: courseList,
            issues,
          },
        };
      }

      /** Human wording for every data issue, kept in one place. */
      function describeIssue(issue, policy) {
        const fmt = (v) => String(v);
        switch (issue.code) {
          case 'missing-course': return 'Course is blank — this row can’t be assigned to a course.';
          case 'missing-id': return 'BITS ID is blank.';
          case 'missing-marks': return 'Total Marks is blank. If this student is getting NC, remove the row from the file.';
          case 'text-marks': return `Total Marks is “${issue.value}”, not a number. Students awarded NC shouldn’t be in the file.`;
          case 'excel-error': return `Total Marks shows an Excel error (${issue.value}). Fix the formula in the sheet.`;
          case 'date-marks': return `Total Marks looks like a date (${issue.value}). Check the cell format in Excel.`;
          case 'out-of-range': return `Total Marks is ${issue.value}, outside the 0–100 scale.`;
          case 'duplicate-identical': return `Listed ${issue.rows.length} times with the same marks (rows ${issue.rows.join(', ')}) — counted once.`;
          case 'duplicate-conflict': return `Listed ${issue.rows.length} times with different marks (${issue.values.map(fmt).join(' vs ')}, rows ${issue.rows.join(', ')}) — excluded until the file is corrected.`;
          case 'decimal': {
            const p = ROUNDING[policy] || ROUNDING.nearest;
            const graded = p.apply(issue.value);
            return graded === issue.value
              ? `Marks ${issue.value} have a decimal — graded as recorded.`
              : `Marks ${issue.value} have a decimal — graded as ${graded} (${p.verb}).`;
          }
          case 'course-variant': return `Also written as ${issue.value.map((v) => `“${v}”`).join(', ')} — merged into “${issue.course}”.`;
          default: return 'Unrecognised problem with this row.';
        }
      }

      /* ------------------------------------------------------------------ grade bands */

      function defaultBands() { return DEFAULT_BANDS.map((b) => ({ ...b })); }
      function cloneBands(bands) { return bands.map((b) => ({ grade: b.grade, min: b.min, max: b.max })); }
      const isWhole = (v) => typeof v === 'number' && Number.isInteger(v);
      const inScale = (v) => isWhole(v) && v >= 0 && v <= 100;

      /**
       * Sets one edge of a band. Edges are shared: a band's minimum is one above the maximum of the
       * band below it, so the neighbouring edge moves too. Invalid input is stored but not propagated.
       */
      function setBandEdge(bands, index, field, value) {
        const next = cloneBands(bands);
        next[index][field] = value;
        if (inScale(value)) {
          if (field === 'min' && index < next.length - 1) next[index + 1].max = value - 1;
          if (field === 'max' && index > 0) next[index - 1].min = value + 1;
        }
        return next;
      }

      function bandsEqual(a, b) {
        return a.length === b.length && a.every((x, i) => x.min === b[i].min && x.max === b[i].max);
      }

      /** Full invariant check. Returns { valid, errors:[{index, field, code, message}] }. */
      function validateBands(bands) {
        const errors = [];
        const push = (index, field, code, message) => errors.push({ index, field, code, message });
        if (!Array.isArray(bands) || bands.length !== GRADES.length) {
          return { valid: false, errors: [{ index: null, field: null, code: 'shape', message: 'Grade bands are incomplete.' }] };
        }
        const ok = bands.map((b) => ({ min: inScale(b.min), max: inScale(b.max) }));
        bands.forEach((b, i) => {
          for (const field of ['min', 'max']) {
            const v = b[field];
            const word = field === 'min' ? 'minimum' : 'maximum';
            if (v === null || v === undefined || v === '' || Number.isNaN(v)) push(i, field, 'missing', `Enter a ${word} for ${b.grade}.`);
            else if (!isWhole(v)) push(i, field, 'not-whole', `${b.grade} ${word} must be a whole number.`);
            else if (v < 0 || v > 100) push(i, field, 'out-of-scale', `${b.grade} ${word} (${v}) must be between 0 and 100.`);
          }
        });
        const top = bands[0], bottom = bands[bands.length - 1];
        if (ok[0].max && top.max !== 100) push(0, 'max', 'top', `${top.grade} must reach 100, or marks ${top.max + 1}–100 get no grade.`);
        if (ok[bands.length - 1].min && bottom.min !== 0) push(bands.length - 1, 'min', 'bottom', `${bottom.grade} must start at 0, or marks 0–${bottom.min - 1} get no grade.`);
        bands.forEach((b, i) => {
          if (ok[i].min && ok[i].max && b.min > b.max) {
            const above = i > 0 ? ` or raise ${bands[i - 1].grade}’s minimum` : '';
            push(i, 'min', 'reversed', `${b.grade} runs from ${b.min} up to ${b.max} — its minimum is above its maximum. Lower ${b.grade}’s minimum${above}.`);
          }
          if (i > 0 && ok[i].max && ok[i - 1].min) {
            const expected = bands[i - 1].min - 1;
            if (b.max < expected) push(i, 'max', 'gap', `Marks ${b.max + 1}–${expected} fall between ${b.grade} and ${bands[i - 1].grade} and would get no grade.`);
            else if (b.max > expected) push(i, 'max', 'overlap', `${b.grade} and ${bands[i - 1].grade} overlap on marks ${bands[i - 1].min}–${b.max}.`);
          }
        });
        return { valid: errors.length === 0, errors };
      }

      /** Accepts band settings from storage only if they are complete and valid. */
      function sanitizeBands(raw) {
        if (!Array.isArray(raw) || raw.length !== GRADES.length) return null;
        const bands = raw.map((b, i) => ({ grade: GRADES[i], min: b && b.min, max: b && b.max }));
        return validateBands(bands).valid ? bands : null;
      }

      function describeBands(bands) {
        return bands.map((b) => `${b.grade} ${b.min}-${b.max}`).join(' | ');
      }

      /** Cutoff grading: the highest band whose minimum the mark reaches. */
      function gradeIndex(mark, bands) {
        for (let i = 0; i < bands.length; i++) if (mark >= bands[i].min) return i;
        return -1;
      }

      /* ------------------------------------------------------------------ statistics */

      function computeStats(values) {
        const n = values.length;
        if (!n) return null;
        const s = [...values].sort((a, b) => a - b);
        const sum = s.reduce((a, b) => a + b, 0);
        const mean = sum / n;
        const median = n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
        const sd = Math.sqrt(s.reduce((a, b) => a + (b - mean) ** 2, 0) / n);
        return { count: n, mean, median, sd, min: s[0], max: s[n - 1], range: s[n - 1] - s[0] };
      }

      /** Grades one course. Distribution fields are null while the bands are invalid. */
      function analyzeCourse(course, bands, policy, window) {
        const p = ROUNDING[policy] || ROUNDING.nearest;
        const w = window == null ? BORDERLINE_WINDOW : window;
        const validation = validateBands(bands);
        const students = course.students.map((s) => ({ ...s, mark: clean(p.apply(s.raw)), grade: null, borderline: null }));
        const stats = computeStats(students.map((s) => s.mark));
        let counts = null, borderline = null;
        if (validation.valid) {
          counts = new Array(bands.length).fill(0);
          borderline = new Array(bands.length).fill(0);
          for (const s of students) {
            const g = gradeIndex(s.mark, bands);
            if (g < 0) throw new Error(`No grade for mark ${s.mark}`);
            s.grade = g;
            counts[g]++;
            if (g > 0) {
              const gap = clean(bands[g - 1].min - s.mark);
              if (gap > 0 && gap <= w) { s.borderline = { to: g - 1, gap }; borderline[g]++; }
            }
          }
        }
        return {
          students, stats, validation, counts, borderline,
          adjusted: students.filter((s) => s.mark !== s.raw).length,
          decimals: students.filter((s) => s.decimal).length,
        };
      }

      /* ------------------------------------------------------------------ export */

      /** RFC 4180 field quoting plus a guard against spreadsheet formula injection. */
      function csvField(v) {
        if (typeof v === 'number') return String(v);
        let s = v == null ? '' : String(v);
        if (/^[=+\-@\t\r]/.test(s)) s = `'${s}`;
        return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
      }
      const csvLine = (cells) => cells.map(csvField).join(',');

      function pad2(n) { return String(n).padStart(2, '0'); }
      function stamp(d) {
        return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
      }

      /**
       * Builds the grade sheet. Refuses to build from invalid bands or ungraded students.
       * Layout keeps the original console's header block (Instructor, Course) followed by the table.
       */
      function buildGradeCsv({ instructor, courseName, bands, analysis, policy, generatedAt, isSample }) {
        if (!String(instructor || '').trim()) throw new Error('Instructor name is required.');
        if (!validateBands(bands).valid) throw new Error('Grade bands are invalid.');
        if (!analysis || !analysis.students.length) throw new Error('There are no students to export.');
        if (analysis.students.some((s) => s.grade === null || s.grade < 0)) throw new Error('Some students have no grade.');
        const lines = [
          csvLine(['Instructor', collapse(instructor)]),
          csvLine(['Course', courseName]),
          csvLine(['Grade bands', describeBands(bands)]),
        ];
        if (analysis.decimals) {
          const p = ROUNDING[policy] || ROUNDING.nearest;
          lines.push(csvLine(['Decimal marks', `${p.label} (${plural(analysis.adjusted, 'mark')} adjusted)`]));
        }
        lines.push(csvLine(['Generated', stamp(generatedAt || new Date())]));
        if (isSample) lines.push(csvLine(['Data source', 'Sample data — not real students']));
        lines.push('');
        lines.push(csvLine(['BITS ID', 'Total Marks', 'Grade']));
        for (const s of analysis.students) lines.push(csvLine([s.id, s.mark, bands[s.grade].grade]));
        return lines.join('\r\n') + '\r\n';
      }

      function buildIssuesCsv(issues, policy) {
        const lines = [csvLine(['Row', 'BITS ID', 'Course', 'Severity', 'Problem'])];
        for (const i of issues) {
          lines.push(csvLine([i.rows ? i.rows.join(' & ') : (i.row || ''), i.id, i.course, i.severity === 'error' ? 'Excluded' : 'Check', describeIssue(i, policy)]));
        }
        return lines.join('\r\n') + '\r\n';
      }

      function slug(s) {
        const out = String(s || '').normalize('NFKD').replace(/[̀-ͯ]/g, '')
          .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60);
        return out || 'course';
      }

      function exportFileName(courseName, date, isSample, kind) {
        const d = date || new Date();
        const day = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
        return `${isSample ? 'SAMPLE_' : ''}${slug(courseName)}_${kind || 'grades'}_${day}.csv`;
      }

      /* ------------------------------------------------------------------ misc formatting */

      function ordinal(n) {
        const v = n % 100;
        if (v >= 11 && v <= 13) return `${n}th`;
        return n + ({ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th');
      }

      function formatDuration(ms) {
        const total = Math.max(0, Math.floor(ms / 1000));
        const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
        if (h) return `${h} h ${pad2(m)} min`;
        if (m) return `${m} min ${pad2(s)} sec`;
        return `${s} sec`;
      }

      function formatClock(ms) {
        const total = Math.max(0, Math.floor(ms / 1000));
        const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
        return h ? `${h}:${pad2(m)}:${pad2(s)}` : `${pad2(m)}:${pad2(s)}`;
      }

      /* ------------------------------------------------------------------ sample data */

      /**
       * A clearly-labelled demonstration workbook (rows as arrays, header first). Deterministic, and it
       * deliberately contains a few problems so the data-quality checks have something to show.
       */
      function sampleRows() {
        let seed = 20260930;
        const rand = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
        const normal = (mean, sd) => {
          const u = Math.max(rand(), 1e-9), v = rand();
          return Math.max(4, Math.min(99, Math.round(mean + sd * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v))));
        };
        const rows = [['BITS ID', 'Course', 'Total Marks']];
        const plan = [
          { course: 'CS F111', branch: 'A7', year: 2025, n: 64, mean: 63, sd: 15 },
          { course: 'MATH F113', branch: 'A4', year: 2025, n: 48, mean: 57, sd: 17 },
          { course: 'ECON F211', branch: 'B3', year: 2024, n: 36, mean: 69, sd: 11 },
        ];
        let serial = 1;
        for (const p of plan) {
          for (let i = 0; i < p.n; i++) {
            const id = `${p.year}${p.branch}PS${String(serial++).padStart(4, '0')}P`;
            rows.push([id, p.course, normal(p.mean, p.sd)]);
          }
        }
        rows[5][2] = 79.5;         // decimal mark on a boundary
        rows[12][2] = 'AB';        // absent student left in the file
        rows.push(rows[20].slice()); // identical duplicate
        rows[70][2] = null;        // blank marks
        return rows;
      }

      return {
        GRADES, DEFAULT_BANDS, ROUNDING, BORDERLINE_WINDOW, ACCEPTED_EXTENSIONS, MAX_FILE_BYTES,
        extensionOf, checkFileMeta, parseFile, parseWorkbook, headerScores, readMarks, readCell,
        describeIssue, defaultBands, cloneBands, setBandEdge, bandsEqual, validateBands, sanitizeBands,
        describeBands, gradeIndex, computeStats, analyzeCourse,
        csvField, buildGradeCsv, buildIssuesCsv, exportFileName, slug, ordinal, formatDuration, formatClock,
        sampleRows,
      };
    });
