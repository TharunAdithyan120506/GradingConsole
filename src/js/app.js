    /*!
     * BITS Pilani Digital · Grading Console — interface
     * Depends on GradingCore (assets/core.js). SheetJS is loaded on demand.
     */
    (function () {
      'use strict';

      const C = window.GradingCore;
      const G = C.GRADES;
      const PAGE_SIZE = 25;
      const ISSUE_RENDER_CAP = 200;

      /* The patched SheetJS build is tried first; public npm mirrors are fallbacks. */
      const SHEETJS_SOURCES = [
        'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
      ];

      /* ================================================================ helpers */

      const $ = (id) => document.getElementById(id);
      const SVGNS = 'http://www.w3.org/2000/svg';

      /** Creates an element. props: class, text, html-free attributes, dataset, style, on* handlers. */
      function h(tag, props, ...children) {
        const el = document.createElement(tag);
        if (props) {
          for (const [k, v] of Object.entries(props)) {
            if (v === undefined || v === null || v === false) continue;
            if (k === 'class') el.className = v;
            else if (k === 'text') el.textContent = v;
            else if (k === 'dataset') Object.assign(el.dataset, v);
            else if (k === 'style') el.style.cssText = v;
            else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
            else el.setAttribute(k, v === true ? '' : v);
          }
        }
        for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) el.append(c.nodeType ? c : String(c));
        return el;
      }
      function svgEl(tag, attrs) {
        const el = document.createElementNS(SVGNS, tag);
        for (const [k, v] of Object.entries(attrs || {})) el.setAttribute(k, v);
        return el;
      }
      function icon(name, cls) {
        const s = svgEl('svg', { 'aria-hidden': 'true' });
        if (cls) s.setAttribute('class', cls);
        s.append(svgEl('use', { href: `#i-${name}` }));
        return s;
      }
      const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
      const plural = (n, one, many) => `${n.toLocaleString('en-IN')} ${n === 1 ? one : (many || one + 's')}`;
      const shown = (g) => g.replace('-', '−');
      const fmt = (v, d = 0) => Number(v).toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d });
      const fmtMark = (v) => (Number.isInteger(v) ? String(v) : fmt(v, 1));
      const nextFrame = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
      const motionOK = () => !reduceMotion.matches;

      function restartClass(el, cls) {
        el.classList.remove(cls);
        void el.offsetWidth; // restart the CSS animation
        el.classList.add(cls);
      }

      const store = {
        get(key, fallback) {
          try { const v = localStorage.getItem(`bdgc:${key}`); return v === null ? fallback : JSON.parse(v); } catch (_) { return fallback; }
        },
        set(key, value) { try { localStorage.setItem(`bdgc:${key}`, JSON.stringify(value)); } catch (_) { /* storage blocked: keep working without it */ } },
        remove(key) { try { localStorage.removeItem(`bdgc:${key}`); } catch (_) { /* storage blocked */ } },
      };

      /* ================================================================ SheetJS loader */

      let sheetjsPromise = null;
      function injectScript(src, timeout) {
        return new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = src;
          s.async = true;
          const timer = setTimeout(() => { s.remove(); reject(new Error('timeout')); }, timeout);
          s.onload = () => { clearTimeout(timer); resolve(); };
          s.onerror = () => { clearTimeout(timer); s.remove(); reject(new Error('blocked')); };
          document.head.append(s);
        });
      }
      function loadSheetJS() {
        if (window.XLSX && window.XLSX.read) return Promise.resolve(window.XLSX);
        if (!sheetjsPromise) {
          sheetjsPromise = (async () => {
            for (const src of SHEETJS_SOURCES) {
              try {
                await injectScript(src, 12000);
                if (window.XLSX && window.XLSX.read) return window.XLSX;
              } catch (_) { /* try the next mirror */ }
            }
            sheetjsPromise = null; // allow a retry on the next attempt
            throw new Error('The spreadsheet reader could not be loaded.');
          })();
        }
        return sheetjsPromise;
      }

      /* ================================================================ state */

      const state = {
        dataset: null,
        courseKey: null,
        bands: C.defaultBands(),
        history: { past: [], future: [], key: null, at: 0 },
        rounding: C.ROUNDING[store.get('rounding', 'nearest')] ? store.get('rounding', 'nearest') : 'nearest',
        analysis: null,
        roster: { query: '', filter: 'all', sort: 'row', dir: 1, page: 1 },
        issueScope: 'course',
        issuesOpen: false,
        ack: false,
        curve: store.get('curve', false) === true,
        tableOpen: false,
        sessions: new Map(),
        busy: false,
      };

      const course = () => (state.dataset && state.courseKey ? state.dataset.courses.find((c) => c.key === state.courseKey) || null : null);
      const session = () => (state.courseKey ? state.sessions.get(state.courseKey) : null);
      const bandsValid = () => !!(state.analysis && state.analysis.validation.valid);

      function analyze() {
        const c = course();
        state.analysis = c ? C.analyzeCourse(c, state.bands, state.rounding) : null;
      }

      /* ================================================================ announcements & toasts */

      const srStatus = $('srStatus');
      let announceTimer = 0;
      function announce(msg) {
        clearTimeout(announceTimer);
        srStatus.textContent = '';
        announceTimer = setTimeout(() => { srStatus.textContent = msg; }, 60);
      }

      const toastsEl = $('toasts');
      function toast(message, actionLabel, action, timeout = 5500) {
        while (toastsEl.children.length >= 3) toastsEl.firstElementChild.remove();
        const t = h('div', { class: 'toast' }, h('span', { class: 'toast-msg', text: message }));
        let timer = 0;
        const close = () => {
          clearTimeout(timer);
          if (!t.isConnected) return;
          if (!motionOK()) { t.remove(); return; }
          t.classList.add('leave');
          setTimeout(() => t.remove(), 220);
        };
        if (actionLabel) t.append(h('button', { type: 'button', class: 'toast-btn', text: actionLabel, onclick: () => { close(); action(); } }));
        const arm = () => { clearTimeout(timer); timer = setTimeout(close, timeout); };
        t.addEventListener('pointerenter', () => clearTimeout(timer));
        t.addEventListener('pointerleave', arm);
        t.addEventListener('focusin', () => clearTimeout(timer));
        t.addEventListener('focusout', arm);
        toastsEl.append(t);
        arm();
      }

      /* ================================================================ theme */

      const root = document.documentElement;
      const themeBtn = $('themeBtn');
      const systemDark = window.matchMedia('(prefers-color-scheme: dark)');
      const effectiveTheme = () => root.dataset.theme || (systemDark.matches ? 'dark' : 'light');
      function renderThemeButton() {
        const dark = effectiveTheme() === 'dark';
        themeBtn.setAttribute('aria-label', dark ? 'Switch to light theme' : 'Switch to dark theme');
        themeBtn.querySelector('use').setAttribute('href', dark ? '#i-sun' : '#i-moon');
      }
      function toggleTheme() {
        const next = effectiveTheme() === 'dark' ? 'light' : 'dark';
        if (motionOK()) {
          root.classList.add('theme-anim');
          setTimeout(() => root.classList.remove('theme-anim'), 420);
        }
        root.dataset.theme = next;
        store.set('theme', next);
        renderThemeButton();
      }

      /* ================================================================ upload */

      const fileInput = $('fileInput');
      const dropzone = $('dropzone');
      const fileCard = $('fileCard');

      function setBusy(on, label, progress) {
        state.busy = on;
        dropzone.classList.toggle('busy', on);
        dropzone.setAttribute('aria-busy', on ? 'true' : 'false');
        fileCard.setAttribute('aria-busy', on ? 'true' : 'false');
        const bar = $('dzProgress');
        bar.hidden = !on;
        if (!on) return;
        if (label) $('dzProgressLabel').textContent = label;
        const track = bar.querySelector('.dz-track');
        track.classList.toggle('indeterminate', progress === null);
        if (progress !== null && progress !== undefined) $('dzFill').style.transform = `scaleX(${clamp(progress, 0, 1)})`;
        if (state.dataset && label) $('fcMeta').textContent = label;
      }

      function readFile(file, onProgress) {
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onprogress = (e) => { if (e.lengthComputable) onProgress(e.loaded / e.total); };
          reader.onload = () => resolve(reader.result);
          reader.onerror = () => reject(reader.error || new Error('The browser could not read the file.'));
          reader.readAsArrayBuffer(file);
        });
      }

      async function handleFile(file) {
        if (state.busy || !file) return;
        hideUploadError();
        const meta = { name: file.name, size: file.size };
        const pre = C.checkFileMeta(meta);
        if (!pre.ok) { showUploadError(pre.error); return; }
        setBusy(true, `Reading ${file.name}…`, 0);
        try {
          const buf = await readFile(file, (p) => setBusy(true, null, p));
          setBusy(true, 'Checking every row…', null);
          let XLSX;
          try { XLSX = await loadSheetJS(); } catch (_) {
            showUploadError({
              title: 'The spreadsheet reader didn’t load',
              detail: 'The console reads Excel files with a small library fetched from the internet, and it could not be reached.',
              hint: 'Check your internet connection, then upload the file again.',
            });
            return;
          }
          await nextFrame();
          const res = C.parseFile(new Uint8Array(buf), meta, XLSX);
          if (!res.ok) { showUploadError(res.error); return; }
          loadDataset(res.dataset);
        } catch (err) {
          console.error(err);
          showUploadError({
            title: 'Something went wrong while reading this file',
            detail: String((err && err.message) || err),
            hint: 'Save the sheet again as .xlsx and retry. Nothing has been changed.',
          });
        } finally {
          setBusy(false);
          fileInput.value = ''; // selecting the same file again must trigger a new upload
          if (state.dataset) renderFileCard(false);
        }
      }

      function showUploadError(err) {
        $('uploadErrorTitle').textContent = err.title;
        $('uploadErrorDetail').textContent = err.detail || '';
        $('uploadErrorHint').textContent = err.hint || '';
        $('uploadErrorHint').hidden = !err.hint;
        const keep = $('uploadErrorKeep');
        keep.hidden = !state.dataset;
        if (state.dataset) keep.textContent = `Your previous file, “${state.dataset.fileName}”, is still loaded.`;
        const box = $('uploadError');
        box.hidden = false;
        restartClass(box, 'alert');
        box.scrollIntoView({ block: 'nearest', behavior: motionOK() ? 'smooth' : 'auto' });
      }
      function hideUploadError() { $('uploadError').hidden = true; }

      async function loadSample() {
        if (state.busy) return;
        hideUploadError();
        setBusy(true, 'Preparing sample data…', null);
        try {
          const XLSX = await loadSheetJS();
          const wb = { SheetNames: ['Marks'], Sheets: { Marks: XLSX.utils.aoa_to_sheet(C.sampleRows()) } };
          const res = C.parseWorkbook(wb, XLSX, { name: 'Sample marks.xlsx', size: 0, isSample: true });
          if (res.ok) loadDataset(res.dataset);
        } catch (_) {
          showUploadError({ title: 'The sample couldn’t be prepared', detail: 'The spreadsheet library did not load.', hint: 'Check your internet connection and try again.' });
        } finally {
          setBusy(false);
        }
      }

      async function downloadSampleSheet() {
        try {
          const XLSX = await loadSheetJS();
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.aoa_to_sheet(C.sampleRows());
          ws['!cols'] = [{ wch: 18 }, { wch: 14 }, { wch: 12 }];
          XLSX.utils.book_append_sheet(wb, ws, 'Marks');
          XLSX.writeFile(wb, 'BITS grading sample.xlsx');
          toast('Downloaded “BITS grading sample.xlsx”. It includes a few planted problems to show the checks.');
        } catch (_) {
          toast('The sample sheet couldn’t be created. Check your connection and try again.');
        }
      }

      function loadDataset(ds) {
        state.dataset = ds;
        state.courseKey = null;
        state.analysis = null;
        state.sessions.clear();
        state.issuesOpen = false;
        state.issueScope = 'course';
        $('courseSearch').value = '';
        renderFileCard(true);
        renderCourses();
        renderWorkspace();
        renderAside();
        renderHero();
        renderSteps();
        const sec = $('sec-course');
        restartClass(sec, 'enter');
        announce(`Loaded ${ds.fileName}: ${plural(ds.gradable, 'student')} ready in ${plural(ds.courses.length, 'course')}.`);
        if (ds.courses.length === 1) selectCourse(ds.courses[0].key, { scroll: true });
        else {
          sec.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'start' });
          $('h-course').focus({ preventScroll: true });
        }
      }

      function removeDataset() {
        const prev = { dataset: state.dataset, courseKey: state.courseKey };
        state.dataset = null;
        state.courseKey = null;
        state.analysis = null;
        state.sessions.clear();
        stopTicker();
        renderAll();
        toast(`Removed “${prev.dataset.fileName}”.`, 'Undo', () => {
          state.dataset = prev.dataset;
          renderAll();
          if (prev.courseKey) selectCourse(prev.courseKey, { scroll: false });
        });
        $('h-upload').focus({ preventScroll: true });
      }

      /* ================================================================ file card & data checks */

      function renderFileCard(entering) {
        const ds = state.dataset;
        $('uploadEmpty').hidden = !!ds;
        fileCard.hidden = !ds;
        if (!ds) return;
        if (entering) restartClass(fileCard, 'enter');
        $('fcName').textContent = ds.fileName;
        $('fcSample').hidden = !ds.isSample;
        const sizeKb = ds.fileSize ? ` · ${ds.fileSize < 1048576 ? `${Math.max(1, Math.round(ds.fileSize / 1024))} KB` : `${(ds.fileSize / 1048576).toFixed(1)} MB`}` : '';
        $('fcMeta').textContent = `${ds.isSample ? 'Demonstration data, not real students' : `Sheet “${ds.sheetName}”`}${sizeKb} · ${plural(ds.rowsRead, 'row')} read`;

        const warnings = ds.issues.filter((i) => i.severity === 'warning').length;
        const chips = [
          [`${plural(ds.gradable, 'student')} ready`, 'ok'],
          [plural(ds.courses.length, 'course'), ''],
        ];
        if (ds.excluded) chips.push([`${plural(ds.excluded, 'row')} excluded`, 'danger']);
        if (warnings) chips.push([`${plural(warnings, 'item')} to check`, 'warn']);
        if (ds.blankRows) chips.push([`${plural(ds.blankRows, 'blank row')} skipped`, '']);
        $('qChips').replaceChildren(...chips.map(([t, cls]) => h('li', { class: `q-chip ${cls}`, text: t })));

        const toggle = $('issuesToggle');
        const hasDetail = ds.issues.length > 0 || ds.notes.length > 0;
        toggle.hidden = !hasDetail;
        toggle.setAttribute('aria-expanded', String(state.issuesOpen));
        toggle.textContent = state.issuesOpen ? 'Hide details' : (ds.issues.length ? `Review ${plural(ds.issues.length, 'row')}` : 'File details');
        $('issuesPanel').hidden = !state.issuesOpen;
        if (state.issuesOpen) renderIssues();
      }

      function scopedIssues() {
        const ds = state.dataset;
        if (!ds) return [];
        const scope = state.courseKey ? state.issueScope : 'all';
        return scope === 'course' ? ds.issues.filter((i) => i.courseKey === state.courseKey) : ds.issues;
      }

      function renderIssues() {
        const ds = state.dataset;
        const notes = [...ds.notes];
        notes.unshift(`Columns used: “${ds.columns.id}”, “${ds.columns.course}”, “${ds.columns.marks}” (header on row ${ds.headerRow}).`);
        $('issuesNotes').textContent = notes.join(' ');
        const scope = state.courseKey ? state.issueScope : 'all';
        document.querySelectorAll('[data-issue-scope]').forEach((b) => {
          b.setAttribute('aria-pressed', String(b.dataset.issueScope === scope));
          b.disabled = b.dataset.issueScope === 'course' && !state.courseKey;
        });
        const list = scopedIssues();
        const rows = list.slice(0, ISSUE_RENDER_CAP).map((i) => h('tr', null,
          h('td', { class: 'num', text: i.rows ? i.rows.join(', ') : (i.row || '—') }),
          h('td', { class: 'mono', text: i.id || '—' }),
          h('td', { text: i.course || '—' }),
          h('td', null, h('span', { class: `sev sev-${i.severity}` }, icon(i.severity === 'error' ? 'x' : 'alert'), i.severity === 'error' ? 'Excluded' : 'Check')),
          h('td', { text: C.describeIssue(i, state.rounding) })));
        if (!list.length) rows.push(h('tr', null, h('td', { colspan: '5', class: 'empty-note', text: scope === 'course' ? 'No problems found in this course.' : 'No problems found in this file.' })));
        if (list.length > ISSUE_RENDER_CAP) rows.push(h('tr', null, h('td', { colspan: '5', class: 'empty-note', text: `Showing the first ${ISSUE_RENDER_CAP} of ${list.length}. Download the list to see them all.` })));
        $('issuesBody').replaceChildren(...rows);
        $('issuesDownload').disabled = !list.length;
      }

      function downloadIssues() {
        const list = scopedIssues();
        if (!list.length) return;
        const c = course();
        const base = state.issueScope === 'course' && c ? c.name : state.dataset.fileName.replace(/\.[^.]+$/, '');
        const name = C.exportFileName(base, new Date(), state.dataset.isSample, 'issues');
        downloadText(C.buildIssuesCsv(list, state.rounding), name);
        toast(`Downloaded ${name}.`);
      }

      /* ================================================================ courses */

      function renderCourses() {
        const ds = state.dataset;
        const sec = $('sec-course');
        sec.hidden = !ds;
        if (!ds) return;
        const wrap = $('courses');
        const q = $('courseSearch').value.trim().toLowerCase();
        $('courseSearchWrap').hidden = ds.courses.length <= 6;
        wrap.classList.toggle('compact', !!state.courseKey);
        $('h-course').textContent = state.courseKey ? `Course · ${ds.courses.length} in this file` : 'Courses in this file';
        const policy = C.ROUNDING[state.rounding];
        const list = ds.courses.filter((c) => !q || c.name.toLowerCase().includes(q));
        const focusedKey = document.activeElement && document.activeElement.classList.contains('course') ? document.activeElement.dataset.key : null;
        wrap.replaceChildren(...list.map((c, i) => {
          const stats = C.computeStats(c.students.map((s) => policy.apply(s.raw)));
          const warnings = ds.issues.filter((x) => x.courseKey === c.key && x.severity === 'warning').length;
          const flag = c.excluded ? h('span', { class: 'course-flag danger' }, icon('alert'), `${c.excluded} excluded`)
            : warnings ? h('span', { class: 'course-flag' }, `${warnings} to check`) : null;
          const selected = c.key === state.courseKey;
          return h('button', {
            type: 'button', class: 'course', 'aria-pressed': String(selected), style: `--i:${Math.min(i, 12)}`,
            dataset: { key: c.key },
            onclick: () => selectCourse(c.key, { scroll: true }),
          },
            h('span', { class: 'course-top' }, h('span', { class: 'course-badge' }, icon('book')), flag),
            h('span', { class: 'course-tick' }, icon('check')),
            h('span', { class: 'course-name', text: c.name }),
            h('span', { class: 'course-meta' }, h('b', { text: String(c.students.length) }), ` ${c.students.length === 1 ? 'student' : 'students'}`, stats ? ` · mean ${fmt(stats.mean, 1)}` : ' · nothing to grade'),
            sparkline(c.students.map((s) => policy.apply(s.raw))));
        }));
        $('courseNone').hidden = list.length > 0;
        if (focusedKey) {
          const again = [...wrap.children].find((b) => b.dataset.key === focusedKey);
          if (again) again.focus({ preventScroll: true });
        }
        if (state.courseKey) {
          const sel = wrap.querySelector('[aria-pressed="true"]');
          if (sel) requestAnimationFrame(() => sel.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'auto' }));
        }
      }

      /** A 20-bin marks histogram (5 marks per bin) for a course card. Decorative; the card text carries the numbers. */
      function sparkline(marks) {
        const bins = new Array(20).fill(0);
        marks.forEach((m) => { bins[Math.min(19, Math.floor(m / 5))]++; });
        const max = Math.max(...bins);
        const s = svgEl('svg', { class: 'course-spark', viewBox: '0 0 200 30', preserveAspectRatio: 'none', 'aria-hidden': 'true' });
        s.append(svgEl('rect', { class: 'spark-base', x: 0, y: 29, width: 200, height: 1 }));
        bins.forEach((n, i) => {
          if (!n) return;
          const hgt = Math.max(2, (n / max) * 28);
          s.append(svgEl('rect', { x: i * 10 + 1, y: 30 - hgt, width: 8, height: hgt, rx: 1.5 }));
        });
        return s;
      }

      function selectCourse(key, { scroll } = {}) {
        if (!state.dataset || state.courseKey === key) return;
        state.courseKey = key;
        const c = course();
        const saved = C.sanitizeBands(store.get(`bands:${key}`, null));
        state.bands = saved || C.defaultBands();
        state.history = { past: [], future: [], key: null, at: 0 };
        state.roster = { query: '', filter: 'all', sort: 'row', dir: 1, page: 1 };
        $('rosterSearch').value = '';
        state.ack = false;
        state.issueScope = 'course';
        $('exportDone').hidden = true;
        if (!state.sessions.has(key)) state.sessions.set(key, { start: Date.now(), exports: 0, frozenAt: null });
        analyze();
        renderCourses();
        renderWorkspace({ fresh: true });
        renderFileCard(false);
        renderAside();
        renderHero();
        renderSteps();
        startTicker();
        if (saved && !C.bandsEqual(saved, C.defaultBands())) {
          toast(`Using your saved bands for ${c.name}.`, 'Use defaults', () => applyBands(C.defaultBands(), { message: 'Restored the default bands.' }));
        }
        announce(`${c.name} opened: ${plural(c.students.length, 'student')}.`);
        if (scroll) {
          const target = $('sec-grade');
          requestAnimationFrame(() => target.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'start' }));
        }
      }

      /* ================================================================ hero & steps */

      const STEP_LABELS = { upload: 'Upload marks', course: 'Choose a course', bands: 'Review the grade bands', export: 'Sign and export' };
      let lastHero = '';
      function renderHero() {
        const ds = state.dataset;
        const c = course();
        const a = state.analysis;
        const name = $('instructor').value.trim();
        const steps = stepStates();
        const order = ['upload', 'course', 'bands', 'export'];
        const at = order.findIndex((k) => steps[k] === 'current');
        const where = at >= 0 ? `Step ${at + 1} of 4 · ${STEP_LABELS[order[at]]}` : 'Grade sheet exported';
        $('heroEyebrow').textContent = name ? `Welcome, ${name} · ${where}` : where;
        let title, sub;
        if (!ds) {
          title = 'Turn raw marks into final grades.';
          sub = 'Upload a marks sheet, tune the grade bands against the real distribution, and export a clean grade sheet. Everything runs in this browser tab; no marks are sent to a server.';
        } else if (!c) {
          title = 'Choose a course to grade.';
          sub = `${ds.fileName} · ${plural(ds.gradable, 'student')} across ${plural(ds.courses.length, 'course')}.${ds.isSample ? ' Sample data, not real students.' : ''}`;
        } else {
          title = c.name;
          if (!a || !a.stats) sub = 'This course has no students that can be graded yet.';
          else {
            const near = a.borderline ? a.borderline.reduce((x, y) => x + y, 0) : 0;
            sub = `${plural(a.stats.count, 'student')} · mean ${fmt(a.stats.mean, 1)} · median ${fmtMark(a.stats.median)}`;
            if (!a.validation.valid) sub += ' · grade bands need fixing';
            else if (near) sub += ` · ${plural(near, 'student')} within ${C.BORDERLINE_WINDOW} marks of the next grade up`;
            if (ds.isSample) sub += ' · sample data';
          }
        }
        const titleEl = $('heroTitle');
        if (!ds) titleEl.replaceChildren('Turn raw marks into ', h('em', { text: 'final grades.' }));
        else titleEl.textContent = title;
        $('heroSub').textContent = sub;
        if (lastHero && lastHero !== title) restartClass(document.querySelector('.hero'), 'swap');
        lastHero = title;
      }

      function stepStates() {
        const ds = !!state.dataset;
        const c = course();
        const hasStudents = !!(c && c.students.length);
        const valid = bandsValid();
        const exported = !!(session() && session().exports > 0);
        const s = {
          upload: ds ? 'done' : 'todo',
          course: c ? 'done' : ds ? 'todo' : 'locked',
          bands: hasStudents && valid ? 'done' : hasStudents ? 'todo' : 'locked',
          export: exported ? 'done' : hasStudents && valid ? 'todo' : 'locked',
        };
        const current = ['upload', 'course', 'bands', 'export'].find((k) => s[k] === 'todo');
        if (current) s[current] = 'current';
        return s;
      }

      function renderSteps() {
        const s = stepStates();
        document.querySelectorAll('.step').forEach((btn) => {
          const st = s[btn.dataset.step];
          btn.dataset.state = st;
          btn.setAttribute('aria-disabled', String(st === 'locked'));
          if (st === 'current') btn.setAttribute('aria-current', 'step'); else btn.removeAttribute('aria-current');
          const label = btn.querySelector('.step-label').textContent;
          btn.setAttribute('aria-label', `${label}: ${st === 'done' ? 'done' : st === 'current' ? 'current step' : st === 'locked' ? 'not available yet' : 'to do'}`);
        });
      }

      function goToStep(btn) {
        if (btn.getAttribute('aria-disabled') === 'true') {
          const msg = { course: 'Upload a marks file first.', bands: 'Open a course with gradable students first.', export: 'Open a course and make sure its bands are valid first.' };
          toast(msg[btn.dataset.step] || 'Not available yet.');
          return;
        }
        const target = $(btn.dataset.target);
        if (!target) return;
        target.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'start' });
        const focusable = btn.dataset.step === 'upload' ? (state.dataset ? $('h-upload') : fileInput)
          : btn.dataset.step === 'course' ? $('h-course') : target;
        focusable.focus({ preventScroll: true });
      }

      /* ================================================================ workspace */

      function renderWorkspace({ fresh = false } = {}) {
        const c = course();
        const sec = $('sec-grade');
        sec.hidden = !c;
        if (!c) return;
        if (fresh) restartClass(sec, 'enter');
        const a = state.analysis;
        const empty = !a.students.length;
        $('emptyCourse').hidden = !empty;
        $('courseBody').hidden = empty;
        if (empty) {
          $('emptyCourseText').textContent = `All ${plural(c.excluded, 'row')} for ${c.name} have problems, so there is nothing to grade. Open “Review rows” above to see what to fix, correct the sheet and upload it again.`;
          return;
        }
        renderStats(fresh);
        renderDecimals();
        if (fresh) buildBands();
        buildChart(fresh);
        renderBands();
        renderRoster(fresh);
      }

      /* ---------------------------------------------------------------- stats */

      function tween(el, to, decimals, fromZero) {
        const prev = fromZero ? 0 : Number(el.dataset.v);
        el.dataset.v = String(to);
        cancelAnimationFrame(el._raf);
        if (!motionOK() || !Number.isFinite(prev) || prev === to) { el.textContent = fmt(to, decimals); return; }
        const start = performance.now();
        const dur = 650;
        const step = (t) => {
          const p = Math.min(1, (t - start) / dur);
          const e = 1 - Math.pow(1 - p, 3);
          el.textContent = fmt(prev + (to - prev) * e, decimals);
          if (p < 1) el._raf = requestAnimationFrame(step);
        };
        el._raf = requestAnimationFrame(step);
      }

      function renderStats(fresh) {
        const a = state.analysis;
        const s = a.stats;
        const c = course();
        const dec = (v) => (Number.isInteger(v) ? 0 : 1);
        const set = (key, v, d) => tween(document.querySelector(`[data-stat="${key}"]`), v, d, fresh);
        set('count', s.count, 0);
        set('mean', s.mean, 1);
        set('median', s.median, dec(s.median));
        set('sd', s.sd, 1);
        set('max', s.max, dec(s.max));
        set('min', s.min, dec(s.min));
        $('statCountNote').textContent = c.excluded ? `${c.excluded} excluded` : 'all rows graded';
        $('statMaxNote').textContent = 'top mark';
        $('statRangeNote').textContent = `range ${fmtMark(s.range)} marks`;
        document.querySelectorAll('.stat').forEach((el, i) => { el.style.setProperty('--i', i); if (fresh) restartClass(el, 'stat'); });
      }

      function renderDecimals() {
        const a = state.analysis;
        const box = $('decimalsBox');
        box.hidden = !a.decimals;
        if (!a.decimals) return;
        $('decimalsTitle').textContent = `${plural(a.decimals, 'mark')} in this course ${a.decimals === 1 ? 'has' : 'have'} decimals`;
        document.querySelectorAll('[data-rounding]').forEach((b) => {
          const on = b.dataset.rounding === state.rounding;
          b.setAttribute('aria-checked', String(on));
          b.tabIndex = on ? 0 : -1;
        });
      }

      function setRounding(r) {
        if (!C.ROUNDING[r] || r === state.rounding) return;
        state.rounding = r;
        store.set('rounding', r);
        analyze();
        resumeSession();
        renderCourses();
        renderWorkspace();
        renderFileCard(false);
        renderAside();
        renderHero();
        renderSteps();
        announce(`Decimal marks: ${C.ROUNDING[r].label}.`);
      }

      /* ================================================================ chart */

      const chartEl = $('chart');
      const chartSvg = $('chartSvg');
      const chartTip = $('chartTip');
      const handlesEl = $('handles');
      const chart = { L: null, cols: null, bars: [], bandBg: [], ribs: [], ribLabels: [], cuts: [], handles: [], hotCol: -1, hot: null, drag: null };

      function layout() {
        const w = chartEl.clientWidth;
        const hgt = chartEl.clientHeight;
        const L = { w, h: hgt, padL: 36, padR: 10, top: 36, ribbonH: 28, axisH: 22 };
        L.ribbonTop = hgt - L.axisH - L.ribbonH;
        L.plotBottom = L.ribbonTop - 8;
        L.plotH = L.plotBottom - L.top;
        L.plotW = Math.max(120, w - L.padL - L.padR);
        L.colW = L.plotW / 101;
        L.x = (m) => L.padL + m * L.colW;
        return L;
      }

      function niceScale(max) {
        if (max <= 4) return { yMax: Math.max(1, max), step: 1 };
        const raw = max / 4;
        const mag = Math.pow(10, Math.floor(Math.log10(raw)));
        const n = raw / mag;
        const step = Math.max(1, Math.round((n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * mag));
        return { yMax: Math.ceil(max / step) * step, step };
      }

      const normalPdf = (x, mu, sd) => Math.exp(-0.5 * ((x - mu) / sd) ** 2) / (sd * Math.sqrt(2 * Math.PI));

      function barPath(x, y, w, base) {
        const r = Math.min(3, w / 2, base - y);
        return `M${x},${base}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${base}Z`;
      }

      function buildChart(animate) {
        const a = state.analysis;
        const desc = $('chartDesc');
        chartSvg.replaceChildren(desc);
        hideTip();
        if (!a || !a.students.length || !chartEl.clientWidth) { chart.L = null; return; }
        const L = layout();
        chart.L = L;
        const cols = Array.from({ length: 101 }, () => []);
        for (const s of a.students) cols[clamp(Math.floor(s.mark), 0, 100)].push(s);
        chart.cols = cols;

        const s = a.stats;
        const curveOK = s.count >= 5 && s.sd > 0;
        const curveOn = state.curve && curveOK;
        const curveToggle = $('curveToggle');
        curveToggle.checked = state.curve;
        curveToggle.disabled = !curveOK;
        $('curveNote').hidden = curveOK;
        $('curveNote').textContent = curveOK ? '' : 'The normal curve needs at least 5 students with different marks.';

        let peak = Math.max(...cols.map((c) => c.length));
        const curvePts = [];
        if (curveOn) {
          for (let t = 0; t <= 101; t += 0.25) {
            const v = s.count * normalPdf(t - 0.5, s.mean, s.sd);
            curvePts.push([t, v]);
            peak = Math.max(peak, v);
          }
        }
        const { yMax, step } = niceScale(peak);
        const y = (v) => L.plotBottom - (v / yMax) * L.plotH;

        const defs = svgEl('defs');
        const clip = svgEl('clipPath', { id: 'ribClip' });
        clip.append(svgEl('rect', { x: L.padL, y: L.ribbonTop, width: L.plotW, height: L.ribbonH, rx: 9 }));
        defs.append(clip);
        const gBands = svgEl('g');
        const gGrid = svgEl('g');
        const colHot = svgEl('rect', { class: 'col-hot', x: 0, y: L.top, width: L.colW, height: L.plotH, opacity: 0 });
        const gBars = svgEl('g');
        const gOver = svgEl('g');
        const gRib = svgEl('g', { 'clip-path': 'url(#ribClip)' });
        const gRibLabels = svgEl('g', { 'aria-hidden': 'true' });
        const gAxis = svgEl('g', { 'aria-hidden': 'true' });

        chart.bandBg = G.map((_, i) => { const r = svgEl('rect', { class: `band-bg b${i}`, x: 0, y: L.top, width: 1, height: L.plotH }); gBands.append(r); return r; });
        for (let v = 0; v <= yMax; v += step) {
          const yy = Math.round(y(v)) + 0.5;
          gGrid.append(svgEl('line', { class: v === 0 ? 'base' : 'grid', x1: L.padL, x2: L.padL + L.plotW, y1: yy, y2: yy }));
          const t = svgEl('text', { x: L.padL - 8, y: yy + 4, 'text-anchor': 'end' });
          t.textContent = String(v);
          gAxis.append(t);
        }
        const unit = svgEl('text', { x: 0, y: L.top - 14, 'text-anchor': 'start' });
        unit.textContent = 'students';
        gAxis.append(unit);

        const gap = Math.min(2, L.colW * 0.28);
        const bw = Math.max(1, L.colW - gap);
        chart.bars = new Array(101).fill(null);
        let order = 0;
        cols.forEach((list, m) => {
          if (!list.length) return;
          const p = svgEl('path', { class: 'bar', d: barPath(L.x(m) + gap / 2, y(list.length), bw, L.plotBottom) });
          if (animate && motionOK()) { p.classList.add('grow'); p.style.setProperty('--delay', `${Math.round(m * 3.2) + order++}ms`); }
          gBars.append(p);
          chart.bars[m] = p;
        });

        if (curveOn) {
          const d = curvePts.map(([t, v], i) => `${i ? 'L' : 'M'}${(L.padL + t * L.colW).toFixed(1)},${y(v).toFixed(1)}`).join('');
          const path = svgEl('path', { class: 'curve', d });
          gOver.append(path);
          if (animate === 'curve' && motionOK()) {
            requestAnimationFrame(() => {
              const len = Math.ceil(path.getTotalLength());
              path.style.setProperty('--len', len);
              path.classList.add('draw');
            });
          }
        }

        // mean and median markers, labels on separate rows when close
        const xm = L.x(s.mean + 0.5), xd = L.x(s.median + 0.5);
        const close = Math.abs(xm - xd) < 92;
        [[xm, `Mean ${fmt(s.mean, 1)}`, 13], [xd, `Median ${fmtMark(s.median)}`, close ? 27 : 13]].forEach(([x, label, ty]) => {
          gOver.append(svgEl('line', { class: 'marker', x1: x, x2: x, y1: ty + 4, y2: L.plotBottom }));
          gOver.append(svgEl('circle', { class: 'marker-dot', cx: x, cy: ty + 4, r: 3.5 }));
          const right = x > L.w - 96;
          const t = svgEl('text', { class: 'marker-label', x: right ? x - 7 : x + 7, y: ty + 8, 'text-anchor': right ? 'end' : 'start' });
          t.textContent = label;
          gOver.append(t);
        });

        chart.cuts = G.slice(0, -1).map(() => { const l = svgEl('line', { class: 'cut', x1: 0, x2: 0, y1: L.top - 4, y2: L.ribbonTop + L.ribbonH }); gOver.append(l); return l; });
        chart.ribs = G.map((_, i) => { const r = svgEl('rect', { class: `rib b${i}`, x: 0, y: L.ribbonTop, width: 1, height: L.ribbonH }); gRib.append(r); return r; });
        chart.ribLabels = G.map((g) => { const t = svgEl('text', { class: 'rib-label', x: 0, y: L.ribbonTop + L.ribbonH / 2 + 4.5 }); t.textContent = shown(g); gRibLabels.append(t); return t; });
        for (let m = 0; m <= 100; m += 10) {
          const t = svgEl('text', { x: L.x(m) + L.colW / 2, y: L.ribbonTop + L.ribbonH + 16, 'text-anchor': 'middle' });
          t.textContent = String(m);
          gAxis.append(t);
        }
        const hit = svgEl('rect', { x: L.padL, y: L.top - 6, width: L.plotW, height: L.plotBottom - L.top + 6, fill: 'transparent' });
        hit.addEventListener('pointermove', onChartPointer);
        hit.addEventListener('pointerdown', onChartPointer);
        hit.addEventListener('pointerleave', (e) => { if (e.pointerType !== 'touch') hideTip(); });

        chartSvg.setAttribute('viewBox', `0 0 ${L.w} ${L.h}`);
        chartSvg.append(defs, gBands, gGrid, colHot, gBars, gOver, gRib, gRibLabels, gAxis, hit);
        chart.colHot = colHot;
        positionHandlesVertically();
        updateChartBands();
        renderChartTable();
      }

      function bandGeometry(b) {
        const lo = Number.isFinite(b.min) ? clamp(b.min, 0, 101) : 0;
        const hi = Number.isFinite(b.max) ? clamp(b.max + 1, 0, 101) : lo;
        return { lo, width: Math.max(0, hi - lo) };
      }

      function updateChartBands() {
        const L = chart.L;
        if (!L) return;
        const valid = bandsValid();
        chartEl.classList.toggle('invalid', !valid);
        state.bands.forEach((b, i) => {
          const { lo, width } = bandGeometry(b);
          const x = L.x(lo);
          const w = width * L.colW;
          chart.bandBg[i].style.transform = `translate(${x}px, 0px) scaleX(${w})`;
          chart.ribs[i].style.transform = `translate(${x + 1}px, 0px) scaleX(${Math.max(0, w - 2)})`;
          const label = chart.ribLabels[i];
          label.style.transform = `translateX(${x + w / 2}px)`;
          label.style.opacity = w >= (G[i].length > 1 ? 36 : 28) ? '1' : '0';
        });
        chart.bars.forEach((p, m) => {
          if (!p) return;
          const g = valid ? C.gradeIndex(m, state.bands) : -1;
          p.setAttribute('class', `${p.classList.contains('grow') ? 'bar grow' : 'bar'}${g >= 0 ? ` b${g}` : ''}${m === chart.hotCol ? ' hot' : ''}`);
        });
        const errs = state.analysis.validation.errors;
        chart.cuts.forEach((line, i) => {
          const v = state.bands[i].min;
          line.style.display = Number.isFinite(v) ? '' : 'none';
          if (Number.isFinite(v)) line.style.transform = `translateX(${L.x(clamp(v, 0, 101))}px)`;
          line.classList.toggle('hot', !!(chart.drag && chart.drag.i === i));
        });
        chart.handles.forEach((hd, i) => {
          const v = state.bands[i].min;
          hd.hidden = !Number.isFinite(v);
          if (!Number.isFinite(v)) return;
          hd.style.transform = `translateX(${L.x(clamp(v, 0, 101))}px)`;
          const { lo, hi } = handleBounds(i);
          hd.setAttribute('aria-valuemin', lo);
          hd.setAttribute('aria-valuemax', hi);
          hd.setAttribute('aria-valuenow', v);
          hd.setAttribute('aria-valuetext', `${shown(G[i])} from ${v} marks`);
          hd.querySelector('.handle-tip').textContent = `${shown(G[i])} ≥ ${v}`;
          hd.classList.toggle('invalid', errs.some((e) => e.index === i || e.index === i + 1));
        });
        const a = state.analysis;
        const parts = valid ? G.map((g, i) => `${shown(g)} ${a.counts[i]}`).join(', ') : 'grade bands are invalid';
        $('chartDesc').textContent = `Marks distribution for ${course().name}: ${plural(a.stats.count, 'student')}, mean ${fmt(a.stats.mean, 1)}, median ${fmtMark(a.stats.median)}. Grades: ${parts}.`;
      }

      function renderChartTable() {
        const wrap = $('chartTableWrap');
        wrap.hidden = !state.tableOpen;
        const btn = $('tableToggle');
        btn.setAttribute('aria-expanded', String(state.tableOpen));
        btn.querySelector('span').textContent = state.tableOpen ? 'Hide table' : 'Show as table';
        if (!state.tableOpen || !chart.cols) return;
        const valid = bandsValid();
        const rows = [];
        for (let m = 100; m >= 0; m--) {
          const n = chart.cols[m].length;
          if (!n) continue;
          const g = valid ? C.gradeIndex(m, state.bands) : -1;
          rows.push(h('tr', null, h('td', { class: 'num', text: state.rounding === 'none' ? `${m}–${m}.9` : String(m) }), h('td', { class: 'num', text: String(n) }), h('td', null, gradeChip(g))));
        }
        $('chartTableBody').replaceChildren(...rows);
      }

      /* hover / tap readout */
      function onChartPointer(e) {
        const L = chart.L;
        if (!L || chart.drag) return;
        const rect = chartSvg.getBoundingClientRect();
        const px = e.clientX - rect.left;
        const m = Math.floor((px - L.padL) / L.colW);
        if (m < 0 || m > 100) { hideTip(); return; }
        showTip(m, px, e.clientY - rect.top);
      }

      function showTip(m, px, py) {
        const L = chart.L;
        const list = chart.cols[m];
        if (chart.hotCol !== m) {
          if (chart.hot) chart.hot.classList.remove('hot');
          chart.hotCol = m;
          chart.hot = chart.bars[m];
          if (chart.hot) chart.hot.classList.add('hot');
        }
        chartEl.classList.add('hovering');
        chart.colHot.setAttribute('x', L.x(m));
        chart.colHot.setAttribute('opacity', '1');
        const valid = bandsValid();
        const g = valid ? C.gradeIndex(m, state.bands) : -1;
        const ids = list.slice(0, 4).map((s) => s.id).join(', ') + (list.length > 4 ? ` +${list.length - 4} more` : '');
        chartTip.className = `chart-tip${g >= 0 ? ` b${g}` : ''}`;
        chartTip.replaceChildren(
          h('div', { class: 'tip-value', text: plural(list.length, 'student') }),
          h('div', { class: 'tip-row' }, g >= 0 ? h('span', { class: 'tip-key' }) : null, `Mark ${state.rounding === 'none' ? `${m}–${m}.9` : m}${g >= 0 ? ` · grade ${shown(G[g])}` : ''}`),
          list.length ? h('div', { class: 'tip-ids', text: ids }) : null);
        chartTip.hidden = false;
        const tw = chartTip.offsetWidth, th = chartTip.offsetHeight;
        let left = px + 14;
        if (left + tw > L.w) left = px - tw - 14;
        const top = clamp(py - th - 12, 0, L.h - th);
        chartTip.style.transform = `translate(${clamp(left, 0, L.w - tw)}px, ${top}px)`;
      }

      function hideTip() {
        chartTip.hidden = true;
        chartEl.classList.remove('hovering');
        if (chart.hot) chart.hot.classList.remove('hot');
        chart.hot = null;
        chart.hotCol = -1;
        if (chart.colHot) chart.colHot.setAttribute('opacity', '0');
      }

      /* draggable cutoff handles (role=slider) */
      function buildHandles() {
        chart.handles = G.slice(0, -1).map((g, i) => {
          const el = h('div', {
            class: 'handle', role: 'slider', tabindex: '0', 'aria-orientation': 'horizontal',
            'aria-label': `Lowest mark for grade ${shown(g)}`,
          }, h('span', { class: 'handle-tip', 'aria-hidden': 'true' }));
          el.addEventListener('pointerdown', (e) => startDrag(e, i));
          el.addEventListener('keydown', (e) => handleKey(e, i));
          handlesEl.append(el);
          return el;
        });
      }
      function positionHandlesVertically() {
        const L = chart.L;
        chart.handles.forEach((hd) => { hd.style.top = `${L.ribbonTop + L.ribbonH / 2}px`; });
      }
      function handleBounds(i) {
        const b = state.bands;
        let lo = i === b.length - 2 ? 1 : b[i + 1].min + 1;
        let hi = i === 0 ? 100 : b[i - 1].min - 1;
        if (!Number.isFinite(lo)) lo = 1;
        if (!Number.isFinite(hi)) hi = 100;
        lo = clamp(lo, 1, 100);
        hi = clamp(hi, 1, 100);
        return lo <= hi ? { lo, hi } : { lo: 1, hi: 100 };
      }
      function startDrag(e, i) {
        if (e.button !== 0 || !chart.L) return;
        e.preventDefault();
        const el = chart.handles[i];
        el.setPointerCapture(e.pointerId);
        el.focus({ preventScroll: true });
        hideTip();
        chart.drag = { i, id: e.pointerId, start: C.cloneBands(state.bands), raf: 0, lastX: e.clientX };
        el.classList.add('active');
        chartEl.classList.add('dragging');
        const move = (ev) => {
          if (ev.pointerId !== chart.drag.id) return;
          chart.drag.lastX = ev.clientX;
          if (chart.drag.raf) return;
          chart.drag.raf = requestAnimationFrame(() => {
            chart.drag.raf = 0;
            const L = chart.L;
            const rect = chartSvg.getBoundingClientRect();
            const { lo, hi } = handleBounds(i);
            const v = clamp(Math.round((chart.drag.lastX - rect.left - L.padL) / L.colW), lo, hi);
            if (v !== state.bands[i].min) {
              state.bands = C.setBandEdge(state.bands, i, 'min', v);
              afterBandsChange({ live: true });
            }
          });
        };
        const end = (ev) => {
          if (ev.pointerId !== chart.drag.id) return;
          cancelAnimationFrame(chart.drag.raf);
          el.removeEventListener('pointermove', move);
          el.removeEventListener('pointerup', end);
          el.removeEventListener('pointercancel', end);
          el.classList.remove('active');
          chartEl.classList.remove('dragging');
          const start = chart.drag.start;
          chart.drag = null;
          if (!C.bandsEqual(start, state.bands)) {
            const hist = state.history;
            hist.past.push(start);
            hist.future = [];
            hist.key = null;
            afterBandsChange({});
            announce(`${shown(G[i])} now starts at ${state.bands[i].min}.`);
          } else updateChartBands();
        };
        el.addEventListener('pointermove', move);
        el.addEventListener('pointerup', end);
        el.addEventListener('pointercancel', end);
      }
      function handleKey(e, i) {
        const cur = state.bands[i].min;
        if (!Number.isFinite(cur)) return;
        const { lo, hi } = handleBounds(i);
        const big = e.shiftKey ? 5 : 1;
        const map = { ArrowLeft: -big, ArrowDown: -big, ArrowRight: big, ArrowUp: big, PageDown: -5, PageUp: 5 };
        let v;
        if (e.key in map) v = cur + map[e.key];
        else if (e.key === 'Home') v = lo;
        else if (e.key === 'End') v = hi;
        else return;
        e.preventDefault();
        v = clamp(v, lo, hi);
        if (v === cur) return;
        applyBands(C.setBandEdge(state.bands, i, 'min', v), { coalesce: `handle${i}` });
      }

      /* ================================================================ bands */

      function buildBands() {
        const wrap = $('bands');
        wrap.replaceChildren(...G.map((g, i) => {
          const minId = `band-${i}-min`, maxId = `band-${i}-max`;
          const top = i === 0, bottom = i === G.length - 1;
          const input = (id, field, fixed) => {
            const el = h('input', {
              id, type: 'number', inputmode: 'numeric', min: '0', max: '100', step: '1',
              dataset: { i: String(i), field }, 'aria-describedby': `band-${i}-err`,
              readonly: fixed ? true : null, 'aria-readonly': fixed ? 'true' : null,
            });
            if (!fixed) {
              el.addEventListener('input', () => scheduleCommit(el));
              el.addEventListener('change', () => commitInput(el));
              el.addEventListener('keydown', (e) => { if (e.key === 'Enter') commitInput(el); });
            }
            return el;
          };
          return h('div', { class: `band b${i}`, style: `--i:${i}`, dataset: { i: String(i) } },
            h('div', { class: 'band-top' },
              h('span', { class: 'band-letter', 'aria-hidden': 'true', text: shown(g) }),
              h('span', { class: 'band-count', id: `band-${i}-count` })),
            h('div', { class: 'band-range', role: 'group', 'aria-label': `Grade ${shown(g)} range` },
              h('div', { class: 'band-field' },
                h('label', { for: minId }, 'Min', h('span', { class: 'sr-only', text: ` mark for grade ${shown(g)}${bottom ? ' (fixed at 0)' : ''}` })),
                input(minId, 'min', bottom)),
              h('span', { class: 'band-dash', 'aria-hidden': 'true', text: '–' }),
              h('div', { class: 'band-field' },
                h('label', { for: maxId }, 'Max', h('span', { class: 'sr-only', text: ` mark for grade ${shown(g)}${top ? ' (fixed at 100)' : ''}` })),
                input(maxId, 'max', top))),
            h('div', { class: 'band-share', 'aria-hidden': 'true' }, h('span')),
            h('p', { class: 'band-note', id: `band-${i}-note` }),
            h('p', { class: 'band-err', id: `band-${i}-err`, hidden: true }));
        }));
      }

      const commitTimers = new Map();
      function scheduleCommit(el) {
        clearTimeout(commitTimers.get(el));
        commitTimers.set(el, setTimeout(() => commitInput(el), 380));
      }
      function commitInput(el) {
        clearTimeout(commitTimers.get(el));
        const i = Number(el.dataset.i);
        const field = el.dataset.field;
        const raw = el.value.trim();
        const value = raw === '' || !Number.isFinite(Number(raw)) ? NaN : Number(raw);
        const cur = state.bands[i][field];
        if (value === cur || (Number.isNaN(value) && Number.isNaN(cur))) return;
        applyBands(C.setBandEdge(state.bands, i, field, value), { coalesce: `${i}${field}`, source: el });
      }

      const bandSig = (b) => b.map((x) => `${x.min}:${x.max}`).join('|');

      /** The single entry point for committed band changes (inputs, keyboard, reset, undo). */
      function applyBands(next, { coalesce = null, source = null, message = null } = {}) {
        if (bandSig(next) === bandSig(state.bands)) return;
        const hist = state.history;
        const now = Date.now();
        if (!(coalesce && hist.key === coalesce && now - hist.at < 1500)) {
          hist.past.push(C.cloneBands(state.bands));
          if (hist.past.length > 200) hist.past.shift();
        }
        hist.future = [];
        hist.key = coalesce;
        hist.at = now;
        state.bands = next;
        afterBandsChange({ source });
        if (message) toast(message, 'Undo', undo);
      }

      function undo() {
        const hist = state.history;
        if (!hist.past.length || !course()) return;
        hist.future.push(C.cloneBands(state.bands));
        state.bands = hist.past.pop();
        hist.key = null;
        afterBandsChange({});
        announce('Undid the last band change.');
      }
      function redo() {
        const hist = state.history;
        if (!hist.future.length || !course()) return;
        hist.past.push(C.cloneBands(state.bands));
        state.bands = hist.future.pop();
        hist.key = null;
        afterBandsChange({});
        announce('Redid the band change.');
      }

      function afterBandsChange({ source = null, live = false } = {}) {
        const c = course();
        if (!c) return;
        analyze();
        if (bandsValid()) {
          if (C.bandsEqual(state.bands, C.defaultBands())) store.remove(`bands:${c.key}`);
          else store.set(`bands:${c.key}`, C.cloneBands(state.bands));
        }
        resumeSession();
        updateChartBands();
        renderBands(source);
        renderMix();
        renderHero();
        renderSteps();
        if (live) return; // roster, table and export catch up when the drag ends
        renderRoster();
        renderChartTable();
        renderExport();
      }

      let prevCounts = null;
      let prevErrorBands = new Set();
      function renderBands(source) {
        const a = state.analysis;
        const v = a.validation;
        const total = a.students.length;
        const errorsByBand = new Map();
        v.errors.forEach((e) => { if (e.index !== null && !errorsByBand.has(e.index)) errorsByBand.set(e.index, e); });
        const active = document.activeElement;
        state.bands.forEach((b, i) => {
          const card = document.querySelector(`.band[data-i="${i}"]`);
          if (!card) return;
          ['min', 'max'].forEach((field) => {
            const input = $(`band-${i}-${field}`);
            const val = b[field];
            const text = Number.isFinite(val) ? String(val) : '';
            if (input !== active && input.value !== text) {
              input.value = text;
              if (source && motionOK()) { restartClass(input, 'flash'); setTimeout(() => input.classList.remove('flash'), 650); }
            }
            const bad = v.errors.some((e) => e.index === i && e.field === field);
            input.setAttribute('aria-invalid', String(bad));
          });
          const err = errorsByBand.get(i);
          const errEl = $(`band-${i}-err`);
          errEl.hidden = !err;
          errEl.textContent = err ? err.message : '';
          card.classList.toggle('has-error', !!err);
          if (err && !prevErrorBands.has(i)) restartClass(card, 'has-error');

          const countEl = $(`band-${i}-count`);
          const note = $(`band-${i}-note`);
          if (v.valid) {
            const n = a.counts[i];
            const pct = total ? Math.round((n / total) * 100) : 0;
            countEl.replaceChildren(String(n), h('small', { text: `${pct}%` }));
            countEl.setAttribute('aria-label', `${plural(n, 'student')}, ${pct} percent`);
            if (prevCounts && prevCounts[i] !== n && motionOK()) restartClass(countEl, 'bump');
            card.querySelector('.band-share span').style.transform = `scaleX(${total ? n / total : 0})`;
            const near = a.borderline[i];
            if (near) {
              note.replaceChildren(icon('arrow-up'), `${near} within ${C.BORDERLINE_WINDOW} of ${shown(G[i - 1])} · `,
                h('button', { type: 'button', text: 'show', onclick: () => showBorderline(i) }));
            } else note.textContent = n ? '' : 'No students';
          } else {
            countEl.replaceChildren('–');
            countEl.setAttribute('aria-label', 'Count unavailable until the bands are fixed');
            card.querySelector('.band-share span').style.transform = 'scaleX(0)';
            note.textContent = '';
          }
        });
        prevCounts = v.valid ? a.counts.slice() : null;
        prevErrorBands = new Set(errorsByBand.keys());

        const box = $('bandErrors');
        box.hidden = v.valid;
        if (!v.valid) $('bandErrorList').replaceChildren(...[...new Set(v.errors.map((e) => e.message))].map((m) => h('li', { text: m })));
        $('undoBtn').disabled = !state.history.past.length;
        $('redoBtn').disabled = !state.history.future.length;
        $('defaultsBtn').disabled = C.bandsEqual(state.bands, C.defaultBands());
      }

      function showBorderline(i) {
        state.roster.filter = `near${i}`;
        state.roster.page = 1;
        renderRoster();
        const panel = $('rosterPanel');
        panel.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'start' });
        $('rosterFilter').focus({ preventScroll: true });
      }

      /* ================================================================ roster */

      function gradeChip(g) {
        return g >= 0 ? h('span', { class: `gchip b${g}`, text: shown(G[g]) }) : h('span', { class: 'gchip none', text: '—', 'aria-label': 'No grade until the bands are fixed' });
      }

      function rosterRows() {
        const a = state.analysis;
        const r = state.roster;
        let rows = a.students;
        const q = r.query.trim().toUpperCase().replace(/\s+/g, '');
        if (q) rows = rows.filter((s) => s.idKey.includes(q));
        const f = r.filter;
        if (/^g\d$/.test(f)) rows = rows.filter((s) => s.grade === Number(f.slice(1)));
        else if (/^near\d$/.test(f)) rows = rows.filter((s) => s.borderline && s.grade === Number(f.slice(4)));
        else if (f === 'near') rows = rows.filter((s) => s.borderline);
        else if (f === 'decimal') rows = rows.filter((s) => s.decimal);
        const key = r.sort;
        const dir = r.dir;
        const val = (s) => (key === 'id' ? s.idKey : key === 'mark' ? s.mark : key === 'grade' ? (s.grade === null ? 99 : s.grade) : s.row);
        return [...rows].sort((x, y) => {
          const vx = val(x), vy = val(y);
          const c = typeof vx === 'string' ? vx.localeCompare(vy, 'en', { numeric: true }) : vx - vy;
          return (c || x.row - y.row) * dir;
        });
      }

      function renderRosterFilter() {
        const a = state.analysis;
        const sel = $('rosterFilter');
        const valid = a.validation.valid;
        const near = a.students.filter((s) => s.borderline).length;
        const opts = [['all', `All students (${a.students.length})`]];
        if (valid) G.forEach((g, i) => opts.push([`g${i}`, `Grade ${shown(g)} (${a.counts[i]})`]));
        if (valid) opts.push(['near', `Within ${C.BORDERLINE_WINDOW} marks of a higher grade (${near})`]);
        if (valid) G.forEach((g, i) => { if (i > 0 && /^near\d$/.test(state.roster.filter) && Number(state.roster.filter.slice(4)) === i) opts.push([`near${i}`, `${shown(g)} within ${C.BORDERLINE_WINDOW} of ${shown(G[i - 1])} (${a.borderline[i]})`]); });
        if (a.decimals) opts.push(['decimal', `Decimal marks (${a.decimals})`]);
        if (!opts.some(([v]) => v === state.roster.filter)) state.roster.filter = 'all';
        sel.replaceChildren(...opts.map(([v, t]) => h('option', { value: v, text: t })));
        sel.value = state.roster.filter;
      }

      function renderRoster(enter) {
        const a = state.analysis;
        if (!a || !a.students.length) return;
        renderRosterFilter();
        const rows = rosterRows();
        const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
        const r = state.roster;
        r.page = clamp(r.page, 1, pages);
        const slice = rows.slice((r.page - 1) * PAGE_SIZE, r.page * PAGE_SIZE);
        const bands = state.bands;
        const trs = slice.map((s, idx) => {
          const markCell = h('td', { class: 'num' });
          if (s.mark !== s.raw) markCell.append(h('span', { class: 'raw', text: `${fmtMark(s.raw)} →` }), h('b', { text: fmtMark(s.mark) }));
          else markCell.append(fmtMark(s.mark));
          const note = h('td', { class: 'col-note' });
          if (s.borderline) note.append(h('span', { class: 'note' }, icon('arrow-up'), `${fmtMark(s.borderline.gap)} below ${shown(bands[s.borderline.to].grade)}`));
          else if (s.decimal) note.append(h('span', { class: 'note', text: s.mark === s.raw ? 'decimal kept' : 'rounded' }));
          const tr = h('tr', { class: enter && motionOK() && idx < 14 ? 'enter' : null, style: `--i:${idx}` },
            h('td', { class: 'col-row', text: String(s.row) }),
            h('td', { class: 'mono', text: s.id }),
            markCell,
            h('td', null, gradeChip(s.grade === null ? -1 : s.grade)),
            note);
          return tr;
        });
        $('rosterBody').replaceChildren(...trs);
        $('rosterEmpty').hidden = rows.length > 0;
        $('pager').hidden = rows.length <= PAGE_SIZE;
        $('pageInfo').textContent = rows.length ? `${(r.page - 1) * PAGE_SIZE + 1}–${Math.min(rows.length, r.page * PAGE_SIZE)} of ${rows.length}` : '';
        $('pagePrev').disabled = r.page <= 1;
        $('pageNext').disabled = r.page >= pages;
        $('rosterSub').textContent = `${plural(a.students.length, 'student')} in file order. Search, filter or sort; the export always lists everyone in file order.`;
        document.querySelectorAll('.roster th[aria-sort]').forEach((th) => {
          const k = th.querySelector('.sort').dataset.sort;
          th.setAttribute('aria-sort', k === r.sort ? (r.dir === 1 ? 'ascending' : 'descending') : 'none');
        });
      }

      /* ================================================================ aside: instructor, timer, mix, export */

      function initials(name) {
        const words = name.replace(/\b(prof|dr|mr|mrs|ms|shri|smt)\.?\s+/gi, '').split(/\s+/).filter(Boolean);
        if (!words.length) return '';
        return (words[0][0] + (words.length > 1 ? words[words.length - 1][0] : '')).toUpperCase();
      }

      function renderInstructor() {
        const name = $('instructor').value.trim();
        const av = $('avatar');
        const ini = initials(name);
        const had = av.classList.contains('named');
        if (ini) {
          if (av.textContent !== ini) av.textContent = ini;
          if (!had) restartClass(av, 'named');
        } else if (had || !av.querySelector('svg')) {
          av.classList.remove('named');
          av.replaceChildren(icon('user'));
        }
      }

      let ticker = 0;
      function startTicker() {
        stopTicker();
        renderTimer();
        ticker = setInterval(renderTimer, 1000);
      }
      function stopTicker() { clearInterval(ticker); ticker = 0; }

      function resumeSession() {
        const s = session();
        if (s && s.frozenAt) {
          s.frozenAt = null;
          $('exportDone').hidden = true;
          renderTimer();
        }
      }

      function renderTimer() {
        const s = session();
        const c = course();
        const box = $('timer');
        if (!s || !c) {
          $('timerValue').textContent = '--:--';
          $('timerArc').style.strokeDashoffset = '125.66';
          $('timerLabel').textContent = 'Grading time';
          $('timerNote').textContent = 'Starts when you open a course.';
          box.classList.remove('done');
          return;
        }
        const elapsed = (s.frozenAt || Date.now()) - s.start;
        $('timerValue').textContent = C.formatClock(elapsed);
        $('timerArc').style.strokeDashoffset = String(125.66 * (1 - Math.min(elapsed / 3600000, 1)));
        box.classList.toggle('done', !!s.frozenAt);
        $('timerLabel').textContent = s.frozenAt ? 'Grading time · paused' : 'Grading time';
        $('timerNote').textContent = s.frozenAt
          ? 'Paused at export. It resumes if you change anything.'
          : `Counting since you opened ${c.name}.`;
      }

      function renderMix() {
        const a = state.analysis;
        const bar = $('mixBar');
        const legend = $('mixLegend');
        const ready = a && a.students.length;
        $('mixEmpty').hidden = !!(ready && a.validation.valid);
        $('mixEmpty').textContent = !ready ? 'Pick a course to see how grades spread.' : 'Fix the grade bands to see the grade mix.';
        bar.hidden = !ready;
        legend.hidden = !ready;
        if (!ready) return;
        const valid = a.validation.valid;
        if (bar.children.length !== G.length) bar.replaceChildren(...G.map((_, i) => h('span', { class: `mix-seg b${i}` })));
        [...bar.children].forEach((seg, i) => { seg.style.flexGrow = valid ? String(a.counts[i]) : '0'; });
        const total = a.students.length;
        legend.replaceChildren(...G.map((g, i) => h('li', { class: `b${i}` },
          h('span', { class: 'sw' }), h('b', { text: shown(g) }),
          h('small', { text: valid ? `${a.counts[i]} · ${Math.round((a.counts[i] / total) * 100)}%` : '–' }))));
      }

      function readiness() {
        const ds = state.dataset;
        const c = course();
        const a = state.analysis;
        const name = $('instructor').value.trim();
        const items = [];
        if (!ds) items.push({ key: 'course', ok: false, text: 'Upload a marks file', target: 'fileInput' });
        else if (!c) items.push({ key: 'course', ok: false, text: 'Open a course to grade', target: 'h-course' });
        else if (!c.students.length) items.push({ key: 'course', ok: false, html: [h('b', { text: c.name }), ' has no gradable students'] });
        else items.push({ key: 'course', ok: true, html: [h('b', { text: c.name }), ` · ${plural(c.students.length, 'student')}`] });
        const gradable = !!(a && a.students.length);
        if (gradable && a.validation.valid) items.push({ key: 'bands', ok: true, html: ['Grade bands are valid'] });
        else if (gradable) items.push({ key: 'bands', ok: false, text: 'Fix the grade bands', target: 'sec-bands' });
        else items.push({ key: 'bands', ok: false, html: ['Valid grade bands'] });
        items.push(name ? { key: 'name', ok: true, html: ['Signed by ', h('b', { text: name })] }
          : { key: 'name', ok: false, text: 'Add the instructor name', target: 'instructor' });
        if (c && c.students.length && c.excluded) items.push({ key: 'ack', ok: state.ack, ack: true, excluded: c.excluded });
        return { items, ok: items.every((i) => i.ok) };
      }

      function exportCsv(instructorName) {
        const c = course();
        return C.buildGradeCsv({
          instructor: instructorName, courseName: c.name, bands: state.bands, analysis: state.analysis,
          policy: state.rounding, generatedAt: new Date(), isSample: state.dataset.isSample,
        });
      }

      function renderExport() {
        const r = readiness();
        const list = $('checklist');
        list.replaceChildren(...r.items.map((it) => {
          if (it.ack) {
            const box = h('input', { type: 'checkbox', id: 'ackBox', checked: state.ack ? true : null, onchange: (e) => { state.ack = e.target.checked; renderExport(); } });
            return h('li', { class: `check ${it.ok ? 'ok' : 'todo'}`, dataset: { key: it.key } },
              h('span', { class: 'check-dot' }, icon(it.ok ? 'check' : 'alert')),
              h('span', { class: 'check-text' },
                h('label', { class: 'ack', for: 'ackBox' }, box, h('span', null, `Export without the ${plural(it.excluded, 'excluded row')}. `)),
                h('button', { type: 'button', text: 'Review them', onclick: openIssues })));
          }
          const text = h('span', { class: 'check-text' });
          if (it.text) text.append(h('button', { type: 'button', text: it.text, onclick: () => focusTarget(it.target) }));
          else text.append(...it.html);
          return h('li', { class: `check ${it.ok ? 'ok' : 'todo'}`, dataset: { key: it.key } },
            h('span', { class: 'check-dot' }, icon(it.ok ? 'check' : 'alert')), text);
        }));

        const c = course();
        const a = state.analysis;
        const canPreview = !!(c && a && a.students.length && a.validation.valid);
        $('exportFile').hidden = !canPreview;
        if (canPreview) {
          const name = $('instructor').value.trim() || '‹instructor name›';
          $('efName').textContent = C.exportFileName(c.name, new Date(), state.dataset.isSample);
          const lines = exportCsv(name).trimEnd().split('\r\n');
          const head = lines.indexOf('BITS ID,Total Marks,Grade');
          const shownLines = lines.slice(0, head + 6);
          const more = lines.length - shownLines.length;
          $('efPreview').textContent = shownLines.join('\n') + (more > 0 ? `\n… ${plural(more, 'more row')}` : '');
        }
        const btn = $('exportBtn');
        btn.setAttribute('aria-disabled', String(!r.ok));
        $('copyBtn').setAttribute('aria-disabled', String(!r.ok));
        $('exportWhy').textContent = r.ok ? 'Opens in Excel or Google Sheets. Students are listed in file order.' : 'Complete the items above to export.';
      }

      function focusTarget(id) {
        const el = $(id);
        if (!el) return;
        el.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'center' });
        el.focus({ preventScroll: true });
        if (id === 'instructor') {
          el.setAttribute('aria-invalid', 'true');
          const help = $('instructorHelp');
          help.textContent = 'Enter your name to sign the grade sheet.';
          help.classList.add('bad');
        }
      }

      function openIssues() {
        state.issuesOpen = true;
        state.issueScope = 'course';
        renderFileCard(false);
        const panel = $('issuesPanel');
        panel.scrollIntoView({ behavior: motionOK() ? 'smooth' : 'auto', block: 'start' });
        $('issuesToggle').focus({ preventScroll: true });
      }

      function guardExport() {
        const r = readiness();
        if (r.ok) return true;
        const first = r.items.find((i) => !i.ok);
        const li = $('checklist').querySelector(`[data-key="${first.key}"]`);
        if (li && motionOK()) { li.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-4px)' }, { transform: 'translateX(4px)' }, { transform: 'translateX(0)' }], { duration: 260 }); }
        if (first.ack) $('ackBox').focus();
        else if (first.target) focusTarget(first.target);
        announce('Export needs attention: ' + (first.ack ? 'confirm the excluded rows.' : li ? li.textContent : ''));
        return false;
      }

      function downloadText(text, name) {
        const blob = new Blob(['﻿', text], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = h('a', { href: url, download: name, rel: 'noopener', style: 'display:none' });
        document.body.append(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
      }

      function doExport() {
        if (!guardExport()) return;
        const c = course();
        let csv;
        try {
          csv = exportCsv($('instructor').value.trim());
        } catch (err) {
          toast(`Export stopped: ${err.message}`);
          return;
        }
        const name = C.exportFileName(c.name, new Date(), state.dataset.isSample);
        downloadText(csv, name);
        const s = session();
        s.exports += 1;
        s.frozenAt = Date.now();
        const took = C.formatDuration(s.frozenAt - s.start);
        const n = state.analysis.students.length;
        const done = $('exportDone');
        done.textContent = s.exports === 1
          ? `Downloaded ${name} with ${plural(n, 'grade')}. You completed grading in ${took} on your first export.`
          : `Downloaded ${name} with ${plural(n, 'grade')}. Total time ${took}, your ${C.ordinal(s.exports)} export of ${c.name}.`;
        done.hidden = false;
        restartClass(done, 'export-done');
        const btn = $('exportBtn');
        btn.classList.add('done');
        $('exportBtnText').textContent = 'Downloaded';
        setTimeout(() => { btn.classList.remove('done'); $('exportBtnText').textContent = 'Download CSV'; }, 2200);
        renderTimer();
        renderSteps();
        renderHero();
        announce(done.textContent);
      }

      async function copyCsv() {
        if (!guardExport()) return;
        let csv;
        try { csv = exportCsv($('instructor').value.trim()); } catch (err) { toast(`Copy stopped: ${err.message}`); return; }
        let ok = false;
        try { await navigator.clipboard.writeText(csv); ok = true; } catch (_) {
          const ta = h('textarea', { readonly: true, style: 'position:fixed;opacity:0;pointer-events:none' });
          ta.value = csv;
          document.body.append(ta);
          ta.select();
          try { ok = document.execCommand('copy'); } catch (__) { ok = false; }
          ta.remove();
        }
        toast(ok ? `Copied ${plural(state.analysis.students.length, 'row')} as CSV. Paste into any spreadsheet.` : 'Copy was blocked by the browser. Use Download CSV instead.');
      }

      function renderAside() {
        renderInstructor();
        renderTimer();
        renderMix();
        renderExport();
      }

      function renderAll() {
        renderFileCard(false);
        renderCourses();
        renderWorkspace();
        renderAside();
        renderHero();
        renderSteps();
      }

      /* ================================================================ events */

      function wire() {
        // steps, theme, guide
        document.querySelectorAll('.step').forEach((b) => b.addEventListener('click', () => goToStep(b)));
        themeBtn.addEventListener('click', toggleTheme);
        systemDark.addEventListener('change', renderThemeButton);
        const guide = $('guide');
        const openGuide = () => { if (typeof guide.showModal === 'function') guide.showModal(); else guide.setAttribute('open', ''); };
        $('guideBtn').addEventListener('click', openGuide);
        document.querySelectorAll('[data-open-guide]').forEach((b) => b.addEventListener('click', openGuide));
        $('guideClose').addEventListener('click', () => guide.close());
        guide.addEventListener('click', (e) => { if (e.target === guide) guide.close(); });

        // upload
        fileInput.addEventListener('change', () => {
          const files = fileInput.files;
          if (files && files.length > 1) toast('Only the first file was used. Upload one marks sheet at a time.');
          if (files && files[0]) handleFile(files[0]);
        });
        $('sampleBtn').addEventListener('click', loadSample);
        $('templateBtn').addEventListener('click', downloadSampleSheet);
        $('removeBtn').addEventListener('click', removeDataset);
        $('uploadErrorClose').addEventListener('click', hideUploadError);
        let dragDepth = 0;
        const hasFiles = (e) => e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files');
        const dropTarget = () => (state.dataset ? fileCard : dropzone);
        window.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; dropTarget().classList.add('drag'); });
        window.addEventListener('dragover', (e) => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
        window.addEventListener('dragleave', (e) => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) { dropzone.classList.remove('drag'); fileCard.classList.remove('drag'); } });
        window.addEventListener('drop', (e) => {
          if (!hasFiles(e)) return;
          e.preventDefault();
          dragDepth = 0;
          dropzone.classList.remove('drag');
          fileCard.classList.remove('drag');
          const files = e.dataTransfer.files;
          if (files.length > 1) toast('Only the first file was used. Upload one marks sheet at a time.');
          if (files[0]) handleFile(files[0]);
        });

        // data checks
        $('issuesToggle').addEventListener('click', () => { state.issuesOpen = !state.issuesOpen; renderFileCard(false); });
        document.querySelectorAll('[data-issue-scope]').forEach((b) => b.addEventListener('click', () => { state.issueScope = b.dataset.issueScope; renderIssues(); }));
        $('issuesDownload').addEventListener('click', downloadIssues);

        // courses
        $('courseSearch').addEventListener('input', renderCourses);

        // rounding radio group
        const radios = [...document.querySelectorAll('[data-rounding]')];
        radios.forEach((b, i) => {
          b.addEventListener('click', () => setRounding(b.dataset.rounding));
          b.addEventListener('keydown', (e) => {
            const d = e.key === 'ArrowRight' || e.key === 'ArrowDown' ? 1 : e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 0;
            if (!d) return;
            e.preventDefault();
            const next = radios[(i + d + radios.length) % radios.length];
            setRounding(next.dataset.rounding);
            next.focus();
          });
        });

        // chart
        $('curveToggle').addEventListener('change', (e) => { state.curve = e.target.checked; store.set('curve', state.curve); buildChart('curve'); });
        $('tableToggle').addEventListener('click', () => { state.tableOpen = !state.tableOpen; renderChartTable(); });
        buildHandles();
        let resizeRaf = 0;
        let lastWidth = 0;
        new ResizeObserver(() => {
          if (resizeRaf) return;
          resizeRaf = requestAnimationFrame(() => {
            resizeRaf = 0;
            const w = chartEl.clientWidth;
            if (w && w !== lastWidth && state.analysis && state.analysis.students.length && !chart.drag) { lastWidth = w; buildChart(false); }
          });
        }).observe(chartEl);
        chartEl.addEventListener('keydown', (e) => { if (e.key === 'Escape') hideTip(); });
        document.addEventListener('pointerdown', (e) => { if (!chartTip.hidden && !chartEl.contains(e.target)) hideTip(); });

        // bands
        $('undoBtn').addEventListener('click', undo);
        $('redoBtn').addEventListener('click', redo);
        $('defaultsBtn').addEventListener('click', () => applyBands(C.defaultBands(), { message: 'Restored the default bands.' }));
        document.addEventListener('keydown', (e) => {
          if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
          const t = e.target;
          if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
          const k = e.key.toLowerCase();
          if (k === 'z' && !e.shiftKey) { if (state.history.past.length) { e.preventDefault(); undo(); } }
          else if ((k === 'z' && e.shiftKey) || k === 'y') { if (state.history.future.length) { e.preventDefault(); redo(); } }
        });

        // roster
        let searchTimer = 0;
        $('rosterSearch').addEventListener('input', (e) => {
          clearTimeout(searchTimer);
          searchTimer = setTimeout(() => { state.roster.query = e.target.value; state.roster.page = 1; renderRoster(); }, 120);
        });
        $('rosterFilter').addEventListener('change', (e) => { state.roster.filter = e.target.value; state.roster.page = 1; renderRoster(); });
        document.querySelectorAll('.roster .sort').forEach((b) => b.addEventListener('click', () => {
          const r = state.roster;
          const k = b.dataset.sort;
          if (r.sort === k) r.dir = -r.dir; else { r.sort = k; r.dir = k === 'mark' ? -1 : 1; }
          r.page = 1;
          renderRoster();
        }));
        $('pagePrev').addEventListener('click', () => { state.roster.page -= 1; renderRoster(); });
        $('pageNext').addEventListener('click', () => { state.roster.page += 1; renderRoster(); });

        // instructor
        const inst = $('instructor');
        let saveTimer = 0;
        inst.addEventListener('input', () => {
          inst.removeAttribute('aria-invalid');
          const help = $('instructorHelp');
          help.classList.remove('bad');
          help.textContent = 'Printed at the top of the exported grade sheet. Remembered on this device.';
          renderInstructor();
          renderHero();
          renderExport();
          clearTimeout(saveTimer);
          saveTimer = setTimeout(() => { const v = inst.value.trim(); if (v) store.set('instructor', v); else store.remove('instructor'); }, 400);
        });

        // export
        $('exportBtn').addEventListener('click', doExport);
        $('copyBtn').addEventListener('click', copyCsv);

        document.addEventListener('visibilitychange', () => { if (!document.hidden && ticker) renderTimer(); });
      }

      /* ================================================================ start */

      function init() {
        const theme = store.get('theme', null);
        if (theme === 'dark' || theme === 'light') root.dataset.theme = theme;
        renderThemeButton();
        const name = store.get('instructor', '');
        if (typeof name === 'string') $('instructor').value = name.slice(0, 80);
        wire();
        renderAll();
        const idle = window.requestIdleCallback || ((fn) => setTimeout(fn, 600));
        idle(() => { loadSheetJS().catch(() => { /* reported when the user uploads */ }); });
      }

      init();
    })();
