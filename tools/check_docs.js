/**
 * check_docs.js
 * Fails when the documentation disagrees with the code.
 *
 * Run with Node from the repo root: `node dashboard/check_docs.js`
 *
 * Docs in this repo have drifted repeatedly and silently: GM_GUIDE said "Two tabs" when
 * there were four, dashboard/README's file list was missing two files, AGENTS.md listed a
 * deployment that had been deleted, and HEALTH_CHECKS.md documented 43 checks of which 22
 * existed. None of that is anyone's fault; prose has no compiler. This is the compiler.
 *
 * It runs in Node rather than Apps Script because it needs the filesystem, which is also
 * why it cannot live in Tests.js.
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const read = p => fs.readFileSync(path.join(ROOT, p), 'utf8');

let failures = 0;
function fail(what, detail) {
  failures++;
  console.log('  FAIL  ' + what + '\n        ' + detail);
}
function ok(what) { console.log('  ok    ' + what); }

// ---------------------------------------------------------------- health checks
// Every id in HEALTH_CATALOGUE must appear in the implemented table, and nothing may
// claim to be implemented that is not. Planned checks live under their own heading and
// are matched separately, so a backlog cannot masquerade as behaviour.
(function healthChecks() {
  const code = read('dashboard/HealthCheck.js');
  const doc = read('dashboard/HEALTH_CHECKS.md');

  const implemented = new Set();
  const re = /^\s{2}([A-Z]\d+):\s*\{\s*severity:\s*'(\w+)'/gm;
  let m;
  while ((m = re.exec(code))) implemented.add(m[1] + '/' + m[2]);

  // The doc's implemented tables are everything before the "Not yet implemented" heading.
  const split = doc.indexOf('## Not yet implemented');
  if (split === -1) {
    fail('HEALTH_CHECKS.md has no "Not yet implemented" heading',
      'Planned checks are indistinguishable from real ones without it.');
    return;
  }
  const documented = new Set();
  const rowRe = /^\|\s*([A-Z]\d+)\s*\|\s*(\w+)\s*\|/gm;
  let r;
  const head = doc.slice(0, split);
  while ((r = rowRe.exec(head))) documented.add(r[1] + '/' + r[2]);

  const missing = [...implemented].filter(x => !documented.has(x));
  const extra = [...documented].filter(x => !implemented.has(x));
  if (missing.length) {
    fail('checks implemented but not documented as such',
      missing.sort().join(', ') + ' — add a row, with the severity the code uses.');
  }
  if (extra.length) {
    fail('checks documented as implemented but absent from HEALTH_CATALOGUE',
      extra.sort().join(', ') + ' — move to "Not yet implemented", or fix the severity.');
  }
  if (!missing.length && !extra.length) {
    ok(implemented.size + ' health checks: ids and severities match HEALTH_CHECKS.md');
  }
})();

// ---------------------------------------------------------------- dashboard file list
(function fileList() {
  const doc = read('dashboard/README.md');
  const actual = fs.readdirSync(path.join(ROOT, 'dashboard'))
    .filter(f => /\.(js|html)$/.test(f) && f !== 'check_docs.js');
  const undocumented = actual.filter(f => doc.indexOf('`' + f + '`') === -1)
    // The three UI files are documented as one grouped row.
    .filter(f => !/^(Index|Stylesheet|JavaScript)\.html$/.test(f) ||
                 doc.indexOf('Index/Stylesheet/JavaScript.html') === -1);
  if (undocumented.length) {
    fail('dashboard files absent from the README file list', undocumented.join(', '));
  } else {
    ok('every dashboard file appears in dashboard/README.md');
  }
})();

// ---------------------------------------------------------------- UI tab count
(function tabs() {
  const index = read('dashboard/Index.html');
  const guide = read('dashboard/GM_GUIDE.md');
  const n = (index.match(/data-view="/g) || []).length;
  const words = { 2: 'Two', 3: 'Three', 4: 'Four', 5: 'Five' };
  if (guide.indexOf('## ' + words[n] + ' tabs') === -1) {
    fail('GM_GUIDE.md does not describe the right number of tabs',
      'Index.html has ' + n + ' tabs, so the heading should read "## ' + words[n] + ' tabs".');
  } else {
    ok('GM_GUIDE.md agrees with Index.html on ' + n + ' tabs');
  }
})();

// ---------------------------------------------------------------- Funding_info keys
// Any key the code reads must be documented, or nobody knows to fill it in.
(function metaKeys() {
  const cfg = read('dashboard/Config.js');
  const doc = read('dashboard/BUDGET_SHEET_TEMPLATE.md');
  const block = cfg.slice(cfg.indexOf('META: {'), cfg.indexOf('},', cfg.indexOf('META: {')));
  const keys = [...block.matchAll(/'([^']+)'/g)].map(m => m[1]);
  const missing = keys.filter(k => {
    const title = k.replace(/\b\w/g, c => c.toUpperCase());   // "exclusivity group" -> "Exclusivity Group"
    return doc.toLowerCase().indexOf('`' + k.toLowerCase() + '`') === -1 &&
           doc.indexOf('`' + title + '`') === -1;
  });
  if (missing.length) {
    fail('Funding_info keys read by code but undocumented', missing.join(', '));
  } else {
    ok(keys.length + ' Funding_info keys read by Config.META are all documented');
  }
})();

// ---------------------------------------------------------------- dead references
// Files the docs point at must exist. This is what caught create_quarterly_budgets.js
// being referenced five times after it was deleted.
(function deadRefs() {
  const docs = ['README.md', 'AGENTS.md', '.agents/skills/SKILL.md',
    'dashboard/README.md', 'dashboard/GM_GUIDE.md', 'dashboard/HEALTH_CHECKS.md',
    'dashboard/BUDGET_SHEET_TEMPLATE.md', 'dashboard/BUDGET_PROCEDURES_ADDENDUM.md',
    'budget_templates/README.md', 'project_reports/README.md',
    'funding_reports/xero-quickstart.md'];
  // Files deleted on purpose, which the docs still name because the lesson outlived the
  // code. SKILL.md's "two files both defined getFolderByName" is still worth knowing even
  // though one of them is gone. Listed explicitly so a genuinely broken reference to a file
  // someone expected to exist still fails.
  const RETIRED_FILES = new Set([
    'create_quarterly_budgets.js',   // retired 2026-08-11, quarterly generation moved into the cockpit
    'loader-budgets-template.js',    // remote code loader, removed 2026-08-10
    'loader_template.js'             // remote code loader, removed 2026-08-10
  ]);

  // Every file in the repo, by basename, so a bare mention in prose resolves.
  const allFiles = new Set();
  (function walk(dir) {
    fs.readdirSync(path.join(ROOT, dir), { withFileTypes: true }).forEach(e => {
      if (e.name === '.git' || e.name === 'node_modules') return;
      const rel = path.join(dir, e.name);
      if (e.isDirectory()) walk(rel); else allFiles.add(e.name);
    });
  })('.');

  const dead = [];
  docs.forEach(d => {
    const text = read(d);
    // Any backticked token that looks like a repo file path.
    [...text.matchAll(/`([A-Za-z0-9_./-]+\.(?:js|html|json|md|xlsx|csv))`/g)].forEach(m => {
      const ref = m[1];
      // URL fragments, and the one legitimate shorthand for the three grouped UI files.
      if (/^https?:|refs\/heads\//.test(ref)) return;
      if (ref === 'Index/Stylesheet/JavaScript.html') return;
      // A bare basename in prose is fine if that file exists anywhere in the repo.
      if (allFiles.has(path.basename(ref))) return;
      if (RETIRED_FILES.has(path.basename(ref))) return;
      dead.push(d + ' -> ' + ref);
    });
  });
  if (dead.length) {
    fail(dead.length + ' reference(s) to files that do not exist',
      [...new Set(dead)].join('\n        '));
  } else {
    ok('every file referenced by the docs exists');
  }
})();

// ---------------------------------------------------------------- sensitive content
// This is a public repository. The rule is that every figure in documentation and test
// fixtures is invented, and the giveaway for a real one is precision: nobody invents
// $66,710. So any money amount that is not a round hundred is treated as suspect until
// it is either rounded or explicitly marked.
//
// Deliberately does not list the real numbers it is looking for. A denylist of actual
// salary figures in a public file would leak the very thing it guards.
(function sensitive() {
  const scan = ['README.md', 'AGENTS.md', '.agents/skills/SKILL.md',
    'dashboard/README.md', 'dashboard/GM_GUIDE.md', 'dashboard/HEALTH_CHECKS.md',
    'dashboard/BUDGET_SHEET_TEMPLATE.md', 'dashboard/BUDGET_PROCEDURES_ADDENDUM.md',
    'budget_templates/README.md', 'dashboard/Tests.js'];

  const ALLOWED_EMAILS = /^(someone|admin|a|b|you|name)@|@example\.(com|org)$/;
  const oddMoney = [], personal = [], accounts = [];

  scan.forEach(f => {
    read(f).split('\n').forEach((line, i) => {
      if (/INVENTED-OK/.test(line)) return;   // explicit escape hatch, use sparingly
      const where = f + ':' + (i + 1);

      // $12,345 or $12345, and bare 5-6 digit numbers in fixtures.
      [...line.matchAll(/\$\s?([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,7})(?:\.[0-9]{2})?/g)]
        .forEach(m => {
          const n = parseInt(m[1].replace(/,/g, ''), 10);
          if (n >= 1000 && n % 100 !== 0) oddMoney.push(where + '  $' + m[1]);
        });

      [...line.matchAll(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g)].forEach(m => {
        if (!ALLOWED_EMAILS.test(m[0])) personal.push(where + '  ' + m[0]);
      });

      // NZ bank account, and anything shaped like one.
      if (/\b\d{2}-\d{3,4}-\d{6,7}-\d{2,3}\b/.test(line)) accounts.push(where);
    });
  });

  if (accounts.length) fail('bank-account-shaped strings', accounts.join(', '));
  else ok('no bank-account-shaped strings');

  if (personal.length) {
    fail(personal.length + ' personal email address(es) in a public repo',
      [...new Set(personal)].join('\n        ') +
      '\n        Use someone@wildlife.ai, or add to ALLOWED_EMAILS if it is a fixture.');
  } else ok('no personal email addresses outside the allowed placeholders');

  if (oddMoney.length) {
    fail(oddMoney.length + ' money figure(s) precise enough to look real',
      [...new Set(oddMoney)].join('\n        ') +
      '\n        Round them, or append INVENTED-OK to the line if genuinely made up.');
  } else ok('every money figure in docs and fixtures is a round hundred');
})();

console.log();
if (failures) {
  console.log(failures + ' documentation check(s) failed');
  process.exitCode = 1;
} else {
  console.log('documentation agrees with the code');
}
