/**
 * sensitive.js
 * One scanner, shared by the docs checker, the commit-msg hook and CI.
 *
 * THE BOUNDARY THIS ENFORCES
 *
 *   open_accounting is public. It holds the structure and the tools: schema, code,
 *   checks, and documentation about how the system works.
 *   wildlife-ai-management and Google Drive hold the content: actual budgets, funder
 *   terms, salary bands, and decisions about specific grants.
 *
 *   The line, stated so it can be checked: a file here may NAME a funding source, and
 *   may never STATE AN AMOUNT. Names like WW_25_TOI are unavoidable, they are Xero
 *   tracking values the code joins on, and funders announce their grants anyway. Amounts
 *   are the part that identifies what someone is paid and what a funder agreed.
 *
 * HOW IT DECIDES
 *
 *   Precision. An invented figure is round; a derived one is not. $48,000 reads as an
 *   illustration; a figure carrying an unrounded remainder reads as something copied out of
 *   a spreadsheet. So any money figure that is not a round hundred is suspect until rounded
 *   or explicitly marked INVENTED-OK.
 *
 *   This docblock originally made that point with a worked pair, one round and one not. The
 *   scanner flagged its own source, correctly: the unrounded half was realistic enough to
 *   read as real, and close enough to a figure being scrubbed from this repo's history at
 *   the time to serve as a hint. An example is not an exemption.
 *
 *   It deliberately holds no list of the real numbers it hunts for. A denylist of actual
 *   salaries, in a public file, would leak exactly what it guards.
 */
'use strict';

// Placeholders that are fine in a public repo.
const ALLOWED_EMAIL =
  /^(someone|admin|a|b|you|name|user|test|noreply|no-reply)@|@example\.(com|org|net)$|@users\.noreply\.github\.com$/;

/**
 * @param {string} text     content to scan
 * @param {string} where    label used in findings, e.g. a path or "commit message"
 * @returns {Array<{kind, where, detail}>}
 */
function scanSensitive(text, where) {
  const out = [];
  String(text).split('\n').forEach((line, i) => {
    if (/INVENTED-OK/.test(line)) return;
    const at = where + (where.indexOf('.') !== -1 ? ':' + (i + 1) : '');

    // Money precise enough to have come from a real sheet.
    const money = line.match(/\$\s?([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,7})(?:\.[0-9]{2})?/g) || [];
    money.forEach(m => {
      const n = parseInt(m.replace(/[$,\s]/g, ''), 10);
      if (n >= 1000 && n % 100 !== 0) {
        out.push({ kind: 'amount', where: at, detail: m.trim() });
      }
    });

    (line.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || []).forEach(e => {
      if (!ALLOWED_EMAIL.test(e)) out.push({ kind: 'email', where: at, detail: e });
    });

    if (/\b\d{2}-\d{3,4}-\d{6,7}-\d{2,3}\b/.test(line)) {
      out.push({ kind: 'bank account', where: at, detail: 'account-shaped number' });
    }

    if (/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.test(line)) {
      out.push({ kind: 'uuid', where: at, detail: 'Xero source id or similar' });
    }
  });
  return out;
}

/** Render findings for a terminal, with the fix spelled out. */
function formatSensitive(findings) {
  if (!findings.length) return '';
  const byKind = {};
  findings.forEach(f => (byKind[f.kind] = byKind[f.kind] || []).push(f));
  let s = '';
  Object.keys(byKind).forEach(k => {
    s += '\n  ' + k + ':\n';
    byKind[k].slice(0, 12).forEach(f => { s += '    ' + f.where + '  ' + f.detail + '\n'; });
    if (byKind[k].length > 12) s += '    ... and ' + (byKind[k].length - 12) + ' more\n';
  });
  s += '\nThis repository is public and holds structure, not content. Amounts, funder\n' +
       'terms and salary bands belong in wildlife-ai-management or Drive.\n' +
       'Round the figure, or append INVENTED-OK to the line if it is genuinely made up.\n';
  return s;
}

module.exports = { scanSensitive, formatSensitive, ALLOWED_EMAIL };
