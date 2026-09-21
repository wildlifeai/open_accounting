/**
 * commit_msg_check.js
 * Called by .githooks/commit-msg with the path to the message being written.
 *
 * Uses the same scanner as the docs checker and CI, so one rule is enforced in all three
 * places rather than three rules drifting apart.
 */
'use strict';
const fs = require('fs');
const { scanSensitive, formatSensitive } = require('./sensitive');

const file = process.argv[2];
if (!file) { console.error('commit_msg_check: no message file given'); process.exit(1); }

// Comment lines are stripped by git before the message is stored, so a figure quoted in
// the diff summary git shows you is not actually being committed.
const msg = fs.readFileSync(file, 'utf8')
  .split('\n').filter(l => l.indexOf('#') !== 0).join('\n');

const findings = scanSensitive(msg, 'commit message');
if (!findings.length) process.exit(0);

console.error('\nCommit blocked: the message looks like it contains content, not structure.');
console.error(formatSensitive(findings));
console.error('Rewrite the message. If the figure is genuinely invented, put INVENTED-OK');
console.error('on that line. To bypass deliberately: git commit --no-verify\n');
process.exit(1);
