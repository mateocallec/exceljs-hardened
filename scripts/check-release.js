'use strict';
// Refuses to publish unless the working tree is clean and HEAD is tagged
// v<package.json version>, so that npm's `gitHead` and the tag always
// identify the exact published source.
const {execSync} = require('child_process');
const {version} = require('../package.json');

const git = cmd => execSync(`git ${cmd}`, {encoding: 'utf8'}).trim();
const errors = [];

if (git('status --porcelain')) {
  errors.push('working tree is not clean (commit or stash your changes)');
}
const tag = `v${version}`;
let tagged = '';
try {
  tagged = git(`rev-list -n 1 ${tag}`);
} catch (e) {
  errors.push(`tag ${tag} does not exist`);
}
if (tagged && tagged !== git('rev-parse HEAD')) {
  errors.push(`tag ${tag} does not point at HEAD`);
}

if (errors.length) {
  console.error(`Release check failed:\n - ${errors.join('\n - ')}`);
  process.exit(1);
}
console.log(`Release check OK: ${tag} @ ${git('rev-parse --short HEAD')}`);
