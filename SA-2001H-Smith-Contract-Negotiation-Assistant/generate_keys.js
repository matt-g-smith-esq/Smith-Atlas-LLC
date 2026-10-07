// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.

const crypto = require('crypto');
const fs = require('fs');

const NUMBER_OF_KEYS = 20000;

const PLAIN_FILE = 'plain_keys_inventory.csv';
const HASHED_FILE = 'hashed_keys.json';

// --- Overwrite protection -------------------------------------------------
// This script OVERWRITES both output files. Re-running it after keys have been
// sold would replace hashed_keys.json with a brand-new set of hashes and
// invalidate every key already in the field. To prevent an accidental,
// unrecoverable overwrite, refuse to run if either output file already exists.
//
// Regenerating is a deliberate act (a new product "generation"). If you truly
// intend it: move or delete the existing files first, then run this script.
const existing = [PLAIN_FILE, HASHED_FILE].filter(f => fs.existsSync(f));
if (existing.length > 0) {
    console.error('\nREFUSING TO RUN: the following file(s) already exist:');
    existing.forEach(f => console.error('   - ' + f));
    console.error(
        '\nRe-running would OVERWRITE these and invalidate every key already sold.\n' +
        'If you really intend to generate a new key pool, move or delete the\n' +
        'existing file(s) to a safe backup location first, then run this again.\n'
    );
    process.exit(1);
}
// -------------------------------------------------------------------------

const plainKeys = [];
const hashedKeys = [];

console.log('Generating Smith-Atlas Security Token Arrays...');

for (let i = 0; i < NUMBER_OF_KEYS; i++) {
    const key = 'SA-' + crypto.randomBytes(2).toString('hex').toUpperCase() +
                '-' + crypto.randomBytes(2).toString('hex').toUpperCase() +
                '-' + crypto.randomBytes(2).toString('hex').toUpperCase();
    plainKeys.push(key);
    const hash = crypto.createHash('sha256').update(key).digest('hex');
    hashedKeys.push(hash);
}

fs.writeFileSync(PLAIN_FILE, plainKeys.join('\n'));
fs.writeFileSync(HASHED_FILE, JSON.stringify(hashedKeys));
console.log(`Successfully completed generation of ${NUMBER_OF_KEYS} product license keys!`);
