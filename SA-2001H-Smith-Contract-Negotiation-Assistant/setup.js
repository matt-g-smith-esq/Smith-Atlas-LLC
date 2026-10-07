// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.
/**
 * setup.js — One-time configuration stamper.
 *
 * Reads NODE_SERVER_URL from the backend .env file and writes the server location
 * into the manifest, which physically requires a literal URL:
 *   1. Contract-Frontend/manifest.xml  (generated from manifest.template.xml)
 *
 * (The client no longer needs stamping: taskpane.js derives its backend origin at
 * runtime from self.location.origin — same-origin — so there is no BACKEND_URL literal
 * to write. Only the manifest, which Word reads before any page loads, needs a literal URL.)
 *
 * .env is the single source of truth. Run this once after setting NODE_SERVER_URL,
 * and again any time that value changes:
 *
 *     npm run setup
 *
 * This runs identically on Windows and macOS (it is plain Node, no shell specifics).
 *
 * Office add-ins require HTTPS for any non-localhost server, so this script refuses
 * to stamp a non-localhost http:// URL and explains why, catching the mistake here
 * rather than as a confusing failure inside Word later.
 */

const fs = require('fs');
const path = require('path');

// Load NODE_SERVER_URL from the backend .env (this script lives in Contract-Backend).
require('dotenv').config();

const BACKEND_DIR = __dirname;
const FRONTEND_DIR = path.join(BACKEND_DIR, '..', 'Contract-Frontend');
const TEMPLATE_PATH = path.join(FRONTEND_DIR, 'manifest.template.xml');
const MANIFEST_OUT = path.join(FRONTEND_DIR, 'manifest.xml');

function fail(message) {
    console.error('\n[setup] ERROR: ' + message + '\n');
    process.exit(1);
}

function isLocalhost(urlObj) {
    const h = urlObj.hostname;
    return h === 'localhost' || h === '127.0.0.1' || h === '[::1]' || h === '::1';
}

function main() {
    const raw = process.env.NODE_SERVER_URL;

    if (!raw || raw.trim() === '' || raw.includes('YOUR_') || raw.includes('<')) {
        fail(
            'NODE_SERVER_URL is not set in your .env file.\n' +
            '       Set it to the address where the Node server will run, for example:\n' +
            '         NODE_SERVER_URL=https://contracts.yourcompany.com\n' +
            '       For local testing on this machine you may use:\n' +
            '         NODE_SERVER_URL=http://localhost:3001'
        );
    }

    const url = raw.trim().replace(/\/+$/, ''); // strip any trailing slash

    let parsed;
    try {
        parsed = new URL(url);
    } catch (e) {
        fail('NODE_SERVER_URL is not a valid URL: "' + url + '"');
    }

    // HTTPS enforcement: Word requires HTTPS for any non-localhost add-in host.
    if (parsed.protocol !== 'https:' && !isLocalhost(parsed)) {
        fail(
            'NODE_SERVER_URL must use https:// for a non-localhost server.\n' +
            '       You provided: ' + url + '\n' +
            '       Microsoft Word will refuse to load an add-in from a non-secure (http://)\n' +
            '       address unless it is localhost. Obtain a valid TLS certificate for your\n' +
            '       server and use an https:// URL (e.g. https://contracts.yourcompany.com).'
        );
    }

    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
        fail('NODE_SERVER_URL must start with http:// or https:// (got "' + parsed.protocol + '").');
    }

    // --- Stamp the manifest from the template ---
    if (!fs.existsSync(TEMPLATE_PATH)) {
        fail('Could not find manifest.template.xml at:\n       ' + TEMPLATE_PATH);
    }
    const template = fs.readFileSync(TEMPLATE_PATH, 'utf8');
    const manifest = template.split('{{NODE_SERVER_URL}}').join(url);
    if (manifest.includes('{{NODE_SERVER_URL}}')) {
        fail('Template still contains an unreplaced placeholder after stamping.');
    }
    fs.writeFileSync(MANIFEST_OUT, manifest, 'utf8');

    const localNote = isLocalhost(parsed)
        ? ' (localhost — fine for local testing; use an https:// address for real deployments)'
        : '';

    console.log('\n[setup] Configuration stamped successfully.');
    console.log('[setup]   Server URL: ' + url + localNote);
    console.log('[setup]   Wrote: ' + MANIFEST_OUT);
    console.log('[setup] Next: deploy/sideload the generated manifest.xml (see the user guide).\n');
}

main();
