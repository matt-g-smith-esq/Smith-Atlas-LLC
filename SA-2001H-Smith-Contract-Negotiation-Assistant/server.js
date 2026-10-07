file:
// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). All rights reserved.
// Load server-side environment variables securely before any routing or logic executes
require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mammoth = require('mammoth');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

// --- --generate-manifest MODE -------------------------------------------------
// In the compiled (packaged) editions there is no loose setup.js for an admin to run,
// so the binary itself stamps the Word manifest. Running the executable with the
// --generate-manifest flag reads NODE_SERVER_URL from .env, stamps it into manifest.xml
// (from manifest.template.xml), then exits WITHOUT starting the web server. This mirrors
// setup.js, including the HTTPS-for-non-localhost rule.
//   Example:  ./smith-atlas-server --generate-manifest
// This runs BEFORE the license file and audit database are loaded below, so stamping a
// manifest never requires a valid license or a writable database — it only needs .env
// and the manifest template.
if (process.argv.includes('--generate-manifest')) {
    try {
        const FRONTEND_DIR = path.join(__dirname, '..', 'Contract-Frontend');
        const TEMPLATE_PATH = path.join(FRONTEND_DIR, 'manifest.template.xml');
        const MANIFEST_OUT = path.join(FRONTEND_DIR, 'manifest.xml');

        const isLocalhost = (u) => {
            const h = u.hostname;
            return h === 'localhost' || h === '127.0.0.1' || h === '[::1]' || h === '::1';
        };
        const die = (msg) => { console.error('\n[generate-manifest] ERROR: ' + msg + '\n'); process.exit(1); };

        const rawUrl = process.env.NODE_SERVER_URL;
        if (!rawUrl || rawUrl.trim() === '' || rawUrl.includes('YOUR_') || rawUrl.includes('<')) {
            die('NODE_SERVER_URL is not set in your .env file.\n' +
                '       Set it to where the server runs, e.g. https://contracts.yourcompany.com\n' +
                '       (or http://localhost:3001 for local testing), then run again.');
        }
        const stampUrl = rawUrl.trim().replace(/\/+$/, '');
        let parsed;
        try { parsed = new URL(stampUrl); }
        catch (e) { die('NODE_SERVER_URL is not a valid URL: "' + stampUrl + '"'); }

        if (parsed.protocol !== 'https:' && !isLocalhost(parsed)) {
            die('NODE_SERVER_URL must use https:// for a non-localhost server.\n' +
                '       You provided: ' + stampUrl + '\n' +
                '       Word refuses to load an add-in from a non-secure (http://) address\n' +
                '       unless it is localhost. Use an https:// URL with a valid certificate.');
        }
        if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
            die('NODE_SERVER_URL must start with http:// or https:// (got "' + parsed.protocol + '").');
        }
        if (!fs.existsSync(TEMPLATE_PATH)) {
            die('Could not find manifest.template.xml at:\n       ' + TEMPLATE_PATH);
        }
        const template = fs.readFileSync(TEMPLATE_PATH, 'utf8');
        const stamped = template.split('{{NODE_SERVER_URL}}').join(stampUrl);
        if (stamped.includes('{{NODE_SERVER_URL}}')) {
            die('Template still contains an unreplaced placeholder after stamping.');
        }
        fs.writeFileSync(MANIFEST_OUT, stamped, 'utf8');

        const localNote = isLocalhost(parsed)
            ? ' (localhost — fine for local testing; use an https:// address for real deployments)'
            : '';
        console.log('\n[generate-manifest] Manifest stamped successfully.');
        console.log('[generate-manifest]   Server URL: ' + stampUrl + localNote);
        console.log('[generate-manifest]   Wrote: ' + MANIFEST_OUT);
        console.log('[generate-manifest] Next: deploy/sideload the generated manifest.xml (see the user guide).\n');
        process.exit(0);
    } catch (err) {
        console.error('\n[generate-manifest] ERROR: ' + err.message + '\n');
        process.exit(1);
    }
}
// Node 24+ ships a built-in SQLite module (no native dependency, no compile step).
// It is marked experimental, which prints a process warning on load; we suppress
// that single warning so the server console stays clean. Everything else is unaffected.
const _origEmitWarning = process.emitWarning;
process.emitWarning = (warning, ...args) => {
    const msg = typeof warning === 'string' ? warning : (warning && warning.message) || '';
    if (msg.includes('SQLite is an experimental feature')) return;
    return _origEmitWarning.call(process, warning, ...args);
};
const { DatabaseSync } = require('node:sqlite');
const OllamaAdapter = require('./adapters/OllamaAdapter');
const CloudAdapter = require('./adapters/CloudAdapter');

// --- DIAGNOSTIC VERBOSITY GATE ------------------------------------------------
// Detailed per-run diagnostics ([DIAG-*], [DIAGNOSTIC] dumps) are invaluable while
// debugging but noisy in a shipped product. They are gated behind a single .env flag,
// DIAG_VERBOSE, which defaults to OFF so a clean customer install has a quiet console.
// To debug a problem in the field, set DIAG_VERBOSE=true in .env and restart — no code
// change or redeploy needed — then set it back to false (or remove it) when done.
// Genuine error lines (console.error) are intentionally NOT gated: a failed run should
// still surface that it failed, even with diagnostics off.
const DIAG_VERBOSE = String(process.env.DIAG_VERBOSE).toLowerCase() === 'true';
function diagLog(...args) { if (DIAG_VERBOSE) console.log(...args); }
function diagWarn(...args) { if (DIAG_VERBOSE) console.warn(...args); }

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, '../Contract-Frontend')));
const validHashes = JSON.parse(fs.readFileSync(path.join(__dirname, 'hashed_keys.json'), 'utf8'));
// --- AUDIT LEDGER (SQLite) ----------------------------------------------------
// The compliance audit trail is stored in a single-file SQLite database. Compared to
// a flat CSV, this gives atomic writes (no interleaving under concurrent users) and
// real queryability. The masked-field design (one-way SHA-256 of salt + value) is
// unchanged: user emails and filenames are never stored in plaintext.
// NOTE (tracked follow-up): for true tamper-evidence (proving no row was altered or
// deleted), a future change can hash-chain each row to the previous one. This baseline
// provides atomic, queryable logging; it does not by itself prove immutability.
const DB_FILE = path.join(__dirname, 'activity_log.db');
const auditDb = new DatabaseSync(DB_FILE);
// WAL (write-ahead logging) lets readers and a writer proceed concurrently and greatly
// reduces "database is locked" errors when multiple users are redlining at once on a
// shared company server. On a single-user local machine this is harmless and invisible.
// It is a persistent property of the database file, so setting it once is sufficient.
try {
    auditDb.exec(`PRAGMA journal_mode = WAL;`);
    auditDb.exec(`PRAGMA busy_timeout = 5000;`);
} catch (e) {
    console.warn('[AUDIT-DB] Could not enable WAL mode (continuing in default mode): ' + e.message);
}
auditDb.exec(`
    CREATE TABLE IF NOT EXISTS audit_log (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp       TEXT NOT NULL,
        masked_user     TEXT,
        client_ip       TEXT,
        masked_filename TEXT,
        playbook_name   TEXT,
        execution_track TEXT,
        active_model    TEXT,
        applied_revisions INTEGER,
        execution_status TEXT,
        review_mode     TEXT,
        prev_hash       TEXT,
        row_hash        TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp);
    CREATE INDEX IF NOT EXISTS idx_audit_status ON audit_log(execution_status);
`);
// Backward-compatible migrations: add any columns that an audit_log table created by an
// earlier version is missing. SQLite has no "ADD COLUMN IF NOT EXISTS", so we probe the
// table info and only ALTER for columns that are absent. review_mode predates the
// tamper-evidence feature; prev_hash/row_hash are the hash-chain columns.
try {
    const cols = auditDb.prepare(`PRAGMA table_info(audit_log)`).all();
    const have = new Set(cols.map(c => c.name));
    if (!have.has('review_mode')) auditDb.exec(`ALTER TABLE audit_log ADD COLUMN review_mode TEXT`);
    if (!have.has('prev_hash'))   auditDb.exec(`ALTER TABLE audit_log ADD COLUMN prev_hash TEXT`);
    if (!have.has('row_hash'))    auditDb.exec(`ALTER TABLE audit_log ADD COLUMN row_hash TEXT`);
} catch (migErr) {
    console.error('Audit schema migration check failed (non-fatal):', migErr.message);
}
// Genesis anchor for the chain. The first ever row links to this constant instead of a
// previous row's hash, so the chain has a well-defined start.
const AUDIT_CHAIN_GENESIS = 'SMITH_ATLAS_AUDIT_CHAIN_GENESIS_v1';
// Canonical serialization of a row's auditable fields, used for BOTH writing and verifying
// so the two can never disagree. Order is fixed and must never change for existing data.
function serializeAuditRow(r) {
    return [
        r.timestamp, r.masked_user, r.client_ip, r.masked_filename, r.playbook_name,
        r.execution_track, r.active_model, String(r.applied_revisions), r.execution_status,
        r.review_mode, r.prev_hash
    ].join('\u241F'); // unit-separator delimiter, unlikely to occur in any field
}
// Computes a row's chain hash from its fields (which already include prev_hash) keyed by the
// audit salt. Without the salt an attacker cannot forge a valid hash even by rewriting rows.
function computeAuditRowHash(r, salt) {
    return crypto.createHash('sha256').update(salt + '\u241F' + serializeAuditRow(r)).digest('hex');
}
const auditInsertStmt = auditDb.prepare(`
    INSERT INTO audit_log
        (timestamp, masked_user, client_ip, masked_filename, playbook_name,
         execution_track, active_model, applied_revisions, execution_status, review_mode,
         prev_hash, row_hash)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);
// Fetches the most recent row's hash (the link the next row chains onto). Returns null
// when the table is empty (the caller substitutes the genesis anchor).
const auditLastHashStmt = auditDb.prepare(`SELECT row_hash FROM audit_log ORDER BY id DESC LIMIT 1`);
/**
 * Formats an elapsed time (in seconds) for user-facing status messages.
 * Under one minute it reads as plain seconds ("45 sec."); at one minute and
 * beyond it switches to minutes + seconds ("1 min. 15 sec.", "2 min."). This keeps
 * long waits readable at a glance rather than showing "75s", "90s", etc.
 * To use a compact clock style instead (e.g. "1:15"), swap the return lines noted below.
 */
function formatElapsed(totalSeconds) {
    let s = Math.max(0, Math.floor(totalSeconds));
    if (s < 60) {
        return `${s} sec.`;
    }
    let m = Math.floor(s / 60);
    let rem = s % 60;
    // --- Spelled-out style (current): "1 min. 15 sec." / "2 min." ---
    if (rem === 0) return `${m} min.`;
    return `${m} min. ${rem} sec.`;
    // --- Compact clock style (alternative): replace the two lines above with:
    //   return `${m}:${String(rem).padStart(2, '0')}`;
}
/**
 * Retrieves setting details, defaulting to local architecture paths.
 */
function getServerConfig() {
    // Server-wide configuration comes from the environment (read-only at runtime).
    // The user's only per-run choice is the provider (cloud vs. local); model, URL, and
    // paths are all server-owned here.
    let defaultProvider = (process.env.DEFAULT_PROVIDER || 'cloud').toLowerCase().trim();
    if (defaultProvider === 'local') defaultProvider = 'ollama';
    if (defaultProvider !== 'cloud' && defaultProvider !== 'ollama') defaultProvider = 'cloud';
    return {
        aiServerUrl: process.env.LOCAL_AI_SERVER_URL || 'http://localhost:11434',
        playbooksFolderPath: process.env.LOCAL_PLAYBOOKS_PATH || '',
        defaultModelName: process.env.LOCAL_DEFAULT_MODEL || 'llama4:maverick',
        defaultProvider: defaultProvider
    };
}
/**
 * Polymorphic adapter selection factory.
 */
function createAdapter(settings) {
    const provider = (settings.provider || 'ollama').toLowerCase();
    switch (provider) {
        case 'ollama':
            return new OllamaAdapter(settings.aiServerUrl, settings.modelName);
        case 'cloud':
            return new CloudAdapter();
        default:
            console.warn(`[ROUTING ALERT] Unknown provider requested: "${provider}". Defaulting to local Ollama node.`);
            return new OllamaAdapter(settings.aiServerUrl, settings.modelName);
    }
}
/**
 * Writes a cryptographically masked transaction row to the SQLite audit ledger.
 * Masking is unchanged: emails and filenames are stored only as one-way SHA-256
 * digests (salt + value), never in plaintext. The comma-stripping of free-text
 * fields is no longer required for delimiter safety (SQLite handles values
 * natively) but is retained so historical exports remain consistent.
 */
function writeAuditLog(userEmail, clientIp, contractFileName, playbook, executionTrack, activeModel, revisionCount, status, reviewMode) {
    try {
        const timestamp = new Date().toISOString();
        const salt = process.env.AUDIT_LOG_SALT || "[placeholder]";
        // Compute secure one-way hashes to protect MNPI and PII from plain-text disclosure
        const maskedUser = userEmail
            ? crypto.createHash('sha256').update(salt + userEmail.trim().toLowerCase()).digest('hex').substring(0, 16)
            : "ANONYMOUS_USER";
        const maskedFilename = contractFileName && contractFileName !== "UNSAVED_DOCUMENT"
            ? crypto.createHash('sha256').update(salt + contractFileName.trim().toLowerCase()).digest('hex').substring(0, 16)
            : "UNSAVED_DOCUMENT";
        const cleanPlaybook = playbook ? String(playbook) : "NONE";
        const cleanModel = activeModel ? String(activeModel) : "UNKNOWN";
        const cleanIP = clientIp ? String(clientIp) : "127.0.0.1";
        const cleanReviewMode = reviewMode ? String(reviewMode) : "unknown";
        const cleanRevisions = Number(revisionCount) || 0;
        // TAMPER-EVIDENCE (hash chain): each row's hash is computed over its own fields PLUS the
        // previous row's hash. Editing or deleting any historical row changes its row_hash, which
        // breaks the prev_hash linkage of every row after it — detectable by verifyAuditChain().
        // The first row links to a fixed genesis anchor. Keyed by the salt so the chain can't be
        // silently recomputed by anyone who doesn't know AUDIT_LOG_SALT.
        let lastRow = null;
        try { lastRow = auditLastHashStmt.get(); } catch (e) { lastRow = null; }
        const prevHash = (lastRow && lastRow.row_hash) ? lastRow.row_hash : AUDIT_CHAIN_GENESIS;
        const rowForHash = {
            timestamp, masked_user: maskedUser, client_ip: cleanIP, masked_filename: maskedFilename,
            playbook_name: cleanPlaybook, execution_track: executionTrack, active_model: cleanModel,
            applied_revisions: cleanRevisions, execution_status: status, review_mode: cleanReviewMode,
            prev_hash: prevHash
        };
        const rowHash = computeAuditRowHash(rowForHash, salt);
        auditInsertStmt.run(
            timestamp,
            maskedUser,
            cleanIP,
            maskedFilename,
            cleanPlaybook,
            executionTrack,
            cleanModel,
            cleanRevisions,
            status,
            cleanReviewMode,
            prevHash,
            rowHash
        );
    } catch (err) {
        console.error('Logging Error:', err);
    }
}
/**
 * Executes a pipeline call protected by an exponential backoff retry loop with jitter.
 * Safe backend processing ceiling set to 15 minutes (900,000 ms) to support frontend interactive checkpoint.
 */
// --- CONCURRENCY GATE ---------------------------------------------------------
// At 20+ simultaneous users, all redline requests funnel into the LLM, which can
// only handle so many at once. This semaphore caps how many model calls run
// concurrently; additional calls wait in line rather than overwhelming the engine
// (and the existing retry/backoff loop covers the queued wait). Tune MAX_CONCURRENT
// to your inference capacity (a single Ollama process ~1-2; a batching server like
// vLLM can be higher).
const MAX_CONCURRENT_MODEL_CALLS = parseInt(process.env.MAX_CONCURRENT_MODEL_CALLS || '2', 10);
let activeModelCalls = 0;
const modelCallQueue = [];
function acquireModelSlot() {
    if (activeModelCalls < MAX_CONCURRENT_MODEL_CALLS) {
        activeModelCalls++;
        return Promise.resolve();
    }
    return new Promise(resolve => modelCallQueue.push(resolve));
}
function releaseModelSlot() {
    const next = modelCallQueue.shift();
    if (next) {
        // Hand the slot directly to the next waiter (count stays the same).
        next();
    } else {
        activeModelCalls = Math.max(0, activeModelCalls - 1);
    }
}

async function executeWithRetry(adapter, systemPrompt, contractText) {
    // Wait for a free model slot before doing any work, and always release it.
    await acquireModelSlot();
    try {
    let attempt = 0;
    let resultText;
    let success = false;
    const startTime = Date.now();
    const MAX_PRIMARY_TIME = 900000; // 15-minute absolute backend safety ceiling
    const calculateBackoff = (n) => (2000 * Math.pow(2, n - 1)) + Math.floor(Math.random() * 1000);
    while (Date.now() - startTime < MAX_PRIMARY_TIME) {
        try {
            attempt++;
            resultText = await adapter.generateResponse(systemPrompt, contractText);
            success = true;
            break;
        } catch (apiError) {
            // Transient conditions worth retrying with backoff:
            //   - 503 / 429: server busy or rate-limited (existing behavior).
            //   - low-level "fetch" failures: dropped/refused connections (existing).
            //   - 502 / 504 / 524: GATEWAY TIMEOUTS. When a reverse proxy (Cloudflare
            //     in front of RunPod, or any enterprise proxy a customer puts ahead of
            //     their endpoint) gives up waiting for a slow origin, it returns one of
            //     these with its own HTML error page. The most common cause is a model
            //     cold-start (weights loading into VRAM on first call). Retrying lets the
            //     now-warm model answer on a subsequent attempt instead of failing hard.
            //     NOTE: this cannot defeat a request that ALWAYS exceeds the proxy's
            //     per-request ceiling — only a cold-start or momentary overrun self-heals.
            const msg = (apiError.message || '').toLowerCase();
            const isRetryable = apiError.status === 503 ||
                                 apiError.status === 429 ||
                                 apiError.status === 502 ||
                                 apiError.status === 504 ||
                                 apiError.status === 524 ||
                                 msg.includes('503') ||
                                 msg.includes('429') ||
                                 msg.includes('502') ||
                                 msg.includes('504') ||
                                 msg.includes('524') ||
                                 msg.includes('gateway') ||
                                 msg.includes('timeout') ||
                                 msg.includes('fetch');
            const elapsed = Date.now() - startTime;
            if (isRetryable && elapsed < MAX_PRIMARY_TIME) {
                let wait = calculateBackoff(attempt);
                if (elapsed + wait > MAX_PRIMARY_TIME) wait = MAX_PRIMARY_TIME - elapsed;
                // Describe the ACTUAL condition rather than the old cloud-API "traffic"
                // wording. For a local model behind a proxy, a retryable failure is almost
                // always a gateway timeout (the proxy giving up while the model loads into
                // VRAM or generates a long response) or a dropped connection — not rate
                // limiting. We surface the underlying error and elapsed time so the log
                // explains itself. A 524 here usually means the model was still cold-loading.
                let reason = apiError.status === 524 || msg.includes('524') ? 'gateway timeout (524) — model likely still loading/generating'
                          : apiError.status === 502 || msg.includes('502') ? 'bad gateway (502)'
                          : apiError.status === 504 || msg.includes('504') ? 'gateway timeout (504)'
                          : apiError.status === 503 || msg.includes('503') ? 'service unavailable (503)'
                          : apiError.status === 429 || msg.includes('429') ? 'rate limited (429)'
                          : msg.includes('fetch') ? 'connection/fetch failure'
                          : 'transient error';
                let elapsedSec = Math.round(elapsed / 1000);
                console.warn(`Retryable ${reason}. Waiting ${wait}ms before attempt ${attempt + 1} (${elapsedSec}s elapsed of ${Math.round(MAX_PRIMARY_TIME/1000)}s budget)...`);
                await new Promise(resolve => setTimeout(resolve, wait));
            } else if (elapsed >= MAX_PRIMARY_TIME) {
                break;
            } else {
                console.error(`Fatal model processing exception:`, apiError.message);
                throw apiError;
            }
        }
    }
    if (!success) {
        throw new Error("AI engine pipeline processing failed to complete within the server resilience threshold.");
    }
    return resultText;
    } finally {
        releaseModelSlot();
    }
}
// --- PORTED COMPLIANCE CONSTRAINTS AND ANTI-COLLATERAL-DAMAGE RULES ---
const UNIVERSAL_WRAPPER = `
### CRITICAL EDITING CONSTRAINTS (THE HUMAN REVIEWER STANDARD)
==========================================================
HOW TO READ THE PLAYBOOK'S RULE TAGS ([Mandatory] / [Optional])
==========================================================
Each rule in the contracting playbook is tagged [Mandatory] or [Optional]. These tags govern WHETHER you must act:
- **[Mandatory]:** This rule MUST be implemented in the contract. If the contract is silent or missing the required language, you MUST ADD it (an insertion). If the contract addresses the topic but non-compliantly, you MUST FIX it (a surgical edit). Do not skip a [Mandatory] rule merely because the contract is silent on it — silence on a [Mandatory] rule is itself the problem to fix. Work through EVERY [Mandatory] rule in the playbook and ensure each one is satisfied in your output.
- **[Optional]:** It is acceptable for the contract to be silent on this rule. Only act if the contract ALREADY contains a clause implicating the rule and that clause is non-compliant — in which case edit it to comply. If the contract is silent on an [Optional] rule, do nothing.
==========================================================
You are redlining like a senior attorney. There are exactly TWO kinds of edits, and you MUST label each revision with an "edit_type" of either "surgical" or "insertion".
==========================================================
THE TARGET_TEXT TEST — DECIDE edit_type BY THIS MECHANICAL TEST, NOT BY INTUITION
==========================================================
Before you choose "surgical", apply this test: Can you copy a SPECIFIC, NON-EMPTY, VERBATIM substring out of the contract and place it in "target_text" — the exact existing words you are striking out or replacing?
- If YES (there are specific existing words to strike or replace) → it is "surgical".
- If NO (your "target_text" would be empty "", OR you are only ADDING new words without removing any specific existing words) → it is NOT surgical. It is an "insertion". Put your new language in "body_text" and leave "target_text" as "".
THERE ARE NO EXCEPTIONS. A "surgical" edit with an empty "target_text" is INVALID and will be discarded — the change will be lost. This is the single most common error: emitting "surgical" for additive language (a new proviso, a new cap, a new sentence, a fallback requirement) that does not replace specific existing words. Adding a "provided that ..." proviso, a new cap sentence, or a new clause is ALWAYS an "insertion", even when it attaches to an existing section. If in doubt: no words to strike ⇒ insertion.
==========================================================
TYPE 1 — "surgical": The concept ALREADY EXISTS in the contract but the language is wrong, insufficient, or non-compliant.
==========================================================
A human attorney fixes existing language by striking the FEWEST possible words and adding the FEWEST possible words. They NEVER delete a run of words and retype those same words back.
RULES FOR SURGICAL EDITS:
1. **Minimal span (Anti-Collateral-Damage Rule):** Your "target_text" must be the SMALLEST span that fully contains the change AND still reads as a COHERENT unit — not the shortest possible fragment, and not the whole surrounding paragraph. Target only the words that are actually changing, plus the minimum adjacent context needed to make the string unique and locatable. Do NOT include compliant surrounding words just because they are adjacent. Two failure modes to avoid in equal measure:
   (a) **OVER-GRABBING (most common):** Do NOT select an entire multi-sentence paragraph as the target when only specific words within it are changing. Swallowing a whole clause produces an unreadable redline (a wall of struck text replaced by a wall of new text) and forces the reviewer to diff it themselves. It also frequently exceeds the editor's locatable length and the edit cannot be applied in place.
   (b) **OVER-FRAGMENTING:** Do NOT chop one coherent change into tiny disconnected scraps (e.g. striking "worldwide, and" on its own) that are not grammatically or legally self-contained. Each target must be a coherent phrase or clause a human attorney would naturally strike as one unit.
   WORKED EXAMPLE — a license-grant sentence reads: "...Contractor shall grant to Customer a perpetual, nonexclusive, worldwide, and royalty-free license to use, reproduce, modify, distribute, and create derivative works of Contractor's Foreground IP for its and its Affiliates' (defined below) own internal research, development, educational, administrative, and non-commercial purposes." The company wants a non-exclusive license for internal COMMERCIAL use. WRONG (over-grab): target_text = the entire ~90-word sentence. RIGHT (two coherent, minimal surgical edits): (1) target_text = "a perpetual, nonexclusive, worldwide, and royalty-free license" → redlined_text = "a non-exclusive, royalty-free, and non-sublicensable (except to Affiliates) license"; and (2) target_text = "internal research, development, educational, administrative, and non-commercial purposes" → redlined_text = "internal commercial purposes". This leaves the rest of the sentence standing in place and shows the reviewer exactly what changed.
   WHEN A LONG TARGET IS CORRECT: If you are STRIKING an entire clause outright, or genuinely REWRITING a whole clause wholesale (where nearly every word changes), then the full clause IS the minimal coherent span — use it. Do not artificially shorten a true full-clause strike or rewrite. Length itself is not the enemy; grabbing more than the change requires is.
2. **Never echo unchanged words:** "target_text" and "redlined_text" must NOT share long identical leading or trailing runs. If your replacement starts or ends with the exact same words as the target, you have selected too large a span — tighten it.
3. **Never identical:** target_text and redlined_text must differ. Do not emit a surgical edit that changes nothing.
3a. **Pure deletions (strikes) are allowed:** When a rule says to STRIKE existing language entirely (e.g., strike a "Work Made for Hire" phrase, strike a liquidated-damages clause, strike a most-favored-nation clause), emit a surgical edit whose "target_text" is the exact existing words to remove and whose "redlined_text" is an empty string "". An empty "redlined_text" is ONLY valid for a surgical edit that has a real, non-empty "target_text" (the words being struck). Do NOT use an empty "redlined_text" on an insertion.
4. **target_text must be a literal, character-for-character, case-exact substring** lifted directly from the contract. Do not normalize capitalization, smart/straight quotes, hyphens, or spacing. If you alter even one character, the edit will fail validation.
5. **context_anchor** is a longer, completely UNCHANGED passage (2-3 sentences or a full paragraph) surrounding the target, used only to locate the right occurrence. It must be copied verbatim and must NOT overlap-edit anything.
6. **Punctuation alignment:** If target_text ends with a period/semicolon/colon, redlined_text must end with the same punctuation.
7. **Defined terms:** Match the contract's existing capitalized defined terms exactly (e.g., if the document says "Customer", do not switch to "Client").
8. **One concept per surgical edit — NEVER span a paragraph boundary (STRICT):** A surgical "target_text" MUST come entirely from ONE single paragraph. NEVER let a target_text run across a paragraph break, a heading, or a section number into the next paragraph. The editor locates text within a single paragraph, so a target that begins in one paragraph and ends in another CANNOT be found and the entire edit is silently lost. This is one of the most damaging errors you can make. WORKED EXAMPLE OF THE ERROR (do not repeat it): a contract has Paragraph 1 ending "...Each Party owns their Foreground IP." and a SEPARATE Paragraph 2 beginning "License Grant; Reservation of Rights. Upon Customer's written request...". Combining an ownership fix and a license-grant fix into ONE target_text that runs "...Each Party owns their Foreground IP. License Grant; Reservation of Rights. Upon Customer's written request..." spans the paragraph break and WILL FAIL. The CORRECT approach is TWO separate revisions: (a) one surgical edit whose target_text is only "Each Party owns their Foreground IP." (from Paragraph 1), and (b) a second surgical edit whose target_text is the license-grant wording from Paragraph 2. Whenever fixing a concept requires changing text in two different paragraphs, ALWAYS emit one revision per paragraph, each with its own minimal, single-paragraph target.
9. **Do not bundle unrelated fixes into one edit:** Each revision addresses ONE rule. Do not, for example, correct a notice email address AND append a new certified-mail requirement in the same surgical edit — fix the email as one surgical edit and add the certified-mail requirement as its own separate edit (surgical on the delivery-methods sentence, or an insertion). Bundling places new language in the wrong location. CONCRETE RULE FOR CONTACT/SIGNATURE VALUES: when correcting a short atomic value that sits in a contact block, signature block, or fill-in line (an email address, a name, a title, a phone number, an address), the "redlined_text" MUST contain ONLY the corrected value and nothing else. Never append a sentence, requirement, or clause to an atomic value fix — emit that requirement as a SEPARATE revision targeting the clause it actually belongs to (for the certified-mail fallback, target the "delivered by (a) certified mail..." sentence in the NOTICES delivery clause, not the contact row).
==========================================================
TYPE 2 — "insertion": The concept is ENTIRELY ABSENT and a mandatory clause must be ADDED.
==========================================================
When a required concept does not exist anywhere in the contract, add it as a brand-new clause. It will render as a pure tracked-changes ADDITION (no strikethrough), exactly as an attorney inserting a missing provision.
RULES FOR INSERTION EDITS:
1. Set "edit_type": "insertion".
2. Leave "target_text" as an empty string "".
3. "context_anchor" is a verbatim, unchanged passage near where the new language belongs.
4. CLASSIFY every insertion with an "insertion_style" of either "continuation" or "new_section":
   - **"continuation"**: the new language EXTENDS an existing section/paragraph (e.g., adding a payment-remedy sentence to the existing PAYMENT section, or extending a confidentiality paragraph). It is NOT its own section and must NOT have its own heading or section number. For continuation, set "context_anchor" to a verbatim passage in the paragraph being extended, put the new sentence(s) in "body_text", and leave "heading_text" as "". To control WHERE in the paragraph the new text lands, optionally set "after_sentence" to the EXACT verbatim sentence (including its ending period) that the new text should immediately follow; the new sentence will be inserted right after that sentence. If "after_sentence" is omitted or cannot be located, the new sentence is appended at the end of the paragraph. Choose "after_sentence" so the addition reads in its most logical position.
   - **"new_section"**: the new language is a genuinely STANDALONE new section that should have its own heading and take its place in the numbering (e.g., a new SURVIVAL section, a new CHANGE ORDER PROCESS section). For new_section, provide "heading_text" (the section's title LABEL only, with no number and no trailing punctuation — e.g., "Survival" or "Change Order Process") and "body_text" (the clause text WITHOUT the heading). Do NOT include the section number or delimiter in either field; the editor detects the document's heading style (capitalization, bold/underline, delimiter) and formats the heading to match the rest of the contract automatically.
5. "body_text" is REQUIRED for all insertions (the clause content). "heading_text" is REQUIRED for "new_section" and empty "" for "continuation".
6. For backward compatibility you may also populate "redlined_text" with the full text, but "body_text"/"heading_text" take precedence when present.
7. PLACEMENT — INTEGRATE BEFORE APPENDING (apply this BEFORE deciding continuation vs. new_section):
   Before concluding that a required rule needs a brand-new section, check whether a section covering the GENERAL TOPIC already exists (e.g., Liability/Limitation of Liability, Intellectual Property, Confidentiality, Payment, Termination, Notices). If such a section exists, do NOT create a duplicate new_section — instead amend that existing section in place: use a SURGICAL edit, or a "continuation" insertion that extends it (e.g., adding a "; provided that ..." sentence to an existing liability section, or a payment-remedy sentence to the existing payment section). DEFAULT TO CONTINUATION whenever the new language naturally belongs inside an existing section.
   CONSOLIDATE EDITS TO THE SAME SECTION: when SEVERAL rules all modify the SAME existing section (e.g., a liability cap, a damages cap, a caps-extend-to-affiliates rule, and a client-indemnification rule that all belong in LIMITATION OF LIABILITY), emit ONE coherent insertion that integrates all of them into a single, self-contained, grammatical passage — NOT several overlapping fragments. Two hard rules for this consolidated passage: (a) it must read as complete, properly punctuated sentences on its own — NEVER begin an insertion with a dangling connective ("Provided, however," "and," "provided that,") that depends on text from a different insertion to make sense; and (b) do NOT repeat the same obligation in more than one insertion (e.g., stating the client-indemnification list twice). If the combined language is too long for one continuation, split it by DISTINCT sub-topic so each insertion is independently coherent and non-overlapping — never by cutting a sentence across two insertions.
8. PLACEMENT — THE ANCHOR WATERFALL (only when the general topic is ENTIRELY ABSENT and a genuinely new standalone section is required):
   Add it as a completely new, sequentially numbered primary section ("insertion_style": "new_section"). To determine WHERE the new section belongs, set "context_anchor" to a verbatim passage at the START of the first available anchor from this waterfall, so the new section lands IMMEDIATELY BEFORE that anchor:
   (1) the "Entire Agreement", "Integration", or "Order of Precedence" clause; else
   (2) the "Miscellaneous" or "General" clause; else
   (3) the "[Signatures and Exhibit(s) Follow]" placeholder; else
   (4) the text introducing the signature block (e.g., "IN WITNESS WHEREOF..."); else
   (5) IF NONE OF THE ABOVE ANCHORS EXIST: append the new section AFTER the LAST existing numbered section of the contract (anchor on the final section's text). This is the normal fallback for contracts that have no integration clause or signature block.
   Appending at the end (or immediately before a closing/boilerplate anchor) preserves every existing section number, so internal cross-references stay accurate. NEVER insert a new section in a position that would force the existing sections to be renumbered.
9. If adding new_section insertions would make an existing "Sections 1 - N" range reference inaccurate, follow the CROSS-REFERENCE RANGES rule below (leave a comment; do not edit the number).
==========================================================
GENERAL STRUCTURAL RULES (BOTH TYPES)
==========================================================
- Do NOT fragment sentences or merge distinct numbered/lettered outline list items into one block. Preserve original list prefix layouts.
- Ignore artificial line-break/PDF layout artifacts; treat continuous thoughts as single paragraphs.
- Do NOT modify language that already satisfies the rule. If it complies, skip it entirely.
==========================================================
TYPE 3 — "comment": The clause is genuinely AMBIGUOUS and you cannot confidently tell whether the existing language already satisfies the rule.
==========================================================
A skilled lawyer does NOT force an edit when unsure — they leave a margin note asking the human to decide. Use this type ONLY when the existing wording is unusual or borderline and you genuinely cannot determine whether it already achieves the rule's goal. Do not use it to avoid normal edits.
RULES FOR COMMENT EDITS:
1. Set "edit_type": "comment".
2. Leave "target_text" and "redlined_text" as empty strings "".
3. Set "context_anchor" to the verbatim passage the comment should attach to.
4. In "issue_summary", explain what is ambiguous and what the human should decide.
### STRENGTHENED ANTI-COLLATERAL-DAMAGE RULE (READ CAREFULLY)
Before emitting ANY surgical edit, ask: "Does the existing language already achieve the GOAL of this playbook rule, even if it uses different words than I would?" Contracts frequently satisfy a rule through unusual phrasing, synonyms, or a different structure. If the existing clause already accomplishes the rule's objective, DO NOT edit it — skip it entirely. Only edit language that genuinely fails to achieve the rule's goal. When you are unsure whether differently-worded language already complies, emit a TYPE 3 "comment" rather than forcing a change. Never "fix" compliant language just because it differs from your preferred wording.
### DEFINED-TERM CONSISTENCY (ENFORCE IN EVERY EDIT)
Every "redlined_text" you write — surgical or insertion — MUST use the contract's existing capitalized defined terms exactly as the document defines them. Scan the GLOBAL DEFINED TERMS DICTIONARY provided below. If the document defines and uses "Customer", never write "Client", "Buyer", or "Purchaser"; if it uses "Contractor", never write "Provider", "Vendor", "Company", or "Consultant" in your inserted text. Match the document's capitalization and terminology precisely so inserted/edited language reads as if drafted by the same hand.
### CROSS-REFERENCE RANGES (e.g., "Order of Precedence")
If the contract contains a textual cross-reference to a RANGE or COUNT of sections (e.g., "Sections 1 - 24 of this Agreement", "Sections 1 through 20") AND your revisions add one or more new standalone sections (insertion_style "new_section"), that range may become inaccurate. Do NOT edit the number yourself (the final count depends on which tracked changes the human accepts). Instead, emit a TYPE 3 "comment" anchored to the cross-reference passage, with issue_summary explaining that new sections were added and the referenced range likely needs to be updated and verified by the reviewer.
### EXPECTED OUTPUT STRUCTURE
Return ONLY a valid JSON object with this exact shape:
{
  "rule_checklist": ["short string describing each rule you evaluated"],
  "revisions": [
    {
      "edit_type": "surgical",
      "context_anchor": "A unique, long, completely unchanged passage surrounding the target, used to locate it.",
      "target_text": "The exact minimal verbatim substring being changed or struck (surgical only). MUST be non-empty for surgical edits. Empty string \"\" for insertions and comments.",
      "redlined_text": "The minimal compliant replacement (surgical) OR empty string \"\" for a pure deletion/strike (surgical with a real target_text) OR empty string for insertions (use body_text instead) and comments.",
      "insertion_style": "For insertions only: 'continuation' (extends an existing section; no heading) or 'new_section' (standalone section with its own heading). Omit for surgical/comment.",
      "heading_text": "For insertion_style 'new_section' only: the section title LABEL with no number and no trailing punctuation (e.g., 'Survival'). Empty otherwise.",
      "body_text": "For insertions: the clause body text WITHOUT any heading. Required for insertions.",
      "after_sentence": "For continuation insertions only (optional): the exact verbatim sentence, including its ending period, that the new text should immediately follow. Omit to append at the paragraph end.",
      "position": "after",
      "rule_reference": "The Group letter and rule name driving this change, e.g. 'Group E — Indemnification & Tort Liability Cap'.",
      "issue_summary": "A plain-language description of what is wrong with the current language (surgical) or what is missing (insertion), written for a human attorney reviewing the margin comment.",
      "drafting_note": "Succinct reason, referencing the specific Playbook rule applied."
    }
  ]
}
For surgical edits, "position" may be omitted or set to "after". Always populate "rule_reference" and "issue_summary" so that, if an automated edit cannot be applied, the human reviewer receives a clear margin comment identifying the rule and the problem. Do not return any conversational text or markdown outside the JSON.
`;
// ============================================================================
//  PASS 1 (Extended default) — TERSE CHECKLIST
// ============================================================================
// A lightweight, Gem-style rule-by-rule checklist: one short line per rule stating
// status and the planned action. This is deliberately terse so it stays small and
// fast, and so the SEPARATE Pass-2 edit call can spend its entire output budget on
// edits. It is the single pre-pass for both Standard and Extended: a short rule-by-rule
// plan that Pass 2 restates and implements. The checklist's job is to make the editor
// reason-before-editing, exactly like the Gem's printed Chain-of-Thought Checklist,
// without consuming edit budget.
const CHECKLIST_WRAPPER = `
You are an expert contract attorney AI producing a SHORT Chain-of-Thought Checklist of a contract
against the Company's internal playbook. You do NOT write redline language here — you produce a concise,
disciplined, rule-by-rule plan that a separate editing step will implement.
### HOW TO BUILD THE CHECKLIST
Walk EVERY rule in the playbook below — both [Mandatory] and [Optional] — and evaluate each against the
ENTIRE contract. For each rule output ONE short line in this exact form:
  "Group X: Rule Name — STATUS — short planned action (and section)"
where STATUS is Compliant, Silent, or Non-Compliant, and the planned action is a brief phrase such as
"strike Section 5", "redline to Net 30 + 1.5% interest", "insert in Section 6 per Placement Rule",
"add new section", or "no action required". Keep each line to a single sentence. Mark already-compliant
clauses Compliant with "no action required" — do NOT invent changes to compliant clauses. For [Mandatory]
rules the contract is silent on, the action is to ADD the missing language. For [Optional] rules the
contract is silent on, the action is "no action required".
### THE PLANNED ACTION MUST BE SPECIFIC (this is what makes the checklist useful)
For every Non-Compliant or Silent-but-Mandatory rule, the planned action MUST name (a) the concrete
change to make AND (b) the specific Section number or clause where it applies. A vague action that omits
the change or the location is NOT acceptable — the separate editing step relies on this specificity to
place the edit correctly.
  GOOD (specific change + specific section):
    "Group B: Customer Data Transfer — Non-Compliant — revise Section 3.3 to remove fees for data export."
    "Group E: Liability Cap — Non-Compliant — revise Section 10.2 cap from fees-paid to 2x total contract value."
  BAD (too vague — do NOT produce lines like these):
    "Group B: Customer Data Transfer — Non-Compliant — needs revision."
    "Group E: Liability Cap — Non-Compliant — fix the cap."
Name the actual change and the actual Section/clause every time the contract contains the relevant
language; only when the contract is entirely silent (the clause does not exist) may you write
"add new section" without an existing section number.
### OUTPUT
Return ONLY a valid JSON object of this exact shape, with no markdown fences or conversational text:
{
  "checklist": [
    "Group A: Governing Law and Venue — Compliant — no action required.",
    "Group A: Equitable Relief Carve-Out — Non-Compliant — add injunctive-relief carve-out in Section 1.",
    "Group C: Liquidated Damages — Non-Compliant — strike Section 5 entirely."
  ]
}
Include a line for EVERY rule in the playbook, including compliant and silent ones, so the editing step
has the complete picture. Keep it tight — this is a checklist, not an essay.`;
/**
 * Runs the terse Pass-1 checklist. Returns { lines: string[], raw: string }.
 * Fails soft: on any parse problem it returns whatever lines it can salvage (possibly []),
 * and the editor proceeds — a checklist problem must never block delivery of a redline.
 */
async function runChecklistPass(adapter, rulesText, definedTermsDictionary, contractText) {
    let systemPrompt = `${CHECKLIST_WRAPPER}
### PLAYBOOK CONTRACTING RULES
=== START RULES ===
${rulesText}
=== END RULES ===
### GLOBAL DEFINED TERMS RECOGNITION DICTIONARY
${definedTermsDictionary || "(none provided)"}`;
    let raw = await executeWithRetry(adapter, systemPrompt, contractText);
    try {
        diagLog('\n========== [DIAG-CHECKLIST-RAW] ==========');
        diagLog('[DIAG-CHECKLIST-RAW] length: ' + (raw ? raw.length : 'null') + ' chars');
        console.log(raw ? raw.substring(0, 2000) : '(null/empty)');
        diagLog('========== [DIAG-CHECKLIST-RAW] END ==========\n');
    } catch (dErr) { diagLog('[DIAG-CHECKLIST-RAW] dump failed: ' + dErr.message); }
    let sanitized = (raw || '').replace(/```json/gi, '').replace(/```/g, '').trim();
    if (sanitized && sanitized[0] !== '{' && sanitized[0] !== '[') {
        let braceIdx = sanitized.indexOf('{');
        let brackIdx = sanitized.indexOf('[');
        let start = (braceIdx === -1) ? brackIdx : (brackIdx === -1) ? braceIdx : Math.min(braceIdx, brackIdx);
        if (start > 0) {
            let lastBrace = sanitized.lastIndexOf('}');
            let lastBrack = sanitized.lastIndexOf(']');
            let end = Math.max(lastBrace, lastBrack);
            if (end > start) sanitized = sanitized.substring(start, end + 1);
        }
    }
    let lines = [];
    try {
        let parsed = JSON.parse(sanitized);
        if (Array.isArray(parsed)) lines = parsed;
        else if (parsed && typeof parsed === 'object') {
            for (const key in parsed) {
                if (Array.isArray(parsed[key])) { lines = parsed[key]; break; }
            }
        }
    } catch (e) {
        diagWarn('[DIAG-CHECKLIST] checklist JSON parse failed (' + e.message + '); proceeding with no checklist lines.');
        lines = [];
    }
    // Keep only non-empty string lines.
    lines = (lines || []).filter(l => typeof l === 'string' && l.trim().length > 0).map(l => l.trim());
    return { lines, raw: raw || '' };
}
/**
 * Builds the directive that injects the terse checklist into the Pass-2 edit prompt and
 * instructs the editor to restate the plan before editing (the visible think-then-act step
 * that measurably improves edits, mirroring the Gem's printed checklist).
 */
function buildChecklistDirective(checklistLines, focusNote) {
    if (!checklistLines || checklistLines.length === 0) {
        return focusNote ? `\n${focusNote}` : "";
    }
    let body = checklistLines.map(l => "- " + l).join("\n");
    let directive = `
### YOUR PRE-APPROVED REVIEW CHECKLIST (follow it; implement every actionable item)
A prior review pass produced this rule-by-rule checklist for THIS contract. Before you write any edits,
read it and BRIEFLY restate (in a "plan" you will then act on) which items require action. Then implement
EVERY actionable item below — every strike, every redline, every required addition — and leave every
"no action required" / Compliant item untouched. Do not edit compliant clauses. Do not skip a Non-Compliant
or required-addition item.
=== START CHECKLIST ===
${body}
=== END CHECKLIST ===`;
    if (focusNote) directive += `\n${focusNote}`;
    return directive;
}
/**
 * Removes duplicate revisions from a merged set (used after the Extended do-over pass, where
 * the do-over re-does the full edit set and therefore re-emits clauses Pass 2 already produced).
 * Two different duplicate tests are applied by edit type, because insertions and surgical edits
 * fail differently:
 *   - INSERTIONS (new clauses): the model paraphrases the same clause slightly between passes,
 *     so we compare body text by token overlap (Jaccard >= 0.75). This catches verbatim repeats
 *     (score 1.00) and light paraphrases (~0.93) while leaving genuinely distinct clauses
 *     (~0.15) alone. The first occurrence is kept.
 *   - SURGICAL edits (in-place changes/strikes): these are precise span replacements the model
 *     reproduces verbatim, so we use an EXACT (whitespace- and smart-quote-normalized) match on
 *     BOTH target_text and redlined_text. Requiring both to match means we only ever collapse a
 *     true redo of the identical edit — two surgical edits that share a target but make different
 *     replacements are preserved. The first occurrence is kept.
 * Comments and anything else pass through untouched.
 */
function dedupeRevisions(revs) {
    // Fuzzy normalization for insertion body comparison.
    const normFuzzy = (s) => String(s || "").toLowerCase().replace(/^[\s.;:]+/, '').replace(/\s+/g, ' ').replace(/[^\w ]/g, '').trim();
    const tokens = (s) => new Set(normFuzzy(s).split(' ').filter(w => w.length > 2));
    const jaccard = (a, b) => {
        let A = tokens(a), B = tokens(b);
        if (A.size === 0 || B.size === 0) return 0;
        let inter = 0; for (let w of A) if (B.has(w)) inter++;
        return inter / (A.size + B.size - inter);
    };
    // Exact-but-tolerant normalization for surgical keys: collapse whitespace and fold smart
    // quotes to straight quotes (the model varies these between passes), but preserve case and
    // other punctuation so we never collapse two genuinely different edits.
    const normExact = (s) => String(s || "")
        .replace(/[\u2018\u2019\u201B\u0060\u00B4]/g, "'")
        .replace(/[\u201C\u201D\u201E]/g, '"')
        .replace(/\s+/g, ' ').trim();
    const isInsertion = (r) => {
        let et = (r.edit_type || "").toLowerCase();
        return et === "insertion" || (et !== "surgical" && et !== "comment" && (!r.target_text || r.target_text === ""));
    };
    const isSurgical = (r) => {
        let et = (r.edit_type || "").toLowerCase();
        return et === "surgical" || (et !== "insertion" && et !== "comment" && r.target_text && r.target_text !== "");
    };
    let kept = [], seenInsertion = [], seenSurgical = new Set();
    let droppedInsertions = 0, droppedSurgical = 0;
    for (let r of revs) {
        if (isInsertion(r)) {
            let body = r.body_text || r.redlined_text || "";
            if (normFuzzy(body).split(' ').length < 5) { kept.push(r); continue; } // too short to judge
            let dup = seenInsertion.some(s => jaccard(body, s) >= 0.75);
            if (dup) { droppedInsertions++; continue; }
            seenInsertion.push(body);
            kept.push(r);
        } else if (isSurgical(r)) {
            let key = normExact(r.target_text) + ' ||=>|| ' + normExact(r.redlined_text);
            if (seenSurgical.has(key)) { droppedSurgical++; continue; }
            seenSurgical.add(key);
            kept.push(r);
        } else {
            kept.push(r); // comments and anything else pass through untouched
        }
    }
    if (droppedInsertions > 0) diagLog('[DIAG-DEDUP] removed ' + droppedInsertions + ' duplicate insertion(s).');
    if (droppedSurgical > 0) diagLog('[DIAG-DEDUP] removed ' + droppedSurgical + ' duplicate surgical edit(s).');
    return kept;
}


/**
 * Runs ONE editing call: builds the editing system prompt (base wrapper + optional analysis
 * directive + optional examples), executes it, and parses the result into { revisions,
 * rule_checklist }. Used for Pass 2 and, in Extended, the optional do-over pass.
 * `directiveOverride` injects the checklist plan (from buildChecklistDirective) so the editor
 * restates and implements it. Returns parsed arrays; throws only on a hard parse failure so
 * the caller can decide how to handle it.
 */
async function runEditingCall(adapter, rulesText, dictionaryResponse, contractText, examplesText, providerLc, directiveOverride, label) {
    let systemPrompt = `You are an expert contract attorney AI assistant representing the Company.
You must review the user's contract text strictly based on the company's internal playbook rules.
### PLAYBOOK CONTRACTING RULES
=== START RULES ===
${rulesText}
=== END RULES ===
### GLOBAL DEFINED TERMS RECOGNITION DICTIONARY
${dictionaryResponse}
${UNIVERSAL_WRAPPER}`;
    if (directiveOverride) {
        systemPrompt += `\n${directiveOverride}`;
    }
    if (examplesText && providerLc !== 'ollama') {
        systemPrompt += `\n\n### APPROVED CORPORATE TONE EXAMPLES\n=== START APPROVED EXAMPLES ===\n${examplesText}\n=== END APPROVED EXAMPLES ===`;
    }
    let raw = await executeWithRetry(adapter, systemPrompt, contractText);
    try {
        diagLog('\n========== [DIAGNOSTIC-RAW]' + (label ? ' ' + label : '') + ' ==========');
        diagLog('[DIAGNOSTIC-RAW] Length: ' + (raw ? raw.length : 'null') + ' chars');
        diagLog('[DIAGNOSTIC-RAW] First 4000 chars of raw model output:');
        console.log(raw ? raw.substring(0, 4000) : '(null)');
        diagLog('========== [DIAGNOSTIC-RAW] END ==========\n');
    } catch (dErr) { diagLog('[DIAGNOSTIC-RAW] failed: ' + dErr.message); }
    let sanitized = (raw || '').replace(/```json/gi, '').replace(/```/g, '').trim();
    let parsed = JSON.parse(sanitized); // may throw — caller handles
    let out = { rule_checklist: [], revisions: [] };
    if (Array.isArray(parsed)) {
        out.revisions = parsed;
    } else if (parsed && typeof parsed === 'object') {
        for (const key in parsed) {
            if (key.toLowerCase() === 'revisions') out.revisions = parsed[key] || [];
            else if (key.toLowerCase() === 'rule_checklist') out.rule_checklist = parsed[key] || [];
        }
    }
    return out;
}
const statusClients = new Map();
app.get('/api/status', (req, res) => {
    const sessionId = req.query.sessionId;
    if (!sessionId) return res.status(400).end();
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    statusClients.set(sessionId, res);
    req.on('close', () => {
        // Only remove this session's stream if it is still the one registered.
        if (statusClients.get(sessionId) === res) {
            statusClients.delete(sessionId);
        }
    });
});
// Send a status message to a single user's stream (identified by sessionId).
function sendStatus(sessionId, message) {
    const client = statusClients.get(sessionId);
    if (client) {
        client.write(`data: ${JSON.stringify({ message })}\n\n`);
    }
}
app.get('/api/init-check', (req, res) => {
    try {
        const cfg = getServerConfig();
        const envKey = process.env.ENTERPRISE_LICENSE_KEY;
        if (envKey && cfg.playbooksFolderPath && envKey.trim() !== "") {
            const inputHash = crypto.createHash('sha256').update(envKey.trim()).digest('hex');
            if (validHashes.includes(inputHash)) {
                return res.json({
                    authenticated: true,
                    licenseKey: envKey.trim(),
                    playbookPath: cfg.playbooksFolderPath
                });
            }
        }
        res.json({ authenticated: false });
    } catch (error) {
        console.error("IT provisioning verification exception:", error);
        res.json({ authenticated: false });
    }
});
// NOTE: The former POST /api/settings route has been removed. Per-user engine and
// model choices are no longer stored on the server; they travel with each
// /api/redline request so concurrent users never overwrite one another.

// Reports which engines and essentials are configured in .env, WITHOUT ever exposing
// secrets (no API keys, URLs, model names, or paths leave the server). The client uses
// these booleans to enable/disable engine choices and to show "contact IT" messages.
// "Configured" means the required values are present and are not unfilled placeholders —
// applied identically to cloud and local. It does NOT verify that a present value is
// actually correct/reachable (a wrong key or dead URL still reads as "configured"); an
// incorrect value surfaces as a clear error at run time. This keeps page load instant.
function isPlaceholder(v) {
    if (!v) return true;
    let s = String(v).trim();
    if (s === "") return true;
    // Common unfilled-template markers: <YOUR_...>, YOUR_..., ROTATE_THIS..., bare <...>
    if (s.includes("<") || s.startsWith("YOUR_") || s.includes("YOUR_") || s.includes("ROTATE_THIS")) return true;
    return false;
}
app.get('/api/config-status', (req, res) => {
    try {
        const cfg = getServerConfig();
        // Cloud is configured when an API key is present and not a placeholder. The endpoint
        // and model have safe defaults, so the key is the determining factor.
        // Cloud is configured when the credentials REQUIRED BY THE SELECTED STRATEGY are present.
        // Different vendors authenticate differently: the single-API-key strategies need
        // CLOUD_API_KEY; Vertex needs a project + service-account credentials; Bedrock needs AWS
        // keys. Checking only CLOUD_API_KEY would wrongly report a correctly-configured Vertex or
        // Bedrock deployment as "not configured" and block the run. This mirrors the adapter's
        // own strategy-aware credential validation.
        const cloudStrategy = (process.env.CLOUD_API_STRATEGY || 'gemini-native').toLowerCase().trim();
        let cloudConfigured;
        if (cloudStrategy === 'gemini-vertex') {
            cloudConfigured = !isPlaceholder(process.env.VERTEX_PROJECT_ID) &&
                              !isPlaceholder(process.env.GOOGLE_APPLICATION_CREDENTIALS);
        } else if (cloudStrategy === 'bedrock') {
            cloudConfigured = !isPlaceholder(process.env.AWS_REGION) &&
                              !isPlaceholder(process.env.AWS_ACCESS_KEY_ID) &&
                              !isPlaceholder(process.env.AWS_SECRET_ACCESS_KEY);
        } else {
            // gemini-native, openai-compatible, anthropic-native — single API key.
            cloudConfigured = !isPlaceholder(process.env.CLOUD_API_KEY);
        }
        // Local is configured when both the server URL and the model are present and not
        // placeholders. (Both have localhost/default fallbacks, so we check the raw env values
        // rather than the defaulted config to avoid a false "ready".)
        const localConfigured = !isPlaceholder(process.env.LOCAL_AI_SERVER_URL) &&
                                !isPlaceholder(process.env.LOCAL_DEFAULT_MODEL);
        const playbooksConfigured = !isPlaceholder(process.env.LOCAL_PLAYBOOKS_PATH) &&
                                    fs.existsSync(cfg.playbooksFolderPath);
        // License is valid when the env key hashes to a known-good license.
        let licenseValid = false;
        const envKey = process.env.ENTERPRISE_LICENSE_KEY;
        if (!isPlaceholder(envKey)) {
            const inputHash = crypto.createHash('sha256').update(String(envKey).trim()).digest('hex');
            licenseValid = validHashes.includes(inputHash);
        }
        // Default engine the IT admin chose. Normalize to the provider strings the code uses
        // ("cloud" / "ollama"). Accept "local" as an alias for "ollama".
        let defaultProvider = (process.env.DEFAULT_PROVIDER || "cloud").toLowerCase().trim();
        if (defaultProvider === "local") defaultProvider = "ollama";
        if (defaultProvider !== "cloud" && defaultProvider !== "ollama") defaultProvider = "cloud";
        res.json({
            cloudConfigured,
            localConfigured,
            playbooksConfigured,
            licenseValid,
            defaultProvider
        });
    } catch (error) {
        console.error('Config status error:', error);
        res.status(500).json({ error: 'Failed to read server configuration status.' });
    }
});
app.post('/api/verify-license', (req, res) => {
    const { licenseKey } = req.body;
    if (!licenseKey) return res.status(400).json({ error: 'No key provided.' });
    const inputHash = crypto.createHash('sha256').update(licenseKey).digest('hex');
    if (validHashes.includes(inputHash)) {
        res.json({ valid: true, message: 'License verified.' });
    } else {
        res.status(401).json({ valid: false, error: 'Invalid Product Key.' });
    }
});
app.post('/api/redline', async (req, res) => {
    let playbookName = "Unknown";
    let currentKey = "Unknown";
    let cotIntervalTimer = null;
    // Capture user network information securely across proxy links
    const clientIp = req.headers['x-forwarded-for'] || req.socket.remoteAddress || "127.0.0.1";
    try {
        const { playbookFolderName, documentBuffer, clientKey, userEmail, contractFileName, twoPass, provider, modelName, aiServerUrl, sessionId, reviewMode } = req.body;
        playbookName = playbookFolderName;
        currentKey = clientKey;
        // Resolve the review mode (Extended / Standard). Default = Extended (most thorough)
        // to favor accuracy. "Fast" was retired once it became identical to Standard (both are
        // a single editing pass); any legacy client still sending "fast" is folded into
        // Resolve review depth. New clients send reviewMode ("extended"/"standard").
        // Legacy clients send only the boolean twoPass: true meant the thorough two-pass
        // review (extended); false meant single-pass (standard). Anything unrecognized
        // defaults to extended.
        let mode = (reviewMode || "").toLowerCase().trim();
        if (mode === "fast") mode = "standard"; // retired alias
        if (mode !== "extended" && mode !== "standard") {
            if (reviewMode === undefined && (twoPass === false || twoPass === "false")) mode = "standard";
            else mode = "extended";
        }
        const runAnalysis = (mode === "extended");
        // Bind status updates to THIS request's session so only this user sees them.
        const broadcastStatus = (message) => sendStatus(sessionId, message);
        const serverConfig = getServerConfig();
        // Per-run settings = shared read-only path (from .env) + this user's choices (from the request)
        // Per-run settings. The PROVIDER (cloud vs. local) is the user's per-run choice from
        // the task pane. Everything else — the local server URL, the model name, the playbooks
        // path, and all cloud credentials — comes exclusively from .env (server-owned). Any
        // model/URL the client may send is intentionally ignored so configuration lives in one
        // place (the .env file, managed by IT) and no endpoint or model is selectable from the UI.
        const runSettings = {
            provider: (provider || serverConfig.defaultProvider || 'cloud').toLowerCase(),
            modelName: serverConfig.defaultModelName,
            aiServerUrl: serverConfig.aiServerUrl,
            playbooksFolderPath: serverConfig.playbooksFolderPath
        };
        const executionTrack = runSettings.provider;
        // activeModel must reflect what each adapter ACTUALLY uses, for both the
        // diagnostic banner and the audit ledger. The cloud adapter reads its model
        // from CLOUD_MODEL_PRIMARY (.env) and ignores runSettings.modelName entirely;
        // only the local Ollama adapter uses runSettings.modelName (LOCAL_DEFAULT_MODEL).
        // Without this split, every cloud run was mislabeled with the local model name.
        const activeModel = (executionTrack === 'cloud')
            ? (process.env.CLOUD_MODEL_PRIMARY || 'cloud-model-unset')
            : runSettings.modelName;
        if (!runSettings.playbooksFolderPath) {
            return res.status(400).json({ error: 'Server playbooks path is not configured.' });
        }
        broadcastStatus(`Loading the selected playbook and rules...`);
        const llmService = createAdapter(runSettings);
        const provider_lc = runSettings.provider;
        const playbookDir = path.join(runSettings.playbooksFolderPath, playbookFolderName);
        const rulesDocPath = path.join(playbookDir, 'rules.docx');
        let rulesText = "";
        if (fs.existsSync(rulesDocPath)) {
            const ruleExtract = await mammoth.convertToMarkdown({ path: rulesDocPath });
            rulesText = ruleExtract.value;
        } else {
            return res.status(404).json({ error: `Rules document not found at: ${rulesDocPath}` });
        }
        const examplesDir = path.join(playbookDir, 'examples');
        let examplesText = "";
        if (fs.existsSync(examplesDir)) {
            const files = fs.readdirSync(examplesDir);
            for (const file of files) {
                if (file.toLowerCase().endsWith('.docx')) {
                    const examplePath = path.join(examplesDir, file);
                    try {
                        const exampleExtract = await mammoth.convertToMarkdown({ path: examplePath });
                        examplesText += `\n--- Approved Example Language from ${file} ---\n${exampleExtract.value}\n`;
                    } catch (err) {
                        console.log(`Failed to read example file: ${file}`, err);
                    }
                }
            }
        }
        if (examplesText) {
            broadcastStatus(`Playbook rules and approved example language loaded.`);
        } else {
            broadcastStatus(`Playbook rules loaded.`);
        }
        const contractText = Buffer.from(documentBuffer, 'base64').toString('utf8');
        broadcastStatus(`Identifying the contract's defined terms...`);
        const definitionSystemPrompt = "You are a professional legal document parser. Extract all capitalized defined terms from the contract text into a brief JSON dictionary. Return ONLY a valid JSON array of objects with 'term' and 'definition' keys. Do not include conversational text or markdown code wrappers outside the JSON array.";
        const dictionaryResponse = await executeWithRetry(llmService, definitionSystemPrompt, contractText);
        broadcastStatus(`Defined terms identified.`);
        // Identify the playbook's rule groups so we can tell the user WHICH areas are being
        // reviewed (truthful — these come straight from this playbook's rules.docx). We show
        // this list once; we do NOT pretend to step through them live, because the redline is
        // a single model call whose internal progress the server cannot observe.
        const groupMatches = rulesText.match(/\*\*Group\s+[A-Z]:[^\*]+\*\*/gi) || [];
        const groupNames = groupMatches.map(g => g.replace(/\*\*/g, '').trim());
        if (groupNames.length > 0) {
            let list = groupNames.join("; ");
            // Keep it readable if there are many groups.
            broadcastStatus(`Reviewing the contract against ${groupNames.length} playbook areas: ${list}.`);
        } else {
            broadcastStatus(`Reviewing the contract against the selected playbook's rules.`);
        }

        // ============================================================
        //  EDITING ORCHESTRATION — Standard and Extended
        // ============================================================
        // Both depths run exactly two model calls in the normal case:
        //   Pass 1 = terse rule-by-rule checklist (cheap, small, guides Pass 2)
        //   Pass 2 = one edit call that implements the checklist
        // They differ ONLY in how a blank insertion stub from Pass 2 is handled:
        //   - STANDARD: drop the blank, leave a reviewer comment, stop. Always 2 calls.
        //   - EXTENDED: run ONE do-over edit pass that re-does the FULL set of needed edits
        //     (both new clauses AND in-place fixes/strikes), then merge with Pass 2 and dedupe.
        //     2 calls normally, 3 at most. Nothing required is silently lost.
        // A blank stub is an insertion whose body_text came back empty — the model identified a
        // clause to add but did not draft it. We never auto-apply a blank; we either comment it
        // (Standard) or try once more to draft the whole set (Extended).
        const isExtended = (mode === "extended");

        // Reviewer comments for clauses that were identified but not auto-drafted. These are
        // appended to the revision set as comment-only items the apply layer already understands.
        let undraftedComments = [];
        const noteUndrafted = (rev) => {
            let rule = (rev.rule_reference && String(rev.rule_reference).trim()) ? String(rev.rule_reference).trim() : "a required";
            undraftedComments.push({
                edit_type: "comment",
                context_anchor: rev.context_anchor || "",
                target_text: "",
                redlined_text: "",
                position: "after",
                rule_reference: rev.rule_reference || "Playbook rule",
                issue_summary: `A required ${rule} clause was identified but could not be auto-drafted. Please add it manually.`,
                drafting_note: `A required ${rule} clause was identified but could not be auto-drafted. Please add it manually.`
            });
        };

        // Identify the blank insertion stubs in a revision set; return { populated, blanks }.
        const splitBlanks = (revisions) => {
            let populated = [], blanks = [];
            for (let r of (revisions || [])) {
                let et = (r.edit_type || "").toLowerCase();
                let isInsertion = et === "insertion" || (et !== "surgical" && et !== "comment" && (!r.target_text || r.target_text === ""));
                if (isInsertion && !((r.body_text || "").toString().trim().length > 0)) {
                    blanks.push(r);
                } else {
                    populated.push(r);
                }
            }
            return { populated, blanks };
        };

        let masterRevisions = [];
        let masterChecklist = [];

        // ---- PASS 1: terse checklist (both depths) ----
        broadcastStatus(`Pass 1 of 2 \u2014 building a rule-by-rule checklist of the contract against the playbook...`);
        let checklistLines = [];
        try {
            let cl = await runChecklistPass(llmService, rulesText, dictionaryResponse, contractText);
            checklistLines = cl.lines || [];
        } catch (clErr) {
            console.error('[DIAG-CHECKLIST] checklist pass error (non-fatal):', clErr.message);
            checklistLines = [];
        }
        diagLog('[DIAG-CHECKLIST] checklist lines: ' + checklistLines.length);
        // If the checklist failed entirely, Pass 2 still runs (unguided); we flag this at the
        // top of the document so the reviewer knows to look more carefully.
        let checklistFailed = (checklistLines.length === 0);
        broadcastStatus(`Pass 1 complete \u2014 ${checklistLines.length} rules checked. Pass 2 of 2 \u2014 drafting the redline edits...`);

        // Heartbeat for the (opaque) editing calls. Plain "still working" reassurance; it does
        // not claim a position the server can't observe.
        let heartbeatCount = 0;
        let heartbeatPhase = "editing"; // "editing" (Pass 2) | "doover" (Extended recovery)
        cotIntervalTimer = setInterval(() => {
            heartbeatCount++;
            let elapsed = formatElapsed(heartbeatCount * 15);
            if (heartbeatPhase === "doover") {
                broadcastStatus(`A few clauses need a second drafting pass \u2014 still completing them (${elapsed} elapsed)...`);
            } else {
                broadcastStatus(`Pass 2 of 2 \u2014 still drafting the redline edits (${elapsed} elapsed). Larger contracts and models take longer...`);
            }
        }, 15000);

        // ---- PASS 2: single edit call, guided by the checklist ----
        let editOut = null;
        try {
            editOut = await runEditingCall(llmService, rulesText, dictionaryResponse, contractText, examplesText, provider_lc, buildChecklistDirective(checklistLines), executionTrack + ' / ' + activeModel + ' [PASS 2 edit]');
        } catch (editErr) {
            console.error('[DIAG-PASS2] edit call parse/exec error:', editErr.message);
            editOut = null;
        }

        if (!editOut || !Array.isArray(editOut.revisions)) {
            // Pass 2 produced nothing usable (e.g. malformed JSON). We cannot recover edits
            // without a usable response; fail cleanly with the same audit/error path as before.
            if (cotIntervalTimer) { clearInterval(cotIntervalTimer); cotIntervalTimer = null; }
            console.error('[DIAG-PASS2] Pass 2 produced no usable revisions.');
            writeAuditLog(userEmail, clientIp, contractFileName, playbookName, executionTrack, activeModel, 0, "PARSE_FAILURE", mode);
            return res.status(500).json({ error: 'The AI generated an improperly formatted response block structure.' });
        }

        masterChecklist = editOut.rule_checklist || [];
        let pass2Split = splitBlanks(editOut.revisions);
        diagLog('[DIAG-PASS2] revisions=' + editOut.revisions.length + ' populated=' + pass2Split.populated.length + ' blanks=' + pass2Split.blanks.length);

        if (pass2Split.blanks.length === 0) {
            // Clean Pass 2: accept it as the complete redline. (Both depths: 2 calls.)
            masterRevisions = pass2Split.populated;
            diagLog('[DIAG-PASS2] no blank stubs \u2014 accepting Pass 2 as final.');
            broadcastStatus(`Review complete \u2014 ${masterRevisions.length} suggested edits. Preparing your document...`);
        } else if (!isExtended) {
            // STANDARD with blank stubs: keep the good edits, comment the blanks, NO do-over.
            masterRevisions = pass2Split.populated;
            pass2Split.blanks.forEach(noteUndrafted);
            diagLog('[DIAG-PASS2] STANDARD \u2014 ' + pass2Split.blanks.length + ' blank stub(s) dropped and flagged for the reviewer (no do-over).');
            broadcastStatus(`Review complete \u2014 ${masterRevisions.length} suggested edits. Preparing your document...`);
        } else {
            // EXTENDED with blank stubs: ONE do-over pass redoing the FULL set, then merge+dedupe.
            diagWarn('[DIAG-PASS2] EXTENDED \u2014 ' + pass2Split.blanks.length + ' blank stub(s); running one full-set do-over pass.');
            broadcastStatus(`A few clauses need a second drafting pass \u2014 completing them now...`);
            heartbeatPhase = "doover";
            // Start from Pass 2's populated edits; the do-over re-drafts everything and the
            // merge + dedupe keeps the best of each (verbatim/paraphrase insertion repeats and
            // exact surgical repeats are collapsed).
            masterRevisions = pass2Split.populated.slice();
            let doOver = null;
            try {
                doOver = await runEditingCall(llmService, rulesText, dictionaryResponse, contractText, examplesText, provider_lc, buildChecklistDirective(checklistLines), executionTrack + ' / ' + activeModel + ' [PASS 2b do-over]');
            } catch (e2) {
                console.error('[DIAG-PASS2] do-over call failed:', e2.message);
                doOver = null;
            }
            if (doOver && Array.isArray(doOver.revisions)) {
                let doSplit = splitBlanks(doOver.revisions);
                diagLog('[DIAG-PASS2] do-over revisions=' + doOver.revisions.length + ' populated=' + doSplit.populated.length + ' blanks=' + doSplit.blanks.length);
                masterRevisions = masterRevisions.concat(doSplit.populated);
                if (Array.isArray(doOver.rule_checklist) && doOver.rule_checklist.length > 0) {
                    masterChecklist = masterChecklist.concat(doOver.rule_checklist);
                }
                // Any clause STILL blank after the do-over is flagged for the reviewer. We match
                // a still-blank clause to a Pass 2 blank by rule_reference so we comment each
                // un-recovered rule once.
                let recoveredRules = new Set(doSplit.populated.map(r => String(r.rule_reference || "").toLowerCase().trim()));
                for (let b of pass2Split.blanks) {
                    let key = String(b.rule_reference || "").toLowerCase().trim();
                    if (!recoveredRules.has(key)) noteUndrafted(b);
                }
                for (let b of doSplit.blanks) {
                    let key = String(b.rule_reference || "").toLowerCase().trim();
                    if (!recoveredRules.has(key)) noteUndrafted(b);
                }
            } else {
                // Do-over failed: fall back to commenting the original blanks (never lose them).
                pass2Split.blanks.forEach(noteUndrafted);
            }
            // Merge cleanup: collapse duplicate insertions (paraphrases) and duplicate surgical
            // edits (exact) introduced by re-doing the full set.
            masterRevisions = dedupeRevisions(masterRevisions);
            diagLog('[DIAG-PASS2] post-do-over merged revision count: ' + masterRevisions.length);
            broadcastStatus(`Review complete \u2014 ${masterRevisions.length} suggested edits. Preparing your document...`);
        }

        // Append the undrafted-clause reviewer comments (if any) to the revision set.
        if (undraftedComments.length > 0) {
            diagLog('[DIAG-UNDRAFTED] ' + undraftedComments.length + ' required clause(s) flagged for manual drafting.');
            masterRevisions = masterRevisions.concat(undraftedComments);
        }

        // If the checklist failed, prepend a top-of-document reviewer note (handled client-side
        // as a comment-only revision anchored to the document start via empty context_anchor).
        if (checklistFailed) {
            masterRevisions.unshift({
                edit_type: "comment",
                context_anchor: "",
                target_text: "",
                redlined_text: "",
                position: "after",
                rule_reference: "Review Note",
                issue_summary: "Note: the rule-by-rule checklist step did not complete, so edits were drafted without it. Please review with extra care.",
                drafting_note: "Note: the rule-by-rule checklist step did not complete, so edits were drafted without it. Please review with extra care."
            });
        }
        if (cotIntervalTimer) {
            clearInterval(cotIntervalTimer);
            cotIntervalTimer = null;
        }
        // Dedupe the checklist in case Pass 2 and the do-over reported overlapping rule names.
        if (Array.isArray(masterChecklist) && masterChecklist.length > 0) {
            masterChecklist = Array.from(new Set(masterChecklist.map(s => String(s))));
        }
        broadcastStatus(`Contract analysis complete. Preparing the suggested changes...`);

        // ============================ DIAGNOSTIC BLOCK ============================
        // TEMPORARY: dumps what the model returned so we can compare anchor text to
        // the document. REMOVE THIS BLOCK (down to the matching end marker) when done.
        try {
            diagLog('\n========== [DIAGNOSTIC] ' + executionTrack + ' / ' + activeModel + ' ==========');
            diagLog('[DIAGNOSTIC] Revision count: ' + (Array.isArray(masterRevisions) ? masterRevisions.length : 'NOT AN ARRAY'));
            if (Array.isArray(masterRevisions)) {
                masterRevisions.forEach((r, i) => {
                    diagLog('\n[DIAGNOSTIC] --- Revision ' + (i + 1) + ' ---');
                    diagLog('[DIAGNOSTIC]   edit_type     : ' + JSON.stringify(r.edit_type));
                    diagLog('[DIAGNOSTIC]   target_text   : ' + JSON.stringify(r.target_text));
                    diagLog('[DIAGNOSTIC]   context_anchor: ' + JSON.stringify(r.context_anchor));
                    diagLog('[DIAGNOSTIC]   redlined_text : ' + JSON.stringify(r.redlined_text));
                    diagLog('[DIAGNOSTIC]   rule_reference: ' + JSON.stringify(r.rule_reference));
                    if (r.heading_text !== undefined) diagLog('[DIAGNOSTIC]   heading_text  : ' + JSON.stringify(r.heading_text));
                    if (r.body_text !== undefined) diagLog('[DIAGNOSTIC]   body_text     : ' + JSON.stringify(r.body_text));
                });
            }
            diagLog('\n========== [DIAGNOSTIC] END ==========\n');
        } catch (diagErr) {
            diagLog('[DIAGNOSTIC] dump failed: ' + diagErr.message);
        }
        // ========================== END DIAGNOSTIC BLOCK =========================

        // (No separate consistency pass — the reviewing attorney reviews the redline. The
        // safety net is the blank-stub handling: Standard flags an undrafted clause for the
        // reviewer; Extended runs one do-over pass to draft it, then dedupes. Standard is
        // always 2 calls (checklist + edit); Extended is 2, or 3 when a do-over is needed.)
        broadcastStatus(`Finalizing and applying tracked changes to your document...`);
        writeAuditLog(userEmail, clientIp, contractFileName, playbookName, executionTrack, activeModel, masterRevisions.length, "SUCCESS", mode);
        res.json({ rule_checklist: masterChecklist, revisions: masterRevisions, reviewMode: mode });
    } catch (error) {
        if (cotIntervalTimer) {
            clearInterval(cotIntervalTimer);
        }
        console.error('Error processing contract:', error);
        const { playbookFolderName, userEmail, contractFileName, reviewMode } = req.body;
        writeAuditLog(userEmail, clientIp, contractFileName, playbookFolderName, "unknown", "unknown", 0, "SYSTEM_ERROR", (reviewMode || "unknown"));
        res.status(500).json({ error: 'Backend failed to process the contract execution targets.' });
    }
});
app.get('/api/playbooks', (req, res) => {
    const settings = getServerConfig();
    if (!settings.playbooksFolderPath || !fs.existsSync(settings.playbooksFolderPath)) {
        return res.json({ playbooks: [] });
    }
    try {
        const folders = fs.readdirSync(settings.playbooksFolderPath, { withFileTypes: true })
            .filter(dirent => dirent.isDirectory())
            .map(dirent => dirent.name);
        res.json({ playbooks: folders });
    } catch (error) {
        res.status(500).json({ error: 'Failed to read playbooks matrix directory targets.' });
    }
});
// Queryable audit ledger read endpoint. Supports optional filters so a compliance
// reviewer can pull a bounded, ordered slice without scanning a flat file:
//   ?limit=100&offset=0&status=SUCCESS&since=2026-01-01T00:00:00Z
// Returns rows newest-first. All stored values are already masked; nothing here
// exposes plaintext PII/MNPI.
app.get('/api/audit-log', (req, res) => {
    try {
        const limit = Math.min(parseInt(req.query.limit || '100', 10) || 100, 1000);
        const offset = Math.max(parseInt(req.query.offset || '0', 10) || 0, 0);
        const conditions = [];
        const params = [];
        if (req.query.status) {
            conditions.push('execution_status = ?');
            params.push(String(req.query.status));
        }
        if (req.query.since) {
            conditions.push('timestamp >= ?');
            params.push(String(req.query.since));
        }
        if (req.query.until) {
            conditions.push('timestamp <= ?');
            params.push(String(req.query.until));
        }
        const where = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';
        const rows = auditDb
            .prepare(`SELECT * FROM audit_log ${where} ORDER BY id DESC LIMIT ? OFFSET ?`)
            .all(...params, limit, offset);
        const total = auditDb
            .prepare(`SELECT COUNT(*) AS n FROM audit_log ${where}`)
            .get(...params).n;
        res.json({ total, limit, offset, rows });
    } catch (error) {
        console.error('Audit query error:', error);
        res.status(500).json({ error: 'Failed to query the audit ledger.' });
    }
});
// Tamper-evidence verification. Walks the audit chain from the first row, recomputing each
// row's hash from its stored fields + the previous row's hash, and comparing to the stored
// row_hash. Any edit, deletion, or reordering of historical rows breaks the chain at the
// first affected row. Returns { intact: true, rows: N } when clean, or { intact: false,
// brokenAtId, reason } at the first discrepancy. Requires the same AUDIT_LOG_SALT that wrote
// the rows. (Read-only; exposes no plaintext PII — only confirms integrity.)
app.get('/api/audit-verify', (req, res) => {
    try {
        const salt = process.env.AUDIT_LOG_SALT || "[placeholder]";
        const rows = auditDb.prepare(`SELECT * FROM audit_log ORDER BY id ASC`).all();
        let expectedPrev = AUDIT_CHAIN_GENESIS;
        for (const row of rows) {
            // The prev_hash stored on the row must match the running chain value.
            if ((row.prev_hash || null) !== expectedPrev) {
                return res.json({
                    intact: false,
                    brokenAtId: row.id,
                    reason: 'prev_hash does not match the previous row\u2019s hash (a row was altered, inserted, or deleted).',
                    rowsChecked: rows.length
                });
            }
            // Recompute this row's hash from its fields and compare to what's stored.
            const recomputed = computeAuditRowHash({
                timestamp: row.timestamp, masked_user: row.masked_user, client_ip: row.client_ip,
                masked_filename: row.masked_filename, playbook_name: row.playbook_name,
                execution_track: row.execution_track, active_model: row.active_model,
                applied_revisions: row.applied_revisions, execution_status: row.execution_status,
                review_mode: row.review_mode, prev_hash: row.prev_hash
            }, salt);
            if (recomputed !== row.row_hash) {
                return res.json({
                    intact: false,
                    brokenAtId: row.id,
                    reason: 'row contents do not match its stored hash (this row was modified).',
                    rowsChecked: rows.length
                });
            }
            expectedPrev = row.row_hash;
        }
        res.json({ intact: true, rows: rows.length });
    } catch (error) {
        console.error('Audit verify error:', error);
        res.status(500).json({ error: 'Failed to verify the audit ledger chain.' });
    }
});
// 3001 for local development so existing local setups are unaffected. Binding 0.0.0.0 makes
// the server reachable from other machines on a company network (required when users' Word
// clients connect to a shared server); on a local machine this still serves localhost fine.
const PORT = parseInt(process.env.PORT || '3001', 10);
const BIND_HOST = process.env.BIND_HOST || '0.0.0.0';

app.listen(PORT, BIND_HOST, () => {
    console.log(`Contract Negotiation Assistant Server Active On Port ${PORT}`);
});
