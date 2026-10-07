// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.
// The backend lives at the SAME origin that served this task pane. Word loads the pane
// from whatever the manifest's SourceLocation points at — http://localhost:3001 during
// local development, or a customer's own https server in production — and the API lives
// on that same server. So we call back to our own origin and never hardcode any host.
// This makes one build work for every customer: each customer's manifest localizes the
// install to their server, and this code follows automatically. The localhost fallback
// only applies if window.location.origin is somehow unavailable (e.g. a file:// context).
const BACKEND_URL = (self.location && self.location.origin && self.location.origin !== "null")
    ? self.location.origin
    : "http://localhost:3001";
// Unique id for this task pane session. It scopes the server's status stream so this
// user only receives their own progress messages (never other concurrent users').
const SESSION_ID = (self.crypto && self.crypto.randomUUID)
    ? self.crypto.randomUUID()
    : 'sess-' + Date.now() + '-' + Math.random().toString(36).slice(2);
// Holds the server's configuration readiness, fetched once on load. The user never enters
// configuration; this only tells the UI which engines are usable and which to default to.
let CONFIG_STATUS = { cloudConfigured: false, localConfigured: false, playbooksConfigured: false, licenseValid: false, defaultProvider: "cloud" };

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("btn-review").onclick = reviewContract;
        document.getElementById("engine-toggle").onchange = handleEngineToggle;
        // Ask the server what is configured in .env (booleans + default only — never secrets).
        try {
            const statusResp = await fetch(`${BACKEND_URL}/api/config-status`);
            CONFIG_STATUS = await statusResp.json();
        } catch (err) {
            console.warn("Could not reach the server to check configuration status.", err);
            // Leave CONFIG_STATUS at its all-false default; the incomplete panel will show.
        }
        // Essentials required to run at all: a valid license, a playbook library, and at least
        // one usable engine. If any is missing, show the "contact IT" panel instead of a dead
        // dashboard.
        const haveAnyEngine = CONFIG_STATUS.cloudConfigured || CONFIG_STATUS.localConfigured;
        const essentialsReady = CONFIG_STATUS.licenseValid && CONFIG_STATUS.playbooksConfigured && haveAnyEngine;
        if (!essentialsReady) {
            showConfigIncomplete();
            return;
        }
        // Dashboard is usable. Default the engine toggle to IT's chosen default, but only if
        // that engine is actually configured; otherwise default to whichever one is.
        document.getElementById("config-incomplete").style.display = "none";
        document.getElementById("main-interface").style.display = "block";
        let pref = CONFIG_STATUS.defaultProvider;
        if (pref === "cloud" && !CONFIG_STATUS.cloudConfigured) pref = "ollama";
        if (pref === "ollama" && !CONFIG_STATUS.localConfigured) pref = "cloud";
        document.getElementById("engine-toggle").value = pref;
        handleEngineToggle();
        await loadPlaybooks();
    }
});
// Renders the "setup incomplete" panel, naming exactly what IT still needs to configure.
function showConfigIncomplete() {
    document.getElementById("main-interface").style.display = "none";
    document.getElementById("config-incomplete").style.display = "block";
    let missing = [];
    if (!CONFIG_STATUS.licenseValid) missing.push("a valid product license");
    if (!CONFIG_STATUS.playbooksConfigured) missing.push("the playbook library");
    if (!CONFIG_STATUS.cloudConfigured && !CONFIG_STATUS.localConfigured) missing.push("at least one AI engine (Cloud or Local)");
    let listEl = document.getElementById("config-missing-list");
    if (missing.length > 0) {
        listEl.innerText = "Still needs configuring: " + missing.join(", ") + ". Please contact your IT administrator.";
    } else {
        listEl.innerText = "The server could not be reached. Please ensure the backend is running and contact your IT administrator if the problem persists.";
    }
}
/**
 * Automatically repairs dropped trailing punctuation between target and patch texts
 */
function repairTrailingPunctuation(target, redline) {
    const trimmedTarget = target.trim();
    const trimmedRedline = redline.trim();
    if (trimmedTarget.endsWith(".") && !trimmedRedline.endsWith(".")) {
        return redline.trimEnd() + ".";
    }
    if (trimmedTarget.endsWith(";") && !trimmedRedline.endsWith(";")) {
        return redline.trimEnd() + ";";
    }
    if (trimmedTarget.endsWith(":") && !trimmedRedline.endsWith(":")) {
        return redline.trimEnd() + ":";
    }
    return redline;
}
// ============================================================
//  SURGICAL REDLINE HELPERS (Type 1 / Type 2 application engine)
// ============================================================
/**
 * Trims shared leading/trailing words between target and redline so that the
 * tracked change strikes/inserts ONLY the words that actually differ.
 * A human attorney never deletes a word and retypes the same word.
 */
function trimToMinimalDiff(target, redline) {
    if (target == null || redline == null) return { target, redline };
    // PRE-CHECK: pure insertion where the target is a character-level PREFIX or SUFFIX of
    // the redline. This catches edits where the model keeps the original word(s) and ADDS
    // more around them — e.g. "fees" -> "fees, costs, and expenses" or "indemnify" ->
    // "defend, indemnify". The word-level trim below splits on whitespace, so punctuation
    // glued to the first/last word in the redline ("fees," vs "fees") made those tokens
    // unequal and the trim fell through to striking-and-retyping the whole target —
    // re-striking words that never changed. Detecting prefix/suffix at the character level
    // fixes that: strike NOTHING, insert only the added text on the correct side.
    if (redline.length > target.length) {
        if (redline.startsWith(target)) {
            let added = redline.slice(target.length);
            if (added.trim().length > 0)
                return { target: target, redline: redline, pureInsert: true, insertText: added, insertSide: "after" };
        }
        if (redline.endsWith(target)) {
            let added = redline.slice(0, redline.length - target.length);
            if (added.trim().length > 0)
                return { target: target, redline: redline, pureInsert: true, insertText: added, insertSide: "before" };
        }
    }
    let tWords = target.split(/(\s+)/);
    let rWords = redline.split(/(\s+)/);
    let start = 0;
    while (start < tWords.length && start < rWords.length && tWords[start] === rWords[start]) {
        start++;
    }
    let tEnd = tWords.length;
    let rEnd = rWords.length;
    while (tEnd > start && rEnd > start && tWords[tEnd - 1] === rWords[rEnd - 1]) {
        tEnd--;
        rEnd--;
    }
    let newTarget = tWords.slice(start, tEnd).join('');
    let newRedline = rWords.slice(start, rEnd).join('');
    if (newTarget.trim() === '' && newRedline.trim() === '') {
        return null;
    }
    if (newTarget.trim() === '') {
        // The target is fully contained in the redline: this is a pure insertion of
        // words with nothing to strike. Determine whether the new words were added
        // at the FRONT (shared text is the suffix) or the BACK (shared text is the
        // prefix) so the caller inserts them on the correct side.
        let side = (start === 0) ? "before" : "after";
        return { target: target, redline: redline, pureInsert: true, insertText: newRedline, insertSide: side };
    }
    return { target: newTarget, redline: newRedline };
}
/**
 * Builds a consistent margin comment. When applied=true it documents the change
 * that was made; when applied=false it tells the reviewer the rule, the issue,
 * and the suggested language so they can apply it by hand.
 */
function buildReviewerComment(rev, applied) {
    let rule = (rev.rule_reference && rev.rule_reference.trim()) ? rev.rule_reference.trim() : "Playbook rule";
    let issue = (rev.issue_summary && rev.issue_summary.trim()) ? rev.issue_summary.trim() : "";
    let note = (rev.drafting_note && rev.drafting_note.trim()) ? rev.drafting_note.trim() : "";
    if (applied) {
        return `[${rule}] ${note || issue}`;
    }
    let lines = [];
    lines.push(`ACTION NEEDED — ${rule}`);
    if (issue) lines.push(`Issue: ${issue}`);
    if (note) lines.push(`Why: ${note}`);
    if (rev.redlined_text && rev.redlined_text.trim()) {
        lines.push(`Suggested language: "${rev.redlined_text.trim()}"`);
    }
    lines.push(`(The exact original text could not be auto-located, so this change was not applied automatically.)`);
    return lines.join("  ");
}
/**
 * Locates the context_anchor range using a tolerance ladder. Never overwrites anything.
 */
async function findAnchorRange(context, body, anchorRaw) {
    let searchAnchor = getSafeSearchText(anchorRaw);
    if (!searchAnchor) return null;
    let results = body.search(searchAnchor, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) return results.items[0];
    let wildcard = searchAnchor
        .replace(/[''&#x2018;`]/g, "?")
        .replace(/["""]/g, "?")
        .replace(/[\-–—]/g, "?")
        .replace(/[\u00A0\s\t]/g, "?");
    results = body.search(wildcard, { matchCase: false, matchWildcards: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) return results.items[0];
    let broad = searchAnchor.replace(/^\s*([0-9A-Za-z✓❑●▪◦•♣♦♥♠]+\.?\s*)+/, "").trim();
    if (broad.length >= 10) {
        results = body.search(getSafeSearchText(broad), { matchCase: false, matchWholeWord: false, ignoreSpace: true });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) return results.items[0];
    }
    let words = searchAnchor.split(' ');
    if (words.length >= 5) {
        results = body.search(words.slice(0, 5).join(' '), { matchCase: false, matchWholeWord: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) return results.items[0];
    }
    return null;
}
/**
 * Locates the target_text range WITHIN a given anchor range. Never overwrites.
 */
async function findTargetInAnchor(context, anchorItem, targetRaw) {
    let targetSearch = getSafeSearchText(targetRaw);
    if (!targetSearch) return null;
    let results = anchorItem.search(targetSearch, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) return results.items[0];
    let wildcard = targetSearch
        .replace(/[''&#x2018;`]/g, "?")
        .replace(/["""]/g, "?")
        .replace(/[\-–—]/g, "?")
        .replace(/[\u00A0\s\t]/g, "?");
    results = anchorItem.search(wildcard, { matchCase: false, matchWildcards: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) return results.items[0];
    return null;
}
/**
 * Word's search() has a hard ~255-character ceiling and throws on longer queries.
 * This returns a search-safe leading fragment (<= 200 chars, trimmed to a word
 * boundary) that is distinctive enough to locate a long target, plus a flag telling
 * the caller whether the query was truncated so it can expand the range afterward.
 */
function makeSearchableFragment(text) {
    let clean = getSafeSearchText(text, 1000); // normalize whitespace, strip trailing dots
    if (clean.length <= 200) return { query: clean, truncated: false };
    let frag = clean.substring(0, 200);
    let lastSpace = frag.lastIndexOf(' ');
    if (lastSpace > 40) frag = frag.substring(0, lastSpace);
    return { query: frag, truncated: true };
}
/**
 * Runs the tolerance ladder for a query string against any searchable scope
 * (body or range). Returns the loaded results collection (possibly empty).
 */
async function runSearchLadder(context, scope, query) {
    let results = scope.search(query, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) return results;
    let wildcard = query
        .replace(/[''&#x2018;`]/g, "?")
        .replace(/["""]/g, "?")
        .replace(/[\-–—]/g, "?")
        .replace(/[\u00A0\s\t]/g, "?");
    // Wildcard queries also must respect the length ceiling.
    if (wildcard.length <= 255) {
        results = scope.search(wildcard, { matchCase: false, matchWildcards: true });
        results.load("items");
        await context.sync();
    }
    return results;
}
/**
 * Locates the target_text by searching the WHOLE BODY (reliable) rather than inside
 * an anchor sub-range. Handles long targets via a leading fragment, and when several
 * matches exist, disambiguates by choosing the one whose surrounding paragraph best
 * matches the provided context_anchor. Returns a Range (possibly expanded to cover
 * the full target) or null.
 */
async function findTargetInBody(context, body, targetRaw, anchorRaw) {
    let normalizedTarget = getSafeSearchText(targetRaw, 1000);
    if (!normalizedTarget) return null;
    let frag = makeSearchableFragment(targetRaw);
    let results = await runSearchLadder(context, body, frag.query);
    // CROSS-PARAGRAPH FALLBACK: Word's search() cannot match a string that spans a
    // paragraph mark. When a model emits an over-long target_text that runs two separate
    // paragraphs together (e.g. "...Foreground IP. License Grant; Reservation of Rights.
    // Upon Customer's..." where those are distinct paragraphs in the document), BOTH the
    // full target and its 200-char leading fragment straddle the boundary and find nothing.
    // Retry on just the FIRST SENTENCE of the target, which lives within a single paragraph
    // and is a stable anchor; the frag.truncated expansion below then grows the range to
    // cover as much of the full target as exists within that paragraph. This only runs when
    // the normal ladder already failed, so it cannot change any currently-succeeding match.
    if (!results.items || results.items.length === 0) {
        let firstSentence = "";
        let m = normalizedTarget.match(/^.*?[.!?](?:\s|$)/);
        if (m && m[0]) firstSentence = m[0].trim();
        // Only worth trying if the first sentence is meaningfully shorter than the whole
        // target (i.e. the target really is multi-sentence / likely cross-paragraph) and
        // long enough to be reasonably unique.
        if (firstSentence && firstSentence.length >= 12 && firstSentence.length < normalizedTarget.length) {
            let sentenceFrag = makeSearchableFragment(firstSentence);
            let sentenceResults = await runSearchLadder(context, body, sentenceFrag.query);
            if (sentenceResults.items && sentenceResults.items.length > 0) {
                results = sentenceResults;
                // Force the expansion path below to try to grow from this first-sentence
                // match toward the full target within the matched paragraph.
                frag = { query: sentenceFrag.query, truncated: true };
            }
        }
    }
    if (!results.items || results.items.length === 0) return null;
    // Pick the match. With one hit, use it. With several, disambiguate via the anchor.
    let chosen = results.items[0];
    if (results.items.length > 1 && anchorRaw) {
        let anchorKey = getSafeSearchText(anchorRaw, 120).toLowerCase();
        let bestScore = -1;
        for (let item of results.items) {
            let para = item.paragraphs.getFirst();
            context.load(para, 'text');
            await context.sync();
            let ptext = (para.text || "").toLowerCase();
            // crude overlap score: how much of the anchor's leading words appear nearby
            let score = 0;
            anchorKey.split(' ').forEach(w => { if (w.length > 3 && ptext.includes(w)) score++; });
            if (score > bestScore) { bestScore = score; chosen = item; }
        }
    }
    // If we searched on a truncated fragment, expand the chosen range so the
    // replacement covers the FULL target, not just the leading fragment.
    if (frag.truncated) {
        try {
            let para = chosen.paragraphs.getFirst();
            context.load(para, 'text');
            await context.sync();
            let paraText = para.text || "";
            // Find the full target inside the paragraph by normalized comparison.
            let normPara = paraText.replace(/\s+/g, ' ');
            let normFull = normalizedTarget.replace(/\s+/g, ' ');
            if (normPara.includes(normFull)) {
                // Re-search within this paragraph for the longest searchable tail to
                // build a range from fragment start to target end.
                let tail = normFull.substring(frag.query.length).trim();
                if (tail.length > 0) {
                    let tailFrag = makeSearchableFragment(tail);
                    let tailResults = await runSearchLadder(context, para, tailFrag.query);
                    if (tailResults.items && tailResults.items.length > 0) {
                        let expanded = chosen.expandTo(tailResults.items[tailResults.items.length - 1]);
                        expanded.load('text');
                        await context.sync();
                        return expanded;
                    }
                }
            } else {
                // CROSS-PARAGRAPH TARGET: the full target does not fit in the matched
                // paragraph (the model ran two paragraphs together in one target_text).
                // Expand the match to cover the portion of the target that DOES live in
                // this paragraph — from the matched fragment to the end of the paragraph —
                // so we strike/replace the in-paragraph portion instead of dropping the whole
                // edit. The remaining (next-paragraph) portion is handled separately or
                // surfaced by the caller's reviewer comment. We only do this when the text
                // from the match to paragraph-end is genuinely a prefix of the target (i.e.
                // the target really continues past this paragraph), to avoid over-striking.
                let normFrag = frag.query.replace(/\s+/g, ' ').trim();
                let fragPos = normPara.indexOf(normFrag);
                if (fragPos !== -1) {
                    let fromFragToEnd = normPara.substring(fragPos).trim();
                    // Confirm the target starts with what this paragraph holds from the match on.
                    if (fromFragToEnd.length > 0 && normFull.startsWith(fromFragToEnd.substring(0, Math.min(fromFragToEnd.length, 40)))) {
                        let paraEndRange = para.getRange('End');
                        let expanded = chosen.expandTo(paraEndRange);
                        expanded.load('text');
                        await context.sync();
                        console.log('[DIAG-CONT] cross-paragraph target: expanded match to paragraph end (in-paragraph portion of an over-long target_text).');
                        return expanded;
                    }
                }
            }
        } catch (e) {
            // Expansion failed; fall back to the fragment match (still better than nothing).
        }
    }
    return chosen;
}
/**
 * Re-locates the EXACT target string inside an already-matched range/paragraph,
 * returning a range whose boundaries precisely cover the target (including its
 * trailing punctuation). This corrects ranges that a normalized search returned
 * slightly short or long, which would otherwise cause partial-replacement garble.
 * Returns a precise Range or null if the exact target can't be pinned.
 */
async function locateExactRange(context, scopeRange, exactTarget) {
    // Get the paragraph that contains the matched range so we can search within it.
    let para;
    try {
        para = scopeRange.paragraphs.getFirst();
        context.load(para, 'text');
        await context.sync();
    } catch (e) {
        para = scopeRange;
    }
    // Search within the paragraph for the exact target. Keep punctuation intact;
    // only collapse internal whitespace via ignoreSpace so wrapping doesn't break it.
    let q = exactTarget.replace(/\s+/g, ' ').trim();
    // Respect Word's 255-char ceiling; if too long, this path can't pin it exactly.
    if (q.length === 0 || q.length > 255) return null;
    let scope = para || scopeRange;
    let results = scope.search(q, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) {
        // Accept ONLY a range whose text fully covers the target (normalized-equal).
        // A range that is merely a prefix/substring of the target must be rejected:
        // replacing a partial range strands the uncovered tail of the original text
        // (the "...internal commercial below) own internal research..." bug). When no
        // fully-covering range exists, return null so the caller falls to a safe comment.
        for (let r of results.items) {
            context.load(r, 'text');
            await context.sync();
            let got = (r.text || "").replace(/\s+/g, ' ').trim();
            if (got === q) return r;
        }
    }
    return null;
}

/**
 * FIX A — Head+tail range building for LONG targets (> Word's 255-char search ceiling).
 * When a surgical target is too long for locateExactRange to pin in a single search(),
 * this builds the replacement range from two short anchors instead: a unique HEAD
 * fragment (< 255, word-boundary trimmed) and a unique TAIL fragment (< 255), then
 * expands a range from the head match to the tail match. It then RE-VERIFIES that the
 * built range's text equals the full intended target (the same anti-garble discipline as
 * locateExactRange) before returning it. If head or tail can't be located, or the built
 * span does not fully cover the target, it returns null so the caller keeps the SAFE
 * comment fallback — we never replace a partial/mismatched span (that causes garble).
 */
async function locateRangeByHeadTail(context, scopeRange, exactTarget) {
    let norm = exactTarget.replace(/\s+/g, ' ').trim();
    if (norm.length <= 255) return null; // short targets are locateExactRange's job
    // Resolve the containing paragraph as the search scope.
    let para;
    try {
        para = scopeRange.paragraphs.getFirst();
        context.load(para, 'text');
        await context.sync();
    } catch (e) {
        para = scopeRange;
    }
    let scope = para || scopeRange;

    // Build a HEAD fragment: a word-boundary prefix <= 200 chars (leaves room under 255).
    const MAXFRAG = 200;
    let head = norm.slice(0, MAXFRAG);
    let hsp = head.lastIndexOf(' ');
    if (hsp > 40) head = head.slice(0, hsp);
    // Build a TAIL fragment: a word-boundary suffix <= 200 chars.
    let tail = norm.slice(norm.length - MAXFRAG);
    let tsp = tail.indexOf(' ');
    if (tsp > -1 && tsp < MAXFRAG - 40) tail = tail.slice(tsp + 1);
    if (!head || !tail || head.length === 0 || tail.length === 0) return null;
    if (head.length > 255 || tail.length > 255) return null;

    try {
        // Locate the head fragment within the paragraph.
        let headResults = scope.search(head, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
        headResults.load("items");
        await context.sync();
        if (!headResults.items || headResults.items.length === 0) return null;

        // Locate the tail fragment within the paragraph.
        let tailResults = scope.search(tail, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
        tailResults.load("items");
        await context.sync();
        if (!tailResults.items || tailResults.items.length === 0) return null;

        // Use the FIRST head match and the LAST tail match, then expand a range that
        // covers from the head's start to the tail's end.
        let headRange = headResults.items[0];
        let tailRange = tailResults.items[tailResults.items.length - 1];
        let built;
        try {
            built = headRange.expandTo(tailRange);
        } catch (e) {
            return null; // expandTo can throw if tail precedes head; bail to safe fallback.
        }
        context.load(built, 'text');
        await context.sync();

        // ANTI-GARBLE GUARD: the built span MUST equal the full target (normalized).
        // If it doesn't fully cover the target, reject so the caller keeps the comment.
        let got = (built.text || "").replace(/\s+/g, ' ').trim();
        if (got === norm) return built;
        return null;
    } catch (e) {
        return null; // any search/expand failure -> safe fallback.
    }
}
/**
 * Inspects existing section headings in the document to learn how this contract
 * formats them, so inserted new-section headings can MATCH the original. Returns a
 * descriptor: { bold, italic, underline, allCaps, smallCaps, delimiter, found }.
 * It samples real heading runs (the leading formatted run of paragraphs that look
 * like section headings) and records the formatting they share plus the delimiter
 * character their label ends with (":", ".", or "-"). This is deterministic — the
 * document itself defines the template, so we never guess typography.
 */
async function detectHeadingStyle(context, body) {
    let style = { bold: false, italic: false, underline: false, allCaps: false, smallCaps: false, delimiter: ":", headingOnOwnLine: false, nativeList: false, found: false };
    try {
        let paras = body.paragraphs;
        paras.load("items");
        await context.sync();
        // PHASE 1 (text-first, reliable): load text for the first ~50 paragraphs and
        // identify heading candidates purely from their text — a short leading label
        // (<= 8 words) terminated by ":", "." or "-". From these we learn the case
        // pattern (ALL-CAPS?) and the delimiter, which is what actually failed before.
        let limit = Math.min(paras.items.length, 50);
        let pool = [];
        for (let i = 0; i < limit; i++) {
            let p = paras.items[i];
            p.load("text");
            pool.push(p);
        }
        await context.sync();
        let candidates = []; // { para, label, delim }
        let delimVotes = { ":": 0, ".": 0, "-": 0, "none": 0 };
        let allCapsVotes = 0;
        let ownLineVotes = 0;   // heading text is short and stands alone on its paragraph (body follows on next paragraph)
        for (let p of pool) {
            let txt = (p.text || "").trim();
            if (!txt) continue;
            // A "heading on its own line" looks like a short line (<= 8 words) that is the
            // WHOLE paragraph, optionally led by a section marker, with NO sentence body after.
            // e.g. "2. Intellectual Property"  (bold, no colon, body in the NEXT paragraph).
            // Detect this BEFORE the delimiter match so a colon-less heading isn't missed.
            let ownLineMatch = txt.match(/^\s*([0-9]+|[A-Za-z]|[IVXLCivxlc]+)?\s*[.)]?\s*([A-Z][A-Za-z &/'-]{1,60})$/);
            if (ownLineMatch && txt.split(/\s+/).length <= 8 && !/[.:;]$/.test(txt)) {
                candidates.push({ para: p, label: ownLineMatch[2].trim(), delim: "none" });
                delimVotes["none"]++;
                ownLineVotes++;
                let lbl = ownLineMatch[2].trim();
                if (lbl === lbl.toUpperCase() && /[A-Z]/.test(lbl)) allCapsVotes++;
                continue;
            }
            let m = txt.slice(0, 90).match(/^([^.:\-]{2,80}?)\s*([:.\-])/);
            if (!m) continue;
            let label = m[1].trim();
            let delim = m[2];
            if (label.split(/\s+/).length > 8) continue;
            if (!/[A-Za-z]/.test(label)) continue;
            // Skip obvious non-headings (lines that are clearly fill-in or contacts).
            if (/^(name|title|address|e-?mail|date|by)\b/i.test(label)) continue;
            candidates.push({ para: p, label, delim });
            delimVotes[delim] = (delimVotes[delim] || 0) + 1;
            // Text-based ALL CAPS: the label's letters are all uppercase.
            if (label === label.toUpperCase() && /[A-Z]/.test(label)) allCapsVotes++;
        }
        if (candidates.length >= 2) {
            style.found = true;
            let n = candidates.length;
            let maj = Math.ceil(n / 2);
            style.allCaps = allCapsVotes >= maj;
            style.headingOnOwnLine = ownLineVotes >= maj;
            let best = ":", bestN = -1;
            for (let d of [":", ".", "-", "none"]) { if ((delimVotes[d] || 0) > bestN) { bestN = delimVotes[d]; best = d; } }
            style.delimiter = best;
            // PHASE 2 (font, best-effort enhancement): sample the heading-label runs of a
            // few candidates to detect bold/italic/underline/small-caps. If this fails or
            // is inconclusive, we still have the reliable text-based case + delimiter.
            try {
                let sample = candidates.slice(0, Math.min(5, candidates.length));
                let fonts = [];
                for (let c of sample) {
                    let f = c.para.getRange().font;
                    f.load(["bold", "italic", "underline", "smallCaps"]);
                    fonts.push(f);
                }
                await context.sync();
                let fb = 0, fi = 0, fu = 0, fsc = 0, fc = fonts.length;
                for (let f of fonts) {
                    if (f.bold) fb++;
                    if (f.italic) fi++;
                    if (f.underline && f.underline !== "None" && f.underline !== false) fu++;
                    if (f.smallCaps) fsc++;
                }
                let fmaj = Math.ceil(fc / 2);
                style.bold = fb >= fmaj;
                style.italic = fi >= fmaj;
                style.underline = fu >= fmaj;
                style.smallCaps = fsc >= fmaj;
            } catch (e) { /* font detection best-effort; text-based style stands */ }
            // PHASE 3 (native-list detection): determine whether the contract's section
            // numbers are NATIVE Word list numbering (Word auto-generates & renumbers) or
            // just PLAIN TYPED text. We sample the heading candidates' paragraphs and ask
            // Word for their listItem. If a majority belong to a real list, the document
            // is native-numbered and Word will assign/renumber for us on insert. If not,
            // the numbers are plain text and we must NOT guess numbers (insert unnumbered,
            // flag for the human). This is the key branch behind the numbering strategy.
            try {
                let sample = candidates.slice(0, Math.min(6, candidates.length));
                let listFlags = [];
                for (let c of sample) {
                    let li = c.para.listItemOrNullObject;
                    li.load("level");
                    listFlags.push(li);
                }
                await context.sync();
                let nativeCount = 0;
                for (let li of listFlags) {
                    if (!li.isNullObject) nativeCount++;
                }
                style.nativeList = nativeCount >= Math.ceil(listFlags.length / 2);
            } catch (e) { /* if listItem probing fails, assume plain text (safer) */ style.nativeList = false; }
        }
    } catch (e) {
        // Detection is best-effort; fall back to safe defaults (caps + colon, plain text).
        style = { bold: false, italic: false, underline: false, allCaps: true, smallCaps: false, delimiter: ":", headingOnOwnLine: false, nativeList: false, found: false };
    }
    try {
        console.log('[DIAG-HEADING] detected style => ' + JSON.stringify({
            found: style.found, delimiter: style.delimiter, headingOnOwnLine: style.headingOnOwnLine,
            bold: style.bold, allCaps: style.allCaps, nativeList: style.nativeList
        }));
    } catch (e) {}
    return style;
}
/**
 * Applies the detected heading style to a heading run (a Word Range covering just
 * the heading label). Sets bold/italic/underline/smallCaps to match the document.
 * Caps are handled by the caller (it uppercases the heading text before insertion).
 */
function applyHeadingFormatting(headingRange, style) {
    try {
        if (style.bold) headingRange.font.bold = true;
        if (style.italic) headingRange.font.italic = true;
        if (style.underline) headingRange.font.underline = "Single";
        if (style.smallCaps) headingRange.font.smallCaps = true;
    } catch (e) { /* formatting best-effort */ }
}
/**
 * Inserts continuation text into a paragraph at a grammatically clean point.
 * Strategy (per design):
 *  1. If the model provided `afterSentence` (a verbatim sentence the new text should
 *     follow), locate that sentence WITH its terminal punctuation inside the paragraph
 *     and insert immediately AFTER it.
 *  2. Otherwise (or if that sentence can't be cleanly located), fall back to appending
 *     at the true END of the paragraph, after its final terminal punctuation.
 * In all cases we guarantee a single clean separator: the prior sentence keeps exactly
 * one terminal period, a single space follows, then the new text (which itself ends in
 * terminal punctuation). This prevents both the "two periods at the end / none in the
 * middle" bug and mid-sentence splits.
 */
/**
 * Normalizes TEXT THAT WILL BE WRITTEN INTO THE DOCUMENT (inserted clause bodies and
 * surgical replacements). Unlike getSafeSearchText (which normalizes a SEARCH query),
 * this fixes characters in the OUTPUT so model artifacts don't reach the contract:
 *   - "Company' s" (stray space in a possessive) -> "Company's"
 *   - smart quotes left as-is (we keep them in output; only repair the broken possessive)
 * We deliberately do NOT collapse all whitespace here (output may legitimately contain it).
 */
function normalizeInsertText(text) {
    if (!text) return text;
    let t = text;
    // Repair a possessive/contraction broken by a stray space: "Company' s" -> "Company's".
    t = t.replace(/([A-Za-z])([\u2018\u2019'])\s+([sStT])\b/g, "$1$2$3");
    return t;
}
/**
 * Detects whether a paragraph is a SECTION HEADING (e.g., "6. Limitation of Liability"
 * — a short line, no sentence-ending punctuation, optionally led by a section marker).
 * Continuation insertions must NEVER append into a heading paragraph (that buries the
 * new clause inside the title — the "6. Limitation of Liability. Provided that..." bug).
 * When the anchor lands on a heading, the caller should redirect to the BODY paragraph
 * that follows it. Returns true if the paragraph text looks like a heading.
 */
function looksLikeHeadingText(txt) {
    let t = (txt || "").trim();
    if (!t) return false;
    if (t.split(/\s+/).length > 9) return false;        // headings are short
    if (/[.;:]$/.test(t)) return false;                  // headings don't end in sentence punctuation
    // Optional leading marker (number/letter/Roman) then a short Title-ish label.
    return /^\s*((\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+)?[A-Z][A-Za-z &/'-]{2,}$/.test(t);
}
/**
 * Strip a leading section heading that the model jammed onto the FRONT of a target_text.
 * Some model runs (observed on gemini-3.1) emit a strike target like:
 *   "5. Liquidated DamagesIn the event the Company fails to deliver..."
 *   "7. PricingThe Company represents and warrants..."
 * i.e. the section number + Title-Case heading run together with the body, with NO space
 * or newline at the seam. That literal string does not exist in the document (the heading
 * and body are separate paragraphs), so the search fails and the strike is silently lost.
 *
 * This detects a leading "<marker>. <Title Case Words>" immediately followed by a capital
 * letter that begins the body sentence (the tell-tale missing separator, e.g. "Damages" +
 * "In"), and returns the body portion alone so it can be located and struck. If no such
 * jammed heading is present, the input is returned unchanged.
 *
 * Conservative: it only strips when (a) there is a section marker like "5." or "B)" at the
 * very start, (b) the words after it are short Title-Case heading words, and (c) those are
 * butted directly against a capital letter with no separating space — the exact signature of
 * the bug. Normal targets (which begin mid-sentence or with their own first word) are left
 * untouched.
 */
// Word's search() cannot match text across a paragraph break. When a target spans
// paragraphs (contains a newline) — e.g. a "Section X  Heading\nBody..." strike — return
// the longest single-paragraph segment so the search runs within one paragraph. Leaves
// single-paragraph targets unchanged.
function largestSingleParagraphSegment(targetRaw) {
    if (!targetRaw || String(targetRaw).indexOf('\n') === -1) return targetRaw;
    let segments = String(targetRaw).split(/\n+/).map(s => s.trim()).filter(s => s.length > 0);
    if (segments.length === 0) return targetRaw;
    let longest = segments.reduce((a, b) => (b.length > a.length ? b : a), segments[0]);
    if (longest.length < 20) return targetRaw;
    return longest;
}

function stripLeadingSectionHeading(targetRaw) {
    if (!targetRaw) return targetRaw;
    let t = String(targetRaw);
    // Marker: digits / single letter / Roman numerals, then . or ) then a space.
    // Then 1..6 Title-Case heading words, the LAST of which is butted against a capital
    // (lowercase-to-uppercase or letter-to-uppercase with no space) that starts the body.
    // We capture the heading and find the seam where a lowercase letter is immediately
    // followed by an uppercase letter (e.g. "Damages|In", "Pricing|The").
    let m = t.match(/^\s*((?:\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+)([A-Z][A-Za-z]*(?:\s+(?:&|and|of|the|for|to|[A-Z][A-Za-z]*)){0,5})/);
    if (!m) return targetRaw;
    let afterMarker = t.slice(m[1].length); // heading + jammed body, no leading marker
    // Find the seam: a lowercase letter directly followed by an uppercase letter, which is
    // where "Damages" meets "In". Only treat it as a jam if that seam exists within the
    // first ~60 chars (headings are short) AND there is no space at the seam.
    let seam = afterMarker.search(/[a-z][A-Z]/);
    if (seam === -1 || seam > 60) return targetRaw;
    // The body begins at seam+1 (the uppercase letter). Everything before is the heading.
    let heading = afterMarker.slice(0, seam + 1);
    let bodyPart = afterMarker.slice(seam + 1);
    // Safety checks: the heading part must actually look like a heading (short, Title-ish),
    // and the body part must be a real sentence-length span (not a stray cap). Otherwise we
    // risk corrupting a legitimate mid-word capital (rare, but be safe).
    if (!looksLikeHeadingText(heading)) return targetRaw;
    if (!bodyPart || bodyPart.trim().length < 20) return targetRaw;
    console.log('[DIAG-STRIPHEAD] removed jammed heading ' + JSON.stringify(m[1] + heading) +
        ' from strike target; searching on body only.');
    return bodyPart.trim();
}
/**
 * Given the paragraph an anchor matched, return the paragraph we should actually
 * continue INTO. If the matched paragraph is a heading, return the next paragraph
 * (the section body). Otherwise return the matched paragraph unchanged.
 */
async function resolveContinuationTarget(context, matchedPara, body) {
    try {
        matchedPara.load("text");
        await context.sync();
        if (!looksLikeHeadingText(matchedPara.text)) return matchedPara;
        // It's a heading. Find the next paragraph (the body) to continue into.
        let paras = body.paragraphs;
        paras.load("items");
        await context.sync();
        // Load text for all items so we can locate matchedPara and step to the next non-empty.
        for (let p of paras.items) p.load("text");
        await context.sync();
        let matchedText = (matchedPara.text || "").trim();
        for (let i = 0; i < paras.items.length; i++) {
            if ((paras.items[i].text || "").trim() === matchedText) {
                // Walk forward to the first non-empty paragraph after the heading.
                for (let j = i + 1; j < paras.items.length; j++) {
                    if ((paras.items[j].text || "").trim()) {
                        console.log('[DIAG-CONT] anchor was a heading; redirecting continuation to body paragraph.');
                        return paras.items[j];
                    }
                }
                break;
            }
        }
    } catch (e) { /* best effort; fall back to the matched paragraph */ }
    return matchedPara;
}
// Detects a heading label embedded at the start of insertion body text ("Survival. The
// provisions..." or "Order of Precedence: ..."). Returns { heading, body } so the clause
// can be promoted to its own subparagraph, else null. A heading is a short Title-Case or
// ALLCAPS phrase (<= 6 words, <= 45 chars) before a . or : followed by the body's first
// capital. The opener guard and case requirement prevent splitting real sentences that
// merely begin with a capitalized word ("Notwithstanding...", "Customer shall...").
function splitEmbeddedHeading(text) {
    if (!text) return null;
    let t = String(text).trim().replace(/^\s*[.;,]\s*/, '').trim();
    let m = t.match(/^([A-Z][A-Za-z]*(?:\s+(?:[A-Z][A-Za-z]*|and|or|of|the|to|&)){0,5})[.:]\s+(?=[A-Z])/);
    if (!m) return null;
    let heading = m[1].trim();
    if (heading.length > 45) return null;
    let sig = heading.split(/\s+/).filter(w => !/^(and|or|of|the|to|&)$/i.test(w));
    let titleOrCaps = sig.every(w => /^[A-Z][a-z]+$/.test(w) || /^[A-Z]{2,}$/.test(w));
    if (!titleOrCaps) return null;
    const openers = new Set(["Notwithstanding","The","Each","Customer","Contractor","Provider","Vendor","Company","Either","Neither","This","All","Any","No","Upon","During","Except","If","In","Subject","Unless"]);
    if (sig.length === 1 && openers.has(sig[0])) return null;
    let body = t.slice(m[0].length).trim();
    if (body.length < 20) return null;
    return { heading, body };
}


async function insertContinuation(context, paragraph, bodyText, afterSentence) {
    // Ensure the new text ends with terminal punctuation and does NOT begin with stray
    // joining punctuation (a leading ". "/"; " from the model would double up at the seam).
    let clean = bodyText.trim().replace(/^\s*[.;,]\s*/, '').trim();
    if (!/[.!?]$/.test(clean)) clean += ".";
    paragraph.load("text");
    await context.sync();
    let paraText = paragraph.text || "";
    // --- Attempt 1: insert after a specific sentence the model named. ---
    if (afterSentence && afterSentence.trim()) {
        let needle = afterSentence.trim();
        // Search for the sentence within the paragraph; accept it only if found with
        // its terminal punctuation so we insert AFTER the period, never before it.
        try {
            let results = paragraph.search(getSafeSearchText(needle), { matchCase: false, matchWholeWord: false, ignoreSpace: true });
            results.load("items");
            await context.sync();
            if (results.items.length > 0) {
                let sentRange = results.items[0];
                sentRange.load("text");
                await context.sync();
                let got = (sentRange.text || "");
                // If the matched range already ends in terminal punctuation, just add " " + text.
                // If not, add the missing period once. Never add a period when one is already there.
                let needsPeriod = !/[.!?:;]$/.test(got.trim());
                let insertText = (needsPeriod ? ". " : " ") + clean;
                sentRange.insertText(insertText, Word.InsertLocation.after);
                await collapseDoubledPunctuation(context, paragraph);
                return true;
            }
        } catch (e) { /* fall through to end-of-paragraph */ }
    }
    // --- Attempt 2: append at the true end of the paragraph. ---
    // The paragraph's text reliably ends with its final punctuation; check it and add
    // exactly one separator. Because `clean` no longer carries a leading period, the only
    // period at the seam is the existing one (if present) or the single one we add here.
    let trimmed = paraText.replace(/\s+$/, "");
    let endsWithTerminal = /[.!?:;]$/.test(trimmed);
    let insertText = (endsWithTerminal ? " " : ". ") + clean;
    paragraph.getRange().insertText(insertText, Word.InsertLocation.end);
    await collapseDoubledPunctuation(context, paragraph);
    return true;
}

/**
 * Defensive cleanup: after a continuation insert, a doubled sentence mark can appear at
 * the seam or tail (e.g. "...notice method.." ) when the model's text and the document's
 * own punctuation meet in a way the pre-insert string logic didn't catch. Word ranges are
 * fiddly and this can't be reproduced reliably outside Word, so we repair it in the live
 * document: find any ".." (or ".." spanning a space) in the paragraph and collapse to one.
 * Conservative — only collapses repeated PERIODS, never touches "..." ellipses authored
 * intentionally is unlikely in contracts, but we still guard by only fixing exactly two.
 */
async function collapseDoubledPunctuation(context, paragraph) {
    try {
        // Search for exactly two consecutive periods (optionally separated by a space).
        let res = paragraph.search("..", { matchCase: false, matchWholeWord: false });
        res.load("items");
        await context.sync();
        if (res.items && res.items.length > 0) {
            for (let i = 0; i < res.items.length; i++) {
                let r = res.items[i];
                r.load("text");
            }
            await context.sync();
            for (let i = 0; i < res.items.length; i++) {
                let r = res.items[i];
                let t = (r.text || "");
                // Only collapse a plain double period "..", not an ellipsis "...".
                if (t === "..") {
                    r.insertText(".", Word.InsertLocation.replace);
                }
            }
            await context.sync();
        }
    } catch (e) { /* best effort; non-fatal */ }
}
async function loadPlaybooks() {
    const select = document.getElementById("playbook-select");
    select.innerHTML = '<option value="">Loading corporate tracks...</option>';
    try {
        const response = await fetch(`${BACKEND_URL}/api/playbooks`);
        const data = await response.json();
        select.innerHTML = "";
        if (data.playbooks && data.playbooks.length > 0) {
            data.playbooks.forEach(playbook => {
                const option = document.createElement("option");
                option.value = playbook; option.text = playbook;
                select.appendChild(option);
            });
        } else {
            select.innerHTML = '<option value="">No playbooks discovered inside directory target</option>';
        }
    } catch (error) {
        select.innerHTML = '<option value="">Network core validation lost</option>';
    }
}
async function handleEngineToggle() {
    const provider = document.getElementById("engine-toggle").value;
    const note = document.getElementById("engine-config-note");
    const statusDiv = document.getElementById("review-status");
    // Both engine options are always visible. If the user selects one that IT has not
    // configured in .env, show a "contact IT" note; the run is blocked in reviewContract().
    let configured = (provider === "cloud") ? CONFIG_STATUS.cloudConfigured : CONFIG_STATUS.localConfigured;
    if (!configured) {
        let label = (provider === "cloud") ? "Cloud" : "Local";
        note.style.display = "block";
        note.innerText = `The ${label} engine has not been configured. Please contact your IT administrator to enable it.`;
    } else {
        note.style.display = "none";
        if (statusDiv) {
            statusDiv.className = "status-message info";
            statusDiv.innerText = `Engine set to ${provider === "cloud" ? "Cloud Engine" : "Local Secure Engine (Ollama)"}.`;
        }
    }
}
// NOTE: commitRuntimeSettings() has been removed. The user's engine and model choice
// is no longer pushed to the server; it is kept locally and sent with each redline
// request so concurrent users never overwrite a shared server-side setting.
function getSafeSearchText(text, maxLength = 250) {
    if (!text) return "";
    let cleanText = text.replace(/[\r\n\t\u00A0\s]+/g, ' ').trim();
    // --- Quote/apostrophe normalization (helps every local model; never changes WHERE a match lands) ---
    // Local models (e.g. Maverick) frequently mangle possessives: they emit smart quotes
    // (U+2018/U+2019/U+201C/U+201D) where the document uses straight quotes, and they
    // sometimes split a possessive with a stray space ("Company' s" instead of "Company's").
    // We fold all single-quote glyphs to a straight apostrophe and all double-quote glyphs
    // to a straight double quote, then repair the "<word>' s" / "<word>' S" split so the
    // search query matches the document's actual characters. This is purely a normalization
    // of the SEARCH STRING; it cannot cause a match in the wrong location.
    cleanText = cleanText
        .replace(/[\u2018\u2019\u201B\u0060\u00B4]/g, "'")   // ‘ ’ ‛ ` ´  -> '
        .replace(/[\u201C\u201D\u201E]/g, '"');               // “ ” „      -> "
    // Repair a possessive/contraction broken by a stray space: "Company' s" -> "Company's".
    cleanText = cleanText.replace(/([A-Za-z])'\s+([sStT])\b/g, "$1'$2");
    cleanText = cleanText.replace(/[\.…]+$/, '').trim();
    if (cleanText.length <= maxLength) return cleanText;
    let truncated = cleanText.substring(0, maxLength);
    let lastSpaceIndex = truncated.lastIndexOf(' ');
    return lastSpaceIndex > 0 ? truncated.substring(0, lastSpaceIndex) : truncated;
}
async function reviewContract() {
    const playbook = document.getElementById("playbook-select").value;
    const statusDiv = document.getElementById("review-status");
    const btn = document.getElementById("btn-review");
    const provider = document.getElementById("engine-toggle").value;
    // The license is validated server-side from .env; the client no longer holds a key.
    const savedKey = "ENV";
    if (!playbook) {
        statusDiv.className = "status-message error";
        statusDiv.innerText = "Please choose an active validation playbook profile.";
        return;
    }
    // Block the run if the selected engine isn't configured (the toggle still lets the user
    // pick it, but it can't run until IT configures it in .env).
    let engineReady = (provider === "cloud") ? CONFIG_STATUS.cloudConfigured : CONFIG_STATUS.localConfigured;
    if (!engineReady) {
        let label = (provider === "cloud") ? "Cloud" : "Local";
        statusDiv.className = "status-message error";
        statusDiv.innerText = `The ${label} engine has not been configured. Please contact your IT administrator to enable it.`;
        return;
    }
    btn.disabled = true;
    statusDiv.className = "status-message info";
    statusDiv.innerText = "Extracting underlying text tracking parameters...";
    let eventSource;
    let clientTimeoutClock;
    let processingAbortedByChoice = false;
    try {
        let rawText = "";
        await Word.run(async (context) => {
            const documentBody = context.document.body;
            context.load(documentBody, 'text');
            await context.sync();
            rawText = documentBody.text;
        });
        if (!rawText || rawText.trim() === "") throw new Error("Target document window contains no usable text.");
        const base64Text = btoa(unescape(encodeURIComponent(rawText)));
        // Securely capture client architecture metadata for centralized corporate audit trails
        let userEmail = "UNKNOWN_USER";
        if (Office.context.userProfile && Office.context.userProfile.email) {
            userEmail = Office.context.userProfile.email;
        }
        let contractFileName = "UNSAVED_DOCUMENT";
        try {
            const fileProps = await new Promise((resolve) => {
                Office.context.document.getFilePropertiesAsync((asyncResult) => {
                    resolve(asyncResult);
                });
            });
            if (fileProps && fileProps.status === Office.AsyncResultStatus.Succeeded && fileProps.value && fileProps.value.url) {
                let pathString = fileProps.value.url;
                contractFileName = pathString.substring(Math.max(pathString.lastIndexOf('/'), pathString.lastIndexOf('\\')) + 1);
                if (!contractFileName || contractFileName.trim() === "") {
                    contractFileName = "UNSAVED_DOCUMENT";
                }
            }
        } catch (metaErr) {
            console.warn("Unable to extract document workspace file properties:", metaErr);
        }
        // Connect to Server-Sent Events to actively stream the Chain of Thought updates
        eventSource = new EventSource(`${BACKEND_URL}/api/status?sessionId=${SESSION_ID}`);
        eventSource.onmessage = function(event) {
            if (processingAbortedByChoice) return;
            const data = JSON.parse(event.data);
            if (statusDiv) {
                statusDiv.innerText = data.message;
            }
        };
        // --- INTERACTIVE 10-MINUTE CHECKPOINT TIMER ---
        clientTimeoutClock = setTimeout(() => {
            if (eventSource) eventSource.close();
            statusDiv.className = "status-message info";
            statusDiv.innerHTML = `
                <div style="font-weight:600; margin-bottom:8px; color:#1e3a8a;">Analysis is running heavy (exceeded 10 mins).</div>
                <div style="font-size:0.95em; margin-bottom:10px; color:#1e40af;">The local AI model is executing its core reasoning checklist passes. Keep waiting?</div>
                <div style="display:flex; gap:8px;">
                    <button id="btn-extend-yes" style="margin:0; padding:6px 10px; background:#1e3a8a; font-size:0.85em; width:auto; display:inline-block;">Keep Waiting</button>
                    <button id="btn-extend-no" style="margin:0; padding:6px 10px; background:#b91c1c; font-size:0.85em; width:auto; display:inline-block;">Cancel Run</button>
                </div>
            `;
            document.getElementById("btn-extend-yes").onclick = () => {
                statusDiv.className = "status-message info";
                statusDiv.innerText = "Resuming session window parameters. Awaiting final engine payloads...";
                eventSource = new EventSource(`${BACKEND_URL}/api/status?sessionId=${SESSION_ID}`);
                eventSource.onmessage = function(event) {
                    const data = JSON.parse(event.data);
                    statusDiv.innerText = data.message;
                };
            };
            document.getElementById("btn-extend-no").onclick = () => {
                processingAbortedByChoice = true;
                if (eventSource) eventSource.close();
                statusDiv.className = "status-message error";
                statusDiv.innerText = "Operation aborted by user request.";
                btn.disabled = false;
            };
        }, 600000);
        // Read the per-review mode selector (Extended / Standard). Default = extended
        // (most thorough) when the control is absent. We also derive the legacy twoPass flag
        // so an older server still behaves sensibly (both current modes => two-pass on).
        let modeEl = document.getElementById("review-mode-select");
        let reviewMode = modeEl ? (modeEl.value || "extended") : "extended";
        if (reviewMode !== "extended" && reviewMode !== "standard") reviewMode = "extended";
        let twoPass = true;
        const response = await fetch(`${BACKEND_URL}/api/redline`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                playbookFolderName: playbook,
                documentBuffer: base64Text,
                clientKey: savedKey,
                userEmail: userEmail,
                contractFileName: contractFileName,
                twoPass: twoPass,
                reviewMode: reviewMode,
                provider: provider,
                sessionId: SESSION_ID
            })
        });
        clearTimeout(clientTimeoutClock);
        if (eventSource) eventSource.close();
        if (processingAbortedByChoice) return;
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || "Processing failed across gateway routing layers.");
        }
        const result = await response.json();
        statusDiv.innerText = "Processing completed. Commencing high-accuracy tracking adjustments...";
        if (result.revisions && result.revisions.length > 0) {
            await Word.run(async (context) => {
                const body = context.document.body;
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
                await context.sync();
                let firstParagraph = body.paragraphs.getFirst();
                // Human-readable label for the review depth, surfaced in the disclaimer so the
                // reviewer (and anyone auditing the document) can see which mode produced it.
                let modeLabel = (result.reviewMode === "standard") ? "Standard Review (definitions + rule-by-rule checklist + tracked changes edits)"
                    : "Extended Review (definitions + rule-by-rule checklist + tracked changes edits + completeness check)";
                firstParagraph.getRange().insertComment("DISCLAIMER: The tracked changes and comments identified below were prepared by generative AI using the \"" + playbook + "\" playbook, in " + modeLabel + ". All suggestions must be verified.");
                // Learn how THIS contract formats section headings, so any new-section
                // insertions can match the document's own style (all-caps/bold/underline/
                // small-caps + the right delimiter). Detected once, reused for all inserts.
                let headingStyle = await detectHeadingStyle(context, body);
                await context.sync();
                let successCount = 0;
                let failCount = 0;
                let commentOnlyCount = 0;
                let reservedCount = 0;
                for (let rev of result.revisions) {
                    // Normalize edit_type; default to surgical when target_text is present, else insertion.
                    let editType = (rev.edit_type || "").toLowerCase().trim();
                    if (editType !== "surgical" && editType !== "insertion" && editType !== "comment") {
                        editType = (rev.target_text && rev.target_text.trim() !== "") ? "surgical" : "insertion";
                    }
                    // EMPTY-TARGET SURGICAL GUARD (matcher-side safety net):
                    // A "surgical" edit MUST have a real target_text to strike/replace. When a
                    // local model mislabels additive language as "surgical" but leaves target_text
                    // empty (observed on Maverick's liability-cap proviso, contract-damages cap, and
                    // notice-email additions), the only coherent interpretation is an insertion: there
                    // is nothing to strike, only new words to add. Rather than discard it as an
                    // incomplete surgical edit (losing the suggestion), reclassify it as an insertion
                    // and route its redlined_text into body_text so the insertion path can place it.
                    // This NEVER strikes existing words, so it cannot corrupt the document; at worst a
                    // human moves an added clause. The prompt-side TARGET_TEXT TEST should make this
                    // rare; this guard catches the stragglers.
                    if (editType === "surgical" && (!rev.target_text || rev.target_text.trim() === "")) {
                        let salvageBody = (rev.body_text && rev.body_text.trim())
                            ? rev.body_text.trim()
                            : (rev.redlined_text || "").trim();
                        if (salvageBody) {
                            console.warn('[DIAG-RECLASS] empty-target surgical reclassified as insertion. rule_reference=' + JSON.stringify(rev.rule_reference) + ' | redlined_text=' + JSON.stringify(rev.redlined_text));
                            editType = "insertion";
                            rev.edit_type = "insertion";
                            if (!rev.body_text || !rev.body_text.trim()) rev.body_text = salvageBody;
                            // The model often emits these as a sentence FRAGMENT meant to flow into a
                            // prior sentence: a leading "; ", ". ", or ", ", and/or a connector word
                            // like "provided that ..." or "and ...". Inserted as-is after a period,
                            // that reads as a lowercase exception to the preceding sentence ("...this
                            // Agreement. provided that...") — which a lawyer flagged as substantively
                            // wrong (it looks like an exception to the consequential-damages waiver).
                            // Convert the fragment into a clean STANDALONE sentence: strip the leading
                            // joining punctuation, drop a leading "and"/"; provided"/"provided that"
                            // connector where it would dangle, and capitalize the first letter.
                            let b = rev.body_text.replace(/^\s*[;.,]\s*/, '').trim();
                            // Drop a dangling leading conjunction that only made sense mid-sentence.
                            b = b.replace(/^(and|but|or)\s+/i, '');
                            // "provided that"/"provided, however," -> "Provided that" reads as an
                            // exception; as a standalone cap sentence we prefer a clean declaration.
                            // Capitalize the first alphabetical character.
                            b = b.replace(/^([a-z])/, (m, c) => c.toUpperCase());
                            // Ensure terminal punctuation.
                            if (!/[.!?]$/.test(b)) b += ".";
                            rev.body_text = b;
                            // No heading provided => continuation; the insertion path will infer style.
                        }
                        // If there is nothing to salvage, fall through; the surgical guard below
                        // will record it as a failed/commented revision.
                    }
                    // --------------------------------------------------------
                    // TYPE 3 — COMMENT-ONLY: consistency findings (second pass) and
                    // ambiguous clauses flagged for human judgment. Never edits the
                    // document; drops an anchored comment at the location of the issue.
                    // --------------------------------------------------------
                    if (editType === "comment") {
                        // A comment with no context_anchor anchors at the top of the document
                        // (first paragraph) — e.g. the "checklist step did not complete" review
                        // note. A comment WITH an anchor is placed at that anchor.
                        let host = (rev.context_anchor && rev.context_anchor.trim() !== "")
                            ? await findAnchorRange(context, body, rev.context_anchor)
                            : null;
                        if (!host) { host = body.paragraphs.getFirst().getRange(); }
                        let rule = (rev.rule_reference && rev.rule_reference.trim()) ? rev.rule_reference.trim() : "Review";
                        let issue = (rev.issue_summary && rev.issue_summary.trim()) ? rev.issue_summary.trim() : (rev.drafting_note || "");
                        // Second-pass (consistency) comments lead with "SECOND PASS:" so reviewers
                        // can immediately distinguish them from first-pass redline comments, e.g.
                        // when a consistency note appears to conflict with an applied edit. The rule
                        // category is retained after the prefix.
                        let commentText = rev.second_pass === true
                            ? `SECOND PASS: [${rule}] ${issue}`
                            : `[${rule}] ${issue}`;
                        try {
                            host.insertComment(commentText);
                            commentOnlyCount++;
                            await context.sync();
                        } catch (e) { /* best effort */ }
                        continue;
                    }
                    // --------------------------------------------------------
                    // TYPE 2 — INSERTION: pure addition, no strikethrough.
                    // --------------------------------------------------------
                    if (editType === "insertion") {
                        let insertionStyle = (rev.insertion_style || "").toLowerCase().trim();
                        if (insertionStyle !== "continuation" && insertionStyle !== "new_section") {
                            insertionStyle = (rev.heading_text && rev.heading_text.trim()) ? "new_section" : "continuation";
                        }
                        // A clause that carries a heading belongs in its own subparagraph, not
                        // jammed inline. If a heading was supplied separately, treat it as a new
                        // section. If the heading is embedded at the start of the body text
                        // ("Survival. The provisions..."), split it out and promote likewise.
                        if (rev.heading_text && rev.heading_text.trim()) {
                            insertionStyle = "new_section";
                        } else {
                            let split = splitEmbeddedHeading(rev.body_text || rev.redlined_text || "");
                            if (split) {
                                rev.heading_text = split.heading;
                                rev.body_text = split.body;
                                insertionStyle = "new_section";
                            }
                        }
                        // Resolve the body text (the clause content) and optional heading label.
                        let bodyText = (rev.body_text && rev.body_text.trim())
                            ? rev.body_text.trim()
                            : (rev.redlined_text || "").trim();
                        bodyText = normalizeInsertText(bodyText);
                        if (!bodyText) { failCount++; continue; }
                        let anchorItem = (rev.context_anchor && rev.context_anchor.trim() !== "")
                            ? await findAnchorRange(context, body, rev.context_anchor)
                            : null;
                        // ---- CONTINUATION: insert the new sentence(s) into the existing paragraph
                        // at a grammatically clean point — after the model-specified sentence if
                        // given (rev.after_sentence), else at the true end of the paragraph. The
                        // helper guarantees single-period separation and never splits mid-sentence. ----
                        if (insertionStyle === "continuation") {
                            if (!anchorItem) {
                                // No anchor to continue from: fall back to a safe comment.
                                failCount++;
                                try {
                                    let host = body.paragraphs.getFirst().getRange();
                                    host.insertComment(buildReviewerComment(rev, false));
                                    await context.sync();
                                } catch (e) {}
                                continue;
                            }
                            // Expand the matched anchor to its FULL containing paragraph, so we
                            // operate on the whole paragraph's text (not the matched sub-range,
                            // which sits mid-paragraph and caused mid-sentence insertion).
                            let para = anchorItem.paragraphs.getFirst();
                            // If the anchor matched a HEADING paragraph (e.g. "6. Limitation of
                            // Liability"), redirect to the body paragraph below it so the new
                            // clause does not get buried inside the section title.
                            para = await resolveContinuationTarget(context, para, body);
                            await insertContinuation(context, para, bodyText, rev.after_sentence);
                            para.getRange().insertComment(buildReviewerComment(rev, true));
                            successCount++;
                            await context.sync();
                            continue;
                        }
                        // ---- NEW SECTION: insert a new top-level section AFTER the anchor's
                        // FULL paragraph (not the mid-paragraph match), so it never splits an
                        // existing section. Two things vary by the detected document style:
                        //   (1) NUMBERING: if the contract uses NATIVE Word list numbering, we
                        //       insert into that list and let Word assign & renumber in the
                        //       contract's own scheme (1./A./IV./Article N). If the numbers are
                        //       PLAIN TYPED text, we do NOT guess — insert unnumbered + flag it.
                        //   (2) HEADING LAYOUT: if headings sit on their OWN line (bold, no colon,
                        //       hard return), put the title on one paragraph and body on the next.
                        //       Otherwise use the detected delimiter inline ("Heading: body").
                        let headingLabel = (rev.heading_text && rev.heading_text.trim()) ? rev.heading_text.trim() : "";
                        if (headingLabel) {
                            if (headingStyle.allCaps) headingLabel = headingLabel.toUpperCase();
                            // Strip any trailing delimiter and any leading number the model added;
                            // the document's own formatting (or Word's list) supplies those.
                            headingLabel = headingLabel.replace(/[:.\-\s]+$/, "");
                            headingLabel = headingLabel.replace(/^\s*([0-9]+|[A-Za-z])[.)]\s*/, "");
                        }
                        let anchorPara = anchorItem ? anchorItem.paragraphs.getFirst() : null;
                        let target = anchorPara || body;
                        let loc = anchorPara ? Word.InsertLocation.after : Word.InsertLocation.end;
                        let ownLine = headingStyle.headingOnOwnLine || headingStyle.delimiter === "none";

                        try {
                            console.log('[DIAG-NEWSEC] inserting new_section. heading=' + JSON.stringify(headingLabel) +
                                ' | nativeList=' + headingStyle.nativeList + ' | ownLine=' + ownLine +
                                ' | delimiter=' + JSON.stringify(headingStyle.delimiter));
                        } catch (e) {}

                        let headingPara;
                        if (ownLine && headingLabel) {
                            // Title on its own paragraph, body on the next. Insert body first, then
                            // the heading immediately before it, so the order reads heading -> body.
                            let bodyPara = target.insertParagraph(bodyText, loc);
                            headingPara = bodyPara.insertParagraph(headingLabel, Word.InsertLocation.before);
                            applyHeadingFormatting(headingPara.getRange(), headingStyle);
                            // Explicitly clear heading formatting on the BODY paragraph so it never
                            // inherits the heading's bold/underline (observed: a multi-sentence body
                            // rendered entirely bold). Body text should be normal weight.
                            try {
                                let bf = bodyPara.getRange().font;
                                bf.bold = false;
                                bf.italic = false;
                                bf.underline = "None";
                                bf.smallCaps = false;
                                await context.sync();
                            } catch (e) { /* best effort */ }
                        } else if (headingLabel) {
                            // Inline "Heading<delim> body" using the detected delimiter.
                            let delim = (headingStyle.delimiter && headingStyle.delimiter !== "none") ? headingStyle.delimiter : ":";
                            let fullText = `${headingLabel}${delim} ${bodyText}`;
                            headingPara = target.insertParagraph(fullText, loc);
                            try {
                                let headingPortion = `${headingLabel}${delim}`;
                                let hr = headingPara.getRange().search(headingPortion, { matchCase: false });
                                hr.load("items");
                                await context.sync();
                                if (hr.items.length > 0) applyHeadingFormatting(hr.items[0], headingStyle);
                            } catch (e) { /* heading formatting best-effort */ }
                        } else {
                            // No heading provided though tagged new_section: insert body alone.
                            headingPara = target.insertParagraph(bodyText, loc);
                        }

                        // NUMBERING: attach the new heading paragraph to the existing section's
                        // native list (Word numbers/renumbers it in the doc's own scheme), or flag
                        // it for manual numbering when the contract uses plain typed numbers.
                        if (headingStyle.nativeList && anchorPara) {
                            try {
                                let anchorList = anchorPara.listItemOrNullObject;
                                anchorList.load("level");
                                await context.sync();
                                if (!anchorList.isNullObject) {
                                    let lvl = anchorList.level || 0;
                                    let hList = headingPara.listItemOrNullObject;
                                    await context.sync();
                                    if (hList.isNullObject) {
                                        anchorPara.list.load("id");
                                        await context.sync();
                                        headingPara.attachToList(anchorPara.list.id, lvl);
                                        await context.sync();
                                    }
                                    console.log('[DIAG-NEWSEC] attached to native list at level ' + lvl + ' (Word numbers/renumbers).');
                                } else {
                                    console.log('[DIAG-NEWSEC] anchor not in a list after all; flagging for manual numbering.');
                                    headingPara.getRange().insertComment("ACTION NEEDED — assign this new section the correct number in this contract's numbering scheme.");
                                }
                            } catch (e) {
                                console.warn('[DIAG-NEWSEC] native-list attach failed; flagging for manual numbering. ' + e.message);
                                headingPara.getRange().insertComment("ACTION NEEDED — assign this new section the correct number in this contract's numbering scheme.");
                            }
                        } else {
                            // Plain-text numbering (or no anchor): do NOT guess a number.
                            console.log('[DIAG-NEWSEC] plain-text numbering => inserted WITHOUT a number; flagged for reviewer.');
                            headingPara.getRange().insertComment("ACTION NEEDED — this contract uses plain (typed) section numbers, so no number was added automatically. Please number this new section per the contract's scheme and update any cross-references.");
                        }

                        headingPara.getRange().insertComment(buildReviewerComment(rev, true));
                        successCount++;
                        await context.sync();
                        continue;
                    }
                    // --------------------------------------------------------
                    // TYPE 1 — SURGICAL: minimal in-place strike + insert.
                    // --------------------------------------------------------
                    // A surgical edit is INCOMPLETE only if it lacks a context_anchor or a
                    // target_text. An empty redlined_text is NOT incomplete — it is a valid
                    // pure DELETION (strike the target words), which the playbook explicitly
                    // calls for (strike Work Made for Hire, strike liquidated damages, strike
                    // most-favored-nation). We handle that deletion case directly below.
                    if (!rev.context_anchor || !rev.target_text || rev.target_text.trim() === "") {
                        // Incomplete surgical edit: never fail silently. Leave a reviewer comment
                        // anchored where we can, so the suggestion isn't lost (License-grant fix).
                        failCount++;
                        try {
                            let host = (rev.context_anchor && rev.context_anchor.trim())
                                ? await findAnchorRange(context, body, rev.context_anchor)
                                : null;
                            if (!host) host = body.paragraphs.getFirst().getRange();
                            host.insertComment(buildReviewerComment(rev, false));
                            await context.sync();
                        } catch (e) { /* best effort */ }
                        continue;
                    }
                    // PURE DELETION (strike): valid target_text, empty redlined_text. Locate the
                    // target and delete it (replace with empty string) as a tracked deletion. We
                    // handle this BEFORE punctuation repair and the bundling guard, both of which
                    // assume a non-empty replacement.
                    if (!rev.redlined_text || rev.redlined_text.trim() === "") {
                        // A cross-paragraph target (heading line + newline + body) can't be
                        // located whole; reduce to its largest single-paragraph segment first.
                        let crossReduced = largestSingleParagraphSegment(rev.target_text);
                        // Some runs jam the heading onto the front of the body
                        // ("5. Liquidated DamagesIn the event..."); strip a run-together heading
                        // so we search the body alone. The [RESERVED] post-pass flags the
                        // now-empty section's heading.
                        let strikeTarget = stripLeadingSectionHeading(crossReduced);
                        let delTarget = await findTargetInBody(context, body, strikeTarget, rev.context_anchor);
                        if (!delTarget && strikeTarget !== rev.target_text) {
                            delTarget = await findTargetInBody(context, body, rev.target_text, rev.context_anchor);
                        }
                        if (!delTarget) {
                            console.warn('[DIAG-FAIL] deletion target NOT found. target_text=' + JSON.stringify(rev.target_text));
                            failCount++;
                            try {
                                let commentHost = await findAnchorRange(context, body, rev.context_anchor);
                                if (!commentHost) commentHost = body.paragraphs.getFirst().getRange();
                                commentHost.insertComment(buildReviewerComment(rev, false));
                                await context.sync();
                            } catch (e) { /* best effort */ }
                            continue;
                        }
                        console.log('[DIAG-OK] deletion target FOUND for target_text=' + JSON.stringify(rev.target_text));
                        delTarget.insertText("", Word.InsertLocation.replace);
                        delTarget.insertComment(buildReviewerComment(rev, true));
                        successCount++;
                        await context.sync();
                        continue;
                    }
                    // Repair dropped trailing punctuation first (existing behavior).
                    rev.redlined_text = repairTrailingPunctuation(rev.target_text, rev.redlined_text);
                    // Repair model artifacts in the replacement text (e.g., "Company' s" -> "Company's").
                    rev.redlined_text = normalizeInsertText(rev.redlined_text);
                    // Preserve sentence-start capitalization: if the ORIGINAL target began with a
                    // capital letter (i.e., it started a sentence/clause) but the replacement begins
                    // lowercase, capitalize the replacement's first letter. This fixes "...invoice. a
                    // late fee..." where the model returned a lowercase replacement for a capitalized
                    // target. Only acts when the target was clearly capitalized, so it never forces
                    // caps onto a mid-sentence replacement that should stay lowercase.
                    try {
                        let tgtTrim = (rev.target_text || "").trim();
                        let rlTrim = (rev.redlined_text || "").trim();
                        if (/^[A-Z]/.test(tgtTrim) && /^[a-z]/.test(rlTrim)) {
                            rev.redlined_text = rev.redlined_text.replace(/^(\s*)([a-z])/, (m, sp, c) => sp + c.toUpperCase());
                        }
                    } catch (e) { /* best effort */ }
                    // BUNDLING GUARD: the model sometimes fixes a small atomic value (an email
                    // address, a name, a phone) AND appends an unrelated new requirement in the
                    // same surgical edit. Applied literally, the appended sentence lands in the
                    // wrong place (e.g., a certified-mail sentence planted in a contact row —
                    // Redline 1). Detect an atomic target (a single token / an email) whose
                    // replacement adds one or more extra sentences, apply ONLY the value swap,
                    // and route the appended language to a reviewer comment for correct placement.
                    {
                        let tgt = rev.target_text.trim();
                        let isAtomic = /^[^\s]+@[^\s]+$/.test(tgt) || tgt.split(/\s+/).length <= 2;
                        if (isAtomic) {
                            let rl = rev.redlined_text.trim();
                            // Sentence count in the replacement (a value swap should be ~1 fragment).
                            let sentenceEnders = (rl.match(/[.!?](\s|$)/g) || []).length;
                            // The replacement's first token (the corrected value) vs. the rest.
                            let firstSentenceMatch = rl.match(/^.*?[.!?](\s|$)/);
                            if (sentenceEnders >= 2 && firstSentenceMatch) {
                                let valueOnly = firstSentenceMatch[0].trim().replace(/[.!?]+$/, "");
                                let appended = rl.slice(firstSentenceMatch[0].length).trim();
                                // Apply only the value swap inline; comment the appended requirement.
                                rev.redlined_text = valueOnly;
                                rev._bundledExtra = appended; // surfaced via comment below
                            }
                        }
                    }
                    // Trim to the minimal differing span so we never strike-and-retype identical words.
                    let diff = trimToMinimalDiff(rev.target_text, rev.redlined_text);
                    if (diff === null) { /* nothing actually changed */ continue; }
                    // Locate the target by searching the WHOLE BODY first. The anchor
                    // disambiguates when there are multiple body matches. A cross-paragraph
                    // target can't be matched whole, so search on its largest single-paragraph
                    // segment; downstream coverage verification still governs whether a clean
                    // replacement is possible (else it falls back to a comment).
                    let searchTarget = diff.pureInsert ? rev.target_text : diff.target;
                    searchTarget = largestSingleParagraphSegment(searchTarget);
                    let targetItem = await findTargetInBody(context, body, searchTarget, rev.context_anchor);
                    if (!targetItem) {
                        console.warn('[DIAG-FAIL] surgical anchor NOT found. searchTarget=' + JSON.stringify(searchTarget) + ' | target_text=' + JSON.stringify(rev.target_text) + ' | context_anchor=' + JSON.stringify(rev.context_anchor));
                        // Safe failure: never blind-overwrite. Attach a margin comment naming the
                        // rule, issue, and suggested change so the reviewer can act manually.
                        failCount++;
                        try {
                            let commentHost = await findAnchorRange(context, body, rev.context_anchor);
                            if (!commentHost) commentHost = body.paragraphs.getFirst().getRange();
                            commentHost.insertComment(buildReviewerComment(rev, /*applied=*/false));
                            await context.sync();
                        } catch (e) { /* comment best-effort */ }
                        continue;
                    }
                    console.log('[DIAG-OK] surgical anchor FOUND for target_text=' + JSON.stringify(rev.target_text));
                    // Guard: skip if the replacement is already present (idempotency).
                    context.load(targetItem, 'text');
                    await context.sync();
                    if (diff.pureInsert) {
                        // Words added with nothing to strike: insert on the correct side of
                        // the matched span (front if prepended, back if appended).
                        let loc = (diff.insertSide === "before") ? Word.InsertLocation.start : Word.InsertLocation.end;
                        targetItem.insertText(diff.insertText, loc);
                        targetItem.insertComment(buildReviewerComment(rev, true));
                    } else {
                        // VERIFY the matched range actually covers the intended target before
                        // replacing. A range found via space-insensitive or fragment search can
                        // end short of the full target, and replacing a partial span leaves a
                        // fragment of the old text behind (the "$1,000,000...$2,000,000 per
                        // occurrence and $4,000,000" garble). If the range text does not match
                        // the target, re-locate the exact target inside the matched paragraph.
                        let matchedText = (targetItem.text || "").replace(/\s+/g, ' ').trim();
                        let wantText = diff.target.replace(/\s+/g, ' ').trim();
                        let exactRange = targetItem;
                        if (matchedText !== wantText) {
                            let exact = await locateExactRange(context, targetItem, diff.target);
                            if (exact) {
                                exactRange = exact;
                            } else {
                                // locateExactRange couldn't pin it (commonly the target exceeds
                                // Word's 255-char search ceiling). Build the range from a short
                                // head + tail anchor instead; it verifies full coverage and
                                // returns null if it can't, so we never replace a partial span.
                                let byHeadTail = await locateRangeByHeadTail(context, targetItem, diff.target);
                                if (byHeadTail) {
                                    exactRange = byHeadTail;
                                } else {
                                    // Could not pin the exact span safely. Do NOT replace a partial
                                    // range (that causes garble). Fall back to a reviewer comment.
                                    failCount++;
                                    try {
                                        targetItem.insertComment(buildReviewerComment(rev, /*applied=*/false));
                                        await context.sync();
                                    } catch (e) { /* best effort */ }
                                    continue;
                                }
                            }
                        }
                        // True surgical replacement of exactly the differing words.
                        // Guard against double punctuation: getSafeSearchText strips a trailing
                        // .;: from the SEARCH, so the matched range can exclude the original's
                        // trailing punctuation. If our redline ends with that same punctuation
                        // and the matched range does NOT already include it, the original mark
                        // survives in the document, producing "enforceable.." Detect this and
                        // drop the redundant trailing punctuation from the replacement.
                        let replacement = diff.redline;
                        let rangeText = (exactRange.text || "");
                        let lastChar = replacement.trim().slice(-1);
                        if ((lastChar === "." || lastChar === ";" || lastChar === ":") &&
                            !rangeText.trim().endsWith(lastChar)) {
                            // The matched range ends without this punctuation; the document's
                            // own punctuation will remain immediately after. Drop ours.
                            replacement = replacement.replace(/[\.;:]+$/, "");
                        }
                        exactRange.insertText(replacement, Word.InsertLocation.replace);
                        exactRange.insertComment(buildReviewerComment(rev, true));
                        // If the bundling guard split off an appended requirement, surface it as a
                        // separate comment so the reviewer can place that language correctly,
                        // rather than letting it land in the wrong spot (e.g., a contact row).
                        if (rev._bundledExtra && rev._bundledExtra.trim()) {
                            try {
                                exactRange.insertComment(`ACTION NEEDED — additional requirement (place in the correct clause): "${rev._bundledExtra.trim()}"`);
                            } catch (e) { /* best effort */ }
                        }
                    }
                    successCount++;
                    await context.sync();
                }

                // ============================================================
                //  POST-PASS: [RESERVED] / partial-strike handling
                // ============================================================
                // Two distinct situations arise when a deletion targets a section's body:
                //  (A) FULLY STRUCK: essentially the entire body was struck. Lawyer convention is
                //      to mark the whole section "[RESERVED]" (title struck, number kept) so
                //      downstream numbering and cross-references stay accurate.
                //  (B) PARTIALLY STRUCK ("gutted"): a large but not complete span was struck,
                //      leaving a grammatically broken stub (e.g. "...by its scheduled date, the
                //      Company ."). Per the drafting decision, we do NOT auto-collapse these
                //      (that would be a legal judgment — the lawyer may want to keep the lead-in
                //      and rewrite). Instead we FLAG the section with a reviewer comment.
                //
                // We score each section by how much of its body the deletions covered, working
                // from the model's deletion targets against the ORIGINAL contract text (reliable),
                // rather than reading tracked-change state from the live doc.
                try {
                    let deletions = (result.revisions || []).filter(r => {
                        let et = (r.edit_type || "").toLowerCase();
                        return et === "surgical" && r.target_text && r.target_text.trim() &&
                               (!r.redlined_text || r.redlined_text.trim() === "");
                    });
                    console.log('[DIAG-RESERVED] post-pass start. deletions found: ' + deletions.length);
                    if (deletions.length > 0) {
                        // Split the ORIGINAL contract text into sections by heading lines.
                        // Word's body.text separates paragraphs with \r (sometimes \r\n or \n),
                        // so split on any of them.
                        let origLines = rawText.split(/\r\n|\r|\n/);
                        let sections = [];
                        let cur = null;
                        let headingRe = /^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+(.+\S)\s*$/;
                        for (let ln of origLines) {
                            let hm = ln.match(headingRe);
                            if (hm && hm[2].split(/\s+/).length <= 8 && !/[.;]$/.test(hm[2])) {
                                if (cur) sections.push(cur);
                                cur = { marker: hm[1], title: hm[2].trim(), headingLine: ln.trim(), body: "" };
                            } else if (cur) {
                                cur.body += (cur.body ? " " : "") + ln.trim();
                            }
                        }
                        if (cur) sections.push(cur);
                        console.log('[DIAG-RESERVED] sections parsed from original text: ' + sections.length +
                            ' => ' + JSON.stringify(sections.map(s => s.marker + ' ' + s.title)));

                        // Normalizer for overlap comparison.
                        let norm = s => (s || "").replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();

                        for (let sec of sections) {
                            let secBody = sec.body.replace(/\s+/g, ' ').trim();
                            if (!secBody) continue;
                            // Compute how many characters of this section's body are covered by
                            // deletion targets (after stripping any heading line the model included).
                            let nbody = norm(secBody);
                            let struckChars = 0;
                            let matchedAny = false;
                            for (let d of deletions) {
                                // Strip a leading section heading from the deletion target so it
                                // matches the section BODY. Two forms occur: (1) heading + newline +
                                // body ("5. Liquidated Damages\nIn the event..."), and (2) heading
                                // jammed directly onto the body with no separator ("5. Liquidated
                                // DamagesIn the event..."). Handle the newline form first, then the
                                // jammed form via the shared helper, so both reduce to body-only text.
                                let raw = (d.target_text || "").replace(/^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+[^\n]{0,60}\n/, '');
                                raw = stripLeadingSectionHeading(raw);
                                let nt = norm(raw);
                                if (nt.length >= 15 && nbody.indexOf(nt) !== -1) {
                                    struckChars = Math.max(struckChars, nt.length);
                                    matchedAny = true;
                                }
                            }
                            if (!matchedAny) continue;
                            let ratio = struckChars / Math.max(1, nbody.length);
                            console.log('[DIAG-RESERVED] section ' + JSON.stringify(sec.marker + ' ' + sec.title) +
                                ' struck-ratio=' + ratio.toFixed(2) + ' (struck ' + struckChars + ' / ' + nbody.length + ' chars)');

                            if (ratio >= 0.85) {
                                // (A) FULLY STRUCK -> FLAG for the reviewer to mark [RESERVED].
                                // We deliberately do NOT auto-strike the title or auto-insert
                                // "[RESERVED]". Earlier attempts to do that surgically struck the
                                // section NUMBER (which must be kept for cross-references), jammed
                                // "[RESERVED]" between the number and title, and left an orphaned
                                // body fragment. Per the drafting decision, the human makes the
                                // final "N. [RESERVED]" edit so numbering is never corrupted.
                                console.log('[DIAG-RESERVED] => FULLY struck; flagging for reviewer to mark [RESERVED] (no auto-collapse).');
                                try {
                                    let titleResults = await runSearchLadder(context, body, getSafeSearchText(sec.title, 120));
                                    let host = (titleResults.items && titleResults.items.length > 0)
                                        ? titleResults.items[0]
                                        : body.paragraphs.getFirst().getRange();
                                    host.insertComment(
                                        "ACTION NEEDED \u2014 the entire body of this section has been struck. " +
                                        "To finish: replace the section title with \"[RESERVED]\" and delete any " +
                                        "remaining struck body text, while KEEPING the section number (e.g. \"" +
                                        sec.marker + ". [RESERVED]\") so cross-references stay valid."
                                    );
                                    await context.sync();
                                    reservedCount++;
                                } catch (e) {
                                    console.warn('[DIAG-RESERVED] flag failed for ' + JSON.stringify(sec.title) + ': ' + e.message);
                                }
                            } else if (ratio >= 0.40) {
                                // (B) PARTIALLY STRUCK / gutted -> FLAG, do not auto-collapse.
                                console.log('[DIAG-RESERVED] => PARTIALLY struck; flagging for reviewer (no auto-collapse).');
                                try {
                                    let titleResults = await runSearchLadder(context, body, getSafeSearchText(sec.title, 120));
                                    let host = (titleResults.items && titleResults.items.length > 0)
                                        ? titleResults.items[0]
                                        : body.paragraphs.getFirst().getRange();
                                    host.insertComment("ACTION NEEDED — a substantial part of this section was struck, which may leave incomplete or grammatically broken language. Consider whether this section should be fully reserved (\"[RESERVED]\") or rewritten. Please review.");
                                    await context.sync();
                                } catch (e) {
                                    console.warn('[DIAG-RESERVED] flag failed for ' + JSON.stringify(sec.title) + ': ' + e.message);
                                }
                            }
                            // ratio < 0.40: a minor deletion; leave it (normal surgical edit).
                        }
                    }
                } catch (e) {
                    console.warn('[DIAG-RESERVED] post-pass error (non-fatal): ' + e.message);
                }

                statusDiv.className = "status-message " + (failCount > 0 ? "error" : "success");
                statusDiv.innerText = "Review complete.";
            });
        } else {
            statusDiv.className = "status-message success";
            statusDiv.innerText = "Analysis Complete. Contract perfectly aligns with playbooks.";
        }
        btn.disabled = false;
    } catch (error) {
        clearTimeout(clientTimeoutClock);
        if (eventSource) eventSource.close();
        if (!processingAbortedByChoice) {
            statusDiv.className = "status-message error";
            statusDiv.innerText = error.message;
            btn.disabled = false;
        }
    }
}
