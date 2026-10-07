// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.

/**
 * @file OllamaAdapter.js
 * @description Local engine integration adapter for Ollama.
 */
const BaseAdapter = require('./BaseAdapter');

class OllamaAdapter extends BaseAdapter {
    constructor(baseUrl, modelName) {
        super(baseUrl || 'http://localhost:11434', modelName);
        // One-line diagnostic so the resolved endpoint + model are visible at startup of
        // each call. If this prints localhost when you expect RunPod, the URL isn't being
        // passed down; if the model string here doesn't EXACTLY match `ollama list`, that
        // mismatch is the likely cause of a 404 from the server.
        // Gated behind DIAG_VERBOSE (default off) so a clean install stays quiet; set
        // DIAG_VERBOSE=true in .env to surface this while debugging an endpoint/model issue.
        if (String(process.env.DIAG_VERBOSE).toLowerCase() === 'true') {
            console.log(`[DIAG-OLLAMA] adapter target -> baseUrl=${JSON.stringify(this.baseUrl)} model=${JSON.stringify(this.model)}`);
        }
        if (!this.model || this.model.trim() === "") {
            throw new Error("CRITICAL FAILURE: No structural AI target model identifier specified for Ollama.");
        }

        // --- IRONCLAD GRAMMAR LOCKS (JSON SCHEMAS) ---

        // Schema 1: Capitalized Defined Terms Dictionary Extraction
        this.definitionSchema = {
            type: "array",
            items: {
                type: "object",
                properties: {
                    term: { type: "string" },
                    definition: { type: "string" }
                },
                required: ["term", "definition"]
            }
        };

        // Schema 2: Two-Step Surgical Patch Format
        this.revisionSchema = {
            type: "object",
            properties: {
                rule_checklist: {
                    type: "array",
                    items: { type: "string" }
                },
                revisions: {
                    type: "array",
                    items: {
                        type: "object",
                        properties: {
                            edit_type: { type: "string", enum: ["surgical", "insertion", "comment"] },
                            context_anchor: { type: "string" },
                            target_text: { type: "string" },
                            redlined_text: { type: "string" },
                            insertion_style: { type: "string" },
                            heading_text: { type: "string" },
                            body_text: { type: "string" },
                            after_sentence: { type: "string" },
                            position: { type: "string" },
                            rule_reference: { type: "string" },
                            issue_summary: { type: "string" },
                            drafting_note: { type: "string" }
                        },
                        required: ["edit_type", "context_anchor", "target_text", "redlined_text", "drafting_note"]
                    }
                }
            },
            required: ["rule_checklist", "revisions"]
        };
    }

    /**
     * Executes response generation with hard token-level constraints and high-capacity context paths.
     *
     * STREAMING (stream: true) is REQUIRED here, not optional. A large model behind a
     * reverse proxy (Cloudflare in front of RunPod, or an AWS/Azure gateway a customer
     * deploys) faces an IDLE-timeout ceiling — Cloudflare's is ~100s. With stream:false,
     * Ollama computes the entire response and sends nothing until it is finished, so a
     * generation that takes longer than the ceiling looks like dead air and the proxy
     * returns a 524 before we ever get a byte. With stream:true, Ollama emits tokens as
     * newline-delimited JSON ("NDJSON") chunks AS THEY ARE GENERATED; bytes keep flowing,
     * the proxy's idle timer keeps resetting, and a multi-minute generation completes
     * without tripping the ceiling. We accumulate the per-chunk message.content fragments
     * back into the single complete string the rest of the pipeline expects, so this change
     * is transparent to every caller (the return type is still the full response string).
     */
    async generateResponse(systemPrompt, userText) {
        const url = `${this.baseUrl}/api/chat`;

        // Determine the required grammar lock based on the operational target.
        // IMPORTANT: match ONLY on the dictionary prompt's unique self-identification
        // ("legal document parser"). We must NOT match on "defined terms" or "parser"
        // alone, because the redline prompt also contains the phrase "DEFINED TERMS"
        // (in its "GLOBAL DEFINED TERMS RECOGNITION DICTIONARY" section). Matching that
        // phrase wrongly locked the redline call to the definition schema, forcing the
        // model to emit term/definition pairs instead of revisions.
        let selectedSchema = this.revisionSchema;
        if (systemPrompt.toLowerCase().includes("legal document parser")) {
            selectedSchema = this.definitionSchema;
        }

        const payload = {
            model: this.model,
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: `Contract Text to Review:\n\n${userText}` }
            ],
            options: {
                temperature: 0.2,
                num_ctx: 32768
            },
            format: selectedSchema,
            // Pin the model in VRAM across calls so a multi-stage pipeline (defined-terms,
            // checklist, analysis, edits) doesn't pay a reload between stages. Mirrors the
            // server-side OLLAMA_KEEP_ALIVE=-1; harmless if that is already set.
            keep_alive: -1,
            stream: true
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            // Surface the actual response body and status code. Ollama returns the real
            // reason (e.g. {"error":"model 'x' not found, try pulling it first"}) in the
            // body; statusText is frequently empty (HTTP/2), which is why earlier failures
            // logged a useless "<none>". Read the body defensively so a non-JSON proxy
            // error page (RunPod 502/503/524, etc.) is still reported instead of swallowed.
            let detail = "";
            try {
                detail = await response.text();
            } catch (e) {
                detail = "(response body could not be read: " + e.message + ")";
            }
            let trimmed = (detail || "").trim();
            if (trimmed.length > 600) trimmed = trimmed.slice(0, 600) + "\u2026(truncated)";
            const statusInfo = `HTTP ${response.status}${response.statusText ? " " + response.statusText : ""}`;
            throw new Error(`AI Engine Router Error (Local Ollama Node): ${statusInfo}` +
                (trimmed ? ` \u2014 ${trimmed}` : " \u2014 (empty response body)"));
        }

        // --- STREAMED RESPONSE ASSEMBLY (NDJSON) ---
        // Ollama streams newline-delimited JSON objects, each shaped like:
        //   {"model":"...","message":{"role":"assistant","content":"<fragment>"},"done":false}
        // and a final object with "done":true (which may also carry an "error" field if the
        // model aborted mid-stream). We read the body as a stream, split on newlines, parse
        // each complete line, and concatenate the content fragments into the full string.
        // A partial line can straddle two network chunks, so we keep a buffer of the
        // unterminated tail and only parse up to the last newline each pass.
        let fullContent = "";
        let buffer = "";
        let sawDone = false;

        const decoder = new TextDecoder();

        // Node 18+/24 fetch returns a web ReadableStream with an async iterator.
        try {
            for await (const chunk of response.body) {
                buffer += decoder.decode(chunk, { stream: true });
                let newlineIndex;
                while ((newlineIndex = buffer.indexOf("\n")) >= 0) {
                    const line = buffer.slice(0, newlineIndex).trim();
                    buffer = buffer.slice(newlineIndex + 1);
                    if (!line) continue;
                    let obj;
                    try {
                        obj = JSON.parse(line);
                    } catch (e) {
                        // A non-JSON line in the stream is unexpected; skip it rather than
                        // aborting the whole assembly for one malformed fragment.
                        continue;
                    }
                    // Ollama can report a mid-stream error in a chunk; fail loudly if so.
                    if (obj.error) {
                        throw new Error(`AI Engine Router Error (Local Ollama Node): stream reported error \u2014 ${obj.error}`);
                    }
                    if (obj.message && typeof obj.message.content === "string") {
                        fullContent += obj.message.content;
                    }
                    if (obj.done === true) {
                        sawDone = true;
                    }
                }
            }
            // Flush any trailing buffered line that wasn't newline-terminated.
            const tail = buffer.trim();
            if (tail) {
                try {
                    const obj = JSON.parse(tail);
                    if (obj.error) {
                        throw new Error(`AI Engine Router Error (Local Ollama Node): stream reported error \u2014 ${obj.error}`);
                    }
                    if (obj.message && typeof obj.message.content === "string") {
                        fullContent += obj.message.content;
                    }
                    if (obj.done === true) sawDone = true;
                } catch (e) {
                    // ignore an unparseable trailing fragment
                }
            }
        } catch (streamErr) {
            // Re-throw with context so the retry loop / caller sees a clear message.
            throw new Error(`AI Engine Router Error (Local Ollama Node): stream read failed \u2014 ${streamErr.message}`);
        }

        if (!fullContent) {
            throw new Error("AI Engine Router Error: Received empty or invalid streamed content from local node." +
                (sawDone ? "" : " (stream ended without a done marker \u2014 connection may have dropped)"));
        }

        return fullContent;
    }
}

module.exports = OllamaAdapter;
