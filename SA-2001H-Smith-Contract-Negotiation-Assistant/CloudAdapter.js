/**
 * @file CloudAdapter.js
 * @description Universal Cloud LLM Engine Router mapping polymorphic payload structures.
 * Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.
 */
const BaseAdapter = require('./BaseAdapter');

class CloudAdapter extends BaseAdapter {
    constructor() {
        // Automatically initialize using generic system environment variables
        super(process.env.CLOUD_BASE_ENDPOINT, process.env.CLOUD_MODEL_PRIMARY);

        this.strategy = (process.env.CLOUD_API_STRATEGY || 'openai-compatible').toLowerCase();
        this.apiKey = process.env.CLOUD_API_KEY;
        this.fallbackModel = process.env.CLOUD_MODEL_FALLBACK;

        // --- Vendor-specific configuration (only the selected strategy's keys are required) ---
        // Vertex AI (Google Cloud) — project/location + service-account credentials, NOT an API key.
        this.vertexProjectId = process.env.VERTEX_PROJECT_ID;
        this.vertexLocation = process.env.VERTEX_LOCATION || 'us-central1';
        // GOOGLE_APPLICATION_CREDENTIALS (a path to the service-account JSON) is read by
        // google-auth-library automatically; we only note its presence for the config check.
        this.googleCredsPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
        // Amazon Bedrock — AWS region + access keys (the AWS SDK reads the standard AWS_* vars).
        this.awsRegion = process.env.AWS_REGION;
        this.awsAccessKeyId = process.env.AWS_ACCESS_KEY_ID;
        this.awsSecretAccessKey = process.env.AWS_SECRET_ACCESS_KEY;

        // Credential validation is STRATEGY-AWARE. Each vendor authenticates differently, so we
        // require only what the selected strategy actually needs and fail fast with a precise
        // message naming the missing variable(s) — never a generic "API key missing".
        switch (this.strategy) {
            case 'gemini-native':
            case 'openai-compatible':
            case 'anthropic-native':
                // Single-API-key vendors (Google AI Studio, OpenAI/Azure/Grok, Anthropic).
                if (!this.apiKey || this.apiKey.trim() === "") {
                    throw new Error(`CRITICAL CONFIGURATION FAILURE: strategy "${this.strategy}" requires CLOUD_API_KEY, which is missing from the environment.`);
                }
                break;
            case 'gemini-vertex':
                // Vertex authenticates with a GCP service account, not an API key.
                if (!this.vertexProjectId || this.vertexProjectId.trim() === "") {
                    throw new Error('CRITICAL CONFIGURATION FAILURE: strategy "gemini-vertex" requires VERTEX_PROJECT_ID.');
                }
                if (!this.googleCredsPath || this.googleCredsPath.trim() === "") {
                    throw new Error('CRITICAL CONFIGURATION FAILURE: strategy "gemini-vertex" requires GOOGLE_APPLICATION_CREDENTIALS (path to the service-account JSON file).');
                }
                break;
            case 'bedrock':
                // Bedrock authenticates with AWS credentials (SigV4), not an API key.
                if (!this.awsRegion || this.awsRegion.trim() === "") {
                    throw new Error('CRITICAL CONFIGURATION FAILURE: strategy "bedrock" requires AWS_REGION.');
                }
                if (!this.awsAccessKeyId || !this.awsSecretAccessKey) {
                    throw new Error('CRITICAL CONFIGURATION FAILURE: strategy "bedrock" requires AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY.');
                }
                break;
            default:
                // Unknown strategy: fail loudly rather than silently falling through to a wrong
                // vendor format. This is what protects a future Bedrock/Vertex customer from a
                // confusing mis-routed request when they typo the strategy name.
                throw new Error(`CRITICAL CONFIGURATION FAILURE: unknown CLOUD_API_STRATEGY "${this.strategy}". Supported: gemini-native, openai-compatible, anthropic-native, gemini-vertex, bedrock.`);
        }

        // Standard Grammar Locks shared symmetrically across all cloud vendors
        this.definitionSchema = { type: "array", items: { type: "object", properties: { term: { type: "string" }, definition: { type: "string" } }, required: ["term", "definition"] } };
        this.revisionSchema = { type: "object", properties: { rule_checklist: { type: "array", items: { type: "string" } }, revisions: { type: "array", items: { type: "object", properties: { edit_type: { type: "string", enum: ["surgical", "insertion", "comment"] }, context_anchor: { type: "string" }, target_text: { type: "string" }, redlined_text: { type: "string" }, insertion_style: { type: "string" }, heading_text: { type: "string" }, body_text: { type: "string" }, after_sentence: { type: "string" }, position: { type: "string" }, rule_reference: { type: "string" }, issue_summary: { type: "string" }, drafting_note: { type: "string" } }, required: ["edit_type", "context_anchor", "target_text", "redlined_text", "drafting_note"] } } }, required: ["rule_checklist", "revisions"] };
    }

    async generateResponse(systemPrompt, userText) {
        // Select the response grammar by the KIND of call:
        //  - defined-terms parser  -> definitionSchema
        //  - editing pass          -> revisionSchema
        // Match the defined-terms call on its full self-identification ("legal document
        // parser"), not the bare word "parser": the redline prompt also contains
        // "parser"-adjacent text, and a loose match would wrongly force redline calls into
        // the definition schema.
        let lower = systemPrompt.toLowerCase();
        let schema;
        if (lower.includes("legal document parser")) {
            schema = this.definitionSchema;
        } else {
            schema = this.revisionSchema;
        }

        try {
            return await this.executePayloadRoute(this.model, systemPrompt, userText, schema);
        } catch (error) {
            if (this.fallbackModel && this.model !== this.fallbackModel) {
                console.warn(`[RECOVERY PIPELINE] Primary cloud engine failed for THIS call. Falling back to: ${this.fallbackModel}`);
                // Use the fallback for THIS call only. Do NOT permanently reassign this.model:
                // the adapter instance is reused across the multi-stage pipeline (defined-terms,
                // checklist, editing), and a single failed call must not silently downgrade
                // every later call in the same request to the fallback engine.
                return await this.executePayloadRoute(this.fallbackModel, systemPrompt, userText, schema);
            }
            throw error;
        }
    }

    /**
     * Normalizes and translates common industry payload structures on the fly.
     * Bedrock is handled by its own method (it uses the AWS SDK + SigV4, not a plain fetch);
     * all HTTP-with-fetch vendors share the unified path below.
     */
    async executePayloadRoute(targetModel, systemPrompt, userText, targetSchema) {
        // Bedrock does not use fetch/URL/headers — delegate to the SDK-based path.
        if (this.strategy === 'bedrock') {
            return await this.executeBedrockRoute(targetModel, systemPrompt, userText);
        }

        let url = "";
        let headers = { 'Content-Type': 'application/json' };
        let bodyPayload = {};

        switch (this.strategy) {
            case 'gemini-native':
                url = `${this.baseUrl}/v1beta/models/${targetModel}:generateContent?key=${this.apiKey}`;
                bodyPayload = {
                    contents: [{ role: "user", parts: [{ text: `Contract Text to Review:\n\n${userText}` }] }],
                    systemInstruction: { parts: [{ text: systemPrompt }] },
                    generationConfig: { temperature: 0.2, responseMimeType: "application/json", responseSchema: targetSchema }
                };
                break;

            case 'gemini-vertex': {
                // Vertex AI (Google Cloud "Gemini Enterprise Agent Platform"). The request and
                // response BODY are identical to gemini-native — we reuse that exact shape. Only
                // two things differ: (1) the regional aiplatform.googleapis.com URL scoped to the
                // customer's project, and (2) OAuth bearer-token auth from a service account
                // (NOT an api-key query param). google-auth-library mints and refreshes the
                // short-lived token for us, so the customer never writes auth code.
                let token = await this.getVertexAccessToken();
                url = `https://${this.vertexLocation}-aiplatform.googleapis.com/v1/projects/${this.vertexProjectId}/locations/${this.vertexLocation}/publishers/google/models/${targetModel}:generateContent`;
                headers['Authorization'] = `Bearer ${token}`;
                bodyPayload = {
                    contents: [{ role: "user", parts: [{ text: `Contract Text to Review:\n\n${userText}` }] }],
                    systemInstruction: { parts: [{ text: systemPrompt }] },
                    generationConfig: { temperature: 0.2, responseMimeType: "application/json", responseSchema: targetSchema }
                };
                break;
            }

            case 'anthropic-native':
                // Anthropic Messages API. Differs from OpenAI in three ways the code must honor:
                // (1) auth header is x-api-key (NOT Authorization: Bearer); (2) anthropic-version
                // header is REQUIRED; (3) max_tokens is REQUIRED and the system prompt is a
                // TOP-LEVEL "system" field, not a message. Anthropic has no responseSchema param,
                // so structured-JSON output is driven by the prompt itself (UNIVERSAL_WRAPPER
                // already instructs "Return ONLY a valid JSON object").
                url = `${this.baseUrl}/v1/messages`;
                headers['x-api-key'] = this.apiKey;
                headers['anthropic-version'] = '2023-06-01';
                bodyPayload = {
                    model: targetModel,
                    max_tokens: 8192,
                    temperature: 0.2,
                    system: systemPrompt,
                    messages: [
                        { role: "user", content: `Contract Text to Review:\n\n${userText}` }
                    ]
                };
                break;

            case 'openai-compatible':
            default:
                // Works out-of-the-box for OpenAI, Grok, Azure, and any local network proxies
                url = `${this.baseUrl}/v1/chat/completions`;
                headers['Authorization'] = `Bearer ${this.apiKey}`;
                bodyPayload = {
                    model: targetModel,
                    messages: [
                        { role: "system", content: systemPrompt },
                        { role: "user", content: `Contract Text to Review:\n\n${userText}` }
                    ],
                    temperature: 0.2,
                    response_format: { type: "json_object", schema: targetSchema }
                };
                break;
        }

        const response = await fetch(url, { method: 'POST', headers: headers, body: JSON.stringify(bodyPayload) });
        if (!response.ok) {
            // Surface the response body, not just the status. Cloud vendors return structured
            // error detail (quota, key, safety-block reasons) that is essential for diagnosis.
            let detail = "";
            try { detail = await response.text(); } catch (e) { detail = "(response body could not be read: " + e.message + ")"; }
            let trimmed = (detail || "").trim();
            if (trimmed.length > 600) trimmed = trimmed.slice(0, 600) + "\u2026(truncated)";
            throw new Error(`Generic Cloud Gateway Exception (${targetModel}): Status ${response.status}` +
                (trimmed ? ` \u2014 ${trimmed}` : ""));
        }
        const data = await response.json();

        // Extract the text by vendor response shape:
        //  - Gemini (AI Studio + Vertex): candidates[0].content.parts[0].text
        //  - Anthropic:                   content[0].text
        //  - OpenAI-compatible:           choices[0].message.content
        if (this.strategy === 'gemini-native' || this.strategy === 'gemini-vertex') {
            return data.candidates[0].content.parts[0].text;
        }
        if (this.strategy === 'anthropic-native') {
            return data.content[0].text;
        }
        return data.choices[0].message.content;
    }

    /**
     * Mints a short-lived Google Cloud OAuth access token from the service-account
     * credentials pointed to by GOOGLE_APPLICATION_CREDENTIALS, for Vertex AI auth.
     * google-auth-library handles signing and caches/refreshes the token (it lives ~1
     * hour), so repeated calls within the pipeline are cheap and the customer never
     * writes or manages auth code. Lazily constructs the GoogleAuth client once.
     */
    async getVertexAccessToken() {
        if (!this._googleAuth) {
            // Required dependency: google-auth-library (bundled in the product, not installed
            // by the customer). Loaded lazily so non-Vertex deployments never touch it.
            const { GoogleAuth } = require('google-auth-library');
            this._googleAuth = new GoogleAuth({
                scopes: ['https://www.googleapis.com/auth/cloud-platform']
            });
        }
        const client = await this._googleAuth.getClient();
        const tokenResponse = await client.getAccessToken();
        let token = tokenResponse && (tokenResponse.token || tokenResponse);
        if (!token) {
            throw new Error('Vertex AI auth failed: could not obtain a Google Cloud access token from the service-account credentials (check GOOGLE_APPLICATION_CREDENTIALS).');
        }
        return token;
    }

    /**
     * Amazon Bedrock invocation path. Bedrock is NOT a plain HTTPS+api-key call — it requires
     * AWS SigV4 request signing, which the AWS SDK performs for us from the standard AWS_REGION /
     * AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY environment variables. The customer supplies
     * those as configuration; the SDK (bundled in the product) does all signing — no customer code.
     *
     * Bedrock fronts many model families; for Anthropic Claude models on Bedrock the request body
     * is the Anthropic Messages format (anthropic_version + max_tokens + system + messages) wrapped
     * in an InvokeModel call, and the response body is the Anthropic shape (content[0].text). This
     * branch targets Claude-on-Bedrock, the most common enterprise Bedrock use for this product.
     *
     * NOTE: written to the documented Bedrock/Anthropic contract but NOT yet verified against a
     * live AWS account — requires an end-to-end test with real Bedrock credentials before shipping.
     */
    async executeBedrockRoute(targetModel, systemPrompt, userText) {
        if (!this._bedrockClient) {
            // Required dependency: @aws-sdk/client-bedrock-runtime (bundled, not customer-installed).
            const { BedrockRuntimeClient } = require('@aws-sdk/client-bedrock-runtime');
            this._bedrockClient = new BedrockRuntimeClient({
                region: this.awsRegion,
                credentials: {
                    accessKeyId: this.awsAccessKeyId,
                    secretAccessKey: this.awsSecretAccessKey
                }
            });
        }
        const { InvokeModelCommand } = require('@aws-sdk/client-bedrock-runtime');
        const body = {
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 8192,
            temperature: 0.2,
            system: systemPrompt,
            messages: [
                { role: 'user', content: `Contract Text to Review:\n\n${userText}` }
            ]
        };
        let response;
        try {
            response = await this._bedrockClient.send(new InvokeModelCommand({
                modelId: targetModel,
                contentType: 'application/json',
                accept: 'application/json',
                body: JSON.stringify(body)
            }));
        } catch (e) {
            throw new Error(`Amazon Bedrock invocation failed (${targetModel}): ${e.message}`);
        }
        // Bedrock returns the model body as a byte array; decode and parse it.
        let decoded;
        try {
            decoded = JSON.parse(new TextDecoder().decode(response.body));
        } catch (e) {
            throw new Error(`Amazon Bedrock returned an unparseable response body (${targetModel}): ${e.message}`);
        }
        // Anthropic-on-Bedrock response shape: { content: [ { type: "text", text: "..." } ], ... }
        if (decoded && Array.isArray(decoded.content) && decoded.content[0] && typeof decoded.content[0].text === 'string') {
            return decoded.content[0].text;
        }
        throw new Error(`Amazon Bedrock response did not contain the expected content[0].text (${targetModel}).`);
    }
}

module.exports = CloudAdapter;
