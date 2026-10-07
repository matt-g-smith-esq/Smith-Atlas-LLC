// Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.

/**
 * @file BaseAdapter.js
 * @description Abstract base class defining the contract for all LLM provider integrations.
 */
class BaseAdapter {
    constructor(baseUrl, modelName) {
        if (this.constructor === BaseAdapter) {
            throw new Error("CRITICAL FAILURE: Cannot instantiate abstract class BaseAdapter directly.");
        }
        this.baseUrl = baseUrl;
        this.model = modelName;
    }

    /**
     * Standardized execution method for payload generation.
     * @param {string} systemPrompt - Structural operational guardrails and guidelines.
     * @param {string} userText - Target contract content to be processed.
     * @returns {Promise<string>} - Validated raw string response (typically structured JSON string).
     */
    async generateResponse(systemPrompt, userText) {
        throw new Error("CRITICAL FAILURE: Method 'generateResponse' must be implemented by subclass.");
    }
}

module.exports = BaseAdapter;
