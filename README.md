# Smith-Atlas LLC — Open-Source AI Tools for Contracting, Compliance, and Research Administration

This repository contains AI tools built by [Smith-Atlas LLC](https://www.smith-atlas.com) for the people who do regulatory, compliance, and contracting work every day: attorneys, contract managers, procurement offices, sponsored programs offices, and research administrators.

Each tool was designed by an attorney and contracts professional with years of experience in commercial contracting and research administration. They are released here as open source under the [Apache 2.0 license](LICENSE), free to use, adapt, and build on.

## Why these tools are here

These tools were originally offered for license on the Smith-Atlas website. We have decided to make them freely available instead, so that more organizations can benefit from them and see how we approach AI in high-stakes work. Our focus at Smith-Atlas is helping organizations *adopt* AI well: the implementation, workflow design, training, and change management that turn a good tool into a working part of how a team operates.

If you'd like help putting any of these tools to work, or building something new, see [Need help?](#need-help) below.

## The tools

| Tool | What it does | What it is |
| --- | --- | --- |
| [**SA-1006 Agent Generator Assistant**](SA-1006-Smith-Atlas-Agent-Generator-Assistant) | Interviews you about your expertise and produces a complete, guardrailed system prompt for your own custom AI assistant. No prompt-writing experience needed. | System instructions for any AI assistant platform |
| [**SA-1007 Policy Evaluation & Drafting Assistant**](SA-1007-Smith-Atlas-Policy-Evaluation-Drafting-Assistant) | Benchmarks peer institutions' policies with traceable citations and drafts institutional policies on your templates, screened against 2 CFR Part 200, the FAR, and NSF/NIH terms. | System instructions for any AI assistant platform |
| [**SA-1008 Federal Cost Allowability Assistant**](SA-1008-Smith-Atlas-Federal-Cost%20Allowability-Assistant) | Answers "can I charge this to my award?" by running the allowable / allocable / reasonable test with subsection-level citations and version dates. | System instructions for any AI assistant platform |
| [**SA-1009 Contract Redline Direct**](SA-1009-Smith-Atlas-Contract%20Redline%20Direct) | Reviews a contract against your playbook and returns your own Word file with native tracked changes and margin comments. | Instructions and a redline engine for a code-capable AI assistant (with an optional Claude Skill) |
| [**SA-1011 Solicitation Navigator**](SA-1011-Smith-Atlas-Solicitation-Navigator) | Extracts requirements from a funding solicitation and its amendments into a 39-element table, with page-level citations for each. | System instructions for any AI assistant platform |
| [**SA-2001H Contract Negotiation Assistant**](SA-2001H-Smith-Contract-Negotiation-Assistant) | A Microsoft Word add-in that redlines the open document against your playbook, using a cloud AI engine or one hosted on your own infrastructure. **Draft release.** | Software (Node.js backend and Word add-in), with a full build guide |

Each folder contains an **About** file describing the tool in detail, the files you need, setup instructions, and a link to a short overview video.

## How they're built

The tools share a few design principles:

- **The human makes the final call.** Every tool keeps a qualified person in the loop and says so in its output. None of them approve, sign, submit, or decide anything on their own.
- **Answers you can check.** Wherever a tool relies on a source, it cites it precisely (the regulation subsection, the solicitation page, the playbook rule) so you can verify it quickly.
- **Your rules, not generic ones.** Contract tools enforce the positions in *your* playbook. Policy and compliance tools work from the documents *you* provide.
- **Works with the AI you already have.** Most tools are platform-agnostic system instructions for Claude, ChatGPT, Gemini, Microsoft Copilot, or an in-house assistant. There's no new subscription to buy.
- **Efficient by design.** Instructions are written to keep token usage, and therefore cost and energy use, down.

## Getting started

1. Open the folder for the tool you want and read its **About** file.
2. Watch the overview video linked at the bottom of the About file.
3. Follow the setup instructions included in the folder. Most of the prompt-based tools take 10 to 20 minutes to set up.

## Important

- **These tools do not provide legal advice**, and using them does not create an attorney-client relationship. All output is AI-generated and must be reviewed by a qualified professional before you rely on it.
- **Included playbooks and configurations are starting points.** Review and adapt them with your own counsel and subject matter experts before use.
- **Protect your data.** You are strongly encouraged to use an enterprise-licensed AI platform, or one with a contractual commitment not to train on your content, before uploading contracts or other sensitive documents.
- **Provided as is.** Under the Apache 2.0 license, these tools are provided without warranty of any kind. The SA-2001H Contract Negotiation Assistant, in particular, is a draft that requires development work before production use.

## Need help?

Smith-Atlas works with organizations to put AI tools to work, including these:

- **Implementation:** setting up and configuring these tools in your environment, adapting playbooks and instructions to your organization's positions, and fitting them into your existing workflows.
- **Custom AI tools:** building new tools for your own workflows, such as policy drafting, RFP development and response, contract data extraction, transactional auditing, and more.
- **Adoption and change management:** consultations, workshops, and training that build AI literacy and help teams, including skeptical ones, actually use the tools.

Visit **[www.smith-atlas.com](https://www.smith-atlas.com)** to learn more or to book a free 15-minute discovery session. You can contact us directly at **matt@smith-atlas.com**.

## License

Copyright 2026 Smith-Atlas LLC. Licensed under the [Apache License, Version 2.0](LICENSE). See the [NOTICE](NOTICE) file for attribution. If you build on these tools, we'd love to hear about it.
