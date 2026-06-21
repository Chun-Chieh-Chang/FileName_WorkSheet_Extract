# Entity: Deep Research

## Description
An autonomous search and research pipeline that deconstructs complex topics, executes multi-turn queries, crawls target web pages, and synthesizes reports with explicit citations.

## Core Capabilities
*   **Deconstructive Queries**: Breaking main queries down into 3+ sub-queries to query multiple search engines/indexes.
*   **Information Crawling**: Programmatic reading of page content using web scraping, text extraction, and token optimization.
*   **Citations Synthesis**: Generating structured markdown findings with inline clickable sources.
*   **Wiki Ingest**: Auto-registering results into the project's wiki entries to support compounding workspace memory.

## Integration
*   **Location**: `skills/core/deep-research/`
*   **Dependencies**: Tavily CLI (`tvly`), internet connectivity.

## Verification
*   Every fact or claim in the research report must trace directly back to a retrieved URL.
