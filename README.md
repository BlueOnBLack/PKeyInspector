# PKeyInspector

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()
[![Releases](https://img.shields.io/badge/releases-v1.0-blue.svg)]()

> POC: Parse Office/Windows `pkeyconfig` and related SKU/license metadata, present interactive HTML reports (searchable) and export to PDF, Excel, CSV, and more.  
> Includes advanced inventory, encoding/decoding, and research‑only features — **disabled by default** and only for authorized research.

---

## Table of contents
- [About](#about)  
- [Capabilities (feature list)](#capabilities-feature-list)  
- [Primary goals](#primary-goals)  
- [Quick start (POC)](#quick-start-poc)  
- [Examples & outputs](#examples--outputs)  
- [Research‑only / sensitive features (gated)](#research-only--sensitive-features-gated)  
- [Security, Legal & Responsible Use](#security-legal--responsible-use)  
- [Contributing & Code of Conduct](#contributing--code-of-conduct)  
- [License](#license)  
- [Credits & references](#credits--references)

---

## About
**PKeyInspector** is a Proof‑Of‑Concept tool for administrators, auditors, and authorized researchers that aggregates metadata from `pkeyconfig` files and other SKU/license sources and produces human‑friendly reports in HTML, PDF, Excel and CSV formats. The project provides rich search/export capabilities and multiple options for how results are presented and exported.

> This repository intentionally separates benign reporting features (openly enabled) from sensitive capabilities (research‑only, disabled by default). The maintainers will not assist or accept contributions that enable unauthorized license extraction, generation, or activation bypass.

---

## Capabilities (feature list)
> The list below documents the full surface of capabilities envisioned for the project. Sensitive items are indicated and are **disabled by default**.

- Parse `pkeyconfig` files (Office/Windows) and extract SKU/product metadata (GUIDs, names, attributes, release info).
- Generate interactive **HTML** reports with client‑side search, filters, paging and sortable columns.
- Export inventory to **PDF**, **Excel (.xlsx)**, **CSV**, and other common data formats.
- Present consolidated lists of the *latest products + associated key metadata* from local files and vendor XML data sources.
- Show keys as either **Genuine** (when verifiable by documented metadata) or **Generated** (when using selected generation options) — see Research‑only section.
- Provide encoding options:
  - Encode using API by SKU or GUID (safe, documented APIs only).
  - Encode by selected *pattern* (research‑only; pattern library driven).
  - Bulk encode from an internal reference list (2000+ entries).
- Provide decoding options:
  - Decode using KeyInfo-like structures and safe parsing utilities.
  - Option to decode data discovered in a `folder` of files (read‑only parse).
- Extraction options (RESEARCH‑ONLY / GATED):
  - Extract candidate key metadata embedded in local binaries (`.exe`, `.dll`) and present extracted artifacts in the report — **disabled by default**.
  - Parse registry data blocks that may contain licensing metadata (two-block structures and kernel policy blocks) — **disabled by default**.
- System information collector:
  - Gather OS Major/Minor/Build/UBR and `EditionID` using documented system calls where possible.
  - Research option: collect the same data using low‑level reads (memory offsets) — **disabled by default**.
- Error mapping & diagnostics:
  - Extract and present error messages via documented APIs when available.
  - Support error lists/mappings for: `CBS`, `BITS`, `HTTP`, `UPDATE`, `NETWORK`, `WIN32`, `NTSTATUS`, `ACTIVATION` (best‑effort and documented).
- Activation/License operations (RESEARCH‑ONLY / GATED):
  - Integration option with external libraries (e.g., `tsforge`) for authorized activation flows for lab/test purposes — **disabled by default**.
  - Auto‑select activation method logic (KMS4K / ZeroCID / AVMA4K / HWID / KMS38) — listed as research/test-only integrations. The project will **not** provide instructions to use these against unauthorized systems.
  - Manage license entries via documented APIs (e.g., `slc.dll` wrappers) in read/write modes — **research‑only and gated**.
- Update matrix generation:
  - Build an update matrix table for any version based on XML data extraction — independent of local API availability — useful for offline or unsupported versions.
- View active license & settings:
  - Report on current OEM defaults, the active SKU, and other system license metadata (read‑only where possible).

---

## Primary goals
- Provide a single, searchable inventory of SKU/product metadata from `pkeyconfig` and vendor datasets.  
- Produce audit‑ready reports in HTML/PDF/Excel/CSV for compliance and documentation.  
- Provide safe, easy‑to-review POC scripts and examples that reproduce parsing and export flows without privileged operations.  
- Offer gated research features for authorized lab work; such features include additional warnings, logging, and explicit authorization steps.

---

## Quick start (POC)
> Non‑actionable, read‑only examples. See `/examples` for runnable POC scripts that are safe and mocked for sensitive functionality.
