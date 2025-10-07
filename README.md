# PKeyInspector

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()
[![Releases](https://img.shields.io/badge/releases-v1.0-blue.svg)]()

> POC: Parse Office/Windows `pkeyconfig` and related SKU/license metadata and present interactive HTML reports (searchable) with export to PDF, Excel, CSV, and more.  
> Uses advanced, low‑level techniques (including undocumented or platform‑internal APIs) to collect richer system metadata — **research‑only** and disabled by default.

---

## Table of contents
- [About](#about)  
- [Why low‑level/undocumented APIs?](#why-low-levelundocumented-apis)  
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
**PKeyInspector** is a Proof‑Of‑Concept tool for administrators, auditors, and authorized researchers that aggregates metadata from `pkeyconfig` files, vendor XMLs, and system sources and produces human‑friendly reports (HTML, PDF, Excel, CSV). It also demonstrates the use of low‑level and undocumented platform internals to collect richer metadata when required.

> Because low‑level and undocumented APIs can access internal structures and behavior that are not guaranteed stable or safe, those capabilities are explicitly **research‑only**, gated, and disabled by default.

---

## Why low‑level / undocumented APIs?
High‑level documented APIs and registry queries are the safest route for typical inventory tasks, but they sometimes omit contextual metadata or vendor annotations. Low‑level techniques are included in this PoC to demonstrate:

- how additional contextual metadata can be discovered in memory, binary structures, or policy blocks,
- methods to recover richer SKU/product information when documented sources are unavailable,
- research techniques for forensic analysis and vendor‑compatibility research.

**Important:** Low‑level methods are brittle, platform‑version dependent, and may change or break with updates. They can also pose safety, privacy, and legal risks — see the Research‑only section and `SECURITY.md`.

---

## Capabilities (feature list)
> Full feature surface. Sensitive features are flagged as research‑only and disabled by default.

- Parse `pkeyconfig` XML files (Office/Windows) and extract SKU/product metadata.
- Interactive **HTML** reports with client‑side search/filter, paging, and sorting.
- Export to **PDF**, **Excel (.xlsx)**, **CSV**, and other formats.
- Consolidate latest product + key metadata from local files and vendor datasets.
- Present keys as **Genuine** or **Generated** (generated keys via pattern/encoding modules are research‑only).
- Encoding options:
  - Encode using documented APIs by SKU/GUID (where supported).
  - Encode by selected **pattern** or bulk encode via internal reference list (2000+ entries) — research‑only where it involves undocumented behavior.
- Decoding options:
  - Decode KeyInfo-like structures; parse binary payloads in read‑only mode.
  - Folder scanning and batch decode (read‑only, non‑destructive).
- Extraction options (RESEARCH‑ONLY / GATED):
  - Candidate extraction from binaries (`.exe`, `.dll`) and other artifacts (disabled by default).
  - Parse registry data blocks including kernel policy blocks for license metadata (disabled by default).
- System information collector:
  - Collect OS Major/Minor/Build/UBR/EditionID via documented APIs.
  - Research-only collectors use low‑level reads of memory/offsets to recover additional context (disabled by default).
- Error mapping & diagnostics: support for CBS, BITS, HTTP, UPDATE, NETWORK, WIN32, NTSTATUS, ACTIVATION code sets (best‑effort).
- Update matrix: build update compatibility tables from XML datasets, including unsupported versions (offline).
- Active license & settings view (OEM defaults, active SKU, and related metadata).
- Activation/license management (RESEARCH‑ONLY / GATED): integration points for lab/test workflows (disabled by default).
- Check Products keys against Offical MS Server API, 3 Api In total
  
---

## Primary goals
- Provide a single, searchable inventory of SKU/product metadata from `pkeyconfig` and vendor datasets.  
- Produce audit‑ready reports in multiple export formats.  
- Demonstrate how additional metadata can be surfaced using advanced low‑level techniques for authorized research and forensic analysis.  
- Keep sensitive features gated, logged, and restricted to lab environments.
