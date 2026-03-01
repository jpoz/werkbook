# Werkbook — Marketing Brief

## One-liner

A zero-dependency Go library for reading, writing, and calculating Excel files — with the only compiled formula engine in the Go ecosystem.

## The Problem

Go developers working with Excel files today face a hard tradeoff:

1. **Existing libraries handle data, not logic.** They can read and write cell values, but formulas are opaque strings — there's no evaluation, no dependency tracking, no recalculation. If your pipeline depends on computed values, you either pre-compute everything outside Excel or shell out to LibreOffice.

2. **Performance is an afterthought.** Most Go Excel libraries re-parse formulas on every evaluation (if they evaluate at all), dispatch functions via reflection, and recalculate the entire workbook when a single cell changes.

3. **Dependency creep.** Many libraries pull in dozens of transitive dependencies for functionality most users never touch, inflating binaries and increasing supply-chain risk.

Teams building financial models, report generators, data pipelines, and SaaS products with Excel import/export are stuck stitching together partial solutions.

## The Product

**Werkbook** is a Go library that treats Excel files as living, calculable documents — not just containers of static data.

### What makes it different

- **Compiled formula engine.** Formulas are parsed once, compiled to bytecode, and executed in a stack-based VM. No reflection. No re-parsing. This is the only Go library with a real formula engine.
- **Incremental recalculation.** A dependency graph tracks which cells depend on which. Changing one cell only recalculates the affected subgraph — not every formula in the workbook.
- **Zero dependencies.** The entire library is built on the Go standard library. Nothing else. Your `go.sum` stays clean.
- **Excel-correct.** Matches Excel behavior for type coercion, operator precedence, date serial numbers (including the 1900 leap year bug), and edge cases. Over 1,400 tests verify parity.

## Tiers

### Community Edition (Free, Open Source)

The Community Edition covers the core use case: reading, writing, and working with Excel files in Go, plus a formula engine with the most commonly used functions.

**Includes:**
- Full XLSX read/write (cells, sheets, shared strings, row/cell iteration)
- Core formula engine with compiled bytecode VM and dependency graph
- ~50 most-used Excel functions across categories:
  - **Math:** SUM, AVERAGE, ROUND, ABS, MOD, INT, and more
  - **Logical:** IF, AND, OR, NOT, IFERROR
  - **Text:** LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, CONCATENATE, FIND, SUBSTITUTE
  - **Lookup:** VLOOKUP, HLOOKUP, INDEX, MATCH
  - **Date:** DATE, DAY, MONTH, YEAR, TODAY, NOW
  - **Statistical:** COUNT, COUNTA, MIN, MAX, MEDIAN
  - **Information:** ISBLANK, ISNUMBER, ISTEXT, ISERROR
- Lazy evaluation and incremental recalculation
- Circular reference detection
- Zero external dependencies
- MIT licensed

**Target users:** Individual developers, open source projects, startups, and teams that need reliable Excel I/O with basic formula support.

### Pro Edition (Paid License)

The Pro Edition is Werkbook with everything turned on. Full Excel function coverage, advanced features, and priority support — for teams where spreadsheet logic is business-critical.

**Includes everything in Community, plus:**
- **All 500+ Excel functions**, including:
  - Full Financial suite (IRR, NPV, PMT, XIRR, XNPV, RATE, NPER, FV, PV, and 45+ more)
  - Full Statistical suite (NORM.DIST, T.TEST, LINEST, FORECAST, CORREL, STDEV, and 80+ more)
  - Full Engineering suite (CONVERT, BIN2DEC, HEX2OCT, COMPLEX, BESSELI, and 50+ more)
  - Database functions (DSUM, DCOUNT, DAVERAGE, and more)
  - Advanced Lookup & Reference (FILTER, UNIQUE, SORT, SORTBY, XLOOKUP, XMATCH, dynamic arrays)
  - Web & Cube functions
- Array formulas and dynamic array support
- Streaming reader/writer for large files (100K+ rows)
- Named ranges and defined names
- Advanced features: merged cells, data validation, conditional formatting, auto-filters
- Charts, pivot tables, and images (read/write)
- Priority support and bug fixes
- Commercial license

**Target users:** SaaS platforms, financial services, data engineering teams, enterprise software — anyone building products where Excel compatibility is a requirement, not a nice-to-have.

## Positioning

| | Werkbook Community | Werkbook Pro | Excelize | Go-xlsx |
|---|---|---|---|---|
| Formula evaluation | Yes (compiled VM) | Yes (compiled VM) | No | No |
| Incremental recalc | Yes | Yes | No | No |
| External dependencies | 0 | 0 | Many | Few |
| Function coverage | ~50 core | 500+ complete | N/A | N/A |
| Streaming I/O | — | Yes | Yes | — |
| Financial functions | — | Yes (55+) | N/A | N/A |
| Commercial support | Community | Priority | Community | Community |

## Audience

### Primary

- **Backend/platform engineers** building SaaS products with Excel import/export (HR platforms, accounting tools, reporting dashboards, ERP systems)
- **Data engineers** processing Excel files in Go pipelines who need computed values, not just raw data
- **Financial software teams** that need server-side Excel calculation (loan modeling, pricing, portfolio analytics)

### Secondary

- **DevOps/infrastructure teams** generating Excel reports from monitoring data
- **Consultancies and agencies** building client-facing tools with Excel output

## Key Messages

1. **"The only Go library that actually calculates."** Other libraries read and write Excel. Werkbook reads, writes, and thinks.

2. **"Zero dependencies. Zero compromise."** Built entirely on the Go standard library. No dependency tree to audit, no supply-chain risk, no bloat.

3. **"Change one cell, recalculate one subgraph."** Werkbook's dependency graph and incremental recalculation mean you're not wasting cycles recomputing an entire workbook.

4. **"Excel-correct, not Excel-adjacent."** 1,400+ tests verify behavior parity with Excel, including the weird parts (1900 date bug, type coercion, operator precedence).

## Pricing Direction

| | Community | Pro |
|---|---|---|
| Price | Free (MIT) | Subscription (per-seat or per-org) |
| Model | Open source | Source-available commercial license |
| Support | GitHub Issues | Priority email/Slack, SLA available |

Specific pricing TBD based on market research. Comparable Go/developer-tool pricing:
- Per-developer seat: $20–50/mo
- Per-organization (unlimited devs): $200–500/mo
- Enterprise (custom SLA, on-prem): Custom pricing

## Go-to-Market

### Phase 1 — Community Launch
- Open source the Community Edition on GitHub
- Write a launch blog post: "We built a formula engine for Go"
- Post to Hacker News, r/golang, Go Weekly, Golang Cafe
- Create examples and documentation (amortization calculator, grade book, report generator)
- Target 500 GitHub stars in the first month

### Phase 2 — Build Credibility
- Publish benchmarks (Werkbook formula eval vs. shelling out to LibreOffice)
- Write deep-dive posts on the compiler/VM architecture
- Contribute to Go ecosystem discussions around Excel handling
- Collect testimonials from early adopters

### Phase 3 — Pro Launch
- Announce Pro Edition with full function coverage
- Offer early-adopter pricing (50% off first year)
- Target teams already using Community Edition who hit the function ceiling
- Sales-assisted for enterprise deals

## Success Metrics

| Metric | 3 months | 6 months | 12 months |
|---|---|---|---|
| GitHub stars | 500 | 2,000 | 5,000 |
| Community users (imports) | 100 | 500 | 2,000 |
| Pro paying customers | — | 10 | 50 |
| Pro ARR | — | $20K | $150K |

## Open Questions

1. **Community function boundary** — Which ~50 functions go in Community vs. Pro? Current thinking: the most-used functions per category stay free; specialized functions (financial, engineering, advanced statistical) are Pro-only.
2. **Licensing model** — MIT for Community is clear. Pro could be source-available (BSL/FSL) or proprietary binary. Source-available builds more trust with Go developers.
3. **Naming** — "Werkbook" is distinctive and memorable. Domain availability TBD.
4. **Cloud/SaaS offering** — A hosted calculation API (POST a spreadsheet, GET computed values) could be a separate product line down the road.
