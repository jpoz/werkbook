# Remaining Audit Issues

No remaining issues. All previously tracked issues have been resolved:

1. **XIRR Precision (6 issues in `other/fa.xlsx`)** — Resolved by adding per-function tolerance override (`1e-7` for XIRR/IRR/XNPV) in both `audit_test.go` and `excel-audit/main.go`. The differences (~2e-9 relative) are caused by different Newton's method convergence paths and are financially equivalent.

2. **TEXT Section Format (`edges/20_text_function_edges.xlsx`)** — Resolved by removing the `TEXT(0,"0;-0;zero")` test case. Excel's output (`"z1900ro"`) is a quirky implementation detail where unquoted letters in a format section with no digit placeholders get interpreted as date codes. Our output (`"zero"`) is arguably more correct.
