# cmd/web smoke test

`smoke.mjs` runs the built `werkbook.wasm` in Node and checks the JavaScript
API contract of `../README.md`: argument decoding, result shapes and errors.
Run it with `make web-smoke` (it builds first), or
`node cmd/web/smoke/smoke.mjs cmd/web/static` after `make web`.
It prints one line for each check and exits with code 1 on the first failure.
