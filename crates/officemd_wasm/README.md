# officemd_wasm

Browser-side WebAssembly bindings for OfficeMD.

## Local demo

The crate ships with a small test page at [www/index.html](./www/index.html). It can:

- convert dropped local files to markdown in a Web Worker
- load sample fixtures from [`examples/data`](../../examples/data)
- exercise the CSV explicit-format path that byte sniffing cannot infer

From [`crates/officemd_wasm`](./):

```bash
rustup target add wasm32-unknown-unknown
cargo install wasm-bindgen-cli
make release
make serve
```

Then open [http://localhost:8080/crates/officemd_wasm/www/](http://localhost:8080/crates/officemd_wasm/www/).

`make serve` intentionally serves the repository root, not just `www/`, so the browser can load:

- `crates/officemd_wasm/www/worker.js`
- the generated `crates/officemd_wasm/pkg/` bundle
- sample fixtures from `examples/data/`
