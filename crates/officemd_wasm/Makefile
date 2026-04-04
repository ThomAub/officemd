# Build the WASM package for browser use.
#
# Prerequisites:
#   rustup target add wasm32-unknown-unknown
#   cargo install wasm-bindgen-cli
#   cargo install wasm-opt            # optional, for size/speed optimisation
#
# Usage:
#   make build          # debug build (fast iteration)
#   make release        # optimised release build
#   make serve          # build + serve the demo on http://localhost:8080

WASM_TARGET = wasm32-unknown-unknown
OUT_DIR     = pkg
CARGO_OUT   = ../../target/$(WASM_TARGET)

# panic=abort removes unwinding tables (~10-15% smaller WASM binary).
RELEASE_FLAGS = CARGO_PROFILE_RELEASE_PANIC=abort

.PHONY: build release serve clean

build:
	cargo build --target $(WASM_TARGET) -p officemd_wasm
	wasm-bindgen --target web --out-dir $(OUT_DIR) \
		$(CARGO_OUT)/debug/officemd_wasm.wasm

release:
	$(RELEASE_FLAGS) cargo build --target $(WASM_TARGET) -p officemd_wasm --release
	wasm-bindgen --target web --out-dir $(OUT_DIR) \
		$(CARGO_OUT)/release/officemd_wasm.wasm
	@# Optional: further optimise with wasm-opt (binaryen)
	@command -v wasm-opt >/dev/null 2>&1 && \
		wasm-opt -O3 $(OUT_DIR)/officemd_wasm_bg.wasm -o $(OUT_DIR)/officemd_wasm_bg.wasm \
		|| echo "wasm-opt not found, skipping (install via: cargo install wasm-opt)"

serve: release
	@echo "Serving at http://localhost:8080"
	python3 -m http.server 8080 --directory www

clean:
	rm -rf $(OUT_DIR)
