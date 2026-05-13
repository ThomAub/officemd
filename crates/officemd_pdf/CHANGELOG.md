# Changelog

All notable changes to this project will be documented in this file.

## [0.1.7] - 2026-05-13

### Bug Fixes

- *(pdf)* Guard blockquotes on multi-column pages ([ba18d72](https://github.com/ThomAub/officemd/commit/ba18d7206c7077906a631d2c8a6c884b9fb4cdc3))
- *(pdf)* Re-port French extraction fixes (3008ad1) on top of re-vendor ([6b73a6f](https://github.com/ThomAub/officemd/commit/6b73a6f5bacafae6e5fa302a16f265a2bdc6ba0e))
- *(pdf)* Widen TJ space threshold inside all-caps runs ([de3eefe](https://github.com/ThomAub/officemd/commit/de3eefec4d38859c1729cf46f2e4e264cfdf1900))
- *(pdf)* Treat U+0000 CMap entries as missing in lookup chain ([33ea515](https://github.com/ThomAub/officemd/commit/33ea515b2ee685fb3db8a20555cfa88b49b0c0cf))
- *(pdf)* Reject CMap entries that miscode uppercase as non-canonical lowercase ([db90237](https://github.com/ThomAub/officemd/commit/db90237038c51113a756cc18e4c2a6b5159437f6))

### Features

- *(pdf)* Improve markdown output with alignment, bold breaks, and blockquotes ([98e6bf3](https://github.com/ThomAub/officemd/commit/98e6bf3ed9be75416ec6c66c239305a4d7758680))
- Add officemd_wasm crate for browser-based document-to-markdown conversion ([205205a](https://github.com/ThomAub/officemd/commit/205205abe5687681e333526310003c7a10b7c86f))
- *(pdf)* Surface ocrmypdf invisible-overlay text from TextBased PDFs ([068783c](https://github.com/ThomAub/officemd/commit/068783c522b5b48f31148aa82fbeab9a9db763b7))
- *(pdf)* Collapse substantial letter-spaced items per-page ([a34e6a7](https://github.com/ThomAub/officemd/commit/a34e6a7f5ddfe9470d305aedec3a922d45eef7d2))

### Miscellaneous

- Release v0.1.6 ([1465e2e](https://github.com/ThomAub/officemd/commit/1465e2ee8903f71ff190f7b69cb35944452d0c2b))
- Fix all clippy lints ([77b0740](https://github.com/ThomAub/officemd/commit/77b0740d3fcd8102e079131b0bd873e791eb116d))
- *(pdf)* Re-vendor pdf-inspector from ThomAub fork @ d196d435 ([5457d2f](https://github.com/ThomAub/officemd/commit/5457d2f59bbed49102d20a0f7bd307fa72f95135))
- *(pdf)* Re-port markdown render fixes after re-vendor ([729524a](https://github.com/ThomAub/officemd/commit/729524a625d5a7d570ce45536e47a10cbe660067))

### Testing

- *(pdf)* Re-enable OCR-gap test with corrected semantics ([4f45dce](https://github.com/ThomAub/officemd/commit/4f45dcef2619e51217d9abd1d572890d9eab44e3))
- *(pdf)* Assert parser extracts invisible Tr=3 text on demand ([80b7e1b](https://github.com/ThomAub/officemd/commit/80b7e1bb4839c51d89d7894d5711d5856f700ae5))
- *(pdf)* Add synthetic fixture for visible + invisible text separation ([61ed297](https://github.com/ThomAub/officemd/commit/61ed2978ad6dc9cdb668fe2cee47001aec879360))


## [0.1.6] - 2026-03-30

### Miscellaneous

- Release ([fe411b1](https://github.com/ThomAub/officemd/commit/fe411b17eea9c1eabaa797fda126a336a7cdb54e))


## [0.1.5] - 2026-03-29

### Miscellaneous

- Release v0.1.4 ([1289d01](https://github.com/ThomAub/officemd/commit/1289d011ddbda12b4891b784eaf2b92bf843f979))

