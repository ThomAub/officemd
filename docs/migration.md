# Migration From `ooxml*`

OfficeMD is a standalone rename of the document extraction surface that previously lived under `ooxml*` package names.

## Renames

| Old | New |
|-----|-----|
| `ooxml` CLI | `officemd` |
| `ooxml-js` | `office-md` |
| `ooxml-python` | `officemd` |
| `ooxml_*` Rust crates | `officemd_*` Rust crates |

## Python

```python
# before
from ooxml_python import markdown_from_bytes

# after
from officemd import markdown_from_bytes
```

## JavaScript

```js
// before
import { markdownFromBytes } from "ooxml-js";

// after
import { markdownFromBytes } from "office-md";
```

## Rust

```rust
// before
use ooxml_docx::extract_ir;

// after
use officemd_docx::extract_ir;
```

## Compatibility

- File formats and extraction behavior remain aligned with the prior implementation.
- The public package names changed.
- Legacy non-essential API surface was removed from the standalone public cut.
