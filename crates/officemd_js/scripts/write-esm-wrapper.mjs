import { copyFile, writeFile } from "node:fs/promises";

const cjsEntry = "index.cjs";
const esmEntry = "index.js";

await copyFile(esmEntry, cjsEntry);

await writeFile(
  esmEntry,
  `import nativeBinding from "./index.cjs";

export const detectFormat = nativeBinding.detectFormat;
export const extractIrJson = nativeBinding.extractIrJson;
export const markdownFromBytes = nativeBinding.markdownFromBytes;
export const markdownFromBytesBatch = nativeBinding.markdownFromBytesBatch;
export const extractSheetNames = nativeBinding.extractSheetNames;
export const extractTablesIrJson = nativeBinding.extractTablesIrJson;
export const extractCsvTablesIrJson = nativeBinding.extractCsvTablesIrJson;
export const inspectPdfJson = nativeBinding.inspectPdfJson;
export const inspectPdfFontsJson = nativeBinding.inspectPdfFontsJson;
export const doclingFromBytes = nativeBinding.doclingFromBytes;
export const createDocumentFromMarkdown = nativeBinding.createDocumentFromMarkdown;
export const applyOoxmlPatchJson = nativeBinding.applyOoxmlPatchJson;

export default nativeBinding;
`,
);
