import test from "node:test";
import assert from "node:assert/strict";

import {
  decodeXmlEntities,
  extractSdkNames,
  normalizeFormula,
  renderSnapshot,
} from "./extract_named_formulas.mjs";

const fixtureXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="SDK_Zeta">_xlfn.LAMBDA(_xlpm.value,_xlpm.value&amp;"!")</definedName>
    <definedName name="Other_Name">Sheet1!$A$1</definedName>
    <definedName name="SDK_Alpha" description="Uses &amp; validates &lt;input&gt;">_xlfn.LAMBDA(_xlpm.board,_xlpm.board&lt;9)</definedName>
  </definedNames>
</workbook>`;

test("decodes named, decimal, and hexadecimal XML entities", () => {
  assert.equal(
    decodeXmlEntities("A&amp;B&lt;C&#10;D&#x21;&quot;&apos;&gt;"),
    "A&B<C\nD!\"'>",
  );
});

test("normalizes storage-only Excel formula prefixes", () => {
  assert.equal(
    normalizeFormula("=_xlfn.LAMBDA(_xlpm.board,_xlpm.board&lt;9)"),
    "=LAMBDA(board,board<9)",
  );
});

test("extracts only non-empty SDK names in stable alphabetical order", () => {
  assert.deepEqual(extractSdkNames(fixtureXml), [
    {
      name: "SDK_Alpha",
      comment: "Uses & validates <input>",
      formula: "=LAMBDA(board,board<9)",
    },
    {
      name: "SDK_Zeta",
      comment: "",
      formula: '=LAMBDA(value,value&"!")',
    },
  ]);
});

test("renders a complete deterministic Markdown snapshot", () => {
  const markdown = renderSnapshot({
    workbookName: "fixture.xlsx",
    version: "1.0.0",
    extractionDate: "2026-07-28",
    sha256: "a".repeat(64),
    names: extractSdkNames(fixtureXml),
  });

  assert.match(markdown, /Formula count: `2`/);
  assert.match(
    markdown,
    /## `SDK_Alpha`[\s\S]*Uses & validates <input>[\s\S]*```excel\n=LAMBDA\(board,board<9\)\n```/,
  );
  assert.ok(markdown.indexOf("SDK_Alpha") < markdown.indexOf("SDK_Zeta"));
  assert.equal(markdown.endsWith("\n\n"), false);
  assert.equal(markdown.endsWith("\n"), true);
});

test("rejects an SDK name with an empty formula", () => {
  const invalidXml = `
    <definedNames>
      <definedName name="SDK_Empty"></definedName>
    </definedNames>`;

  assert.throws(
    () => extractSdkNames(invalidXml),
    /SDK_Empty has an empty formula/,
  );
});
