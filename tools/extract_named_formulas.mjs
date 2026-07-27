import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync, writeFileSync } from "node:fs";
import { basename, resolve } from "node:path";
import { pathToFileURL } from "node:url";

export function decodeXmlEntities(text) {
  return text.replace(
    /&(?:#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g,
    (entity) => {
      if (entity.startsWith("&#x")) {
        return String.fromCodePoint(Number.parseInt(entity.slice(3, -1), 16));
      }
      if (entity.startsWith("&#")) {
        return String.fromCodePoint(Number.parseInt(entity.slice(2, -1), 10));
      }
      return {
        "&amp;": "&",
        "&lt;": "<",
        "&gt;": ">",
        "&quot;": '"',
        "&apos;": "'",
      }[entity];
    },
  );
}

export function normalizeFormula(formula) {
  const decoded = decodeXmlEntities(formula.trim())
    .replaceAll("_xlfn.", "")
    .replaceAll("_xlpm.", "");
  return decoded.startsWith("=") ? decoded : `=${decoded}`;
}

function readAttribute(attributes, name) {
  const match = attributes.match(
    new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`),
  );
  return match ? decodeXmlEntities(match[1] ?? match[2]) : "";
}

export function extractSdkNames(xml) {
  const names = [];
  const pattern = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g;

  for (const match of xml.matchAll(pattern)) {
    const attributes = match[1];
    const name = readAttribute(attributes, "name");
    if (!name.startsWith("SDK_")) continue;

    const storedFormula = match[2].trim();
    if (!storedFormula) {
      throw new Error(`${name} has an empty formula`);
    }

    names.push({
      name,
      comment:
        readAttribute(attributes, "comment") ||
        readAttribute(attributes, "description"),
      formula: normalizeFormula(storedFormula),
    });
  }

  return names.sort((left, right) => left.name.localeCompare(right.name, "en"));
}

export function renderSnapshot({
  workbookName,
  version,
  extractionDate,
  sha256,
  names,
}) {
  const lines = [
    "# Sudoku Named LAMBDA Snapshot",
    "",
    "Generated file — do not edit manually. The Excel workbook is the canonical executable artifact; formulas are normalized for display by removing storage-only `_xlfn.` and `_xlpm.` prefixes.",
    "",
    `- Workbook: \`${workbookName}\``,
    `- Workbook version: \`${version}\``,
    `- Workbook SHA-256: \`${sha256}\``,
    `- Extraction date: \`${extractionDate}\``,
    `- Formula count: \`${names.length}\``,
    "",
  ];

  for (const item of names) {
    lines.push(`## \`${item.name}\``, "");
    if (item.comment) lines.push(item.comment, "");
    lines.push("```excel", item.formula, "```", "");
  }

  return `${lines.join("\n").trimEnd()}\n`;
}

function parseArguments(argv) {
  const values = new Map();
  for (let index = 0; index < argv.length; index += 2) {
    const flag = argv[index];
    const value = argv[index + 1];
    if (!flag?.startsWith("--") || value === undefined) {
      throw new Error(`Invalid argument sequence near ${flag ?? "end of input"}`);
    }
    values.set(flag.slice(2), value);
  }

  const required = [
    "workbook",
    "output",
    "version",
    "date",
    "expected-count",
  ];
  for (const key of required) {
    if (!values.has(key)) throw new Error(`Missing required --${key}`);
  }
  if (values.size !== required.length) {
    const unknown = [...values.keys()].filter((key) => !required.includes(key));
    throw new Error(`Unknown argument: --${unknown[0]}`);
  }

  const expectedCount = Number.parseInt(values.get("expected-count"), 10);
  if (!Number.isInteger(expectedCount) || expectedCount < 1) {
    throw new Error("--expected-count must be a positive integer");
  }

  return {
    workbookPath: resolve(values.get("workbook")),
    outputPath: resolve(values.get("output")),
    version: values.get("version"),
    extractionDate: values.get("date"),
    expectedCount,
  };
}

function runCli(argv) {
  const options = parseArguments(argv);
  execFileSync("unzip", ["-tqq", options.workbookPath], { stdio: "pipe" });
  const workbookXml = execFileSync(
    "unzip",
    ["-p", options.workbookPath, "xl/workbook.xml"],
    { encoding: "utf8" },
  );
  if (!workbookXml.includes("<definedNames>")) {
    throw new Error("xl/workbook.xml has no defined-name collection");
  }

  const names = extractSdkNames(workbookXml);
  if (names.length !== options.expectedCount) {
    throw new Error(
      `Found ${names.length} SDK_* names; expected ${options.expectedCount}`,
    );
  }

  const workbookBytes = readFileSync(options.workbookPath);
  const sha256 = createHash("sha256").update(workbookBytes).digest("hex");
  if (!sha256) throw new Error("Could not calculate workbook SHA-256");

  const markdown = renderSnapshot({
    workbookName: basename(options.workbookPath),
    version: options.version,
    extractionDate: options.extractionDate,
    sha256,
    names,
  });
  writeFileSync(options.outputPath, markdown, "utf8");
  process.stdout.write(
    `Wrote ${names.length} formulas to ${options.outputPath}\n`,
  );
}

if (
  process.argv[1] &&
  import.meta.url === pathToFileURL(resolve(process.argv[1])).href
) {
  try {
    runCli(process.argv.slice(2));
  } catch (error) {
    process.stderr.write(`${error.message}\n`);
    process.exitCode = 1;
  }
}
