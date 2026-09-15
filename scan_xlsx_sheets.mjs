import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

function decodeXml(s = "") {
  return s.replace(/&quot;/g, '"').replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
}

const root = process.cwd();
const output = execFileSync("powershell", [
  "-NoProfile",
  "-Command",
  "Get-ChildItem -Recurse -Filter *.xlsx | ForEach-Object { $_.FullName }",
], { encoding: "utf8" });

for (const file of output.split(/\r?\n/).filter(Boolean)) {
  const temp = path.join(root, "outputs", `tmp_sheet_scan_${Math.random().toString(36).slice(2)}`);
  const zip = `${temp}.zip`;
  fs.mkdirSync(temp, { recursive: true });
  try {
    fs.copyFileSync(file, zip);
    execFileSync("powershell", [
      "-NoProfile",
      "-Command",
      `Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${temp.replace(/'/g, "''")}' -Force`,
    ], { stdio: "ignore" });
    const wb = fs.readFileSync(path.join(temp, "xl", "workbook.xml"), "utf8");
    const names = [...wb.matchAll(/<sheet\b[^>]*name="([^"]+)"/g)].map((m) => decodeXml(m[1]));
    console.log(`${file} => ${names.join(" | ")}`);
  } catch (error) {
    console.log(`${file} => ERROR: ${error.message}`);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
    fs.rmSync(zip, { force: true });
  }
}
