import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const bytes = fs.readFileSync(new URL("./pdf-layout-output.pdf", import.meta.url));
const outputDir = new URL("./pdf-jpegs/", import.meta.url);
const outputPath = fileURLToPath(outputDir);
fs.mkdirSync(outputPath, { recursive: true });
let offset = 0;
let page = 0;
while (offset < bytes.length - 1) {
  const start = bytes.indexOf(Buffer.from([0xff, 0xd8]), offset);
  if (start < 0) break;
  const end = bytes.indexOf(Buffer.from([0xff, 0xd9]), start + 2);
  if (end < 0) break;
  page++;
  fs.writeFileSync(path.join(outputPath, `page-${String(page).padStart(2, "0")}.jpg`), bytes.subarray(start, end + 2));
  offset = end + 2;
}
console.log(`extracted=${page}`);
