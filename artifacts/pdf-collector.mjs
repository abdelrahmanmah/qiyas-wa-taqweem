import http from "node:http";
import fs from "node:fs";

http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.end(); return; }
  if (req.method !== "POST") { res.statusCode = 405; res.end(); return; }
  const chunks = [];
  req.on("data", chunk => chunks.push(chunk));
  req.on("end", () => {
    fs.writeFileSync(new URL("./pdf-layout-output.pdf", import.meta.url), Buffer.concat(chunks));
    fs.writeFileSync(new URL("./pdf-logical-pages.txt", import.meta.url), req.url);
    res.end("ok");
  });
}).listen(5175, "127.0.0.1");
