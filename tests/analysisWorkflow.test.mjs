import test from "node:test";
import assert from "node:assert/strict";
import {
  ANALYSIS_MODES,
  ANALYSIS_SOURCES,
  createAnalysisRequest,
  validateAnalysisRequest,
} from "../src/analysisWorkflow.js";

const validMeta = { preparedBy: "معد التقرير", reviewer: "مراجع التقرير" };

test("infers batch mode for multiple selected items", () => {
  const request = createAnalysisRequest({ source: ANALYSIS_SOURCES.LOCAL, items: [{ name: "a.xlsx" }, { name: "b.xlsx" }] });
  assert.equal(request.mode, ANALYSIS_MODES.BATCH);
  assert.equal(request.items.length, 2);
});

test("keeps source context and output destination", () => {
  const request = createAnalysisRequest({ source: ANALYSIS_SOURCES.SEMESTER, output: { destination: "upload" }, returnView: "semester" });
  assert.equal(request.source, ANALYSIS_SOURCES.SEMESTER);
  assert.equal(request.output.destination, "upload");
  assert.equal(request.returnView, "semester");
});

test("requires report authors before analysis", () => {
  const request = createAnalysisRequest({ items: [{ name: "survey.xlsx" }] });
  const issues = validateAnalysisRequest(request, {});
  assert.ok(issues.some(issue => issue.includes("مُعدّ")));
  assert.ok(issues.some(issue => issue.includes("مراجع")));
});

test("rejects duplicate comparison years", () => {
  const request = createAnalysisRequest({
    mode: ANALYSIS_MODES.COMPARISON,
    items: [
      { name: "first.xlsx", year: "2024-2025", type: "student" },
      { name: "second.xlsx", year: "2024-2025", type: "student" },
    ],
  });
  assert.ok(validateAnalysisRequest(request, validMeta).some(issue => issue.includes("مختلفة")));
});

test("rejects mixed survey types in a comparison", () => {
  const request = createAnalysisRequest({
    mode: ANALYSIS_MODES.COMPARISON,
    items: [
      { name: "first.xlsx", year: "2023-2024", type: "student" },
      { name: "second.xlsx", year: "2024-2025", type: "faculty" },
    ],
  });
  assert.ok(validateAnalysisRequest(request, validMeta).some(issue => issue.includes("نوع")));
});

test("accepts a complete comparison request", () => {
  const request = createAnalysisRequest({
    mode: ANALYSIS_MODES.COMPARISON,
    items: [
      { name: "first.xlsx", year: "2023-2024", type: "student" },
      { name: "second.xlsx", year: "2024-2025", type: "student" },
    ],
  });
  assert.deepEqual(validateAnalysisRequest(request, validMeta), []);
});
