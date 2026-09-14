import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  loadGisScript, initSemesterTokenClient,
  SEMESTER_TOKEN_KEY, saveStoredToken, getStoredToken,
  TEMPLATE_FOLDER_ID, ROOT_SURVEYS_FOLDER_ID,
  SEMESTERS, loadDepartments, saveDepartments, expandDepartmentUnits,
  listFormsInFolder, listSubfolders,
  getForm, listAllResponses, computeResponseStats,
  buildGenerationJobs, runGenerationJobs, listSemesterSurveysWithStats,
  yearFolderName, isValidAcademicYear,
  editorResponsesUrl, responsesToRows, departmentFromSurveyName,
} from "./engine/semesterSurveyModel.js";
import { analyze } from "./engine/analyze.js";
import { DEFAULT_SETTINGS } from "./engine/buildDocx.js";
import { buildBrandedReportPdf } from "./engine/buildReportPdf.js";
import { getAllAnalysisSchemas, detectAnySurveyType } from "./engine/customSurveyModel.js";
import { GoogleDriveIcon, InlineNotice } from "./UiElements.jsx";

const SETTINGS_KEY = "eruQA_settings_v1"; // same key App.jsx's SettingsPanel writes to

function loadReportSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_SETTINGS };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Shared "links export" helpers — used by both the Generate results table and the
// Dashboard's survey table, wherever a list of { name, formUrl, responsesUrl } exists.
function buildOrganizedMessage(rows, title) {
  const lines = [`📋 ${title}`, ""];
  rows.forEach((r, i) => {
    lines.push(`${i + 1}. ${r.name}`);
    lines.push(`   🔗 النموذج: ${r.formUrl}`);
    lines.push(`   📊 الردود: ${r.responsesUrl}`);
    lines.push("");
  });
  return lines.join("\n").trim();
}

function buildLinksWorkbookBlob(rows) {
  const data = rows.map(r => ({
    "الاستبيان": r.name,
    "رابط النموذج": r.formUrl,
    "رابط الردود": r.responsesUrl,
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = [{ wch: 45 }, { wch: 55 }, { wch: 55 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "روابط الاستبيانات");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function LinksExportButtons({ rows, title, filename, pushToast }) {
  if (!rows.length) return null;
  return (
    <div style={{ display: "flex", gap: 10 }}>
      <button className="btn btn-ghost btn-sm" onClick={() => {
        copyToClipboard(buildOrganizedMessage(rows, title), pushToast);
      }}>📋 نسخ الروابط كرسالة</button>
      <button className="btn btn-ghost btn-sm" onClick={() => {
        downloadBlob(buildLinksWorkbookBlob(rows), filename);
        pushToast("تم تنزيل ملف Excel.", "success");
      }}>⬇ تنزيل Excel</button>
    </div>
  );
}

const CSS = `
@keyframes ssgShimmer{0%{background-position:-300px 0}100%{background-position:300px 0}}
@keyframes ssgToastIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes ssgProgressFlow{0%{background-position:0 0}100%{background-position:40px 0}}
@keyframes ssgTipFade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
@keyframes ssgOrbit{to{transform:rotate(360deg)}}
@keyframes ssgPulse{0%,100%{transform:scale(.92);opacity:.45}50%{transform:scale(1);opacity:1}}
@keyframes ssgDot{0%,60%,100%{transform:translateY(0);opacity:.35}30%{transform:translateY(-4px);opacity:1}}
@keyframes ssgCardIn{from{opacity:0;transform:translateY(12px) scale(.985)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes ssgGlowMove{0%{transform:translateX(110%)}100%{transform:translateX(-110%)}}
@keyframes ssgSuccessPop{0%{transform:scale(.7);opacity:0}70%{transform:scale(1.08)}100%{transform:scale(1);opacity:1}}
.ssg-home{max-width:1080px;margin:0 auto;padding:20px 0 8px}
.ssg-home-hero{text-align:center;padding:22px 18px 28px}
.ssg-home-badge{display:inline-flex;align-items:center;gap:7px;padding:7px 12px;border-radius:999px;
  color:#5eead4;background:rgba(26,188,156,.1);border:1px solid rgba(94,234,212,.18);font-size:11px;font-weight:800}
.ssg-action-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
.ssg-action-card{position:relative;min-height:210px;padding:24px;border:1px solid rgba(255,255,255,.11);
  border-radius:20px;background:linear-gradient(145deg,rgba(255,255,255,.075),rgba(255,255,255,.035));
  color:#fff;text-align:right;font-family:inherit;cursor:pointer;overflow:hidden;transition:border-color .2s ease,transform .2s ease,background .2s ease}
.ssg-action-card:before{content:"";position:absolute;inset:auto -45px -55px auto;width:150px;height:150px;border-radius:50%;
  background:var(--ssg-glow,rgba(26,188,156,.13));filter:blur(2px);transition:transform .25s ease}
.ssg-action-card:hover{transform:translateY(-3px);border-color:rgba(94,234,212,.35);background:linear-gradient(145deg,rgba(255,255,255,.1),rgba(255,255,255,.045))}
.ssg-action-card:hover:before{transform:scale(1.16)}
.ssg-action-icon{width:48px;height:48px;display:grid;place-items:center;border-radius:15px;margin-bottom:22px;
  color:var(--ssg-accent,#5eead4);background:color-mix(in srgb,var(--ssg-accent,#5eead4) 14%,transparent);border:1px solid color-mix(in srgb,var(--ssg-accent,#5eead4) 24%,transparent)}
.ssg-action-title{font-size:17px;font-weight:900;margin-bottom:8px;position:relative}
.ssg-action-copy{font-size:12px;line-height:1.9;color:rgba(255,255,255,.56);position:relative}
.ssg-action-arrow{position:absolute;left:20px;bottom:18px;color:var(--ssg-accent,#5eead4);font-size:18px}
.ssg-secondary-row{display:flex;justify-content:center;margin-top:16px}
.ssg-create-hero{display:flex;align-items:center;justify-content:space-between;gap:18px;padding:22px 24px;margin-bottom:16px;border-radius:20px;
  border:1px solid rgba(94,234,212,.16);background:linear-gradient(125deg,rgba(26,188,156,.12),rgba(40,116,166,.08))}
.ssg-create-kicker{color:#67e8d0;font-size:10px;font-weight:900;margin-bottom:5px}.ssg-create-hero h2{color:#fff;font-size:20px;font-weight:950;margin:0 0 5px}
.ssg-create-hero p{color:rgba(255,255,255,.48);font-size:11.5px;line-height:1.75;margin:0}
.ssg-step-pills{display:flex;gap:7px;flex-wrap:wrap}.ssg-step-pill{display:flex;align-items:center;gap:6px;padding:7px 10px;border-radius:999px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);color:rgba(255,255,255,.52);font-size:10px;font-weight:800;white-space:nowrap}
.ssg-step-pill b{width:20px;height:20px;display:grid;place-items:center;border-radius:50%;color:#bafff1;background:rgba(26,188,156,.15)}
.ssg-create-layout{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(280px,.65fr);gap:16px;align-items:start}
.ssg-panel{padding:20px;border-radius:18px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.045)}
.ssg-panel-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:15px}.ssg-panel-title{color:#fff;font-size:14px;font-weight:900}.ssg-panel-copy{color:rgba(255,255,255,.42);font-size:10.5px;margin-top:2px}
.ssg-template-tools{display:flex;align-items:center;gap:8px;margin-bottom:11px}.ssg-template-search{flex:1;min-width:0;padding:9px 12px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(4,18,34,.35);color:#fff;font-family:inherit;font-size:11.5px}
.ssg-template-list{display:flex;flex-direction:column;gap:7px;max-height:470px;overflow:auto;padding-left:3px}
.ssg-template-card{display:grid;grid-template-columns:auto minmax(0,1fr) 138px;align-items:center;gap:11px;padding:11px 12px;border-radius:12px;border:1px solid rgba(255,255,255,.075);background:rgba(255,255,255,.035);transition:.18s}
.ssg-template-card.selected{border-color:rgba(94,234,212,.24);background:rgba(26,188,156,.075)}
.ssg-template-name{color:#e8f0fe;font-size:11.5px;font-weight:700;line-height:1.65}.ssg-template-card select{width:100%;padding:7px 8px;border-radius:8px;border:1px solid rgba(255,255,255,.13);background:#142d47;color:#e8f0fe;font-family:inherit;font-size:10px}
.ssg-create-aside{position:sticky;top:94px}.ssg-field-label{display:block;color:rgba(255,255,255,.54);font-size:10.5px;font-weight:700;margin-bottom:6px}.ssg-field{width:100%;padding:10px 12px;border-radius:10px;border:1px solid rgba(255,255,255,.14);background:rgba(4,18,34,.35);color:#fff;font-family:inherit;font-size:12px}
.ssg-selection-summary{display:flex;justify-content:space-between;align-items:center;padding:12px;margin:15px 0;border-radius:11px;background:rgba(26,188,156,.08);color:rgba(255,255,255,.6);font-size:11px}.ssg-selection-summary strong{color:#7cebd5;font-size:18px}
.ssg-create-submit{width:100%;justify-content:center}.ssg-results-panel{margin-top:16px}
.ssg-generation-stage{position:relative;overflow:hidden;padding:24px;border-radius:18px;border:1px solid rgba(94,234,212,.16);background:linear-gradient(130deg,rgba(6,28,47,.74),rgba(17,71,68,.42))}
.ssg-generation-stage:after{content:"";position:absolute;inset:0 auto 0 0;width:36%;background:linear-gradient(90deg,transparent,rgba(94,234,212,.035),transparent);animation:ssgGlowMove 2.6s linear infinite;pointer-events:none}
.ssg-generation-widget{position:fixed;left:24px;bottom:82px;width:min(620px,calc(100vw - 48px));max-height:calc(100vh - 120px);z-index:160;
  box-sizing:border-box;background:linear-gradient(135deg,rgba(6,28,47,.97),rgba(13,66,61,.96));border-color:rgba(94,234,212,.3);
  box-shadow:0 24px 70px rgba(0,0,0,.46),0 0 0 1px rgba(94,234,212,.06);backdrop-filter:blur(20px);animation:ssgCardIn .3s ease both}
.ssg-generation-widget .ssg-job-grid{max-height:min(210px,32vh)}
.ssg-generation-widget.collapsed{width:min(450px,calc(100vw - 48px));padding:16px 18px}
.ssg-generation-widget.collapsed .ssg-generation-head{margin-bottom:10px}
.ssg-generation-widget.collapsed .ssg-current-job,.ssg-generation-widget.collapsed .ssg-job-grid{display:none}
.ssg-widget-toggle{position:relative;z-index:2;width:30px;height:30px;display:grid;place-items:center;flex:0 0 30px;border-radius:9px;
  border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.06);color:#d8fff7;font-family:inherit;font-size:17px;line-height:1;cursor:pointer}
.ssg-widget-toggle:hover{background:rgba(94,234,212,.12);border-color:rgba(94,234,212,.25)}
.ssg-generation-head{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:18px}.ssg-generation-title{display:flex;align-items:center;gap:13px;color:#fff;font-size:14px;font-weight:900}.ssg-generation-percent{color:#67e8d0;font-size:24px;font-weight:950}
.ssg-current-job{padding:11px 13px;margin-top:12px;border-radius:11px;color:#d8fff7;background:rgba(26,188,156,.09);font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ssg-job-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:7px;margin-top:14px;max-height:210px;overflow:auto}.ssg-job-chip{display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:10px;background:rgba(255,255,255,.035);color:rgba(255,255,255,.48);font-size:10.5px;white-space:nowrap;overflow:hidden}.ssg-job-chip span:last-child{overflow:hidden;text-overflow:ellipsis}.ssg-job-chip.done{color:#9ff5df;background:rgba(16,185,129,.07)}.ssg-job-chip.active{color:#fff;background:rgba(59,130,246,.1)}
.ssg-status-dot{width:8px;height:8px;flex:0 0 8px;border-radius:50%;background:rgba(255,255,255,.22)}.ssg-job-chip.done .ssg-status-dot{background:#34d399}.ssg-job-chip.active .ssg-status-dot{background:#60a5fa;box-shadow:0 0 0 5px rgba(96,165,250,.1);animation:ssgPulse 1.2s infinite}.ssg-job-chip.error .ssg-status-dot{background:#fb7185}
.ssg-success-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:16px;padding-bottom:15px;border-bottom:1px solid rgba(255,255,255,.07)}.ssg-success-copy{display:flex;align-items:center;gap:11px;color:#a7f3d0;font-weight:900}.ssg-success-icon{width:38px;height:38px;display:grid;place-items:center;border-radius:50%;background:rgba(16,185,129,.14);color:#6ee7b7;animation:ssgSuccessPop .45s ease-out}
.ssg-created-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:10px}.ssg-created-card{padding:15px;border-radius:14px;border:1px solid rgba(255,255,255,.085);background:rgba(255,255,255,.035);animation:ssgCardIn .35s ease both}.ssg-created-name{color:#f1f5f9;font-size:11.5px;font-weight:800;line-height:1.75;min-height:40px}.ssg-created-actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:11px}.ssg-link-action{display:inline-flex;align-items:center;gap:5px;padding:6px 9px;border-radius:8px;border:1px solid rgba(94,234,212,.14);background:rgba(26,188,156,.065);color:#70e8d1;font-family:inherit;font-size:9.5px;font-weight:700;text-decoration:none;cursor:pointer}
.ssg-dashboard-toolbar{display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap;padding:17px 18px;margin-bottom:14px;border-radius:17px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.045)}
.ssg-dashboard-actions{display:flex;align-items:center;gap:8px;margin-right:auto;flex-wrap:wrap}.ssg-selection-count{display:inline-flex;align-items:center;gap:6px;color:#83ead6;font-size:11px;font-weight:900;padding:7px 10px;border-radius:999px;background:rgba(26,188,156,.09)}
.ssg-stat-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:15px}.ssg-stat-card{animation:ssgCardIn .35s ease both;position:relative;overflow:hidden}.ssg-stat-card:after{content:"";position:absolute;width:80px;height:80px;border-radius:50%;left:-30px;bottom:-42px;background:rgba(94,234,212,.055)}
.ssg-survey-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(315px,1fr));gap:11px}.ssg-survey-card{position:relative;min-height:190px;display:flex;flex-direction:column;padding:17px;border-radius:17px;border:1px solid rgba(255,255,255,.09);background:linear-gradient(145deg,rgba(255,255,255,.055),rgba(255,255,255,.025));cursor:pointer;transition:transform .2s ease,border-color .2s ease,background .2s ease,box-shadow .2s ease;animation:ssgCardIn .4s ease both;outline:none}
.ssg-survey-card:hover,.ssg-survey-card:focus-visible{transform:translateY(-3px);border-color:rgba(96,165,250,.3);box-shadow:0 15px 34px rgba(2,12,27,.2)}.ssg-survey-card.selected{border-color:rgba(94,234,212,.48);background:linear-gradient(145deg,rgba(26,188,156,.14),rgba(40,116,166,.06));box-shadow:inset 0 0 0 1px rgba(94,234,212,.12)}
.ssg-card-select{position:absolute;left:13px;top:13px;width:25px;height:25px;display:grid;place-items:center;border-radius:9px;color:rgba(255,255,255,.34);background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.09);transition:.18s}.ssg-survey-card.selected .ssg-card-select{color:#052e2b;background:#5eead4;border-color:#5eead4;animation:ssgSuccessPop .3s ease}
.ssg-survey-type{width:max-content;max-width:calc(100% - 35px);padding:4px 8px;border-radius:999px;color:#8eead9;background:rgba(26,188,156,.08);font-size:8.5px;font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.ssg-survey-name{color:#f4f8ff;font-size:12.5px;font-weight:900;line-height:1.75;margin:12px 0;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.ssg-survey-metrics{display:flex;gap:8px;margin-top:auto}.ssg-survey-metric{flex:1;padding:9px;border-radius:10px;background:rgba(5,19,34,.28)}.ssg-survey-metric strong{display:block;color:#fff;font-size:17px}.ssg-survey-metric span{display:block;color:rgba(255,255,255,.38);font-size:8.5px;margin-top:1px}
.ssg-survey-card-actions{display:flex;align-items:center;gap:6px;margin-top:11px;opacity:0;transform:translateY(3px);transition:.18s}.ssg-survey-card:hover .ssg-survey-card-actions,.ssg-survey-card:focus-within .ssg-survey-card-actions,.ssg-survey-card.selected .ssg-survey-card-actions{opacity:1;transform:none}.ssg-card-action{padding:6px 9px;border-radius:8px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.055);color:#dce8f6;font-family:inherit;font-size:9px;font-weight:750;cursor:pointer;text-decoration:none}.ssg-card-action.primary{margin-left:auto;color:#8cf2dc;border-color:rgba(94,234,212,.2);background:rgba(26,188,156,.1)}
.ssg-subnav{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;margin-bottom:18px;
  border:1px solid rgba(255,255,255,.09);border-radius:15px;background:rgba(255,255,255,.035)}
.ssg-subnav-title{display:flex;align-items:center;gap:10px;color:#fff;font-size:14px;font-weight:850}
.ssg-loading-glyph{position:relative;width:58px;height:58px;display:grid;place-items:center;margin:0 auto 14px;color:#5eead4}
.ssg-loading-glyph.compact{width:34px;height:34px;margin:0}
.ssg-loading-ring{position:absolute;inset:0;border-radius:50%;border:2px solid rgba(94,234,212,.14);border-top-color:#2dd4bf;animation:ssgOrbit 1.1s linear infinite}
.ssg-loading-ring.second{inset:6px;border-top-color:#60a5fa;animation-duration:1.55s;animation-direction:reverse}
.ssg-loading-core{width:12px;height:12px;border-radius:4px;background:linear-gradient(135deg,#2dd4bf,#3b82f6);animation:ssgPulse 1.4s ease-in-out infinite}
.ssg-loading-line{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.ssg-loading-dots{display:inline-flex;gap:3px;margin-inline-start:5px}
.ssg-loading-dots i{width:4px;height:4px;border-radius:50%;background:#5eead4;animation:ssgDot 1.15s ease-in-out infinite}
.ssg-loading-dots i:nth-child(2){animation-delay:.16s}.ssg-loading-dots i:nth-child(3){animation-delay:.32s}
.ssg-progress-track{background:rgba(255,255,255,.08);border-radius:6px;height:10px;overflow:hidden}
.ssg-progress-fill{height:100%;border-radius:6px;transition:width .4s ease;
  background-image:linear-gradient(135deg,rgba(255,255,255,.2) 25%,transparent 25%,transparent 50%,
    rgba(255,255,255,.2) 50%,rgba(255,255,255,.2) 75%,transparent 75%,transparent);
  background-color:#1abc9c;background-size:28px 28px;animation:ssgProgressFlow 1s linear infinite}
.ssg-tip{animation:ssgTipFade .35s ease-out}
.ssg-skel-row{height:38px;border-radius:10px;margin-bottom:8px;
  background:linear-gradient(90deg,rgba(255,255,255,.05) 25%,rgba(255,255,255,.12) 37%,rgba(255,255,255,.05) 63%);
  background-size:600px 100%;animation:ssgShimmer 1.4s infinite linear}
.ssg-tpl-row{display:flex;align-items:center;justify-content:space-between;gap:14px;
  padding:12px 16px;border-radius:12px;background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.08);margin-bottom:8px}
.ssg-tpl-row input[type=checkbox]{width:17px;height:17px;accent-color:#1abc9c;cursor:pointer}
.ssg-tpl-dept{display:flex;align-items:center;gap:8px;font-size:12px;color:rgba(255,255,255,.55);white-space:nowrap}
.ssg-stat-card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);
  border-radius:14px;padding:16px 18px}
.ssg-toast{animation:ssgToastIn .25s ease-out;padding:12px 18px;border-radius:12px;
  font-size:13px;font-weight:700;color:#fff;box-shadow:0 8px 24px rgba(0,0,0,.35);min-width:220px}
.ssg-departments{max-width:1120px;margin:0 auto;padding-bottom:44px}
.ssg-dept-hero{display:grid;grid-template-columns:minmax(0,1fr) auto;align-items:center;gap:24px;padding:24px 26px;margin-bottom:16px;border-radius:22px;border:1px solid rgba(94,234,212,.18);background:linear-gradient(125deg,rgba(26,188,156,.14),rgba(40,116,166,.08))}
.ssg-dept-kicker{display:flex;align-items:center;gap:7px;color:#72ead4;font-size:11px;font-weight:900;margin-bottom:7px}.ssg-dept-hero h2{color:#fff;font-size:23px;font-weight:950;margin:0 0 7px}.ssg-dept-hero p{max-width:680px;color:rgba(255,255,255,.64);font-size:13px;line-height:1.85;margin:0}
.ssg-dept-stats{display:grid;grid-template-columns:repeat(2,minmax(105px,1fr));gap:9px}.ssg-dept-stat{padding:13px 15px;border-radius:14px;border:1px solid rgba(255,255,255,.09);background:rgba(4,18,34,.28)}.ssg-dept-stat strong{display:block;color:#fff;font-size:22px;line-height:1.2}.ssg-dept-stat span{color:rgba(255,255,255,.55);font-size:10.5px}
.ssg-dept-add{display:grid;grid-template-columns:minmax(0,1fr) auto;align-items:end;gap:12px;padding:18px 20px;margin-bottom:22px;border-radius:18px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.045)}.ssg-dept-add-copy strong{display:block;color:#fff;font-size:14px}.ssg-dept-add-copy span{display:block;color:rgba(255,255,255,.55);font-size:11.5px;margin-top:3px}.ssg-dept-add-form{display:flex;gap:8px;min-width:min(480px,48vw)}
.ssg-dept-input{width:100%;min-height:44px;padding:10px 13px;border-radius:11px;border:1px solid rgba(255,255,255,.15);background:rgba(4,18,34,.42);color:#fff;font-family:inherit;font-size:13px;outline:none}.ssg-dept-input:focus{border-color:rgba(94,234,212,.5);box-shadow:0 0 0 3px rgba(26,188,156,.1)}
.ssg-dept-section-head{display:flex;align-items:end;justify-content:space-between;gap:12px;margin:0 2px 12px}.ssg-dept-section-head h3{color:#fff;font-size:17px;margin:0}.ssg-dept-section-head span{color:rgba(255,255,255,.52);font-size:11.5px}
.ssg-dept-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:13px}.ssg-dept-card{--dept-accent:#5eead4;position:relative;overflow:hidden;padding:18px;border-radius:18px;border:1px solid rgba(255,255,255,.1);background:linear-gradient(145deg,rgba(255,255,255,.065),rgba(255,255,255,.03))}.ssg-dept-card:before{content:"";position:absolute;right:0;top:0;width:4px;height:100%;background:var(--dept-accent)}
.ssg-dept-card-head{display:grid;grid-template-columns:auto minmax(0,1fr) auto;align-items:center;gap:10px;padding-bottom:14px;border-bottom:1px solid rgba(255,255,255,.07)}.ssg-dept-index{width:38px;height:38px;display:grid;place-items:center;border-radius:12px;color:var(--dept-accent);background:color-mix(in srgb,var(--dept-accent) 13%,transparent);border:1px solid color-mix(in srgb,var(--dept-accent) 24%,transparent);font-size:12px;font-weight:900}.ssg-dept-name{min-width:0;background:transparent;border:1px solid transparent;border-radius:9px;padding:7px 9px;color:#fff;font-family:inherit;font-size:14px;font-weight:900;outline:none}.ssg-dept-name:hover{background:rgba(255,255,255,.035);border-color:rgba(255,255,255,.08)}.ssg-dept-name:focus{background:rgba(4,18,34,.38);border-color:rgba(94,234,212,.34)}
.ssg-dept-delete{width:38px;height:38px;display:grid;place-items:center;border-radius:11px;border:1px solid rgba(248,113,113,.17);background:rgba(239,68,68,.07);color:#fca5a5;cursor:pointer}.ssg-dept-delete:hover{background:rgba(239,68,68,.15);border-color:rgba(248,113,113,.3)}
.ssg-program-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:14px 0 10px;color:rgba(255,255,255,.65);font-size:11.5px}.ssg-program-count{padding:3px 8px;border-radius:999px;color:#88eedb;background:rgba(26,188,156,.09);font-size:10px;font-weight:800}.ssg-program-list{display:flex;flex-wrap:wrap;gap:7px;min-height:38px}.ssg-program-chip{display:inline-flex;align-items:center;gap:7px;padding:6px 8px 6px 11px;border-radius:999px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.06);color:#e8f0fe;font-size:11.5px}.ssg-program-remove{width:20px;height:20px;display:grid;place-items:center;border:0;border-radius:50%;background:rgba(239,68,68,.13);color:#ffaaa2;cursor:pointer;font-size:10px}.ssg-program-empty{width:100%;padding:10px 12px;border-radius:10px;border:1px dashed rgba(255,255,255,.1);color:rgba(255,255,255,.42);font-size:11.5px;text-align:center}.ssg-program-add{display:flex;gap:7px;margin-top:13px}.ssg-program-add .ssg-dept-input{min-height:40px;font-size:12px}
.ssg-delete-confirm{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:13px;padding:10px 11px;border-radius:11px;border:1px solid rgba(248,113,113,.2);background:rgba(239,68,68,.08);color:#ffd1cd;font-size:11.5px}.ssg-delete-actions{display:flex;gap:6px;flex-shrink:0}
.ssg-dept-empty{grid-column:1/-1;padding:42px 20px;border-radius:18px;border:1px dashed rgba(255,255,255,.13);color:rgba(255,255,255,.55);text-align:center}
@media(max-width:920px){.ssg-create-layout{grid-template-columns:1fr}.ssg-create-aside{position:static}.ssg-create-hero{align-items:flex-start;flex-direction:column}.ssg-stat-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:820px){.ssg-action-grid{grid-template-columns:1fr}.ssg-action-card{min-height:170px}.ssg-home-hero{padding-top:8px}.ssg-subnav{align-items:flex-start}.ssg-subnav-title{font-size:12px}.ssg-template-card{grid-template-columns:auto minmax(0,1fr)}.ssg-template-card select{grid-column:2}.ssg-dept-hero{grid-template-columns:1fr}.ssg-dept-stats{max-width:300px}.ssg-dept-add{grid-template-columns:1fr}.ssg-dept-add-form{min-width:0}.ssg-dept-grid{grid-template-columns:1fr}}
@media(max-width:560px){.ssg-stat-grid{grid-template-columns:1fr 1fr}.ssg-survey-grid,.ssg-created-grid{grid-template-columns:1fr}.ssg-dashboard-actions{margin-right:0}.ssg-survey-card-actions{opacity:1;transform:none}.ssg-generation-widget,.ssg-generation-widget.collapsed{left:10px;bottom:74px;width:calc(100vw - 20px);padding:16px}.ssg-generation-title{font-size:12px}.ssg-generation-percent{font-size:20px}.ssg-generation-widget .ssg-job-grid{grid-template-columns:1fr;max-height:26vh}.ssg-dept-hero{padding:19px}.ssg-dept-add{padding:15px}.ssg-dept-add-form,.ssg-program-add{flex-direction:column}.ssg-dept-card{padding:15px}.ssg-dept-card-head{grid-template-columns:auto minmax(0,1fr)}.ssg-dept-delete{grid-column:1/-1;width:100%;height:38px}.ssg-delete-confirm{align-items:flex-start;flex-direction:column}.ssg-delete-actions{width:100%}.ssg-delete-actions .btn{flex:1}}
`;

// ---------- toasts ----------
function useToasts() {
  const [toasts, setToasts] = useState([]);
  const push = useCallback((text, kind = "success") => {
    const id = Date.now() + Math.random();
    setToasts(t => [...t, { id, text, kind }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 4000);
  }, []);
  const dismiss = useCallback(id => setToasts(t => t.filter(x => x.id !== id)), []);
  return { toasts, push, dismiss };
}

function ToastStack({ toasts, onDismiss }) {
  if (!toasts.length) return null;
  return (
    <div style={{ position: "fixed", bottom: 24, left: 24, zIndex: 9999, display: "flex", flexDirection: "column", gap: 10 }}>
      {toasts.map(t => (
        <div
          key={t.id}
          className="ssg-toast"
          style={{ background: t.kind === "error" ? "linear-gradient(135deg,#314a68,#243b55)" : "linear-gradient(135deg,#1abc9c,#16a085)", border: t.kind === "error" ? "1px solid rgba(147,197,253,.24)" : "none", cursor: "pointer" }}
          onClick={() => onDismiss(t.id)}
        >
          {t.text}
        </div>
      ))}
    </div>
  );
}

// ---------- auth ----------
// Persists the access token across reloads (localStorage, ~1hr lifetime) and, once it
// expires, tries a silent requestAccessToken({prompt:""}) in the background before ever
// falling back to the visible "الاتصال بـ Google" button — so a returning user only sees
// the account picker once, not on every visit. See semesterSurveyModel.js for why.
function useSemesterAuth(enabled = true) {
  const [token, setToken] = useState(() => getStoredToken(SEMESTER_TOKEN_KEY));
  const [connecting, setConnecting] = useState(false);
  const [authError, setAuthError] = useState("");
  const clientRef = useRef(null);
  const explicitRef = useRef(false);
  const triedSilent = useRef(false);

  const ensureClient = useCallback(async () => {
    await loadGisScript();
    if (!clientRef.current) {
      clientRef.current = initSemesterTokenClient({
        onToken: (t, expiresIn) => {
          setToken(t);
          saveStoredToken(SEMESTER_TOKEN_KEY, t, expiresIn);
          setConnecting(false);
        },
        onError: e => {
          setConnecting(false);
          // A silent background attempt failing is normal (no prior session/consent) —
          // only surface an error banner for an explicit, user-clicked connect.
          if (explicitRef.current) setAuthError("فشل الاتصال بـ Google: " + e);
        },
      });
    }
    return clientRef.current;
  }, []);

  useEffect(() => {
    if (!enabled || token || triedSilent.current) return;
    triedSilent.current = true;
    (async () => {
      try {
        const client = await ensureClient();
        explicitRef.current = false;
        client.requestAccessToken({ prompt: "" });
      } catch { /* ignore — falls through to the manual connect button */ }
    })();
  }, [enabled, token, ensureClient]);

  const connect = useCallback(async () => {
    setAuthError(""); setConnecting(true);
    try {
      const client = await ensureClient();
      explicitRef.current = true;
      client.requestAccessToken({ prompt: "select_account" });
    } catch (e) {
      setAuthError(e.message); setConnecting(false);
    }
  }, [ensureClient]);

  return { token, connecting, authError, connect };
}

// ---------- small shared pieces ----------
function Skeleton({ rows = 4 }) {
  return (
    <div>
      {Array.from({ length: rows }).map((_, i) => <div key={i} className="ssg-skel-row" />)}
    </div>
  );
}

function LoadingGlyph({ compact = false }) {
  return (
    <span className={`ssg-loading-glyph${compact ? " compact" : ""}`} aria-hidden="true">
      <span className="ssg-loading-ring" />
      <span className="ssg-loading-ring second" />
      <span className="ssg-loading-core" />
    </span>
  );
}

function LoadingDots() {
  return <span className="ssg-loading-dots" aria-hidden="true"><i /><i /><i /></span>;
}

function DashboardLoading({ label = "جاري تجهيز لوحة التحكم" }) {
  return (
    <div className="ssg-panel" style={{ padding: 28, textAlign: "center" }} aria-live="polite">
      <LoadingGlyph />
      <div style={{ color: "#eafffb", fontSize: 13, fontWeight: 850 }}>{label}<LoadingDots /></div>
      <div style={{ color: "rgba(255,255,255,.36)", fontSize: 10.5, marginTop: 5 }}>نجمع أحدث الاستبيانات والاستجابات من Google Drive.</div>
      <div className="ssg-created-grid" style={{ marginTop: 20, textAlign: "right" }}><Skeleton rows={3} /><Skeleton rows={3} /><Skeleton rows={3} /></div>
    </div>
  );
}

function EmptyState({ icon = "📭", text }) {
  return (
    <div style={{ textAlign: "center", padding: "36px 20px", color: "rgba(255,255,255,.45)" }}>
      <div style={{ fontSize: 34, marginBottom: 8 }}>{icon}</div>
      <div style={{ fontSize: 13 }}>{text}</div>
    </div>
  );
}

function ErrorBanner({ text }) {
  return <InlineNotice text={text} style={{ marginTop: 10 }} />;
}

function BarChart({ data, gradient = "linear-gradient(90deg,#1abc9c,#2874a6)" }) {
  const max = Math.max(1, ...data.map(d => d.value));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {data.map(d => (
        <div key={d.label}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "rgba(255,255,255,.7)", marginBottom: 4 }}>
            <span>{d.label}</span><span>{d.value}</span>
          </div>
          <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 8, overflow: "hidden" }}>
            <div style={{ height: "100%", borderRadius: 3, transition: "width .5s", background: gradient, width: `${(d.value / max) * 100}%` }} />
          </div>
        </div>
      ))}
    </div>
  );
}

// Cycles through short status lines while a long batch operation runs, so the wait
// doesn't feel like a frozen screen.
function useRotatingTip(active, tips, intervalMs = 2600) {
  const [idx, setIdx] = useState(0);
  useEffect(() => {
    if (!active) { setIdx(0); return; }
    const t = setInterval(() => setIdx(i => (i + 1) % tips.length), intervalMs);
    return () => clearInterval(t);
  }, [active, tips, intervalMs]);
  return tips[idx];
}

function ProgressBar({ done, total, label, tip }) {
  const pct = total ? Math.round((done / total) * 100) : 0;
  return (
    <div style={{ marginBottom: 16 }}>
      <div className="ssg-loading-line" style={{ justifyContent: "space-between", fontSize: 12, color: "rgba(255,255,255,.7)" }}>
        <span style={{ display: "flex", alignItems: "center", gap: 9 }}><LoadingGlyph compact /> <span>{label}<LoadingDots /></span></span>
        <span>{done} / {total} — {pct}%</span>
      </div>
      <div className="ssg-progress-track">
        <div className="ssg-progress-fill" style={{ width: `${pct}%` }} />
      </div>
      {tip && <div key={tip} className="ssg-tip" style={{ textAlign: "center", color: "rgba(255,255,255,.5)", fontSize: 12, marginTop: 10 }}>{tip}</div>}
    </div>
  );
}

function copyToClipboard(text, pushToast) {
  navigator.clipboard?.writeText(text)
    .then(() => pushToast("تم نسخ الرابط.", "success"))
    .catch(() => pushToast("تعذّر نسخ الرابط.", "error"));
}

function formatWhen(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  const now = new Date();
  const days = Math.floor((now - d) / 86400000);
  if (days <= 0) return "اليوم";
  if (days === 1) return "أمس";
  return d.toLocaleDateString("ar-EG", { year: "numeric", month: "short", day: "numeric" });
}

// ---------- Generate Surveys ----------
const GENERATE_TIPS = [
  "🔄 جاري نسخ النماذج من القوالب...",
  "📁 جاري تنظيم مجلدات السنة والفصل الدراسي على Google Drive...",
  "✏️ جاري تحديث عنوان كل استبيان...",
  "🚀 جاري نشر الاستبيانات ليصبح بإمكانها استقبال الردود...",
  "⏳ لحظات ونكون قد انتهينا...",
];

function GenerateSurveysView({ token, pushToast, onReconnect, onGoDashboard }) {
  const [templates, setTemplates] = useState(null); // null = loading
  const [loadError, setLoadError] = useState("");
  const [year, setYear] = useState("");
  const [yearOptions, setYearOptions] = useState([]);
  const [semester, setSemester] = useState(SEMESTERS[0]);
  const [validationMsg, setValidationMsg] = useState("");
  const [jobs, setJobs] = useState(null);
  const [generating, setGenerating] = useState(false);
  const [results, setResults] = useState(null);
  const [templateSearch, setTemplateSearch] = useState("");
  const [templateReloadKey, setTemplateReloadKey] = useState(0);
  const [progressCollapsed, setProgressCollapsed] = useState(false);
  const tip = useRotatingTip(generating, GENERATE_TIPS);

  useEffect(() => {
    let cancelled = false;
    setTemplates(null); setLoadError("");
    listFormsInFolder(token, TEMPLATE_FOLDER_ID)
      .then(files => { if (!cancelled) setTemplates(files.map(f => ({ ...f, selected: false, mode: "general" }))); })
      .catch(e => { if (!cancelled) { setTemplates([]); setLoadError("فشل تحميل القوالب: " + e.message); } });
    return () => { cancelled = true; };
  }, [token, templateReloadKey]);

  // Suggestions only — the year field stays free-text so a new year can always be typed.
  useEffect(() => {
    listSubfolders(token, ROOT_SURVEYS_FOLDER_ID)
      .then(fs => setYearOptions(fs.map(f => f.name)))
      .catch(() => {});
  }, [token]);

  const allSelected = !!templates?.length && templates.every(t => t.selected);
  const toggleAll = () => setTemplates(ts => ts.map(t => ({ ...t, selected: !allSelected })));
  const toggleOne = id => setTemplates(ts => ts.map(t => t.id === id ? { ...t, selected: !t.selected } : t));
  const setTemplateMode = (id, mode) => setTemplates(ts => ts.map(t => t.id === id ? { ...t, mode } : t));
  const selectedCount = templates?.filter(t => t.selected).length ?? 0;
  const visibleTemplates = templates?.filter(t => t.name.toLowerCase().includes(templateSearch.trim().toLowerCase())) ?? [];
  const finishedJobs = jobs?.filter(j => j.status === "done" || j.status === "error").length ?? 0;
  const generationPct = jobs?.length ? Math.round((finishedJobs / jobs.length) * 100) : 0;
  const activeJob = jobs?.find(j => j.status === "active");

  async function handleGenerate() {
    setValidationMsg("");
    if (selectedCount === 0) { setValidationMsg("الرجاء اختيار استبيان واحد على الأقل."); return; }
    if (!isValidAcademicYear(year)) { setValidationMsg("الرجاء إدخال السنة الدراسية بالصيغة 2026/2027."); return; }

    const departments = loadDepartments();
    const selected = templates.filter(t => t.selected);
    const jobList = buildGenerationJobs(selected, departments);

    setJobs(jobList.map(j => ({
      key: j.key, status: "pending",
      label: j.department ? `${j.template.name} — ${j.department}` : j.template.name,
    })));
    setResults(null);
    setProgressCollapsed(false);
    setGenerating(true);

    try {
      const { successRows, failCount } = await runGenerationJobs(
        token, jobList, { year, semester },
        (key, status) => setJobs(js => js.map(j => j.key === key ? { ...j, status } : j)),
      );
      setResults(successRows);
      if (successRows.length) pushToast(`تم إنشاء ${successRows.length} استبيان بنجاح.`, "success");
      if (failCount) pushToast(`فشل إنشاء ${failCount} استبيان.`, "error");
    } catch (e) {
      pushToast("فشل إنشاء مجلدات الفصل الدراسي: " + e.message, "error");
    } finally {
      setGenerating(false);
    }
  }

  return (
    <div className="ssg-create-page">
      <div className="ssg-create-hero">
        <div>
          <div className="ssg-create-kicker">مسار إنشاء سريع ومنظم</div>
          <h2>جهّز استبيانات الفصل الدراسي</h2>
          <p>اختر القوالب وحدد طريقة توزيع كل نموذج، ثم راجع الفصل وابدأ الإنشاء.</p>
        </div>
        <div className="ssg-step-pills" aria-label="خطوات الإنشاء">
          <span className="ssg-step-pill"><b>1</b> القوالب</span>
          <span className="ssg-step-pill"><b>2</b> الفصل الدراسي</span>
          <span className="ssg-step-pill"><b>3</b> الإنشاء</span>
        </div>
      </div>

      <div className="ssg-create-layout">
        <section className="ssg-panel">
          <div className="ssg-panel-head">
            <div><div className="ssg-panel-title">اختيار القوالب</div><div className="ssg-panel-copy">حدد الاستبيانات المطلوبة وطريقة إنشاء النسخ.</div></div>
            <span className="ssg-home-badge">{selectedCount} محدد</span>
          </div>
        {loadError ? (
          <div style={{ padding: "18px", borderRadius: 14, background: "rgba(96,165,250,.045)", border: "1px solid rgba(147,197,253,.12)" }}>
            <ErrorBanner text={loadError} />
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
              <button type="button" className="btn btn-primary btn-sm" onClick={() => setTemplateReloadKey(k => k + 1)}>إعادة تحميل القوالب</button>
              {onReconnect && <button type="button" className="btn btn-ghost btn-sm" onClick={onReconnect}>تحديث اتصال Google</button>}
            </div>
          </div>
        ) : templates === null ? (
          <Skeleton rows={5} />
        ) : templates.length === 0 ? (
          <EmptyState text="لا توجد نماذج Google Forms داخل مجلد القوالب المحدد." />
        ) : (
          <>
            <div className="ssg-template-tools">
              <input className="ssg-template-search" value={templateSearch} onChange={e => setTemplateSearch(e.target.value)} placeholder="ابحث باسم الاستبيان..." />
              <button type="button" className="btn btn-ghost btn-sm" onClick={toggleAll}>{allSelected ? "إلغاء الكل" : "تحديد الكل"}</button>
            </div>
            <div className="ssg-template-list">
            {visibleTemplates.map(t => (
              <div key={t.id} className={`ssg-template-card ${t.selected ? "selected" : ""}`}>
                <input type="checkbox" checked={t.selected} onChange={() => toggleOne(t.id)} aria-label={`تحديد ${t.name}`} style={{ width: 17, height: 17, accentColor: "#1abc9c", cursor: "pointer" }} />
                <label className="ssg-template-name" onClick={() => toggleOne(t.id)} style={{ cursor: "pointer" }}>{t.name}</label>
                <select
                  value={t.mode} disabled={!t.selected} onChange={e => setTemplateMode(t.id, e.target.value)}
                  title="نسخة عامة واحدة، أو نسخة لكل قسم، أو نسخة لكل قسم وبرنامج (للأقسام التي لها برامج)"
                  style={{ opacity: t.selected ? 1 : .42, cursor: t.selected ? "pointer" : "default" }}
                >
                  <option value="general">نسخة عامة</option>
                  <option value="departments">نسخة لكل قسم</option>
                  <option value="programs">الأقسام والبرامج</option>
                </select>
              </div>
            ))}
            {!visibleTemplates.length && <EmptyState icon="⌕" text="لا توجد قوالب مطابقة للبحث." />}
            </div>
          </>
        )}
        </section>

        <aside className="ssg-create-aside">
          <div className="ssg-panel">
            <div className="ssg-panel-head"><div><div className="ssg-panel-title">بيانات الفصل</div><div className="ssg-panel-copy">ستُستخدم في أسماء المجلدات والنماذج.</div></div></div>
            <label className="ssg-field-label">السنة الدراسية</label>
            <input
              value={year} onChange={e => setYear(e.target.value)} placeholder="2026/2027"
              disabled={generating} list="ssg-year-options" autoComplete="off"
              className="ssg-field"
            />
            <datalist id="ssg-year-options">
              {yearOptions.map(y => <option key={y} value={y} />)}
            </datalist>
            <label className="ssg-field-label" style={{ marginTop: 13 }}>الفصل الدراسي</label>
            <select
              value={semester} onChange={e => setSemester(e.target.value)} disabled={generating}
              className="ssg-field"
            >
              {SEMESTERS.map(s => <option key={s} value={s} style={{ color: "#000" }}>{s}</option>)}
            </select>
            <div className="ssg-selection-summary"><span>القوالب المختارة</span><strong>{selectedCount}</strong></div>
            <button className="btn btn-primary ssg-create-submit" disabled={generating || selectedCount === 0} onClick={handleGenerate}>
              {generating ? <span>جاري الإنشاء<LoadingDots /></span> : `إنشاء ${selectedCount ? `${selectedCount} قالب` : "الاستبيانات"}`}
            </button>
            <ErrorBanner text={validationMsg} />
          </div>
        </aside>
      </div>

      {generating && jobs && (
        <div className={`ssg-generation-stage ssg-generation-widget ${progressCollapsed ? "collapsed" : ""}`} role="status" aria-live="polite" aria-label="تقدم تجهيز استبيانات الفصل">
          <div className="ssg-generation-head">
            <div className="ssg-generation-title"><LoadingGlyph compact /><div><div>جاري تجهيز استبيانات الفصل</div><small style={{ color: "rgba(255,255,255,.42)", fontWeight: 600 }}>{tip}</small></div></div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div className="ssg-generation-percent">{generationPct}%</div>
              <button type="button" className="ssg-widget-toggle" onClick={() => setProgressCollapsed(v => !v)} aria-label={progressCollapsed ? "توسيع نافذة التقدم" : "تصغير نافذة التقدم"} title={progressCollapsed ? "توسيع" : "تصغير"}>{progressCollapsed ? "+" : "−"}</button>
            </div>
          </div>
          <div className="ssg-progress-track"><div className="ssg-progress-fill" style={{ width: `${generationPct}%` }} /></div>
          <div className="ssg-current-job">{activeJob ? `نعمل الآن على: ${activeJob.label}` : "جاري تنظيم مجلدات Google Drive..."}</div>
          <div className="ssg-job-grid">
            {jobs.map(j => <div key={j.key} className={`ssg-job-chip ${j.status}`}><span className="ssg-status-dot"/><span>{j.label}</span></div>)}
          </div>
        </div>
      )}

      {!generating && results?.length > 0 && (
        <div className="ssg-panel ssg-results-panel">
          <div className="ssg-success-head">
            <div className="ssg-success-copy"><span className="ssg-success-icon">✓</span><div><div>تم إنشاء {results.length} استبيان بنجاح</div><small style={{ color: "rgba(255,255,255,.42)", fontWeight: 600 }}>النماذج جاهزة لاستقبال الردود على Google Drive.</small></div></div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <LinksExportButtons rows={results.map(r => ({ name: r.name, formUrl: r.formUrl, responsesUrl: editorResponsesUrl(r.formId) }))} title={`استبيانات ${semester} ${year}`} filename={`روابط_استبيانات_${semester}_${yearFolderName(year)}.xlsx`} pushToast={pushToast} />
              <button type="button" className="btn btn-primary btn-sm" onClick={onGoDashboard}>اذهب إلى لوحة التحكم ←</button>
            </div>
          </div>
          <div className="ssg-created-grid">
            {results.map((r, index) => (
              <div key={r.key} className="ssg-created-card" style={{ animationDelay: `${Math.min(index * 35, 280)}ms` }}>
                <div className="ssg-created-name">{r.name}</div>
                <div className="ssg-created-actions">
                  <a className="ssg-link-action" href={r.formUrl} target="_blank" rel="noreferrer">فتح النموذج</a>
                  <a className="ssg-link-action" href={editorResponsesUrl(r.formId)} target="_blank" rel="noreferrer">عرض الردود</a>
                  <button className="ssg-link-action" onClick={() => copyToClipboard(r.formUrl, pushToast)}>نسخ الرابط</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

const ANALYZE_TIPS = [
  "🔎 جاري جلب ردود كل استبيان...",
  "📊 جاري حساب المؤشرات الإحصائية...",
  "📄 جاري تصميم تقرير PDF لكل استبيان...",
  "☁️ جاري رفع التقارير على Google Drive...",
  "⏳ لحظات ونكون قد انتهينا...",
];

function SurveyDashboardCard({ survey, selected, index, onToggle, onAnalyze, onRefresh }) {
  const stop = fn => e => { e.stopPropagation(); fn?.(); };
  const toggleByKey = e => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onToggle(); }
  };
  return (
    <article className={`ssg-survey-card ${selected ? "selected" : ""}`} role="button" tabIndex={0} aria-pressed={selected} onClick={onToggle} onKeyDown={toggleByKey} style={{ animationDelay: `${Math.min(index * 35, 350)}ms` }}>
      <span className="ssg-card-select" aria-hidden="true">{selected ? "✓" : "+"}</span>
      <span className="ssg-survey-type">{survey.surveyType || "استبيان فصلي"}</span>
      <h3 className="ssg-survey-name">{survey.name}</h3>
      <div className="ssg-survey-metrics">
        <div className="ssg-survey-metric"><strong>{survey.responses}</strong><span>إجمالي الردود</span></div>
        <div className="ssg-survey-metric"><strong style={{ fontSize: 12 }}>{formatWhen(survey.last)}</strong><span>آخر استجابة</span></div>
      </div>
      <div className="ssg-survey-card-actions">
        <button type="button" className="ssg-card-action primary" onClick={stop(() => onAnalyze("download"))}>تحليل وتحميل</button>
        <button type="button" className="ssg-card-action" onClick={stop(() => onAnalyze("upload"))}>تحليل ورفع</button>
        <a className="ssg-card-action" href={survey.formUrl} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()}>النموذج</a>
        <a className="ssg-card-action" href={editorResponsesUrl(survey.id)} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()}>الردود</a>
        <button type="button" className="ssg-card-action" onClick={stop(onRefresh)}>تحديث</button>
      </div>
    </article>
  );
}

// ---------- Dashboard ----------
function DashboardView({ token, pushToast, onAnalyzeForms, quickMode = false }) {
  const [years, setYears] = useState(null);
  const [year, setYear] = useState("");
  const [semesters, setSemesters] = useState(null);
  const [semester, setSemester] = useState("");
  const [surveys, setSurveys] = useState(null);
  const [loadError, setLoadError] = useState("");
  const [refreshKey, setRefreshKey] = useState(0);
  const [selectedIds, setSelectedIds] = useState(() => new Set());

  useEffect(() => {
    let cancelled = false;
    setYears(null); setLoadError("");
    listSubfolders(token, ROOT_SURVEYS_FOLDER_ID)
      .then(fs => { if (cancelled) return; setYears(fs); setYear(y => y || fs[0]?.name || ""); })
      .catch(e => { if (!cancelled) setLoadError("فشل تحميل السنوات الدراسية: " + e.message); });
    return () => { cancelled = true; };
  }, [token, refreshKey]);

  const yearFolder = years?.find(f => f.name === year) ?? null;

  useEffect(() => {
    if (!yearFolder) { setSemesters(null); return; }
    let cancelled = false;
    setSemesters(null);
    listSubfolders(token, yearFolder.id)
      .then(fs => { if (cancelled) return; setSemesters(fs); setSemester(s => (fs.some(f => f.name === s) ? s : (fs[0]?.name || ""))); })
      .catch(e => { if (!cancelled) setLoadError("فشل تحميل الفصول الدراسية: " + e.message); });
    return () => { cancelled = true; };
  }, [yearFolder?.id, refreshKey]);

  const semFolder = semesters?.find(f => f.name === semester) ?? null;

  useEffect(() => { setSelectedIds(new Set()); }, [year, semester]);

  const loadSurveys = useCallback(() => {
    if (!semFolder) { setSurveys([]); return; }
    setSurveys(null); setLoadError("");
    // Each survey type now lives in its own subfolder under the semester folder
    // (see GenerateSurveysView) — listSemesterSurveysWithStats walks that structure.
    listSemesterSurveysWithStats(token, semFolder.id)
      .then(rows => setSurveys(rows))
      .catch(e => setLoadError("فشل تحميل الاستبيانات: " + e.message));
  }, [token, semFolder?.id]);

  useEffect(() => { loadSurveys(); }, [loadSurveys, refreshKey]);

  async function refreshOne(id) {
    try {
      const responses = await listAllResponses(token, id);
      const stats = computeResponseStats(responses);
      setSurveys(rows => rows.map(r => r.id === id ? { ...r, responses: stats.count, last: stats.lastSubmittedTime, error: null } : r));
      pushToast("تم تحديث بيانات الاستبيان.", "success");
    } catch (e) {
      pushToast("فشل تحديث الاستبيان: " + e.message, "error");
    }
  }

  const totalSurveys = surveys?.length ?? 0;
  const totalResponses = surveys?.reduce((s, r) => s + r.responses, 0) ?? 0;
  const avg = totalSurveys ? Math.round((totalResponses / totalSurveys) * 10) / 10 : 0;
  const lastResponse = surveys?.reduce((max, r) => (r.last && (!max || r.last > max)) ? r.last : max, null) ?? null;

  // ---- Analyze all: reuse the app's existing analysis engine and PDF builder by
  // converting each form's responses into the
  // same [header, ...rows] shape a real Excel export would have. ----
  const [analyzing, setAnalyzing] = useState(false);
  const [analyzeProgress, setAnalyzeProgress] = useState(null);
  const [reportAuthors, setReportAuthors] = useState({ preparedBy: "", reviewer: "" });
  const analysisCancelRef = useRef(false);
  const [analysisCancelRequested, setAnalysisCancelRequested] = useState(false);
  const analyzeTip = useRotatingTip(analyzing, ANALYZE_TIPS);

  async function runAnalyzeAll(targetSurveys = surveys, reportMeta) {
    analysisCancelRef.current = false;
    setAnalysisCancelRequested(false);
    setAnalyzing(true);
    setAnalyzeProgress(targetSurveys.map(s => ({ id: s.id, name: s.name, status: "pending" })));
    const schemas = getAllAnalysisSchemas();
    const settings = loadReportSettings();
    let successCount = 0, skipCount = 0;
    const reports = [];

    for (const survey of targetSurveys) {
      if (analysisCancelRef.current) break;
      setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "active" } : x));
      try {
        const [form, responses] = await Promise.all([getForm(token, survey.id), listAllResponses(token, survey.id)]);
        if (analysisCancelRef.current) break;
        const rows = responsesToRows(form, responses);
        const schemaId = detectAnySurveyType(survey.name, rows[0]);
        const schema = schemaId ? schemas[schemaId] : null;
        if (!schema) {
          setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "skipped" } : x));
          skipCount++;
          continue;
        }
        const result = analyze(rows, schema);
        const meta = { year, program: departmentFromSurveyName(survey.name), ...reportMeta };
        const builtPdf = await buildBrandedReportPdf(
          result, meta, settings, undefined,
          { shouldCancel: () => analysisCancelRef.current }
        );
        if (analysisCancelRef.current) break;
        reports.push({ blob: builtPdf.blob, filename: builtPdf.filename });
        setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "done" } : x));
        successCount++;
      } catch (e) {
        if (e?.name === "AbortError" && analysisCancelRef.current) break;
        setAnalyzeProgress(p => p.map(x => x.id === survey.id ? { ...x, status: "error" } : x));
      }
    }

    setAnalyzing(false);
    if (!analysisCancelRef.current && reports.length === 1) {
      downloadBlob(reports[0].blob, reports[0].filename);
    } else if (!analysisCancelRef.current && reports.length > 1) {
      const zip = new JSZip();
      reports.forEach(report => zip.file(report.filename, report.blob));
      downloadBlob(await zip.generateAsync({ type: "blob" }), `تقارير_تحليل_${year}.zip`);
    }
    if (successCount) pushToast(`تم إنشاء ${successCount} تقرير بنجاح.`, "success");
    if (skipCount) pushToast(`تعذّر التعرف على نوع ${skipCount} استبيان — لم يتم تحليله.`, "error");
  }

  async function analyzeSelection(items, delivery = "upload") {
    if (!items.length || analyzing) return;
    const preparedBy = reportAuthors.preparedBy.trim();
    const reviewer = reportAuthors.reviewer.trim();
    if (!preparedBy || !reviewer) {
      setLoadError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء أي تقرير.");
      return;
    }
    const reportMeta = { year, preparedBy, reviewer };
    if (!onAnalyzeForms) { await runAnalyzeAll(items, reportMeta); return; }
    setAnalyzing(true); setLoadError("");
    try {
      await onAnalyzeForms(items, token, { delivery, reportMeta });
    } catch (e) {
      setLoadError("تعذّر تحليل الاستبيانات: " + e.message);
    } finally {
      setAnalyzing(false);
    }
  }

  return (
    <div className="ssg-dashboard-page">
      {quickMode && (
        <div className="ssg-create-hero" style={{ marginBottom: 14 }}>
          <div>
            <div className="ssg-create-kicker">تحليل سريع</div>
            <h2>اختر السنة والفصل ثم شغّل التحليل</h2>
            <p>حلّل كل استبيانات الفصل دفعة واحدة، أو حدّد الاستبيانات المطلوبة فقط. سيتم إنشاء ملفات PDF ورفعها تلقائياً على Google Drive.</p>
          </div>
          <div className="ssg-step-pills">
            <span className="ssg-step-pill"><b>١</b> السنة والفصل</span>
            <span className="ssg-step-pill"><b>٢</b> الكل أو المحدد</span>
            <span className="ssg-step-pill"><b>٣</b> رفع PDF</span>
          </div>
        </div>
      )}
      <div className="ssg-dashboard-toolbar">
          <div style={{ flex: "1 1 180px" }}>
            <label className="ssg-field-label">السنة الدراسية</label>
            <select className="ssg-field" value={year} onChange={e => setYear(e.target.value)}>
              {(years ?? []).map(f => <option key={f.id} value={f.name} style={{ color: "#000" }}>{f.name}</option>)}
            </select>
          </div>
          <div style={{ flex: "1 1 180px" }}>
            <label className="ssg-field-label">الفصل الدراسي</label>
            <select className="ssg-field" value={semester} onChange={e => setSemester(e.target.value)}>
              {(semesters ?? []).map(f => <option key={f.id} value={f.name} style={{ color: "#000" }}>{f.name}</option>)}
            </select>
          </div>
          <div style={{ flex: "1 1 190px" }}>
            <label className="ssg-field-label">مُعدّ التحليل <span aria-hidden="true">*</span></label>
            <input className="ssg-field" value={reportAuthors.preparedBy} required placeholder="الاسم الكامل" onChange={e => setReportAuthors(a => ({ ...a, preparedBy: e.target.value }))} />
          </div>
          <div style={{ flex: "1 1 190px" }}>
            <label className="ssg-field-label">مراجع التحليل <span aria-hidden="true">*</span></label>
            <input className="ssg-field" value={reportAuthors.reviewer} required placeholder="الاسم الكامل" onChange={e => setReportAuthors(a => ({ ...a, reviewer: e.target.value }))} />
          </div>
          <div className="ssg-dashboard-actions">
            {!!selectedIds.size && <span className="ssg-selection-count">✓ {selectedIds.size} محدد</span>}
            {!!surveys?.length && <button className="btn btn-ghost btn-sm" onClick={() => setSelectedIds(selectedIds.size === surveys.length ? new Set() : new Set(surveys.map(s => s.id)))}>{selectedIds.size === surveys.length ? "إلغاء تحديد الكل" : "تحديد الكل"}</button>}
            {!!selectedIds.size && <button className="btn btn-ghost btn-sm" disabled={analyzing} onClick={() => analyzeSelection(surveys.filter(s => selectedIds.has(s.id)), "download")}>تحليل وتحميل المحدد ({selectedIds.size})</button>}
            {!!selectedIds.size && <button className="btn btn-primary btn-sm" disabled={analyzing} onClick={() => analyzeSelection(surveys.filter(s => selectedIds.has(s.id)), "upload")}>تحليل ورفع المحدد ({selectedIds.size})</button>}
            <button className="btn btn-ghost btn-sm" onClick={() => setRefreshKey(k => k + 1)}>تحديث البيانات</button>
            {!!surveys?.length && <button className="btn btn-ghost btn-sm" disabled={analyzing} onClick={() => analyzeSelection(surveys, "download")}>تحليل وتحميل الكل ({surveys.length})</button>}
            {!!surveys?.length && <button className="btn btn-primary btn-sm" disabled={analyzing} onClick={() => analyzeSelection(surveys, "upload")}>تحليل ورفع الكل ({surveys.length})</button>}
          </div>
        <ErrorBanner text={loadError} />

        {analyzeProgress && (
          <div style={{ marginTop: 18, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,.1)" }}>
            {analyzing && (
              <>
                <ProgressBar
                  done={analyzeProgress.filter(p => p.status === "done" || p.status === "error" || p.status === "skipped").length}
                  total={analyzeProgress.length}
                  label={analysisCancelRequested ? "جارٍ إلغاء العملية..." : "جاري تحليل الاستبيانات..."}
                  tip={analyzeTip}
                />
                <button type="button" className="btn btn-ghost btn-sm" disabled={analysisCancelRequested} onClick={() => { analysisCancelRef.current = true; setAnalysisCancelRequested(true); }} style={{ marginTop: 10, color: "#ffb4b4", borderColor: "rgba(239,68,68,.35)" }}>{analysisCancelRequested ? "جارٍ الإلغاء…" : "إلغاء العملية"}</button>
              </>
            )}
            {analyzeProgress.map(p => (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 4px", fontSize: 13, color: "rgba(255,255,255,.8)" }}>
                <span>{p.status === "done" ? "✔" : p.status === "error" ? "✖" : p.status === "skipped" ? "⚠" : p.status === "active" ? "⏳" : "○"}</span>
                <span>{p.name}</span>
                {p.status === "skipped" && <span style={{ color: "rgba(255,255,255,.4)", fontSize: 11 }}>(نوع غير معروف)</span>}
              </div>
            ))}
          </div>
        )}
      </div>

      {years === null || (year && semesters === null) ? (
        <DashboardLoading label="جاري تحميل السنوات والفصول" />
      ) : !years.length ? (
        <div className="card" style={{ padding: 24 }}>
          <EmptyState text="لا توجد سنوات دراسية بعد داخل مجلد الاستبيانات الرئيسي." />
        </div>
      ) : surveys === null ? (
        <DashboardLoading />
      ) : surveys.length === 0 ? (
        <div className="card" style={{ padding: 24 }}>
          <EmptyState text="لا توجد استبيانات لهذا الفصل بعد." />
        </div>
      ) : (
        <>
          <div className="ssg-stat-grid">
            <div className="ssg-stat-card" style={{ animationDelay: "0ms" }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>إجمالي الاستبيانات</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{totalSurveys}</div>
            </div>
            <div className="ssg-stat-card" style={{ animationDelay: "55ms" }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>إجمالي الردود</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{totalResponses}</div>
            </div>
            <div className="ssg-stat-card" style={{ animationDelay: "110ms" }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>متوسط الردود لكل استبيان</div>
              <div style={{ color: "#fff", fontSize: 26, fontWeight: 900 }}>{avg}</div>
            </div>
            <div className="ssg-stat-card" style={{ animationDelay: "165ms" }}>
              <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12 }}>آخر رد مستلم</div>
              <div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>{formatWhen(lastResponse)}</div>
            </div>
          </div>

          <section className="ssg-panel">
            <div className="ssg-panel-head">
              <div><div className="ssg-panel-title">استبيانات الفصل الدراسي</div><div className="ssg-panel-copy">اضغط على البطاقة لتحديدها، أو مرّر عليها للوصول إلى التحليل والروابط.</div></div>
              <LinksExportButtons
                rows={surveys.map(r => ({ name: r.name, formUrl: r.formUrl, responsesUrl: editorResponsesUrl(r.id) }))}
                title={`استبيانات ${semester} ${year}`}
                filename={`روابط_استبيانات_${semester}_${yearFolderName(year)}.xlsx`}
                pushToast={pushToast}
              />
            </div>
            <div className="ssg-survey-grid">
              {surveys.map((survey, index) => (
                <SurveyDashboardCard key={survey.id} survey={survey} index={index} selected={selectedIds.has(survey.id)}
                  onToggle={() => setSelectedIds(prev => { const next = new Set(prev); if (next.has(survey.id)) next.delete(survey.id); else next.add(survey.id); return next; })}
                  onAnalyze={delivery => analyzeSelection([survey], delivery)} onRefresh={() => refreshOne(survey.id)} />
              ))}
            </div>
          </section>
        </>
      )}
    </div>
  );
}

// ---------- root ----------
// ---------- Departments & Programs settings ----------
function DepartmentsView({ pushToast }) {
  const [departments, setDepartments] = useState(loadDepartments);
  const [newDeptName, setNewDeptName] = useState("");
  const [programDrafts, setProgramDrafts] = useState({}); // deptId -> current input text
  const [pendingDelete, setPendingDelete] = useState(null);
  const accents = ["#5eead4", "#60a5fa", "#c4b5fd", "#fbbf24", "#fb7185"];
  const totalPrograms = departments.reduce((sum, dept) => sum + (dept.programs?.length ?? 0), 0);

  function persist(next, message) {
    setDepartments(next);
    saveDepartments(next);
    if (message) pushToast(message, "success");
  }

  function addDepartment() {
    const name = newDeptName.trim();
    if (!name) return;
    if (departments.some(d => d.name === name)) { pushToast("هذا القسم موجود بالفعل.", "error"); return; }
    persist([...departments, { id: `dept-${Date.now()}`, name, programs: [] }], "تم إضافة القسم.");
    setNewDeptName("");
  }

  function renameDepartment(id, name) {
    setDepartments(ds => ds.map(d => d.id === id ? { ...d, name } : d));
  }

  function commitRename(id) {
    const department = departments.find(d => d.id === id);
    const name = department?.name.trim() ?? "";
    if (!name) {
      setDepartments(loadDepartments());
      pushToast("لا يمكن حفظ قسم بدون اسم.", "error");
      return;
    }
    if (departments.some(d => d.id !== id && d.name.trim() === name)) {
      setDepartments(loadDepartments());
      pushToast("يوجد قسم آخر بنفس الاسم.", "error");
      return;
    }
    const next = departments.map(d => d.id === id ? { ...d, name } : d);
    persist(next, "تم حفظ اسم القسم.");
  }

  function removeDepartment(id) {
    persist(departments.filter(d => d.id !== id), "تم حذف القسم.");
    setPendingDelete(null);
  }

  function addProgram(deptId) {
    const text = (programDrafts[deptId] || "").trim();
    if (!text) return;
    const department = departments.find(d => d.id === deptId);
    if (department?.programs?.some(program => program === text)) {
      pushToast("هذا البرنامج موجود بالفعل داخل القسم.", "error");
      return;
    }
    persist(departments.map(d => d.id === deptId ? { ...d, programs: [...d.programs, text] } : d));
    setProgramDrafts(p => ({ ...p, [deptId]: "" }));
    pushToast("تمت إضافة البرنامج.", "success");
  }

  function removeProgram(deptId, program) {
    persist(departments.map(d => d.id === deptId ? { ...d, programs: d.programs.filter(p => p !== program) } : d));
  }

  return (
    <section className="ssg-departments" aria-labelledby="departments-title">
      <div className="ssg-dept-hero">
        <div>
          <div className="ssg-dept-kicker"><SemesterIcon name="settings" size={16} /> إعداد الهيكل الأكاديمي</div>
          <h2 id="departments-title">إدارة الأقسام والبرامج</h2>
          <p>نظّم الأقسام والبرامج التي يستخدمها النظام عند إنشاء استبيانات الفصل. القسم الذي لا يحتوي على برامج سيُعامل كوحدة واحدة، بينما يُنشأ نموذج مستقل لكل برنامج عند إضافته.</p>
        </div>
        <div className="ssg-dept-stats" aria-label="ملخص الأقسام والبرامج">
          <div className="ssg-dept-stat"><strong>{departments.length}</strong><span>قسم أكاديمي</span></div>
          <div className="ssg-dept-stat"><strong>{totalPrograms}</strong><span>برنامج مسجل</span></div>
        </div>
      </div>

      <div className="ssg-dept-add">
        <div className="ssg-dept-add-copy"><strong>إضافة قسم جديد</strong><span>أضف القسم أولًا، ثم أضف برامجه من البطاقة الخاصة به.</span></div>
        <form className="ssg-dept-add-form" onSubmit={e => { e.preventDefault(); addDepartment(); }}>
          <input className="ssg-dept-input" value={newDeptName} onChange={e => setNewDeptName(e.target.value)} placeholder="مثال: قسم إدارة الأعمال" aria-label="اسم القسم الجديد" />
          <button type="submit" className="btn btn-primary btn-sm" disabled={!newDeptName.trim()}>إضافة القسم +</button>
        </form>
      </div>

      <div className="ssg-dept-section-head">
        <div><h3>الأقسام الحالية</h3><span>اضغط على اسم القسم لتعديله؛ الحفظ يتم عند الخروج من الحقل.</span></div>
        <span>{departments.length} قسم</span>
      </div>

      <div className="ssg-dept-grid">
        {departments.map((d, index) => (
        <article key={d.id} className="ssg-dept-card" style={{ "--dept-accent": accents[index % accents.length] }}>
          <div className="ssg-dept-card-head">
            <span className="ssg-dept-index" aria-hidden="true">{String(index + 1).padStart(2, "0")}</span>
            <input className="ssg-dept-name" value={d.name} aria-label={`اسم القسم ${index + 1}`}
              onChange={e => renameDepartment(d.id, e.target.value)} onBlur={() => commitRename(d.id)}
              onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }} />
            <button type="button" className="ssg-dept-delete" onClick={() => setPendingDelete(d.id)} aria-label={`حذف قسم ${d.name}`} title="حذف القسم">✕</button>
          </div>

          <div className="ssg-program-head"><span>البرامج التابعة للقسم</span><span className="ssg-program-count">{d.programs.length} برنامج</span></div>
          <div className="ssg-program-list">
            {d.programs.map(p => (
              <span key={p} className="ssg-program-chip">
                {p}
                <button type="button" className="ssg-program-remove" onClick={() => removeProgram(d.id, p)} aria-label={`إزالة البرنامج ${p}`} title={`إزالة ${p}`}>✕</button>
              </span>
            ))}
            {!d.programs.length && <div className="ssg-program-empty">لا توجد برامج—سيُنشأ استبيان واحد للقسم.</div>}
          </div>

          <form className="ssg-program-add" onSubmit={e => { e.preventDefault(); addProgram(d.id); }}>
            <input className="ssg-dept-input" value={programDrafts[d.id] || ""} onChange={e => setProgramDrafts(p => ({ ...p, [d.id]: e.target.value }))} placeholder="اسم برنامج جديد (اختياري)" aria-label={`برنامج جديد في ${d.name}`} />
            <button type="submit" className="btn btn-ghost btn-sm" disabled={!(programDrafts[d.id] || "").trim()}>إضافة برنامج +</button>
          </form>

          {pendingDelete === d.id && (
            <div className="ssg-delete-confirm" role="alert">
              <span>حذف القسم سيحذف قائمة برامجه من إعدادات الإنشاء.</span>
              <div className="ssg-delete-actions">
                <button type="button" className="btn btn-danger btn-sm" onClick={() => removeDepartment(d.id)}>تأكيد الحذف</button>
                <button type="button" className="btn btn-ghost btn-sm" onClick={() => setPendingDelete(null)}>تراجع</button>
              </div>
            </div>
          )}
        </article>
      ))}
        {!departments.length && <div className="ssg-dept-empty">لا توجد أقسام حتى الآن. أضف أول قسم من النموذج بالأعلى.</div>}
      </div>
    </section>
  );
}

function SemesterIcon({ name, size = 24 }) {
  const common = { width: size, height: size, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round", strokeLinejoin: "round" };
  if (name === "create") return <svg {...common}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z"/><path d="m14 6 4 4"/></svg>;
  if (name === "analysis") return <svg {...common}><path d="M4 19V9"/><path d="M10 19V5"/><path d="M16 19v-7"/><path d="M22 19V3"/><path d="M2 19h22"/></svg>;
  if (name === "dashboard") return <svg {...common}><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></svg>;
  if (name === "settings") return <svg {...common}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .34 1.88l.06.06-2.83 2.83-.06-.06A1.7 1.7 0 0 0 15 19.4a1.7 1.7 0 0 0-1 .6 1.7 1.7 0 0 0-.4 1.1V21h-4v-.1A1.7 1.7 0 0 0 8.6 19.4a1.7 1.7 0 0 0-1.88.34l-.06.06-2.83-2.83.06-.06A1.7 1.7 0 0 0 4.6 15a1.7 1.7 0 0 0-.6-1 1.7 1.7 0 0 0-1.1-.4H3v-4h.1A1.7 1.7 0 0 0 4.6 8.6a1.7 1.7 0 0 0-.34-1.88l-.06-.06 2.83-2.83.06.06A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-.6 1.7 1.7 0 0 0 .4-1.1V3h4v.1A1.7 1.7 0 0 0 15.4 4.6a1.7 1.7 0 0 0 1.88-.34l.06-.06 2.83 2.83-.06.06A1.7 1.7 0 0 0 19.4 9c.16.37.37.7.6 1 .3.27.68.4 1.1.4h.1v4h-.1c-.42 0-.8.13-1.1.4-.23.3-.44.63-.6 1Z"/></svg>;
  if (name === "back") return <svg {...common}><path d="m15 18-6-6 6-6"/></svg>;
  return <svg {...common}><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/></svg>;
}

const SEMESTER_ACTIONS = [
  {
    id: "generate", icon: "create", accent: "#5eead4", glow: "rgba(20,184,166,.16)",
    title: "إنشاء استبيانات فصل جديد",
    copy: "اختر القوالب والسنة والفصل، ثم أنشئ النماذج منظمةً تلقائياً على Google Drive.",
  },
  {
    id: "quick-analysis", icon: "analysis", accent: "#60a5fa", glow: "rgba(59,130,246,.16)",
    title: "تحليل سريع",
    copy: "حدّد السنة والفصل، ثم حلّل كل الاستبيانات أو اختر بعضها، وارفع تقارير PDF على Drive تلقائياً.",
  },
  {
    id: "dashboard", icon: "dashboard", accent: "#c4b5fd", glow: "rgba(139,92,246,.15)",
    title: "لوحة التحكم",
    copy: "استعرض كل الاستبيانات، وعدد الردود، والحالة والروابط من شاشة واحدة سهلة.",
  },
];

function SemesterHome({ onSelect }) {
  return (
    <section className="ssg-home" aria-labelledby="semester-home-title">
      <div className="ssg-home-hero">
        <span className="ssg-home-badge"><SemesterIcon name="dashboard" size={15} /> استبيانات الفصل الدراسي</span>
        <h1 id="semester-home-title" style={{ color: "#fff", fontSize: "clamp(23px,3vw,34px)", margin: "15px 0 8px", fontWeight: 950 }}>ماذا تريد أن تنجز؟</h1>
        <p style={{ color: "rgba(255,255,255,.54)", fontSize: 13, margin: 0 }}>الإنشاء والتحليل والمتابعة في نقطة بداية واحدة.</p>
      </div>

      <div className="ssg-action-grid">
        {SEMESTER_ACTIONS.map(action => (
          <button
            key={action.id}
            type="button"
            className="ssg-action-card"
            style={{ "--ssg-accent": action.accent, "--ssg-glow": action.glow }}
            onClick={() => onSelect(action.id)}
          >
            <span className="ssg-action-icon"><SemesterIcon name={action.icon} /></span>
            <span className="ssg-action-title">{action.title}</span>
            <span className="ssg-action-copy">{action.copy}</span>
            <span className="ssg-action-arrow" aria-hidden="true">←</span>
          </button>
        ))}
      </div>

      <div className="ssg-secondary-row">
        <button type="button" className="btn btn-ghost btn-sm" onClick={() => onSelect("departments")} style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
          <SemesterIcon name="settings" size={17} /> إدارة الأقسام والبرامج
        </button>
      </div>
    </section>
  );
}

const TAB_TITLES = {
  generate: "إنشاء استبيانات فصل دراسي جديد",
  "quick-analysis": "تحليل سريع",
  dashboard: "لوحة تحكم الاستبيانات",
  departments: "إدارة الأقسام والبرامج",
};

export default function SemesterSurveys({ onOpenAnalysis, onAnalyzeForms, onSectionChange, googleAuth, initialTab = "home" }) {
  const [tab, setTab] = useState(initialTab);
  const localAuth = useSemesterAuth(!googleAuth);
  const auth = googleAuth ?? localAuth;
  const { toasts, push, dismiss } = useToasts();
  const configured = !!TEMPLATE_FOLDER_ID && !!ROOT_SURVEYS_FOLDER_ID;

  return (
    <div>
      <style>{CSS}</style>
      <ToastStack toasts={toasts} onDismiss={dismiss} />

      {tab === "home" ? (
        <SemesterHome onSelect={next => { setTab(next); onSectionChange?.(next); }} />
      ) : (
        <>
          <div className="ssg-subnav">
            <div className="ssg-subnav-title">
              <span className="ssg-action-icon" style={{ width: 36, height: 36, borderRadius: 11, margin: 0 }}>
                <SemesterIcon name={tab === "generate" ? "create" : tab === "dashboard" ? "dashboard" : tab === "quick-analysis" ? "analysis" : "settings"} size={19} />
              </span>
              {TAB_TITLES[tab]}
            </div>
            <button type="button" className="btn btn-ghost btn-sm" onClick={() => { setTab("home"); onSectionChange?.("home"); }} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <SemesterIcon name="back" size={17} /> رجوع للاختيارات
            </button>
          </div>

          {tab === "departments" ? (
            <DepartmentsView pushToast={push} />
          ) : !configured ? (
            <div className="card" style={{ padding: 24 }}>
              <InlineNotice>لم يتم ضبط معرفات مجلدات Google Drive. يرجى إضافة <code>VITE_GOOGLE_TEMPLATE_FOLDER_ID</code> و<code>VITE_GOOGLE_ROOT_SURVEYS_FOLDER_ID</code> إلى ملف <code>.env.local</code> وإعادة تشغيل الخادم.</InlineNotice>
            </div>
          ) : !auth.token ? (
            <div className="card" style={{ padding: 32, textAlign: "center" }}>
              {auth.connecting ? <LoadingGlyph /> : <div style={{ width: 62, height: 62, display: "grid", placeItems: "center", margin: "0 auto 14px", borderRadius: 18, background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)" }}><GoogleDriveIcon size={34} /></div>}
              <div style={{ color: "#fff", fontSize: 15, fontWeight: 800, marginBottom: 7 }}>الاتصال بحساب Google</div>
              <div style={{ color: "rgba(255,255,255,.48)", fontSize: 12, marginBottom: 17 }}>مطلوب للوصول إلى قوالب واستبيانات الفصل الدراسي على Drive.</div>
              <button className="btn btn-primary" disabled={auth.connecting} onClick={auth.connect}>
                {auth.connecting ? <span>جاري الاتصال<LoadingDots /></span> : "اتصال آمن بـ Google"}
              </button>
              <ErrorBanner text={auth.authError} />
            </div>
          ) : tab === "generate" ? (
            <GenerateSurveysView token={auth.token} pushToast={push} onReconnect={auth.connect} onGoDashboard={() => { setTab("dashboard"); onSectionChange?.("dashboard"); }} />
          ) : (
            <DashboardView token={auth.token} pushToast={push} onAnalyzeForms={onAnalyzeForms} quickMode={tab === "quick-analysis"} />
          )}
        </>
      )}
    </div>
  );
}
