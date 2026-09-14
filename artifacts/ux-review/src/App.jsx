import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import JSZip from "jszip";
import { analyze, analyzeRows, prepareData, readExcel, buildComparison } from "./engine/analyze.js";
import { buildAnnualDocx, buildComparisonDocx, DEFAULT_SETTINGS } from "./engine/buildDocx.js";
import { buildBrandedReportPdf } from "./engine/buildReportPdf.js";
import AiChat, { PROVIDERS, DEFAULT_AI_SETTINGS } from "./AiChat.jsx";
import EnhancedReportView from "./EnhancedReportView.jsx";
import SurveyManagement from "./SurveyManagement.jsx";
import SemesterSurveys from "./SemesterSurveys.jsx";
import SemesterFormPicker from "./SemesterFormPicker.jsx";
import CourseEvaluationHub from "./CourseEvaluationHub.jsx";
import { GoogleDriveIcon, InlineNotice } from "./UiElements.jsx";
import { getAllAnalysisSchemas as allSchemas, detectAnySurveyType as detectSurveyType } from "./engine/customSurveyModel.js";
import { saveStoredToken, getStoredToken, clearStoredToken, getForm, listAllResponses, responsesToRows, departmentFromSurveyName, SEMESTER_SCOPE, SEMESTER_TOKEN_KEY } from "./engine/semesterSurveyModel.js";

// ── Styles ────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
*{box-sizing:border-box;margin:0;padding:0}

/* ── Design tokens ──
   Named once here so color/spacing intent stays consistent across App.jsx and every
   sibling view (SurveyManagement.jsx, SemesterSurveys.jsx, CourseSplitter.jsx,
   CourseEvaluationHub.jsx) that reuses these same class names. */
:root{
  --bg-1:#0f2035; --bg-2:#1a3a5c; --bg-3:#0d3b2e;
  --surface:rgba(255,255,255,.07);       --surface-hover:rgba(255,255,255,.12);
  --surface-soft:rgba(255,255,255,.04);  --surface-strong:rgba(255,255,255,.1);
  --border:rgba(255,255,255,.12);        --border-soft:rgba(255,255,255,.08);
  --text:#e8f0fe;    --text-strong:#fff;
  --text-muted:rgba(255,255,255,.55);    --text-faint:rgba(255,255,255,.4);
  --accent:#1abc9c;      --accent-dark:#16a085;    --accent-soft:rgba(26,188,156,.15);
  --accent-2:#2874a6;    --accent-2-dark:#1a3a5c;
  --success:#0d6e3a;     --warning:#ffc107;         --danger:#e74c3c;
  --radius-sm:8px; --radius-md:14px; --radius-lg:20px; --radius-pill:50px;
  --space-1:4px; --space-2:8px; --space-3:12px; --space-4:16px; --space-5:24px; --space-6:32px;
}
body{font-family:'Cairo',sans-serif;direction:rtl}

/* ── Buttons ── */
.btn{display:inline-flex;align-items:center;gap:8px;padding:12px 28px;border-radius:var(--radius-pill);
  font-family:'Cairo',sans-serif;font-weight:700;font-size:15px;cursor:pointer;border:none;
  transition:all .22s;white-space:nowrap}
.btn:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.btn-primary{background:linear-gradient(135deg,var(--accent),var(--accent-dark));color:#fff}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(26,188,156,.4)}
.btn-blue{background:linear-gradient(135deg,var(--accent-2),var(--accent-2-dark));color:#fff}
.btn-blue:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(40,116,166,.4)}
.btn-ghost{background:var(--surface-strong);color:var(--text);border:1px solid var(--border)}
.btn-ghost:hover{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.3)}
.btn-danger{background:rgba(231,76,60,.15);color:#e87c70;border:1px solid rgba(231,76,60,.25)}
.btn-danger:hover{background:rgba(231,76,60,.28);color:#ff6b5b}
.btn-sm{padding:7px 16px;font-size:13px}

/* ── Cards ── */
.card{background:var(--surface);backdrop-filter:blur(16px);
  border:1px solid var(--border);border-radius:var(--radius-lg)}

/* ── Upload zone ── */
.upload-zone{border:2px dashed rgba(255,255,255,.3);border-radius:16px;padding:48px 32px;
  text-align:center;cursor:pointer;transition:all .25s;background:rgba(255,255,255,.04)}
.upload-zone:hover,.upload-zone.drag{border-color:#1abc9c;background:rgba(26,188,156,.08)}

/* ── Mini tables ── */
.mini-table{width:100%;border-collapse:collapse;font-size:13px;color:#e8f0fe}
.mini-table th{background:rgba(255,255,255,.1);padding:9px 10px;text-align:center;font-weight:700;color:#d0e8ff}
.mini-table td{padding:8px 10px;border-bottom:1px solid rgba(255,255,255,.06);text-align:center}
.mini-table tr:nth-child(even) td{background:rgba(255,255,255,.03)}
.mini-table tr:hover td{background:rgba(26,188,156,.06)}
.mini-table .rtl-td{text-align:right}

/* ── Form inputs ── */
.input{
  width:100%;padding:11px 16px;border-radius:10px;
  border:1px solid rgba(255,255,255,.18);
  background:#0d1f33;color:#e8f0fe;
  font-size:14px;font-family:'Cairo',sans-serif;
  outline:none;direction:rtl;
  transition:border-color .2s,box-shadow .2s,background .2s}
.input:focus{border-color:#1abc9c;background:#0f2438;box-shadow:0 0 0 3px rgba(26,188,156,.15)}
.input:hover:not(:focus){border-color:rgba(255,255,255,.3);background:#0f2236}
.input::placeholder{color:rgba(255,255,255,.28)}
.input[type="number"]{-moz-appearance:textfield}
.input[type="number"]::-webkit-inner-spin-button,.input[type="number"]::-webkit-outer-spin-button{opacity:.5}

/* ── Select dropdowns — the key fix ── */
select.input{
  appearance:none;-webkit-appearance:none;cursor:pointer;
  padding-left:38px;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='11' height='7' viewBox='0 0 11 7'%3E%3Cpath d='M1 1l4.5 4.5L10 1' stroke='%231abc9c' stroke-width='1.8' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;
  background-position:left 14px center;
}
select.input option{
  background:#0d1f33;
  color:#e8f0fe;
  padding:10px 14px;
  font-family:'Cairo',sans-serif;
  font-size:14px;
}
select.input option:checked{background:#1a5276;color:#fff}

/* ── Textarea ── */
.textarea{
  width:100%;padding:11px 16px;border-radius:10px;
  border:1px solid rgba(255,255,255,.18);
  background:#0d1f33;color:#e8f0fe;
  font-size:13px;font-family:'Cairo',sans-serif;
  outline:none;direction:rtl;resize:vertical;min-height:80px;line-height:1.7;
  transition:border-color .2s,box-shadow .2s}
.textarea:focus{border-color:#1abc9c;background:#0f2438;box-shadow:0 0 0 3px rgba(26,188,156,.15)}
.textarea::placeholder{color:rgba(255,255,255,.28)}

/* ── Labels ── */
.label{color:rgba(200,220,255,.75);font-size:12.5px;display:block;margin-bottom:7px;font-weight:700;letter-spacing:.3px}

/* ── Stat cards ── */
.stat-card{
  background:rgba(255,255,255,.08);border-radius:16px;padding:18px 22px;text-align:center;
  border:1px solid var(--border-soft);transition:background .2s}
.stat-card:hover{background:var(--surface-hover)}

/* ── Stat chip (compact inline stat, used in File Health Summary etc.) ── */
.stat-chip{background:var(--surface-soft);border:1px solid var(--border-soft);border-radius:var(--radius-sm);
  padding:10px 14px;display:flex;flex-direction:column;gap:2px;min-width:110px}
.stat-chip-value{color:var(--accent);font-weight:900;font-size:16px}
.stat-chip-label{color:var(--text-faint);font-size:11px}
.stat-chip.warn .stat-chip-value{color:var(--warning)}

/* ── Spinner ── */
.spin{animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── Tutorial overlay ── */
@keyframes tutSlideIn{from{opacity:0;transform:translateY(32px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes tutSlideNext{from{opacity:0;transform:translateX(-28px)}to{opacity:1;transform:translateX(0)}}
@keyframes tutSlidePrev{from{opacity:0;transform:translateX(28px)}to{opacity:1;transform:translateX(0)}}
.tut-card{animation:tutSlideIn .38s cubic-bezier(.22,1,.36,1)}
.tut-slide-next{animation:tutSlideNext .28s ease}
.tut-slide-prev{animation:tutSlidePrev .28s ease}
.tut-dot{width:8px;height:8px;border-radius:50%;transition:all .22s;cursor:pointer}
.tut-nav{width:40px;height:40px;border-radius:50%;border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.07);color:rgba(255,255,255,.7);font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .18s}
.tut-nav:hover{background:rgba(255,255,255,.15);border-color:rgba(255,255,255,.3);color:#fff}
.tut-nav:disabled{opacity:.25;cursor:not-allowed}

/* ── Processing overlay ── */
@keyframes fadeInUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:translateY(0)}}
@keyframes stepPop{0%{transform:scale(1)}45%{transform:scale(1.18)}100%{transform:scale(1)}}
@keyframes activePulse{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(26,188,156,.4)}60%{opacity:.7;box-shadow:0 0 0 6px rgba(26,188,156,0)}}
@keyframes progressFlow{0%{background-position:0% 50%}100%{background-position:200% 50%}}
@keyframes loaderOrbit{to{transform:rotate(360deg)}}
@keyframes loaderFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-7px)}}
@keyframes loaderDot{0%,60%,100%{opacity:.25;transform:translateY(0)}30%{opacity:1;transform:translateY(-4px)}}
@keyframes rowShimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}
.proc-overlay{animation:fadeInUp .35s cubic-bezier(.22,1,.36,1)}
.proc-step-active .proc-icon{animation:activePulse 1.1s ease infinite}
.proc-step-done   .proc-icon{animation:stepPop .3s ease}
.loader-visual{position:relative;width:78px;height:78px;margin:0 auto 16px;display:grid;place-items:center;animation:loaderFloat 2.4s ease-in-out infinite}
.loader-ring{position:absolute;inset:0;border:2px solid rgba(26,188,156,.15);border-top-color:#5eead4;border-right-color:#2874a6;border-radius:50%;animation:loaderOrbit 1.25s linear infinite}
.loader-ring.inner{inset:10px;animation-duration:.9s;animation-direction:reverse;border-top-color:#a78bfa;border-right-color:#5eead4}
.loader-dots{display:inline-flex;gap:4px;margin-inline-start:7px}.loader-dots i{width:4px;height:4px;border-radius:50%;background:#71e8d4;animation:loaderDot 1.2s infinite}.loader-dots i:nth-child(2){animation-delay:.15s}.loader-dots i:nth-child(3){animation-delay:.3s}
.batch-row.active{background:linear-gradient(90deg,rgba(26,188,156,.12),rgba(255,255,255,.035),rgba(26,188,156,.12));background-size:220% 100%;animation:rowShimmer 2s linear infinite}

/* ── Badges ── */
.badge{display:inline-block;padding:4px 14px;border-radius:20px;font-size:12px;font-weight:700}

/* ── Analysis journey ── */
.step-bar{display:flex;gap:8px;margin-bottom:32px;padding:8px;border-radius:18px;
  border:1px solid var(--border-soft);background:rgba(5,18,35,.36)}
.step-item{flex:1;min-width:0;padding:10px 6px 11px;text-align:center;font-size:11px;font-weight:700;
  background:transparent;color:var(--text-muted);transition:all .22s;border:0;position:relative;
  display:flex;flex-direction:column;align-items:center;gap:6px;border-radius:12px}
.step-item-icon{width:31px;height:31px;border-radius:10px;display:grid;place-items:center;
  color:rgba(255,255,255,.42);background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.08);transition:all .22s}
.step-item-icon svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:1.9;stroke-linecap:round;stroke-linejoin:round}
.step-item-num{font-size:9px;font-weight:900;line-height:1;opacity:.72}
.step-item.active{background:rgba(26,188,156,.14);color:#d8fff7;box-shadow:inset 0 0 0 1px rgba(26,188,156,.24)}
.step-item.active .step-item-icon{color:#fff;background:linear-gradient(135deg,#1abc9c,#16977f);border-color:transparent;box-shadow:0 5px 13px rgba(26,188,156,.24)}
.step-item.active::after{content:'';position:absolute;bottom:3px;width:24px;height:3px;border-radius:3px;background:#1abc9c}
.step-item.done{color:rgba(156,238,221,.8)}
.step-item.done .step-item-icon{color:#8be5d2;background:rgba(26,188,156,.1);border-color:rgba(26,188,156,.2)}
.workflow-header{display:flex;align-items:center;gap:13px;margin-bottom:20px;padding:14px 16px;border-radius:15px;background:linear-gradient(120deg,rgba(26,188,156,.1),rgba(40,116,166,.07));border:1px solid rgba(26,188,156,.16)}
.workflow-header-icon{width:40px;height:40px;padding:10px;display:grid;place-items:center;flex:0 0 auto;border-radius:12px;color:#d7fff7;background:linear-gradient(135deg,#1abc9c,#167d94);box-shadow:0 7px 16px rgba(26,188,156,.16)}
.workflow-header-kicker{color:#71e8d4;font-size:10px;font-weight:900}.workflow-header-title{color:#fff;font-size:18px;font-weight:900}.workflow-header-copy{color:rgba(255,255,255,.5);font-size:11.5px;margin-top:2px}

/* ── Type/mode cards ── */
.type-card{flex:1;padding:24px 18px;border-radius:16px;border:2px solid rgba(255,255,255,.1);
  background:rgba(255,255,255,.04);cursor:pointer;text-align:center;transition:all .22s}
.type-card:hover{border-color:rgba(26,188,156,.45);background:rgba(26,188,156,.06);
  transform:translateY(-2px)}
.type-card.selected{border-color:#1abc9c;background:rgba(26,188,156,.11);
  box-shadow:0 4px 20px rgba(26,188,156,.15)}

/* ── Settings ── */
.settings-section{
  background:rgba(255,255,255,.04);border-radius:14px;padding:22px 24px;margin-bottom:18px;
  border:1px solid rgba(255,255,255,.07)}
.settings-section-title{
  color:#1abc9c;font-size:14px;font-weight:900;margin-bottom:18px;
  padding-bottom:10px;border-bottom:1px solid rgba(26,188,156,.18);letter-spacing:.3px}
.toggle-row{display:flex;align-items:center;justify-content:space-between;
  padding:10px 0;border-bottom:1px solid rgba(255,255,255,.05)}
.toggle-row:last-child{border-bottom:none}
.toggle-row:hover{background:rgba(255,255,255,.02);border-radius:8px;padding-inline:6px;margin-inline:-6px}
.toggle{width:46px;height:25px;border-radius:13px;border:none;cursor:pointer;transition:background .2s;
  position:relative;flex-shrink:0}
.toggle.on{background:#1abc9c}
.toggle.off{background:rgba(255,255,255,.18)}
.toggle::after{content:'';position:absolute;width:19px;height:19px;border-radius:50%;
  background:#fff;top:3px;transition:left .18s;box-shadow:0 1px 3px rgba(0,0,0,.3)}
.toggle.on::after{left:24px}
.toggle.off::after{left:3px}

/* ── Logo preview ── */
.logo-preview{width:60px;height:60px;object-fit:contain;border-radius:8px;
  border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.05)}

/* ── Scrollbar ── */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:rgba(255,255,255,.04)}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,.2);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,.35)}

/* ── Mobile responsiveness ── */
html,body,#root{overflow-x:clip;width:100%;max-width:100%;min-height:100%;background:#091a2d}
body{min-height:100vh;background:linear-gradient(145deg,#091a2d 0%,#102f4b 58%,#0b352f 100%) fixed;overscroll-behavior-y:none}
.quality-app,.app-shell,.app-content,.app-main{max-width:100%;min-width:0}
.quality-app{min-height:100vh;min-height:100dvh;background-color:#091a2d}
.table-scroll{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}
@media (max-width:640px){
  .app-header{padding:14px 16px !important;gap:10px !important}
  .app-header-actions{gap:8px !important}
  .app-header-actions .btn{padding:8px 14px !important;font-size:12.5px !important}
  .app-main{padding:16px !important}
  .card{padding:18px !important}
  .proc-overlay{padding:24px 20px !important}
  .settings-section{padding:16px 14px !important}
  .type-card-row{flex-direction:column !important}
  .type-card{padding:16px 18px !important}
  .step-bar{flex-wrap:wrap}
  .step-item{font-size:10px;padding:9px 4px;flex-basis:31%}
  .mini-table{font-size:11.5px}
  .mini-table th,.mini-table td{padding:6px 6px}
}

/* ── Respect reduced-motion preference ── */
@media (prefers-reduced-motion:reduce){
  *{animation-duration:.001ms !important;animation-iteration-count:1 !important;
    transition-duration:.001ms !important;scroll-behavior:auto !important}
}

/* ── Focus visibility (keyboard nav) ── */
a:focus-visible,button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,[tabindex]:focus-visible{
  outline:2px solid var(--accent);outline-offset:2px;border-radius:4px}

/* ── Header nav group (view-toggle buttons, separated visually from the reset action) ── */
.header-nav-group{display:flex;gap:8px;align-items:center;flex-wrap:wrap;
  background:var(--surface-soft);border:1px solid var(--border-soft);border-radius:var(--radius-pill);padding:5px}
.header-nav-group .btn{padding:7px 16px;font-size:12.5px}

/* ── Collapsible tool sidebar ── */
.app-shell{display:block;direction:rtl;min-height:calc(100vh - 75px)}
.tool-sidebar{position:fixed;right:0;top:0;bottom:0;width:254px;max-width:calc(100vw - 20px);height:auto;box-sizing:border-box;padding:0 13px 18px;
  display:flex;flex-direction:column;border-left:1px solid rgba(255,255,255,.08);background:rgba(5,20,36,.48);
  backdrop-filter:blur(18px);transition:width .24s ease,transform .24s ease;z-index:35;overflow:hidden}
.tool-sidebar.collapsed{width:76px;padding-inline:10px}
.sidebar-brand{height:75px;display:flex;align-items:center;gap:9px;flex:0 0 75px;border-bottom:1px solid rgba(255,255,255,.07);margin-bottom:17px;white-space:nowrap}
.sidebar-brand-mark{width:34px;height:34px;display:grid;place-items:center;flex:0 0 34px;border-radius:11px;color:#fff;background:linear-gradient(145deg,#20c5a4,#197ca4)}
.sidebar-brand-copy{min-width:0;flex:1;color:#fff;font-size:11.5px;font-weight:900;line-height:1.45;overflow:hidden}
.sidebar-brand-copy small{display:block;color:rgba(255,255,255,.34);font-size:7px;letter-spacing:.3px}
.tool-sidebar.collapsed .sidebar-brand{justify-content:center}
.tool-sidebar.collapsed .sidebar-brand-mark,.tool-sidebar.collapsed .sidebar-brand-copy{display:none}
.sidebar-heading{padding:0 10px 12px;color:rgba(255,255,255,.58);font-size:10.5px;font-weight:900;white-space:nowrap;letter-spacing:.2px}
.sidebar-nav{display:flex;flex-direction:column;gap:6px}
.sidebar-item{width:100%;min-height:46px;display:flex;align-items:center;gap:11px;padding:8px 11px;border-radius:13px;border:1px solid transparent;
  color:rgba(232,240,254,.75);background:transparent;font-family:inherit;font-size:13px;font-weight:750;text-align:right;cursor:pointer;white-space:nowrap;transition:.18s ease}
.sidebar-item:hover{color:#fff;background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.07)}
.sidebar-item.active{color:#eafffb;background:linear-gradient(125deg,rgba(26,188,156,.2),rgba(40,116,166,.12));border-color:rgba(94,234,212,.22)}
.sidebar-icon{width:32px;height:32px;display:grid;place-items:center;flex:0 0 32px;color:currentColor}
.sidebar-icon svg{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.sidebar-label{overflow:hidden;opacity:1;transition:opacity .16s ease}
.tool-sidebar.collapsed .sidebar-label,.tool-sidebar.collapsed .sidebar-heading{display:none}
.tool-sidebar.collapsed .sidebar-item{justify-content:center;padding-inline:0;gap:0}
.sidebar-separator{height:1px;background:rgba(255,255,255,.07);margin:13px 8px}
.sidebar-bottom{margin-top:auto}
.sidebar-toggle{width:40px;height:40px;display:grid;place-items:center;border-radius:12px;border:1px solid rgba(255,255,255,.12);
  color:#dffbf5;background:rgba(255,255,255,.055);cursor:pointer;transition:.18s}
.sidebar-toggle:hover{background:rgba(26,188,156,.14);border-color:rgba(94,234,212,.24)}
.app-content{min-width:0;margin-right:254px;transition:margin-right .24s ease}
.app-shell.sidebar-closed .app-content{margin-right:76px}
.quality-header.sidebar-open{margin-right:254px}.quality-header.sidebar-closed{margin-right:76px}
.quality-header{min-height:75px;min-width:0;transition:margin-right .24s ease}
.google-global-btn{display:inline-flex;align-items:center;gap:9px;min-height:40px;padding:8px 13px;border-radius:12px;border:1px solid rgba(255,255,255,.12);font-family:inherit;font-size:11.5px;font-weight:800;cursor:pointer;transition:.18s}
.google-global-btn.disconnected{color:#e8f0fe;background:rgba(255,255,255,.06)}
.google-global-btn.connected{color:#9ff5df;background:rgba(26,188,156,.1);border-color:rgba(94,234,212,.22);cursor:pointer}
.google-status-dot{width:7px;height:7px;border-radius:50%;background:#34d399;box-shadow:0 0 0 4px rgba(52,211,153,.1)}
.mobile-nav-toggle{display:none}
.upload-workspace{padding:24px 28px!important;text-align:right!important;overflow:visible}
.upload-source-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:18px 0}
.upload-source-card{display:flex;align-items:center;gap:12px;padding:14px 16px;border-radius:14px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.035);color:rgba(255,255,255,.52);font-family:inherit;text-align:right;cursor:pointer;transition:.2s}
.upload-source-card:hover{background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.16)}.upload-source-card.active{color:#eafffb;border-color:rgba(94,234,212,.3);background:linear-gradient(125deg,rgba(26,188,156,.15),rgba(40,116,166,.08));box-shadow:inset 0 -2px #1abc9c}
.upload-source-icon{width:42px;height:42px;display:grid;place-items:center;flex:0 0 42px;border-radius:12px;color:#7cebd5;background:rgba(26,188,156,.1)}.upload-source-icon>svg{width:22px;height:22px}.upload-source-card strong{display:block;font-size:13.5px}.upload-source-card small{display:block;color:rgba(255,255,255,.66);font-size:11.5px;margin-top:2px}
.drive-control-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:11px 13px;margin-bottom:12px;border-radius:13px;border:1px solid rgba(94,234,212,.18);background:rgba(26,188,156,.07)}
.drive-view-switch{display:flex;gap:4px;padding:4px;border-radius:10px;background:rgba(4,18,34,.3);border:1px solid rgba(255,255,255,.075)}.drive-view-btn{padding:6px 10px;border:0;border-radius:7px;background:transparent;color:rgba(255,255,255,.42);font-family:inherit;font-size:9.5px;font-weight:800;cursor:pointer;white-space:nowrap}.drive-view-btn.active{color:#bafff1;background:rgba(26,188,156,.17)}
.drive-filter-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr)) auto;gap:8px;margin-bottom:9px}.drive-search-field{width:100%;background:rgba(4,18,34,.48);color:#e8f0fe;border:1px solid rgba(255,255,255,.12);border-radius:11px;padding:11px 13px;font-family:inherit;font-size:11.5px;direction:rtl;margin-bottom:10px;outline:none}.drive-search-field:focus{border-color:rgba(94,234,212,.42);box-shadow:0 0 0 3px rgba(26,188,156,.07)}
.drive-batch-toolbar{position:sticky;top:88px;z-index:8;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;padding:11px 13px;margin-bottom:11px;border-radius:13px;border:1px solid rgba(94,234,212,.18);background:rgba(9,31,49,.94);backdrop-filter:blur(16px);box-shadow:0 9px 25px rgba(2,12,27,.14)}
.drive-file-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:9px;max-height:440px;overflow-y:auto;padding:2px 3px 8px}
.drive-file-card{position:relative;min-height:150px;display:flex;flex-direction:column;padding:14px;border-radius:15px;border:1px solid rgba(255,255,255,.085);background:linear-gradient(145deg,rgba(255,255,255,.05),rgba(255,255,255,.025));cursor:pointer;transition:.2s;text-align:right;animation:fadeInUp .3s ease both;outline:none}.drive-file-card:hover,.drive-file-card:focus-visible{transform:translateY(-2px);border-color:rgba(96,165,250,.28);box-shadow:0 12px 28px rgba(2,12,27,.18)}.drive-file-card.selected{border-color:rgba(94,234,212,.48);background:linear-gradient(145deg,rgba(26,188,156,.13),rgba(40,116,166,.05));box-shadow:inset 0 0 0 1px rgba(94,234,212,.1)}
.drive-file-select{position:absolute;left:12px;top:12px;width:25px;height:25px;display:grid;place-items:center;border-radius:8px;color:rgba(255,255,255,.34);background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.08)}.drive-file-card.selected .drive-file-select{color:#052e2b;background:#5eead4;border-color:#5eead4}.drive-file-kind{width:36px;height:36px;display:grid;place-items:center;border-radius:11px;color:#8bd8ff;background:rgba(59,130,246,.1)}
.drive-file-name{color:#edf5ff;font-size:13px;font-weight:800;line-height:1.7;margin:10px 0;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}.drive-file-meta{display:flex;gap:5px;flex-wrap:wrap;margin-top:auto}.drive-file-tag{font-size:11px;color:#72e5cf;background:rgba(26,188,156,.09);border-radius:6px;padding:3px 7px}.drive-file-foot{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:10px;padding-top:9px;border-top:1px solid rgba(255,255,255,.055)}.drive-file-date{color:rgba(255,255,255,.58);font-size:11px}.drive-file-analyze{opacity:0;transform:translateY(2px);padding:7px 10px;border-radius:8px;border:1px solid rgba(94,234,212,.2);background:rgba(26,188,156,.1);color:#8cf2dc;font-family:inherit;font-size:11px;font-weight:800;cursor:pointer;transition:.18s}.drive-file-card:hover .drive-file-analyze,.drive-file-card:focus-within .drive-file-analyze,.drive-file-card.selected .drive-file-analyze{opacity:1;transform:none}
.upload-drop-modern{max-width:720px!important;min-height:280px;margin:0 auto!important;display:grid;place-items:center;text-align:center;background:radial-gradient(circle at 50% 15%,rgba(26,188,156,.1),transparent 45%),rgba(255,255,255,.025)!important}.upload-drop-icon{width:66px;height:66px;display:grid;place-items:center;margin:0 auto 15px;border-radius:20px;color:#bafff1;background:linear-gradient(145deg,rgba(26,188,156,.2),rgba(40,116,166,.16));border:1px solid rgba(94,234,212,.2)}
@media(max-width:760px){
  .app-shell,.app-content,.app-shell.sidebar-closed .app-content{margin-right:0;width:100%}
  .quality-header.sidebar-open,.quality-header.sidebar-closed{margin-right:0}
  .tool-sidebar{top:10px;right:10px;bottom:10px;height:auto;border:1px solid rgba(255,255,255,.11);border-radius:18px;box-shadow:0 28px 70px rgba(0,0,0,.48);background:rgba(7,25,42,.96)}
  .tool-sidebar.collapsed{display:none;transform:none;width:254px;padding-inline:13px}
  .tool-sidebar.open{display:flex;transform:none}
  .tool-sidebar.collapsed .sidebar-label,.tool-sidebar.collapsed .sidebar-heading{display:block}
  .tool-sidebar.collapsed .sidebar-item{justify-content:flex-start;padding-inline:11px;gap:11px}
  .tool-sidebar.collapsed .sidebar-brand-mark{display:grid}.tool-sidebar.collapsed .sidebar-brand-copy{display:block}
  .mobile-nav-toggle{display:grid;flex:0 0 40px}
  .quality-header{padding:11px 12px!important;gap:8px!important}
  .quality-header>div:first-child{min-width:0;flex:1;overflow:hidden}
  .quality-header>div:first-child>span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .google-global-btn{flex:0 0 auto;min-height:40px;padding:7px 10px;font-size:11px}
  .google-global-btn .google-global-label{max-width:92px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .upload-workspace{padding:16px!important}.upload-source-grid{grid-template-columns:1fr}.drive-control-bar{align-items:flex-start;flex-direction:column}.drive-view-switch{width:100%;overflow-x:auto}.drive-view-btn{flex:1}.drive-filter-grid{grid-template-columns:1fr 1fr}.drive-filter-grid>*:nth-child(3){grid-column:1/-1}.drive-batch-toolbar{top:8px;align-items:stretch;flex-direction:column}.drive-file-grid{grid-template-columns:1fr;max-height:none}.drive-file-analyze{opacity:1;transform:none}.upload-drop-modern{min-height:230px}.app-main{padding:18px 12px!important}
  .step-bar{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px}
  .step-item{width:100%;font-size:11px;min-width:0}
  .workflow-header{align-items:flex-start;padding:13px}
  .workflow-header-title{font-size:16px}.workflow-header-copy{font-size:12px;line-height:1.65}
  .btn{min-height:44px;justify-content:center}
}
@media(max-width:460px){.drive-filter-grid{grid-template-columns:1fr}.drive-filter-grid>*:nth-child(3){grid-column:auto}.drive-batch-toolbar .btn{width:100%;justify-content:center}.upload-source-card small{display:block;font-size:12px}.google-global-btn .google-global-label{max-width:74px}}

/* ── Quality hub shell ── */
.quality-app{position:relative;isolation:isolate}
.quality-app::before{content:'';position:fixed;inset:0;z-index:-1;pointer-events:none;
  background:radial-gradient(circle at 84% 8%,rgba(26,188,156,.13),transparent 30%),
    radial-gradient(circle at 8% 88%,rgba(59,130,246,.11),transparent 27%)}
.quality-header{position:sticky;top:0;z-index:40;background:rgba(10,27,45,.86);
  backdrop-filter:blur(20px);box-shadow:0 10px 35px rgba(2,12,27,.16)}
.brand-mark{width:46px;height:46px;border-radius:15px;display:grid;place-items:center;
  color:#fff;background:linear-gradient(145deg,#20c5a4,#197ca4);box-shadow:0 10px 24px rgba(26,188,156,.22)}
.hub{animation:fadeInUp .4s cubic-bezier(.22,1,.36,1)}
.hub-hero{position:relative;overflow:hidden;display:grid;grid-template-columns:minmax(0,1fr);
  gap:32px;padding:42px;border:1px solid rgba(255,255,255,.1);border-radius:28px;
  background:linear-gradient(125deg,rgba(15,49,74,.98),rgba(13,66,61,.94));
  box-shadow:0 28px 70px rgba(2,12,27,.22)}
.hub-hero::after{content:'';position:absolute;width:360px;height:360px;border-radius:50%;left:-110px;top:-180px;
  border:70px solid rgba(255,255,255,.025);box-shadow:0 0 0 40px rgba(26,188,156,.025)}
.hub-eyebrow{display:inline-flex;align-items:center;gap:8px;width:max-content;padding:6px 12px;border-radius:999px;
  color:#6ee7cf;background:rgba(26,188,156,.1);border:1px solid rgba(110,231,207,.22);font-size:11px;font-weight:800}
.hub-title{color:#fff;font-size:clamp(28px,4vw,46px);line-height:1.25;font-weight:900;margin:16px 0 10px;letter-spacing:-.8px}
.hub-subtitle{color:rgba(232,240,254,.78);font-size:15px;line-height:1.9;max-width:680px}
.hub-hero-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:25px}
.hub-snapshot{position:relative;z-index:1;align-self:stretch;display:flex;flex-direction:column;justify-content:space-between;
  padding:22px;border:1px solid rgba(255,255,255,.1);border-radius:20px;background:rgba(4,19,33,.27)}
.hub-snapshot-top{display:flex;align-items:center;justify-content:space-between;color:#fff;font-size:13px;font-weight:800}
.hub-live{display:inline-flex;align-items:center;gap:6px;color:#76e6d0;font-size:10px}
.hub-live::before{content:'';width:7px;height:7px;border-radius:50%;background:#34d399;box-shadow:0 0 0 5px rgba(52,211,153,.1)}
.hub-stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:22px}
.hub-stat{padding:15px;border-radius:14px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.07)}
.hub-stat strong{display:block;color:#fff;font-size:21px;line-height:1.2}
.hub-stat span{color:rgba(255,255,255,.45);font-size:10px}
.hub-section-head{display:flex;align-items:end;justify-content:space-between;gap:20px;margin:34px 2px 17px}
.hub-section-head h2{color:#fff;font-size:20px;font-weight:900;margin:0}
.hub-section-head p{color:rgba(255,255,255,.43);font-size:12px;margin-top:3px}
.tool-grid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:15px}
.tool-card{grid-column:span 4;position:relative;overflow:hidden;min-height:230px;padding:24px;text-align:right;
  color:inherit;font-family:inherit;border-radius:20px;border:1px solid rgba(255,255,255,.1);
  background:linear-gradient(145deg,rgba(255,255,255,.085),rgba(255,255,255,.035));cursor:pointer;
  transition:transform .24s,border-color .24s,background .24s,box-shadow .24s;display:flex;flex-direction:column;align-items:stretch}
.tool-card:hover{transform:translateY(-5px);border-color:rgba(110,231,207,.38);
  background:linear-gradient(145deg,rgba(255,255,255,.12),rgba(255,255,255,.055));box-shadow:0 20px 42px rgba(2,12,27,.2)}
.tool-card.featured{grid-column:span 6;background:linear-gradient(145deg,rgba(26,188,156,.17),rgba(255,255,255,.045))}
.tool-card-icon{width:48px;height:48px;border-radius:14px;display:grid;place-items:center;margin-bottom:20px;
  color:var(--tool-color,#65d8c0);background:color-mix(in srgb,var(--tool-color,#65d8c0) 14%,transparent);
  border:1px solid color-mix(in srgb,var(--tool-color,#65d8c0) 27%,transparent)}
.tool-card-icon svg{width:24px;height:24px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.tool-card h3{display:block;color:#fff;font-size:17px;font-weight:900;line-height:1.5;margin:0 0 7px}
.tool-card p{display:block;color:rgba(255,255,255,.7);font-size:13.5px;line-height:1.85;margin:0 0 18px;max-width:440px}
.tool-card-foot{position:relative;margin-top:auto;padding-top:12px;display:flex;align-items:center;justify-content:space-between;border-top:1px solid rgba(255,255,255,.055)}
.tool-tag{color:rgba(255,255,255,.66);font-size:11.5px;font-weight:700}
.tool-arrow{width:30px;height:30px;display:grid;place-items:center;border-radius:50%;color:#fff;
  background:rgba(255,255,255,.08);transition:transform .2s,background .2s}
.tool-card:hover .tool-arrow{transform:translateX(-4px);background:rgba(26,188,156,.22)}
.hub-note{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-top:16px;padding:18px 22px;
  border:1px solid rgba(255,255,255,.08);border-radius:17px;background:rgba(255,255,255,.035)}
.hub-note-copy{display:flex;align-items:center;gap:13px;color:rgba(255,255,255,.75);font-size:13px}
.hub-note-icon{width:36px;height:36px;border-radius:11px;display:grid;place-items:center;flex:0 0 auto;
  background:rgba(96,165,250,.12);color:#93c5fd}
@media(max-width:900px){.hub-hero{grid-template-columns:1fr;padding:30px}.tool-card,.tool-card.featured{grid-column:span 6}}
@media(max-width:640px){.quality-header .brand-copy small{display:none}.quality-header{position:sticky}.hub-hero{padding:22px;border-radius:22px}
  .hub-title{font-size:30px;letter-spacing:0}.hub-subtitle{font-size:14px}.hub-snapshot{display:none}.tool-card,.tool-card.featured{grid-column:1/-1;min-height:190px;padding:20px}.hub-section-head{align-items:start;flex-direction:column;gap:3px}.hub-note{align-items:flex-start;flex-direction:column}}
`;

const STEPS = [
  { label: "رفع الملف", icon: "upload" },
  { label: "التحقق", icon: "shield" },
  { label: "معاينة البيانات", icon: "table" },
  { label: "بيانات التقرير", icon: "edit" },
  { label: "النتائج", icon: "chart" },
];
const SETTINGS_KEY = "eruQA_settings_v1";
const AI_KEY       = "eruQA_ai_v1";

const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID ?? "";
const DRIVE_SCOPE      = SEMESTER_SCOPE;
const DRIVE_TOKEN_KEY  = SEMESTER_TOKEN_KEY;
const GSHEETS_MIME     = "application/vnd.google-apps.spreadsheet";
const DRIVE_FILE_MIMES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  GSHEETS_MIME,
  "text/csv",
];

function loadGisScript() {
  return new Promise(resolve => {
    if (window.google?.accounts?.oauth2) { resolve(); return; }
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.onload = resolve;
    document.head.appendChild(s);
  });
}

const PROGRAMS = [
  "محاسبة",
  "اقتصاد",
  "إدارة الأعمال",
  "تحليل الأعمال",
  "نظم المعلومات الإدارية",
  "الذكاء التسويقي",
  "التكنولوجيا المالية الرقمية",
  "علوم سياسية",
  "تكنولوجيا أعمال",
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function dirColor(tier) {
  if (tier === "excellent") return "#0d6e3a";
  if (tier === "good")      return "#1a5276";
  if (tier === "neutral")   return "#784212";
  return "#922b21";
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function readFileAsBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload  = e => resolve(e.target.result);
    r.onerror = () => reject(new Error("فشل قراءة الملف"));
    r.readAsArrayBuffer(file);
  });
}

async function downloadDriveBuffer(file, token) {
  const isGSheet = file.mimeType === GSHEETS_MIME;
  const url = isGSheet
    ? `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
    : `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`تعذر تحميل الملف (${res.status})`);
  return {
    buffer: await res.arrayBuffer(),
    filename: isGSheet && !file.name.endsWith(".xlsx") ? `${file.name}.xlsx` : file.name,
  };
}

async function uploadReportNextToSource(blob, filename, sourceFile, token) {
  const mimeType = blob.type || (filename.toLowerCase().endsWith(".pdf") ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  const metadata = { name: filename, mimeType };
  if (sourceFile.parents?.[0]) metadata.parents = [sourceFile.parents[0]];
  const body = new FormData();
  body.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json; charset=UTF-8" }));
  body.append("file", blob, filename);
  const res = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body,
  });
  const data = await res.json();
  if (!res.ok || data.error) {
    const suffix = res.status === 403 ? " — افصل Drive وأعد ربطه لمنح صلاحية إنشاء التقارير." : "";
    throw new Error((data.error?.message || `فشل رفع التقرير (${res.status})`) + suffix);
  }
  return data;
}

function safeReportName(sourceName) {
  const base = String(sourceName || "استبيان").replace(/\.(xlsx|xls|csv)$/i, "").trim();
  return `تقرير تحليل - ${base}.docx`;
}

function safePdfReportName(sourceName) {
  const base = String(sourceName || "استبيان").replace(/\.(xlsx|xls|csv)$/i, "").trim();
  return `تقرير تحليل - ${base}.pdf`;
}

async function downloadReports(reports, year) {
  if (reports.length === 1) {
    downloadBlob(reports[0].blob, reports[0].filename);
    return;
  }
  if (reports.length > 1) {
    const zip = new JSZip();
    reports.forEach(report => zip.file(report.filename, report.blob));
    downloadBlob(await zip.generateAsync({ type: "blob" }), `تقارير_تحليل_${year || "الاستبيانات"}.zip`);
  }
}

// Filename → program detector. Inline Arabic normalization so we don't depend on engine internals.
function detectProgramFromFilename(filename) {
  const n = String(filename ?? "")
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .toLowerCase();
  if (n.includes("تكنولوجيا") && n.includes("اعمال"))  return "تكنولوجيا أعمال";
  if (n.includes("علوم")     && n.includes("سياسي")) return "علوم سياسية";
  if (n.includes("سياسي")) return "علوم سياسية";
  if (n.includes("محاسب"))  return "محاسبة";
  if (n.includes("اقتصاد")) return "اقتصاد";
  if (n.includes("ادار"))   return "إدارة";
  return null;
}

function detectYearFromFilename(filename) {
  const m = String(filename).match(/(\d{4}[-_]\d{4}|\d{4}[-_]\d{2})/);
  if (!m) return null;
  const parts = m[1].split(/[-_]/);
  if (parts[1].length === 2) return `${parts[0]}-20${parts[1]}`;
  return `${parts[0]}-${parts[1]}`;
}

function detectTypeHintFromFilename(filename) {
  const n = String(filename ?? "")
    .replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").toLowerCase();
  for (const s of Object.values(allSchemas())) {
    if (s.fileHints?.some(h =>
      n.includes(h.replace(/[أإآ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").toLowerCase())
    )) return s.id;
  }
  return null;
}

function uniqueCol(rows, colIdx) {
  if (colIdx == null) return [];
  const set = new Set();
  rows.forEach(r => {
    const v = r[colIdx];
    if (v != null && String(v).trim()) set.add(String(v));
  });
  return [...set].sort((a, b) => a.localeCompare(b, "ar"));
}

function computeFilteredRows(allRows, removed, metaCols, filters) {
  if (!allRows) return [];
  const { deptIdx, degreeIdx } = metaCols ?? {};
  return allRows.filter((row, idx) => {
    if (removed.has(idx)) return false;
    if (filters.dept   && deptIdx   != null && String(row[deptIdx])   !== filters.dept)   return false;
    if (filters.degree && degreeIdx != null && String(row[degreeIdx]) !== filters.degree) return false;
    return true;
  });
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_SETTINGS };
}

function saveSettings(s) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

function loadAiSettings() {
  try {
    const raw = localStorage.getItem(AI_KEY);
    if (raw) return { ...DEFAULT_AI_SETTINGS, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_AI_SETTINGS };
}

function saveAiSettings(s) {
  try { localStorage.setItem(AI_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

// ── SettingsPanel ─────────────────────────────────────────────────────────────
function SettingsPanel({ settings, onChange, aiSettings, onAiChange }) {
  const logoInputRef = useRef();

  const set = (key, value) => {
    const next = { ...settings, [key]: value };
    onChange(next);
    saveSettings(next);
  };

  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => set("logoDataUrl", ev.target.result);
    reader.readAsDataURL(file);
  };

  const reset = () => {
    onChange({ ...DEFAULT_SETTINGS });
    saveSettings({ ...DEFAULT_SETTINGS });
  };

  const F = (key, label, placeholder, type = "text") => (
    <div>
      <label className="label">{label}</label>
      <input
        className="input"
        type={type}
        value={settings[key] ?? ""}
        placeholder={placeholder}
        onChange={e => set(key, e.target.value)}
      />
    </div>
  );

  const T = (key, label, placeholder) => (
    <div>
      <label className="label">{label}</label>
      <textarea
        className="textarea"
        value={settings[key] ?? ""}
        placeholder={placeholder}
        onChange={e => set(key, e.target.value)}
      />
    </div>
  );

  const Toggle = (key, label) => (
    <div className="toggle-row">
      <span style={{ color: "#e8f0fe", fontSize: 14 }}>{label}</span>
      <button
        className={`toggle ${settings[key] ? "on" : "off"}`}
        onClick={() => set(key, !settings[key])}
      />
    </div>
  );

  return (
    <div>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 28 }}>
        <div style={{ color: "#fff", fontSize: 22, fontWeight: 900 }}>⚙ الإعدادات</div>
        <button className="btn btn-danger btn-sm" onClick={reset}>إعادة تعيين للافتراضي</button>
      </div>

      {/* Section 1 — Institutional Identity */}
      <div className="settings-section">
        <div className="settings-section-title">🏛 هوية المؤسسة (ترويسة التقرير)</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          {F("uniName",       "اسم الجامعة",   "Egyptian Russian University")}
          {F("facultyName",   "اسم الكلية",    "Faculty of Management...")}
          {F("unitName",      "اسم الوحدة",    "Quality Assurance unit (QAU)")}
          {F("committeeName", "اسم اللجنة",    "Measurement and Evaluation Committee")}
        </div>

        {/* Logo */}
        <div style={{ marginTop: 16 }}>
          <label className="label">شعار الجامعة</label>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            {settings.logoDataUrl
              ? <img src={settings.logoDataUrl} className="logo-preview" alt="logo" />
              : <div className="logo-preview" style={{ display: "flex", alignItems: "center", justifyContent: "center", color: "rgba(255,255,255,.3)", fontSize: 11 }}>لا يوجد</div>
            }
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <button className="btn btn-ghost btn-sm" onClick={() => logoInputRef.current.click()}>
                📎 رفع شعار
              </button>
              {settings.logoDataUrl && (
                <button className="btn btn-danger btn-sm" onClick={() => set("logoDataUrl", null)}>
                  حذف الشعار
                </button>
              )}
            </div>
            <input ref={logoInputRef} type="file" accept="image/*" style={{ display: "none" }}
              onChange={handleLogoUpload} />
          </div>
          <div style={{ color: "rgba(255,255,255,.3)", fontSize: 11, marginTop: 6 }}>
            PNG أو JPG — يظهر في الترويسة على كل صفحة. إذا لم يُرفع يُستخدم logo.png الافتراضي.
          </div>
        </div>
      </div>

      {/* Section 2 — Signature Block */}
      <div className="settings-section">
        <div className="settings-section-title">✍ كتلة التوقيع (نهاية التقرير)</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          {F("qmName", "رئيس وحدة القياس والتقويم", "د/ ...")}
          {F("quName", "رئيس وحدة الجودة",           "د/ ...")}
        </div>
      </div>

      {/* Section 3 — Footer */}
      <div className="settings-section">
        <div className="settings-section-title">📄 تذييل الصفحة</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {T("vision",  "رؤية الوحدة",       "رؤية الوحدة: ...")}
          {T("mission", "رسالة الوحدة",      "رسالة الوحدة: ...")}
          {F("email",   "البريد الإلكتروني", "E-mail: ...")}
        </div>
      </div>

      {/* Section 4 — Report Options */}
      <div className="settings-section">
        <div className="settings-section-title">📋 خيارات التقرير</div>
        {Toggle("includeEvaluatorsTable", "تضمين جدول القائم بالتقييم")}
        {Toggle("includeParticipants",    "تضمين توزيع المشاركين")}
        {Toggle("includeRecommendations", "تضمين قسم التوصيات")}
        {settings.includeRecommendations && (
          <div style={{ marginTop: 14 }}>
            <label className="label">عدد المحاور في التوصيات</label>
            <input
              className="input"
              type="number"
              min={1}
              max={17}
              style={{ width: 120 }}
              value={settings.recommendationsCount ?? 5}
              onChange={e => set("recommendationsCount", Math.max(1, parseInt(e.target.value) || 5))}
            />
          </div>
        )}
      </div>

      <div className="settings-section">
        <div className="settings-section-title">🧩 إعدادات كل استبيان</div>
        <div style={{ color: "rgba(255,255,255,.45)", fontSize: 12, marginBottom: 14 }}>
          تحكم مستقل في صفحة الرؤية والرسالة ولون التقرير لكل استبيان، ويُطبق على Word وPDF.
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {Object.values(allSchemas()).map(sc => {
            const opts = settings.surveyReportOptions?.[sc.id] ?? {};
            const includeVM = opts.includeVisionMission ?? settings.includeVisionMission ?? true;
            const theme = opts.colorTheme ?? settings.colorTheme ?? "default";
            const update = patch => set("surveyReportOptions", {
              ...(settings.surveyReportOptions ?? {}),
              [sc.id]: { ...opts, ...patch },
            });
            return (
              <div key={sc.id} style={{ display: "grid", gridTemplateColumns: "minmax(210px,1.5fr) minmax(180px,1fr) minmax(180px,1fr)", gap: 12, alignItems: "center", padding: "12px 14px", borderRadius: 12, border: "1px solid rgba(255,255,255,.08)", background: "rgba(255,255,255,.025)" }}>
                <div style={{ color: "#eef5ff", fontSize: 13, fontWeight: 700 }}>{sc.icon ?? "📋"} {sc.label}</div>
                <label style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, color: "rgba(255,255,255,.7)", fontSize: 12 }}>
                  إظهار الرؤية والرسالة
                  <button className={`toggle ${includeVM ? "on" : "off"}`} onClick={() => update({ includeVisionMission: !includeVM })} />
                </label>
                <select className="input" value={theme} onChange={e => update({ colorTheme: e.target.value })} style={{ cursor: "pointer" }} aria-label={`لون تقرير ${sc.label}`}>
                  <option value="default">أزرق</option>
                  <option value="green">أخضر</option>
                  <option value="purple">بنفسجي</option>
                  <option value="dark">رمادي</option>
                  <option value="red">أحمر</option>
                </select>
              </div>
            );
          })}
        </div>
      </div>

      {/* Section 4b — Branded PDF Options */}
      <div className="settings-section">
        <div className="settings-section-title">🖨️ تقرير PDF المصمم</div>
        {Toggle("includePdfCharts", "تضمين الرسوم البيانية (الأعمدة والأشرطة الملوّنة)")}
        <div style={{ marginTop: 14 }}>
          <label className="label">حد التوصيات — أقل من (%)</label>
          <input
            className="input"
            type="number"
            min={1}
            max={100}
            style={{ width: 120 }}
            value={settings.recommendationsThreshold ?? 70}
            onChange={e => set("recommendationsThreshold", Math.min(100, Math.max(1, parseInt(e.target.value) || 70)))}
          />
          <div style={{ color: "rgba(255,255,255,.4)", fontSize: 11.5, marginTop: 6 }}>
            أي عبارة تسجّل نسبة موافقة أقل من هذه القيمة تظهر في جدول التوصيات بتقرير الـ PDF (المحور، العبارة، النسبة).
          </div>
        </div>
      </div>

      {/* Section 5 — Report Design */}
      <div className="settings-section">
        <div className="settings-section-title">🎨 التصميم الافتراضي للتقرير (Word وPDF)</div>

        {/* Font */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">الخط</label>
          <select
            className="input"
            value={settings.reportFont ?? "Arial"}
            onChange={e => set("reportFont", e.target.value)}
            style={{ cursor: "pointer" }}
          >
            {["Arial", "Times New Roman", "Calibri", "Tahoma"].map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </div>

        {/* Direction */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">اتجاه التقرير</label>
          <div style={{ display: "flex", gap: 10 }}>
            {[{ val: "rtl", label: "من اليمين للشمال (RTL) ← عربي" },
              { val: "ltr", label: "من الشمال لليمين (LTR) → English" }].map(opt => (
              <button key={opt.val}
                onClick={() => set("reportDirection", opt.val)}
                style={{
                  flex: 1, padding: "9px 8px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
                  border: `2px solid ${settings.reportDirection === opt.val ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                  background: settings.reportDirection === opt.val ? "rgba(26,188,156,.15)" : "rgba(255,255,255,.04)",
                  color: settings.reportDirection === opt.val ? "#1abc9c" : "rgba(255,255,255,.55)",
                }}>{opt.label}</button>
            ))}
          </div>
        </div>

        {/* Table text & number alignment */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14, marginBottom: 16 }}>
          <div>
            <label className="label">محاذاة نص الجداول</label>
            <select className="input" value={settings.textAlign ?? "right"}
              onChange={e => set("textAlign", e.target.value)} style={{ cursor: "pointer" }}>
              <option value="right">يمين</option>
              <option value="center">وسط</option>
              <option value="left">يسار</option>
            </select>
          </div>
          <div>
            <label className="label">محاذاة الأرقام في الجداول</label>
            <select className="input" value={settings.numAlign ?? "center"}
              onChange={e => set("numAlign", e.target.value)} style={{ cursor: "pointer" }}>
              <option value="center">وسط</option>
              <option value="right">يمين</option>
              <option value="left">يسار</option>
            </select>
          </div>
        </div>

        {/* Color theme */}
        <div>
          <label className="label">ثيم الألوان</label>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {[
              { id: "default", label: "أزرق (افتراضي)", primary: "#1F3864", secondary: "#2E75B6" },
              { id: "green",   label: "أخضر",           primary: "#1a4731", secondary: "#2d6a4f" },
              { id: "purple",  label: "بنفسجي",         primary: "#4a235a", secondary: "#7d3c98" },
              { id: "dark",    label: "رمادي",          primary: "#1c1c1c", secondary: "#444444" },
              { id: "red",     label: "أحمر",           primary: "#b3373a", secondary: "#cf5a5d" },
            ].map(t => (
              <button key={t.id}
                onClick={() => set("colorTheme", t.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 8,
                  padding: "9px 14px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                  border: `2px solid ${settings.colorTheme === t.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                  background: settings.colorTheme === t.id ? "rgba(26,188,156,.12)" : "rgba(255,255,255,.04)",
                  color: settings.colorTheme === t.id ? "#1abc9c" : "rgba(255,255,255,.6)",
                }}>
                <span style={{
                  display: "inline-block", width: 18, height: 18, borderRadius: 4,
                  background: `linear-gradient(135deg,${t.primary},${t.secondary})`,
                  border: "1px solid rgba(255,255,255,.2)", flexShrink: 0,
                }} />
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Section 6 — Survey Responsible Persons */}
      <div className="settings-section">
        <div className="settings-section-title">👤 المسئولون عن الاستبيانات</div>
        <div style={{ color: "rgba(255,255,255,.4)", fontSize: 12, marginBottom: 16 }}>
          يُملأ حقل "أعده" تلقائياً عند اكتشاف نوع الاستبيان
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {Object.values(allSchemas()).map(sc => (
            <div key={sc.id} style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, alignItems: "center" }}>
              <div style={{ color: "#e8f0fe", fontSize: 13, fontWeight: 600 }}>
                {sc.icon ?? "📋"} {sc.label}
              </div>
              <input
                className="input"
                value={settings.surveyResponsible?.[sc.id] ?? ""}
                placeholder="اسم المسئول — مثال: د/ أحمد محمد"
                onChange={e => set("surveyResponsible", {
                  ...(settings.surveyResponsible ?? {}),
                  [sc.id]: e.target.value,
                })}
              />
            </div>
          ))}
        </div>
      </div>

      {/* Section 7 — AI Assistant */}
      <div className="settings-section">
        <div className="settings-section-title">🤖 المساعد الذكي</div>

        {/* Provider selector */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">المزود</label>
          <div style={{ display: "flex", gap: 8 }}>
            {Object.entries(PROVIDERS).map(([id, p]) => (
              <button
                key={id}
                onClick={() => {
                  const next = { ...aiSettings, provider: id, model: p.default };
                  onAiChange(next);
                  saveAiSettings(next);
                }}
                style={{
                  flex: 1, padding: "9px 6px", borderRadius: 10, cursor: "pointer",
                  fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                  border: `2px solid ${aiSettings.provider === id ? p.color : "rgba(255,255,255,.15)"}`,
                  background: aiSettings.provider === id ? `${p.color}22` : "rgba(255,255,255,.04)",
                  color: aiSettings.provider === id ? p.color : "rgba(255,255,255,.55)",
                  transition: "all .2s",
                }}
              >
                {p.icon} {p.label}
              </button>
            ))}
          </div>
        </div>

        {/* Model selector */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">الموديل</label>
          <input
            className="input"
            list={`model-list-${aiSettings.provider}`}
            value={aiSettings.model}
            placeholder="اكتب اسم الموديل أو اختر من القائمة"
            onChange={e => {
              const next = { ...aiSettings, model: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
            style={{ direction: "ltr", textAlign: "left" }}
          />
          <datalist id={`model-list-${aiSettings.provider}`}>
            {(PROVIDERS[aiSettings.provider]?.models || []).map(m => (
              <option key={m.id} value={m.id}>{m.label}</option>
            ))}
          </datalist>
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            اكتب اسم أي موديل مدعوم أو اختر من الاقتراحات
          </div>
        </div>

        {/* API key */}
        <div style={{ marginBottom: 16 }}>
          <label className="label">مفتاح API</label>
          <input
            className="input"
            type="password"
            value={aiSettings.key}
            placeholder={`sk-... أو gsk-... حسب المزود`}
            onChange={e => {
              const next = { ...aiSettings, key: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
          />
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            يُخزّن في المتصفح فقط · بديلاً يمكن تعيين متغير البيئة{" "}
            <code style={{ color: "#1abc9c" }}>{aiSettings.provider.toUpperCase()}_API_KEY</code>
            {" "}في ملف .env
          </div>
        </div>

        {/* Surveys folder */}
        <div>
          <label className="label">مجلد الاستبيانات (اختياري)</label>
          <input
            className="input"
            type="text"
            value={aiSettings.surveysFolder ?? ""}
            placeholder="مثال: C:\جودة\استبيانات"
            onChange={e => {
              const next = { ...aiSettings, surveysFolder: e.target.value };
              onAiChange(next);
              saveAiSettings(next);
            }}
          />
          <div style={{ color: "rgba(255,255,255,.28)", fontSize: 11, marginTop: 6 }}>
            المسار الكامل للمجلد المحلي الذي يحتوي ملفات Excel —
            يتيح للمساعد قراءة الملفات مباشرة دون رفع يدوي
          </div>
        </div>
      </div>
    </div>
  );
}

// ── BatchItem ─────────────────────────────────────────────────────────────────
function BatchItem({ item, year, settings, reportMeta }) {
  const [downloading, setDownloading] = useState(false);

  const doDownload = async () => {
    if (!item.result || !item.type) return;
    setDownloading(true);
    try {
      const meta = {
        year,
        program:    item.program ?? "",
        preparedBy: reportMeta.preparedBy,
        reviewer: reportMeta.reviewer,
      };
      const blob = await buildAnnualDocx(item.result, meta, settings);
      const prog  = item.program ? `_${item.program}` : "";
      const label = allSchemas()[item.type]?.label ?? "تقرير";
      downloadBlob(blob, `تقرير_${label}${prog}_${year}.docx`);
    } catch (e) { console.error(e); }
    finally    { setDownloading(false); }
  };

  const bg     = item.status === "done"       ? "rgba(26,188,156,.08)"
               : item.status === "error"      ? "rgba(96,165,250,.06)"
               : item.status === "processing" ? "rgba(255,255,255,.05)"
               :                               "rgba(255,255,255,.02)";
  const border = item.status === "done"       ? "rgba(26,188,156,.32)"
               : item.status === "error"      ? "rgba(147,197,253,.18)"
               : item.status === "processing" ? "rgba(255,255,255,.18)"
               :                               "rgba(255,255,255,.07)";

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 14,
      background: bg, border: `1px solid ${border}`,
      borderRadius: 12, padding: "13px 18px", transition: "background .3s",
    }}>
      {/* Status icon */}
      <div style={{ fontSize: 20, flexShrink: 0, width: 24, textAlign: "center" }}>
        {item.status === "done"       ? "✅"
       : item.status === "error"      ? "❌"
       : item.status === "processing" ? (
          <svg className="spin" width="20" height="20" viewBox="0 0 20 20" style={{ display: "block" }}>
            <circle cx="10" cy="10" r="7" fill="none" stroke="#1abc9c" strokeWidth="2.5" strokeDasharray="32 12"/>
          </svg>
        ) : "🕐"}
      </div>

      {/* Details */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ color: "#fff", fontWeight: 700, fontSize: 13,
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {item.file.name}
        </div>
        <div style={{ fontSize: 11, marginTop: 3 }}>
          {item.status === "done" && (
            <span style={{ display: "flex", gap: 10, flexWrap: "wrap", color: "rgba(255,255,255,.5)" }}>
              <span style={{ color: "#1abc9c", fontWeight: 700 }}>
                {allSchemas()[item.type]?.icon} {allSchemas()[item.type]?.label}
              </span>
              <span>·</span><span>{item.result.n} استجابة</span>
              {item.program && <><span>·</span><span>{item.program}</span></>}
              <span>·</span>
              <span style={{ color: "#1abc9c", fontWeight: 700 }}>{item.result.overallAgreePct}% موافقة</span>
            </span>
          )}
          {item.status === "processing" && <span style={{ color: "#1abc9c" }}>جاري التحليل…</span>}
          {item.status === "error"      && <span style={{ color: "#dce8f7" }}>{item.error}</span>}
          {item.status === "pending"    && <span style={{ color: "rgba(255,255,255,.3)" }}>في الانتظار</span>}
        </div>
      </div>

      {/* Download */}
      {item.status === "done" && (
        <button className="btn btn-ghost btn-sm" onClick={doDownload} disabled={downloading}
          style={{ flexShrink: 0, fontSize: 12, whiteSpace: "nowrap", opacity: downloading ? 0.6 : 1 }}>
          {downloading
            ? <svg className="spin" width="14" height="14" viewBox="0 0 14 14">
                <circle cx="7" cy="7" r="5" fill="none" stroke="white" strokeWidth="2" strokeDasharray="22 8"/>
              </svg>
            : "⬇ Word"}
        </button>
      )}
    </div>
  );
}

// ── BatchProcessor ────────────────────────────────────────────────────────────
function BatchProcessor({ files, year, onYearChange, settings, reportMeta, onReportMetaChange, onBack }) {
  const [items, setItems] = useState(() =>
    files.map(f => ({ file: f, status: "pending", result: null, error: null, type: null, program: null }))
  );
  const [running,        setRunning]        = useState(false);
  const [allDone,        setAllDone]        = useState(false);
  const [downloadingAll, setDownloadingAll] = useState(false);
  const [authorError, setAuthorError] = useState("");
  const cancelRef = useRef(false);
  const [cancelRequested, setCancelRequested] = useState(false);

  const doneCount  = items.filter(i => i.status === "done").length;
  const errorCount = items.filter(i => i.status === "error").length;
  const totalDone  = doneCount + errorCount;

  const startProcessing = async () => {
    if (!String(reportMeta.preparedBy ?? "").trim() || !String(reportMeta.reviewer ?? "").trim()) {
      setAuthorError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل بدء التحليل.");
      return;
    }
    setAuthorError("");
    cancelRef.current = false;
    setCancelRequested(false);
    setRunning(true);
    for (let idx = 0; idx < files.length; idx++) {
      if (cancelRef.current) break;
      setItems(prev => prev.map((it, i) => i === idx ? { ...it, status: "processing" } : it));
      await new Promise(r => setTimeout(r, 30));  // let React paint
      try {
        const buf          = await readFileAsBuffer(files[idx]);
        if (cancelRef.current) break;
        const rows         = readExcel(buf);
        const detectedType = detectSurveyType(files[idx].name, rows[0]) ?? "faculty";
        const s            = allSchemas()[detectedType];
        const result       = analyze(rows, s);
        const program      = detectProgramFromFilename(files[idx].name);
        setItems(prev => prev.map((it, i) => i === idx
          ? { ...it, status: "done", result, type: detectedType, program }
          : it
        ));
      } catch (err) {
        setItems(prev => prev.map((it, i) => i === idx
          ? { ...it, status: "error", error: err.message ?? "خطأ في القراءة" }
          : it
        ));
      }
    }
    setRunning(false);
    setAllDone(true);
  };

  const downloadAll = async () => {
    setDownloadingAll(true);
    const reports = [];
    for (const it of items) {
      if (it.status === "done" && it.result) {
        try {
          const meta  = { year, program: it.program ?? "", preparedBy: reportMeta.preparedBy, reviewer: reportMeta.reviewer };
          const blob  = await buildAnnualDocx(it.result, meta, settings);
          const prog  = it.program ? `_${it.program}` : "";
          const label = allSchemas()[it.type]?.label ?? "تقرير";
          reports.push({ blob, filename: `تقرير_${label}${prog}_${year}.docx` });
        } catch (e) { console.error(e); }
      }
    }
    await downloadReports(reports, year);
    setDownloadingAll(false);
  };

  return (
    <div className="card" style={{ padding: 36 }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between",
                    marginBottom: 24, gap: 16 }}>
        <div>
          <div style={{ color: "#fff", fontSize: 20, fontWeight: 900 }}>
            ⚡ معالجة دفعية — {files.length} ملف
          </div>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 13, marginTop: 4 }}>
            يُكتشف نوع كل استبيان تلقائياً · التحليل متسلسل
          </div>
        </div>
        {!running && (
          <button className="btn btn-ghost btn-sm" onClick={onBack}>→ إعادة الاختيار</button>
        )}
      </div>

      {/* Year input (before start) */}
      {!running && !allDone && (
        <div style={{ marginBottom: 24, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14 }}>
          <div>
          <label className="label">العام الأكاديمي (مشترك لجميع الملفات)</label>
          <input className="input" value={year}
            onChange={e => onYearChange(e.target.value)}
            placeholder="مثال: 2024-2025" />
          </div>
          <div>
            <label className="label">مُعدّ التحليل *</label>
            <input className="input" value={reportMeta.preparedBy ?? ""} onChange={e => onReportMetaChange({ ...reportMeta, preparedBy: e.target.value })} placeholder="الاسم الكامل" />
          </div>
          <div>
            <label className="label">مراجع التحليل *</label>
            <input className="input" value={reportMeta.reviewer ?? ""} onChange={e => onReportMetaChange({ ...reportMeta, reviewer: e.target.value })} placeholder="الاسم الكامل" />
          </div>
        </div>
      )}
      {!!authorError && <InlineNotice text={authorError} style={{ marginBottom: 18 }} />}

      {/* Progress bar */}
      {(running || allDone) && (
        <div style={{ marginBottom: 22 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 7 }}>
            <span style={{ color: "rgba(255,255,255,.6)", fontSize: 12.5 }}>
              {totalDone} من {files.length} ملف
            </span>
            <span style={{ fontSize: 12.5, fontWeight: 700 }}>
              <span style={{ color: "#1abc9c" }}>{doneCount} ناجح </span>
              {errorCount > 0 && <span style={{ color: "#fbbf24" }}>· {errorCount} فشل</span>}
            </span>
          </div>
          <div style={{ background: "rgba(255,255,255,.1)", borderRadius: 6, height: 7, overflow: "hidden" }}>
            <div style={{
              height: "100%", borderRadius: 6, transition: "width .4s ease",
              background: "linear-gradient(90deg,#1abc9c,#2874a6)",
              width: `${files.length ? (totalDone / files.length) * 100 : 0}%`,
            }} />
          </div>
        </div>
      )}

      {/* File list */}
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 28,
                    maxHeight: 480, overflowY: "auto" }}>
        {items.map((it, i) => (
          <BatchItem key={i} item={it} year={year} settings={settings} reportMeta={reportMeta} />
        ))}
      </div>

      {/* Actions */}
      <div style={{ display: "flex", justifyContent: "center", gap: 12, flexWrap: "wrap" }}>
        {running && (
          <button type="button" className="btn btn-ghost" disabled={cancelRequested}
            style={{ color: "#ffb4b4", borderColor: "rgba(239,68,68,.35)" }}
            onClick={() => { cancelRef.current = true; setCancelRequested(true); }}>
            {cancelRequested ? "جارٍ الإلغاء…" : "إلغاء العملية"}
          </button>
        )}
        {!running && !allDone && (
          <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 44px" }}
            onClick={startProcessing}>
            ▶ بدء التحليل
          </button>
        )}
        {allDone && doneCount > 1 && (
          <button className="btn btn-blue" style={{ fontSize: 15, padding: "12px 32px" }}
            disabled={downloadingAll} onClick={downloadAll}>
            {downloadingAll
              ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحميل…</>
              : `⬇ تحميل الكل (${doneCount} ملف Word)`}
          </button>
        )}
        {allDone && (
          <button className="btn btn-ghost" style={{ fontSize: 15, padding: "12px 28px" }}
            onClick={onBack}>
            ↩ رفع ملفات جديدة
          </button>
        )}
      </div>
    </div>
  );
}

// ── DriveDashboard ────────────────────────────────────────────────────────────
function DriveDashboard({ files, token, onSelectFile }) {
  const [stats, setStats] = useState(() =>
    files.map(f => ({
      id: f.id, name: f.name, mimeType: f.mimeType,
      modifiedTime: f.modifiedTime,
      type: detectTypeHintFromFilename(f.name),
      year: detectYearFromFilename(f.name),
      program: detectProgramFromFilename(f.name),
      count: null, status: "pending",
    }))
  );
  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(0);

  const loadStats = async () => {
    if (loading) return;
    setLoading(true);
    setDone(0);
    setStats(prev => prev.map(s => ({ ...s, status: "loading", count: null })));
    const BATCH = 4;
    for (let i = 0; i < files.length; i += BATCH) {
      const batch = files.slice(i, i + BATCH);
      await Promise.all(batch.map(async (f, bi) => {
        const idx = i + bi;
        try {
          const isGSheet = f.mimeType === GSHEETS_MIME;
          const url = isGSheet
            ? `https://www.googleapis.com/drive/v3/files/${f.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
            : `https://www.googleapis.com/drive/v3/files/${f.id}?alt=media`;
          const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
          if (!res.ok) throw new Error();
          const buf = await res.arrayBuffer();
          const rows = readExcel(buf);
          const count = Math.max(0, rows.length - 1);
          setStats(prev => prev.map((s, si) => si === idx ? { ...s, count, status: "done" } : s));
        } catch {
          setStats(prev => prev.map((s, si) => si === idx ? { ...s, status: "error" } : s));
        }
        setDone(prev => prev + 1);
      }));
    }
    setLoading(false);
  };

  const doneStats   = stats.filter(s => s.status === "done");
  const totalResp   = doneStats.reduce((a, s) => a + (s.count ?? 0), 0);

  const aggregate = (key) => {
    const map = {};
    stats.forEach(s => {
      const k = s[key] ?? "غير محدد";
      if (!map[k]) map[k] = { files: 0, responses: 0 };
      map[k].files++;
      if (s.count != null) map[k].responses += s.count;
    });
    return map;
  };
  const byType    = aggregate("type");
  const byProgram = aggregate("program");
  const byYear    = aggregate("year");

  const maxProg = Math.max(...Object.values(byProgram).map(v => v.files), 1);
  const maxYear = Math.max(...Object.values(byYear).map(v => v.files), 1);

  const barBtn = (active, label, onClick) => (
    <button onClick={onClick} style={{
      padding: "6px 12px", borderRadius: 8, border: "none", cursor: "pointer",
      fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
      background: active ? "rgba(26,188,156,.25)" : "rgba(255,255,255,.08)",
      color: active ? "#1abc9c" : "rgba(255,255,255,.5)",
      borderBottom: `2px solid ${active ? "#1abc9c" : "transparent"}`,
    }}>{label}</button>
  );

  return (
    <div style={{ textAlign: "right" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 13 }}>
            {files.length} ملف
          </span>
          {doneStats.length > 0 && (
            <span style={{ color: "#1abc9c", fontWeight: 700, fontSize: 13, marginRight: 10 }}>
              · {totalResp.toLocaleString()} استجابة إجمالاً
            </span>
          )}
          {loading && (
            <div style={{ color: "rgba(255,255,255,.4)", fontSize: 11, marginTop: 2 }}>
              جاري تحليل {done} من {files.length}…
            </div>
          )}
        </div>
        {!loading && done < files.length && (
          <button onClick={loadStats} style={{
            padding: "8px 16px", background: "linear-gradient(135deg,#1abc9c,#16a085)",
            border: "none", borderRadius: 20, color: "#fff",
            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12, cursor: "pointer",
          }}>
            📊 تحليل الملفات ({files.length})
          </button>
        )}
        {loading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <svg className="spin" width="16" height="16" viewBox="0 0 16 16">
              <circle cx="8" cy="8" r="6" fill="none" stroke="#1abc9c" strokeWidth="2.5" strokeDasharray="24 8"/>
            </svg>
            <span style={{ color: "rgba(255,255,255,.5)", fontSize: 12 }}>
              {Math.round(done / files.length * 100)}%
            </span>
          </div>
        )}
      </div>

      {/* Progress bar */}
      {(loading || done > 0) && (
        <div style={{ background: "rgba(255,255,255,.1)", borderRadius: 4, height: 4, marginBottom: 14, overflow: "hidden" }}>
          <div style={{
            height: "100%", background: "linear-gradient(90deg,#1abc9c,#2874a6)", borderRadius: 4,
            transition: "width .4s", width: `${Math.round(done / files.length * 100)}%`,
          }} />
        </div>
      )}

      {/* Cards by survey type */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 8, letterSpacing: .4 }}>
          حسب نوع الاستبيان
        </div>
        <div style={{ display: "flex", gap: 8, overflowX: "auto", paddingBottom: 4 }}>
          {Object.entries(byType).map(([tid, d]) => {
            const sc = allSchemas()[tid];
            return (
              <div key={tid} style={{
                background: "rgba(255,255,255,.06)", border: "1px solid rgba(255,255,255,.1)",
                borderRadius: 12, padding: "12px 14px", textAlign: "center", minWidth: 110, flexShrink: 0,
              }}>
                <div style={{ fontSize: 20, marginBottom: 3 }}>{sc?.icon ?? "📋"}</div>
                <div style={{ color: "#e8f0fe", fontWeight: 700, fontSize: 11, marginBottom: 4,
                              whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 100 }}>
                  {sc?.label ?? "غير معروف"}
                </div>
                <div style={{ color: "#1abc9c", fontWeight: 900, fontSize: 20 }}>{d.files}</div>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10 }}>ملف</div>
                {d.responses > 0 && (
                  <>
                    <div style={{ color: "#f0b27a", fontWeight: 700, fontSize: 14, marginTop: 4 }}>
                      {d.responses.toLocaleString()}
                    </div>
                    <div style={{ color: "rgba(255,255,255,.35)", fontSize: 10 }}>استجابة</div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Bar charts */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14, marginBottom: 16 }}>
        {/* By program */}
        <div style={{ background: "rgba(255,255,255,.04)", borderRadius: 12, padding: "12px 14px", border: "1px solid rgba(255,255,255,.07)" }}>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 10, letterSpacing: .4 }}>البرامج</div>
          {Object.entries(byProgram).sort((a, b) => b[1].files - a[1].files).map(([prog, d]) => (
            <div key={prog} style={{ marginBottom: 7 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                <span style={{ color: "#e8f0fe", fontSize: 11 }}>{prog}</span>
                <span style={{ color: "#1abc9c", fontSize: 11, fontWeight: 700 }}>
                  {d.files}{d.responses > 0 ? ` · ${d.responses.toLocaleString()}` : ""}
                </span>
              </div>
              <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 6, overflow: "hidden" }}>
                <div style={{
                  height: "100%", borderRadius: 3, transition: "width .5s",
                  background: "linear-gradient(90deg,#1abc9c,#2874a6)",
                  width: `${d.files / maxProg * 100}%`,
                }} />
              </div>
            </div>
          ))}
        </div>

        {/* By year */}
        <div style={{ background: "rgba(255,255,255,.04)", borderRadius: 12, padding: "12px 14px", border: "1px solid rgba(255,255,255,.07)" }}>
          <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 10, letterSpacing: .4 }}>السنوات الدراسية</div>
          {Object.entries(byYear).sort((a, b) => b[0].localeCompare(a[0])).map(([yr, d]) => (
            <div key={yr} style={{ marginBottom: 7 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                <span style={{ color: "#e8f0fe", fontSize: 11 }}>{yr}</span>
                <span style={{ color: "#a29bfe", fontSize: 11, fontWeight: 700 }}>
                  {d.files}{d.responses > 0 ? ` · ${d.responses.toLocaleString()}` : ""}
                </span>
              </div>
              <div style={{ background: "rgba(255,255,255,.08)", borderRadius: 3, height: 6, overflow: "hidden" }}>
                <div style={{
                  height: "100%", borderRadius: 3, transition: "width .5s",
                  background: "linear-gradient(90deg,#8e44ad,#2874a6)",
                  width: `${d.files / maxYear * 100}%`,
                }} />
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Detailed table */}
      <div style={{ color: "rgba(255,255,255,.45)", fontSize: 11, fontWeight: 700, marginBottom: 8, letterSpacing: .4 }}>
        تفاصيل الملفات
      </div>
      <div style={{ maxHeight: 260, overflowY: "auto", border: "1px solid rgba(255,255,255,.1)", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "rgba(255,255,255,.08)", position: "sticky", top: 0, zIndex: 1 }}>
              {["الملف","النوع","البرنامج","السنة","الاستجابات",""].map((h, i) => (
                <th key={i} style={{ padding: "7px 10px", color: "rgba(255,255,255,.55)", fontWeight: 700,
                                     textAlign: i === 0 ? "right" : "center", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {stats.map((s, i) => (
              <tr key={s.id} style={{ borderTop: "1px solid rgba(255,255,255,.05)", cursor: "default" }}>
                <td style={{ padding: "7px 10px", color: "#e8f0fe", maxWidth: 170,
                             overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {s.name}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center" }}>
                  {s.type
                    ? <span style={{ fontSize: 10, color: "#1abc9c", background: "rgba(26,188,156,.12)", borderRadius: 6, padding: "2px 7px", whiteSpace: "nowrap" }}>
                        {allSchemas()[s.type]?.label ?? s.type}
                      </span>
                    : <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center", color: "#d6eaf8", fontSize: 11, whiteSpace: "nowrap" }}>
                  {s.program ?? <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center", color: "#d6eaf8", fontSize: 11, whiteSpace: "nowrap" }}>
                  {s.year ?? <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "7px 10px", textAlign: "center" }}>
                  {s.status === "loading"
                    ? <svg className="spin" width="13" height="13" viewBox="0 0 13 13">
                        <circle cx="6.5" cy="6.5" r="5" fill="none" stroke="#1abc9c" strokeWidth="2" strokeDasharray="18 8"/>
                      </svg>
                    : s.status === "done"
                    ? <span style={{ color: "#1abc9c", fontWeight: 700 }}>{s.count?.toLocaleString()}</span>
                    : s.status === "error"
                    ? <span style={{ color: "#dce8f7", fontSize: 10 }}>تعذّر التحميل</span>
                    : <span style={{ color: "rgba(255,255,255,.2)" }}>—</span>}
                </td>
                <td style={{ padding: "6px 8px", textAlign: "center" }}>
                  <button onClick={() => onSelectFile(files[i])} style={{
                    padding: "3px 10px", background: "rgba(26,188,156,.13)",
                    border: "1px solid rgba(26,188,156,.3)", borderRadius: 7,
                    color: "#1abc9c", cursor: "pointer", fontSize: 11,
                    fontFamily: "'Cairo',sans-serif", fontWeight: 700,
                  }}>تحليل</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── LoadingOverlay ────────────────────────────────────────────────────────────
// ── TutorialOverlay ───────────────────────────────────────────────────────────
const TUTORIAL_SLIDES = [
  {
    icon: "🎉",
    color: "#1abc9c",
    title: "مرحباً بك في محلل الاستبيانات",
    subtitle: "دليل سريع – 6 خطوات",
    body: "هذا التطبيق يحلّل نتائج استبيانات وحدة الجودة ويولّد تقارير Word كاملة بالإحصاءات والتوصيات.",
    tips: [
      "يدعم 5 أنواع استبيانات: طلاب، هيئة تدريس، هيئة معاونة، خريجين، منسقي برامج",
      "يولّد تقارير Word جاهزة للطباعة مع كل الجداول والمحاور",
      "يدعم تقارير المقارنة بين 2 أو 3 سنوات دراسية",
      "يمكن ربطه بـ Google Drive مباشرةً",
    ],
  },
  {
    icon: "📂",
    color: "#2874a6",
    title: "الخطوة ١ – رفع ملف الاستبيان",
    subtitle: "رفع ملف · Google Drive",
    body: "ارفع ملف Excel الصادر من Google Forms عبر السحب والإفلات، أو اربط حسابك على Google Drive واختر الملف مباشرةً.",
    tips: [
      "يدعم .xlsx / .xls / .csv",
      "يكتشف نوع الاستبيان تلقائياً من اسم الملف والأعمدة",
      "لو اخترت أكثر من ملف تفتح وضع المعالجة الدفعية",
      "في Drive يمكن تصفية الملفات حسب النوع والسنة والبرنامج",
    ],
  },
  {
    icon: "✅",
    color: "#1abc9c",
    title: "الخطوة ٢ – التحقق من نوع الاستبيان",
    subtitle: "مراجعة الاكتشاف التلقائي",
    body: "يعرض التطبيق نوع الاستبيان الذي اكتشفه مع مستوى الثقة. يمكنك تأكيده أو تغييره يدوياً.",
    tips: [
      "الاكتشاف يعتمد على اسم الملف وترويسات الأعمدة",
      "لو النوع مش صح، غيّره من القائمة المنسدلة",
      "اختر نمط الاستبيان: مقياس 3 درجات أو 5 درجات",
    ],
  },
  {
    icon: "📋",
    color: "#9b59b6",
    title: "الخطوة ٣ – نوع التقرير",
    subtitle: "سنوي · مقارنة سنتين · مقارنة 3 سنوات",
    body: "اختر إذا كنت تريد تقريراً لسنة واحدة، أو مقارنة نتائج سنتين أو ثلاث سنوات في نفس التقرير.",
    tips: [
      "التقرير السنوي يحلّل ملفاً واحداً بالكامل",
      "تقرير المقارنة يرصد التطور ويظهر أسهم تحسّن/تراجع",
      "في المقارنة ارفع ملف لكل سنة",
    ],
  },
  {
    icon: "👁",
    color: "#e67e22",
    title: "الخطوة ٤ – معاينة وتنقية البيانات",
    subtitle: "مراجعة الصفوف قبل التحليل",
    body: "راجع بيانات المشاركين وأزِل الصفوف الخاطئة أو المكررة قبل الشروع في التحليل.",
    tips: [
      "التكرارات تُحذف تلقائياً في الخطوة السابقة",
      "يمكن تصفية البيانات حسب القسم أو الدرجة الوظيفية",
      "حدّد الصفوف يدوياً ثم اضغط 'حذف المحدد'",
    ],
  },
  {
    icon: "📄",
    color: "#1abc9c",
    title: "الخطوات ٥ و ٦ – البيانات والنتائج",
    subtitle: "أدخل معلومات التقرير ثم حمّله",
    body: "أدخل بيانات التقرير (السنة، البرنامج، المُعد، المراجع) ثم شاهد النتائج الكاملة وحمّل ملف Word.",
    tips: [
      "البيانات المُدخلة تُحفظ تلقائياً لأول مرة تُعبّأ",
      "التقرير يشمل: المقدمة، المشاركين، النتائج بالمحاور، الملخص، التوصيات",
      "يمكن تخصيص الألوان والخطوط وتوقيع التقرير من الإعدادات ⚙",
      "زر 'عرض محسّن' يفتح عرضاً تفاعلياً داخل التطبيق",
    ],
  },
];

function TutorialOverlay({ onClose }) {
  const [idx,  setIdx]  = useState(0);
  const [dir,  setDir]  = useState("next"); // "next" | "prev"
  const [anim, setAnim] = useState(false);
  const slide = TUTORIAL_SLIDES[idx];
  const total = TUTORIAL_SLIDES.length;

  const go = (next) => {
    if (next < 0 || next >= total) return;
    setDir(next > idx ? "next" : "prev");
    setAnim(false);
    // micro-delay so the class is stripped then re-added
    requestAnimationFrame(() => {
      setIdx(next);
      setAnim(true);
    });
  };

  return (
    <div
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}
      style={{
        position: "fixed", inset: 0, zIndex: 9100,
        background: "rgba(4,12,24,.88)", backdropFilter: "blur(16px)",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontFamily: "'Cairo',sans-serif", padding: 20,
      }}
    >
      <div className="tut-card" style={{
        background: "linear-gradient(155deg,rgba(255,255,255,.08),rgba(255,255,255,.03))",
        border: "1px solid rgba(255,255,255,.13)",
        borderRadius: 28, width: "100%", maxWidth: 560,
        boxShadow: "0 48px 96px rgba(0,0,0,.55), 0 0 0 1px rgba(255,255,255,.04)",
        overflow: "hidden",
        position: "relative",
      }}>
        {/* Accent top bar */}
        <div style={{ height: 4, background: `linear-gradient(90deg,${slide.color},${slide.color}88)` }} />

        {/* Close */}
        <button onClick={onClose} aria-label="إغلاق الدليل التعليمي" style={{
          position: "absolute", top: 16, left: 16,
          width: 32, height: 32, borderRadius: 8,
          border: "1px solid rgba(255,255,255,.15)",
          background: "rgba(255,255,255,.07)",
          color: "rgba(255,255,255,.5)", fontSize: 16,
          cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
          transition: "all .18s",
        }}
          onMouseOver={e => { e.currentTarget.style.background = "rgba(231,76,60,.2)"; e.currentTarget.style.color = "#ff6b6b"; }}
          onMouseOut={e  => { e.currentTarget.style.background = "rgba(255,255,255,.07)"; e.currentTarget.style.color = "rgba(255,255,255,.5)"; }}
        >✕</button>

        {/* Slide content */}
        <div
          key={idx}
          className={anim ? (dir === "next" ? "tut-slide-next" : "tut-slide-prev") : ""}
          style={{ padding: "28px 36px 24px", textAlign: "center" }}
        >
          {/* Icon */}
          <div style={{
            width: 72, height: 72, borderRadius: 20, margin: "0 auto 18px",
            background: `linear-gradient(135deg,${slide.color}30,${slide.color}12)`,
            border: `1px solid ${slide.color}45`,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 34,
          }}>{slide.icon}</div>

          {/* Title */}
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 20, marginBottom: 4, lineHeight: 1.4 }}>
            {slide.title}
          </div>
          <div style={{ color: slide.color, fontSize: 12, fontWeight: 700, marginBottom: 16, letterSpacing: .5 }}>
            {slide.subtitle}
          </div>

          {/* Body */}
          <div style={{
            color: "rgba(255,255,255,.65)", fontSize: 14, lineHeight: 1.8,
            marginBottom: 20, textAlign: "right",
          }}>{slide.body}</div>

          {/* Tips */}
          <div style={{
            background: "rgba(255,255,255,.05)", borderRadius: 14,
            border: "1px solid rgba(255,255,255,.08)",
            padding: "14px 18px", textAlign: "right",
          }}>
            {slide.tips.map((t, i) => (
              <div key={i} style={{
                display: "flex", alignItems: "flex-start", gap: 10,
                padding: "6px 0",
                borderBottom: i < slide.tips.length - 1 ? "1px solid rgba(255,255,255,.05)" : "none",
              }}>
                <div style={{
                  width: 20, height: 20, borderRadius: 6, flexShrink: 0, marginTop: 1,
                  background: `${slide.color}25`, border: `1px solid ${slide.color}40`,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: 10, color: slide.color, fontWeight: 900,
                }}>{i + 1}</div>
                <div style={{ color: "rgba(255,255,255,.7)", fontSize: 13, lineHeight: 1.7 }}>{t}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer navigation */}
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "16px 36px 24px",
          borderTop: "1px solid rgba(255,255,255,.07)",
        }}>
          {/* Prev */}
          <button className="tut-nav" onClick={() => go(idx - 1)} disabled={idx === 0}>‹</button>

          {/* Dots */}
          <div style={{ display: "flex", gap: 7, alignItems: "center" }}>
            {TUTORIAL_SLIDES.map((_, i) => (
              <div
                key={i}
                className="tut-dot"
                onClick={() => go(i)}
                style={{
                  background: i === idx ? slide.color : "rgba(255,255,255,.2)",
                  transform: i === idx ? "scale(1.35)" : "scale(1)",
                }}
              />
            ))}
          </div>

          {/* Next / Done */}
          {idx < total - 1 ? (
            <button className="tut-nav" onClick={() => go(idx + 1)}>›</button>
          ) : (
            <button
              onClick={onClose}
              style={{
                padding: "8px 20px", borderRadius: 22, border: "none",
                background: `linear-gradient(135deg,#1abc9c,#16a085)`,
                color: "#fff", fontFamily: "'Cairo',sans-serif",
                fontWeight: 700, fontSize: 13, cursor: "pointer",
              }}
            >ابدأ الآن ✓</button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── ProcessingOverlay ─────────────────────────────────────────────────────────
function ProcessingOverlay({ steps, filename }) {
  const doneCount = steps.filter(s => s.status === "done").length;
  const pct = Math.round((doneCount / steps.length) * 100);
  const activeLabel = steps.find(s => s.status === "active")?.label || "إنهاء المعالجة";

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 8900,
      background: "rgba(5,14,28,.93)", backdropFilter: "blur(14px)",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Cairo',sans-serif",
    }}>
      <div className="proc-overlay" style={{
        background: "linear-gradient(145deg,rgba(255,255,255,.07),rgba(255,255,255,.03))",
        border: "1px solid rgba(255,255,255,.13)",
        borderRadius: 24, padding: "36px 44px",
        width: "min(460px, calc(100vw - 32px))", maxWidth: 460,
        boxShadow: "0 40px 80px rgba(0,0,0,.6), 0 0 0 1px rgba(26,188,156,.08)",
      }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div className="loader-visual">
            <span className="loader-ring"/><span className="loader-ring inner"/>
            <span style={{ width: 26, height: 26, color: "#d8fff7" }}><StepIcon name="chart" /></span>
          </div>
          <div style={{ color: "#fff", fontWeight: 900, fontSize: 18, marginBottom: 6 }}>
            {activeLabel}<span className="loader-dots"><i/><i/><i/></span>
          </div>
          {filename && (
            <div style={{
              color: "rgba(255,255,255,.35)", fontSize: 11,
              maxWidth: 320, margin: "0 auto",
              overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
            }}>{filename}</div>
          )}
        </div>

        {/* Steps */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 24 }}>
          {steps.map((s, i) => (
            <div
              key={i}
              className={`proc-step-${s.status}`}
              style={{
                display: "flex", alignItems: "center", gap: 14,
                padding: "10px 16px", borderRadius: 12,
                background: s.status === "done"   ? "rgba(26,188,156,.1)"  :
                            s.status === "active" ? "rgba(255,255,255,.07)" : "transparent",
                border: `1px solid ${
                  s.status === "done"   ? "rgba(26,188,156,.28)"  :
                  s.status === "active" ? "rgba(255,255,255,.12)" : "transparent"
                }`,
                transition: "background .3s, border-color .3s",
              }}
            >
              {/* Icon */}
              <div className="proc-icon" style={{
                width: 28, height: 28, borderRadius: 8, flexShrink: 0,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 14,
                background: s.status === "done"   ? "rgba(26,188,156,.22)" :
                            s.status === "active" ? "rgba(255,255,255,.1)"  : "rgba(255,255,255,.04)",
                border: `1px solid ${
                  s.status === "done"   ? "rgba(26,188,156,.4)"  :
                  s.status === "active" ? "rgba(255,255,255,.18)" : "rgba(255,255,255,.07)"
                }`,
              }}>
                {s.status === "done"   ? <span style={{ color: "#1abc9c", fontWeight: 900, fontSize: 13 }}>✓</span> :
                 s.status === "active" ? (
                   <svg className="spin" width="14" height="14" viewBox="0 0 14 14">
                     <circle cx="7" cy="7" r="5.5" fill="none" stroke="rgba(255,255,255,.2)" strokeWidth="2"/>
                     <path d="M7 1.5a5.5 5.5 0 0 1 5.5 5.5" fill="none" stroke="#1abc9c" strokeWidth="2" strokeLinecap="round"/>
                   </svg>
                 ) : <span style={{ color: "rgba(255,255,255,.18)", fontSize: 11 }}>○</span>}
              </div>

              {/* Label */}
              <div style={{
                color: s.status === "done"   ? "#1abc9c"               :
                       s.status === "active" ? "#fff"                   : "rgba(255,255,255,.28)",
                fontSize: 14,
                fontWeight: s.status === "active" ? 700 : 600,
                transition: "color .3s",
                flex: 1,
              }}>{s.label}</div>

              {/* Timing badge for done */}
              {s.status === "done" && (
                <div style={{
                  fontSize: 10, color: "rgba(26,188,156,.6)",
                  background: "rgba(26,188,156,.1)", borderRadius: 6,
                  padding: "2px 7px", fontWeight: 700,
                }}>تم</div>
              )}
            </div>
          ))}
        </div>

        {/* Progress bar */}
        <div style={{ height: 5, background: "rgba(255,255,255,.07)", borderRadius: 3, overflow: "hidden" }}>
          <div style={{
            height: "100%", borderRadius: 3,
            background: "linear-gradient(90deg,#1abc9c,#2874a6,#1abc9c)",
            backgroundSize: "200% 100%",
            width: `${pct}%`,
            transition: "width .45s cubic-bezier(.4,0,.2,1)",
            animation: pct < 100 ? "progressFlow 1.8s linear infinite" : "none",
          }}/>
        </div>
        <div style={{ textAlign: "center", marginTop: 8, color: "rgba(255,255,255,.3)", fontSize: 11 }}>
          {pct}%
        </div>
      </div>
    </div>
  );
}

function LoadingOverlay({ message = "جاري المعالجة…", progress, onCancel, cancelling = false }) {
  const hasProgress = progress && progress.total > 0;
  const pct = hasProgress ? Math.round((progress.current / progress.total) * 100) : 0;
  return (
    <div role="status" aria-live="polite" aria-busy={!cancelling} style={{
      position: "fixed", inset: 0, zIndex: 8500,
      background: "rgba(9,20,38,.90)", backdropFilter: "blur(12px)",
      display: "flex", flexDirection: "column", alignItems: "center",
      justifyContent: "center", gap: 24, fontFamily: "'Cairo',sans-serif",
    }}>
      <div className="loader-visual" style={{ margin: 0 }}>
        <span className="loader-ring"/><span className="loader-ring inner"/>
        <span style={{ width: 26, height: 26, color: "#d8fff7" }}><StepIcon name="edit" /></span>
      </div>
      <div style={{ color: "#fff", fontSize: 18, fontWeight: 700, textAlign: "center" }}>
        {cancelling ? "جارٍ إيقاف إنشاء PDF" : message}<span className="loader-dots"><i/><i/><i/></span>
      </div>
      {hasProgress ? (
        <div style={{ width: 280 }}>
          <div style={{ height: 8, background: "rgba(255,255,255,.1)", borderRadius: 4, overflow: "hidden" }}>
            <div style={{
              height: "100%", borderRadius: 4, width: `${pct}%`,
              background: "linear-gradient(90deg,#1abc9c,#2874a6,#1abc9c)",
              backgroundSize: "200% 100%",
              animation: "progressFlow 1.8s linear infinite",
              transition: "width .3s ease",
            }} />
          </div>
          <div style={{ textAlign: "center", marginTop: 10, color: "rgba(255,255,255,.55)", fontSize: 13, fontWeight: 700 }}>
            صفحة {progress.current} من {progress.total} ({pct}%)
          </div>
        </div>
      ) : (
        <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>يرجى الانتظار…</div>
      )}
      {onCancel && <button type="button" className="btn btn-ghost btn-sm" onClick={onCancel} disabled={cancelling}
        style={{ color: cancelling ? "#ffb4b4" : undefined, borderColor: cancelling ? "rgba(239,68,68,.4)" : undefined }}>
        {cancelling ? "جارٍ الإلغاء…" : "إلغاء العملية"}
      </button>}
    </div>
  );
}

function DriveBatchOverlay({ state, onClose, onCancel }) {
  if (!state) return null;
  const doneCount = state.items.filter(item => item.status === "done").length;
  const errorCount = state.items.filter(item => item.status === "error").length;
  const downloadableItems = state.items.filter(item => item.status === "done" && item.localBlob);
  const finished = doneCount + errorCount;
  const pct = state.total ? Math.round((finished / state.total) * 100) : 0;
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 8950, background: "rgba(5,14,28,.94)", backdropFilter: "blur(15px)", display: "grid", placeItems: "center", padding: 18 }}>
      <div className="proc-overlay" style={{ width: "min(650px,100%)", maxHeight: "min(720px,92vh)", display: "flex", flexDirection: "column", background: "linear-gradient(150deg,#10253e,#0b1b2f)", border: "1px solid rgba(255,255,255,.13)", borderRadius: 24, padding: 26, boxShadow: "0 36px 90px rgba(0,0,0,.58)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 18 }}>
          <div className="loader-visual" style={{ width: 58, height: 58, margin: 0, flexShrink: 0 }}>
            <span className="loader-ring" style={state.done ? { animation: "none", borderColor: errorCount ? "#f59e0b" : "#1abc9c" } : undefined}/>
            <span style={{ width: 21, height: 21, color: state.done ? (errorCount ? "#fbbf24" : "#71e8d4") : "#d8fff7" }}><StepIcon name={state.done ? "shield" : "upload"}/></span>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ color: "#fff", fontSize: 18, fontWeight: 900 }}>{state.done ? "اكتملت تقارير PDF" : "تحليل الاستبيانات وإنشاء PDF"}</div>
            <div style={{ color: "rgba(255,255,255,.48)", fontSize: 12, marginTop: 3 }}>
              {state.done ? `${state.cancelled ? "تم إلغاء العملية بعد" : state.delivery === "download" ? "تم تنزيل" : "تم رفع"} ${doneCount} تقرير${errorCount ? ` وتعذر ${errorCount}` : ""}` : state.cancelRequested ? "جارٍ إلغاء العملية…" : `جاري معالجة الملف ${state.current} من ${state.total}`}
            </div>
          </div>
          {state.done && <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
            {!!downloadableItems.length && <button className="btn btn-primary btn-sm" onClick={() => downloadableItems.forEach((item, index) => setTimeout(() => downloadBlob(item.localBlob, item.reportName || `report-${index + 1}.pdf`), index * 180))}>تحميل الكل ({downloadableItems.length})</button>}
            <button className="btn btn-ghost btn-sm" onClick={onClose}>إغلاق</button>
          </div>}
          {!state.done && <button type="button" className="btn btn-ghost btn-sm" onClick={onCancel} disabled={state.cancelRequested} style={{ color: "#ffb4b4", borderColor: "rgba(239,68,68,.35)" }}>{state.cancelRequested ? "جارٍ الإلغاء…" : "إلغاء العملية"}</button>}
        </div>
        <div style={{ height: 7, borderRadius: 7, background: "rgba(255,255,255,.07)", overflow: "hidden", marginBottom: 16 }}>
          <div style={{ width: `${pct}%`, height: "100%", borderRadius: 7, transition: "width .35s ease", background: errorCount ? "linear-gradient(90deg,#1abc9c,#f59e0b)" : "linear-gradient(90deg,#1abc9c,#2874a6,#1abc9c)", backgroundSize: "200% 100%", animation: !state.done ? "progressFlow 1.8s linear infinite" : "none" }}/>
        </div>
        <div style={{ overflowY: "auto", display: "flex", flexDirection: "column", gap: 7, paddingLeft: 3 }}>
          {state.items.map(item => (
            <div key={item.id} className={`batch-row ${item.status}`} style={{ display: "flex", alignItems: "center", gap: 11, padding: "10px 12px", borderRadius: 11, border: `1px solid ${item.status === "done" ? "rgba(26,188,156,.25)" : item.status === "error" ? "rgba(231,76,60,.3)" : "rgba(255,255,255,.07)"}`, backgroundColor: item.status === "done" ? "rgba(26,188,156,.08)" : item.status === "error" ? "rgba(231,76,60,.08)" : "rgba(255,255,255,.025)" }}>
              <span style={{ width: 25, height: 25, display: "grid", placeItems: "center", flexShrink: 0, color: item.status === "done" ? "#5eead4" : item.status === "error" ? "#ff8a80" : "rgba(255,255,255,.35)" }}>
                {item.status === "active" ? <svg className="spin" viewBox="0 0 20 20" style={{ width: 18, height: 18 }}><circle cx="10" cy="10" r="7" fill="none" stroke="currentColor" strokeWidth="2" strokeDasharray="28 12"/></svg> : item.status === "done" ? "✓" : item.status === "error" ? "!" : "○"}
              </span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ color: "#e8f0fe", fontSize: 12.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{item.name}</div>
                <div style={{ color: item.status === "error" ? "#ff9d96" : "rgba(255,255,255,.38)", fontSize: 10.5, marginTop: 2 }}>{item.stage}</div>
              </div>
              {item.status === "done" && (
                <div style={{ display: "flex", alignItems: "center", gap: 7, flexShrink: 0 }}>
                  {item.report?.webViewLink && <a href={item.report.webViewLink} target="_blank" rel="noreferrer" style={{ color: "#71e8d4", fontSize: 10.5, textDecoration: "none" }}>فتح على Drive</a>}
                  {item.localBlob && <button type="button" onClick={() => downloadBlob(item.localBlob, item.reportName || "report.pdf")} style={{ padding: "6px 9px", borderRadius: 8, border: "1px solid rgba(96,165,250,.22)", background: "rgba(59,130,246,.1)", color: "#a9d1ff", fontFamily: "inherit", fontSize: 10, fontWeight: 800, cursor: "pointer" }}>تحميل PDF</button>}
                </div>
              )}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── StepBar ───────────────────────────────────────────────────────────────────
function StepIcon({ name }) {
  const paths = {
    layout: <><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M3 9h18M9 9v12"/></>,
    upload: <><path d="M12 16V4M8 8l4-4 4 4"/><path d="M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4"/></>,
    shield: <><path d="M12 3l7 3v5c0 4.4-3 7.7-7 10-4-2.3-7-5.6-7-10V6l7-3Z"/><path d="m9 12 2 2 4-4"/></>,
    table: <><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16M15 10v10"/></>,
    edit: <><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4L16.5 3.5Z"/></>,
    chart: <><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></>,
  };
  return <svg viewBox="0 0 24 24" aria-hidden="true" style={{ width: "100%", height: "100%", fill: "none", stroke: "currentColor", strokeWidth: 1.9, strokeLinecap: "round", strokeLinejoin: "round" }}>{paths[name] ?? paths.layout}</svg>;
}

function StepBar({ step }) {
  const visibleStep = Math.max(0, step - 1);
  return (
    <div className="step-bar" role="list" aria-label="خطوات إنشاء التقرير">
      {STEPS.map(({ label, icon }, i) => {
        const status = i === visibleStep ? "active" : i < visibleStep ? "done" : "pending";
        return (
          <div
            key={i}
            role="listitem"
            className={`step-item ${status}`}
            aria-current={status === "active" ? "step" : undefined}
          >
            <span className="step-item-icon"><StepIcon name={status === "done" ? "shield" : icon} /></span>
            <span className="step-item-num">{status === "done" ? "مكتمل" : `خطوة ${i + 1}`}</span>
            <span>{label}</span>
          </div>
        );
      })}
    </div>
  );
}

function WorkflowHeader({ step, icon, title, description }) {
  const visibleStep = Math.max(1, Number(step) - 1);
  return <div className="workflow-header">
    <span className="workflow-header-icon"><StepIcon name={icon}/></span>
    <div><div className="workflow-header-kicker">خطوة {visibleStep} من 5</div><div className="workflow-header-title">{title}</div><div className="workflow-header-copy">{description}</div></div>
  </div>;
}

// ── ReportModePicker ──────────────────────────────────────────────────────────
function ReportModePicker({ value, onChange }) {
  const MODES = [
    { id: "annual",   icon: "chart", label: "تقرير سنة واحدة", desc: "تحليل عام دراسي محدد" },
    { id: "compare2", icon: "table", label: "مقارنة سنتين",     desc: "مقارنة عامين دراسيين" },
    { id: "compare3", icon: "layout", label: "مقارنة 3 سنوات",  desc: "مقارنة ثلاثة أعوام" },
  ];
  return (
    <div className="card" style={{ padding: 36 }}>
      <WorkflowHeader step="1" icon="layout" title="نوع التقرير" description="اختر نطاق التحليل؛ يمكنك الرجوع وتغييره لاحقًا." />
      <div className="type-card-row" style={{ display: "flex", gap: 16 }}>
        {MODES.map(m => (
          <div key={m.id} className={`type-card ${value === m.id ? "selected" : ""}`} onClick={() => onChange(m.id)}>
            <div style={{ width: 34, height: 34, margin: "0 auto 10px", color: value === m.id ? "#71e8d4" : "rgba(255,255,255,.58)" }}><StepIcon name={m.icon}/></div>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, marginBottom: 6 }}>{m.label}</div>
            <div style={{ color: "rgba(255,255,255,.45)", fontSize: 12 }}>{m.desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── FileSlot ──────────────────────────────────────────────────────────────────
function FileSlot({ slot, onChange, label, driveToken, driveFiles, driveLoading, onConnectDrive }) {
  const ref = useRef();
  const [slotTab, setSlotTab] = useState("local");
  const [driveSearch, setDriveSearch] = useState("");
  const [downloading, setDownloading] = useState(false);
  const [slotDriveSelected, setSlotDriveSelected] = useState(null);

  const pickFromDrive = async (file) => {
    if (!driveToken || downloading) return;
    setDownloading(true);
    try {
      const isGSheet = file.mimeType === GSHEETS_MIME;
      const url = isGSheet
        ? `https://www.googleapis.com/drive/v3/files/${file.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
        : `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${driveToken}` } });
      if (!res.ok) throw new Error(`خطأ ${res.status}`);
      const arrayBuf = await res.arrayBuffer();
      const filename = isGSheet ? (file.name.endsWith(".xlsx") ? file.name : file.name + ".xlsx") : file.name;
      let detectedType = detectTypeHintFromFilename(filename);
      try {
        const rows = readExcel(arrayBuf);
        detectedType = detectSurveyType(filename, rows[0]) ?? detectedType;
      } catch { /* keep hint */ }
      const prog = detectProgramFromFilename(filename);
      const f = new File([arrayBuf], filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const yr = detectYearFromFilename(filename);
      onChange({ ...slot, _file: f, fileName: filename, ...(yr ? { year: yr } : {}), _type: detectedType, _program: prog });
      setSlotDriveSelected(file);
    } catch {
      // silent
    } finally {
      setDownloading(false);
    }
  };

  const filtered = (driveFiles ?? []).filter(f =>
    !driveSearch || f.name.toLowerCase().includes(driveSearch.toLowerCase())
  );

  return (
    <div style={{
      background: "rgba(255,255,255,.06)", borderRadius: 14, padding: 18,
      border: slot._file ? "1px solid #1abc9c" : "1px solid rgba(255,255,255,.1)"
    }}>
      <div style={{ marginBottom: 10 }}>
        <label className="label">{label}</label>
        <input className="input" value={slot.year}
          onChange={e => onChange({ ...slot, year: e.target.value })}
          placeholder="مثال: 2024-2025" />
      </div>

      {/* Source tabs */}
      <div style={{
        display: "flex", borderRadius: 8, overflow: "hidden",
        border: "1px solid rgba(255,255,255,.1)", background: "rgba(255,255,255,.04)", marginBottom: 10,
      }}>
        {[{ id: "local", label: "📁 جهازك" }, { id: "drive", label: "☁️ Drive" }].map(t => (
          <button key={t.id} onClick={() => setSlotTab(t.id)} style={{
            flex: 1, padding: "7px 4px", border: "none", cursor: "pointer",
            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 12,
            background: slotTab === t.id ? "rgba(26,188,156,.2)" : "transparent",
            color: slotTab === t.id ? "#1abc9c" : "rgba(255,255,255,.45)",
            borderBottom: slotTab === t.id ? "2px solid #1abc9c" : "2px solid transparent",
            transition: "all .2s",
          }}>{t.label}</button>
        ))}
      </div>

      {/* Local upload */}
      {slotTab === "local" && (
        <>
          <div onClick={() => ref.current.click()} style={{
            border: "1.5px dashed rgba(255,255,255,.25)", borderRadius: 10, padding: "16px 10px",
            textAlign: "center", cursor: "pointer", background: "rgba(255,255,255,.03)",
            borderColor: slot._file ? "#1abc9c" : "rgba(255,255,255,.25)",
          }}>
            <div style={{ fontSize: 28, marginBottom: 6 }}>{slot._file ? "✅" : "📂"}</div>
            <div style={{ color: slot._file ? "#1abc9c" : "rgba(255,255,255,.6)", fontSize: 12, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {slot.fileName || "اضغط لاختيار ملف"}
            </div>
            {slot._file && (
              <div style={{ marginTop: 6 }}>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap", justifyContent: "center", marginBottom: slot._file ? 5 : 0 }}>
                  {slot._type && allSchemas()[slot._type] ? (
                    <span style={{ fontSize: 10, color: "#1abc9c", background: "rgba(26,188,156,.15)", borderRadius: 6, padding: "2px 7px", fontWeight: 700 }}>
                      {allSchemas()[slot._type].icon} {allSchemas()[slot._type].label}
                    </span>
                  ) : (
                    <span style={{ fontSize: 10, color: "#ffd54f", background: "rgba(255,193,7,.12)", borderRadius: 6, padding: "2px 7px", fontWeight: 700 }}>
                      ⚠ نوع غير محدد
                    </span>
                  )}
                  {slot._program && (
                    <span style={{ fontSize: 10, color: "#d6eaf8", background: "rgba(255,255,255,.08)", borderRadius: 6, padding: "2px 7px" }}>
                      {slot._program}
                    </span>
                  )}
                </div>
              </div>
            )}
          </div>
          <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={async e => {
              const f = e.target.files[0];
              if (!f) return;
              const yr = detectYearFromFilename(f.name);
              const prog = detectProgramFromFilename(f.name);
              let detectedType = detectTypeHintFromFilename(f.name);
              try {
                const buf = await f.arrayBuffer();
                const rows = readExcel(buf);
                detectedType = detectSurveyType(f.name, rows[0]) ?? detectedType;
              } catch { /* keep hint */ }
              onChange({ ...slot, _file: f, fileName: f.name, ...(yr ? { year: yr } : {}), _program: prog, _type: detectedType });
            }} />
          {slot._file && (
            <div style={{ marginTop: 8 }}>
              <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10, marginBottom: 4, textAlign: "center" }}>
                النوع غير صحيح؟ اختر:
              </div>
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "center" }}>
                {Object.values(allSchemas()).map(sc => (
                  <button key={sc.id} onClick={e => { e.stopPropagation(); onChange({ ...slot, _type: sc.id }); }}
                    style={{
                      padding: "3px 9px", borderRadius: 8, cursor: "pointer",
                      fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 10,
                      border: `1.5px solid ${slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                      background: slot._type === sc.id ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                      color: slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.5)",
                      transition: "all .15s",
                    }}>
                    {sc.icon} {sc.label}
                  </button>
                ))}
              </div>
            </div>
          )}
        </>
      )}

      {/* Drive picker */}
      {slotTab === "drive" && (
        !driveToken ? (
          <div style={{ textAlign: "center", padding: "14px 0" }}>
            <button onClick={onConnectDrive} style={{
              display: "inline-flex", alignItems: "center", gap: 8,
              padding: "10px 20px", borderRadius: 50, border: "1px solid #dadce0",
              background: "#fff", color: "#3c4043", cursor: "pointer",
              fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
            }}>
              <svg width="16" height="16" viewBox="0 0 48 48">
                <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.31-8.16 2.31-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
              </svg>
              ربط Google Drive
            </button>
          </div>
        ) : driveLoading ? (
          <div style={{ textAlign: "center", padding: "14px 0", color: "rgba(255,255,255,.5)", fontSize: 12 }}>
            جاري تحميل الملفات…
          </div>
        ) : (
          <>
            <input value={driveSearch} onChange={e => setDriveSearch(e.target.value)}
              placeholder="🔍 ابحث…"
              style={{
                width: "100%", background: "#0d1f33", color: "#e8f0fe",
                border: "1px solid rgba(255,255,255,.18)", borderRadius: 8,
                padding: "7px 10px", fontFamily: "'Cairo',sans-serif",
                fontSize: 12, direction: "rtl", marginBottom: 6, outline: "none",
              }} />
            <div style={{ maxHeight: 160, overflowY: "auto", display: "flex", flexDirection: "column", gap: 4 }}>
              {filtered.length === 0 ? (
                <div style={{ color: "rgba(255,255,255,.3)", fontSize: 12, textAlign: "center", padding: 12 }}>لا توجد ملفات</div>
              ) : filtered.map(f => {
                const isSel = slotDriveSelected?.id === f.id;
                return (
                  <div key={f.id} onClick={() => pickFromDrive(f)} style={{
                    display: "flex", alignItems: "center", gap: 8,
                    background: isSel ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                    border: `1px solid ${isSel ? "rgba(26,188,156,.5)" : "rgba(255,255,255,.08)"}`,
                    borderRadius: 8, padding: "7px 10px",
                    cursor: downloading ? "not-allowed" : "pointer", transition: "all .15s",
                  }}>
                    <span style={{ fontSize: 14 }}>{f.mimeType === GSHEETS_MIME ? "📊" : "📄"}</span>
                    <div style={{ flex: 1, overflow: "hidden" }}>
                      <div style={{ color: "#e8f0fe", fontSize: 11, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</div>
                    </div>
                    {isSel && <span style={{ color: "#1abc9c", fontSize: 14 }}>✅</span>}
                    {downloading && isSel && (
                      <svg className="spin" width="12" height="12" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="#1abc9c" strokeWidth="2" strokeDasharray="22 8"/></svg>
                    )}
                  </div>
                );
              })}
            </div>
            {slot._file && (
              <div style={{ marginTop: 8 }}>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 10, marginBottom: 4, textAlign: "center" }}>
                  النوع غير صحيح؟ اختر:
                </div>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "center" }}>
                  {Object.values(allSchemas()).map(sc => (
                    <button key={sc.id} onClick={() => onChange({ ...slot, _type: sc.id })}
                      style={{
                        padding: "3px 9px", borderRadius: 8, cursor: "pointer",
                        fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 10,
                        border: `1.5px solid ${slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                        background: slot._type === sc.id ? "rgba(26,188,156,.18)" : "rgba(255,255,255,.04)",
                        color: slot._type === sc.id ? "#1abc9c" : "rgba(255,255,255,.5)",
                        transition: "all .15s",
                      }}>
                      {sc.icon} {sc.label}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </>
        )
      )}
    </div>
  );
}

// ── MetadataForm ──────────────────────────────────────────────────────────────
function MetadataForm({ meta, onChange }) {
  const F = (key, label, placeholder) => (
    <div>
      <label className="label">{label}</label>
      <input className="input" value={meta[key] ?? ""} placeholder={placeholder}
        onChange={e => onChange({ ...meta, [key]: e.target.value })} />
    </div>
  );
  return (
    <div className="card" style={{ padding: 32 }}>
      <WorkflowHeader step="5" icon="edit" title="بيانات التقرير" description="أكمل البيانات التي ستظهر في غلاف التقرير والتوقيعات." />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 16 }}>
        <div>
          <label className="label">اسم البرنامج / القسم</label>
          <ProgramPicker value={meta.program ?? ""}
            onChange={v => onChange({ ...meta, program: v })} />
        </div>
        {F("year",       "العام الأكاديمي", "مثال: 2024-2025")}
        {F("preparedBy", "أعده *",          "اسم معد التقرير")}
        {F("reviewer",   "راجعه *",         "اسم المراجع")}
      </div>
    </div>
  );
}

// ── ProgramPicker ─────────────────────────────────────────────────────────────
function ProgramPicker({ value, onChange }) {
  const isCustom = !!value && !PROGRAMS.includes(value);
  const [mode, setMode] = useState(isCustom ? "other" : "list");

  useEffect(() => {
    // Re-sync if parent value changes (e.g. auto-detect from filename)
    if (value && PROGRAMS.includes(value) && mode !== "list") setMode("list");
    if (value && !PROGRAMS.includes(value) && mode !== "other") setMode("other");
  }, [value]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div>
      <select className="input"
        value={mode === "other" ? "__other__" : value}
        onChange={e => {
          if (e.target.value === "__other__") { setMode("other"); onChange(""); }
          else { setMode("list"); onChange(e.target.value); }
        }}>
        <option value="">— اختر البرنامج —</option>
        {PROGRAMS.map(p => <option key={p} value={p}>{p}</option>)}
        <option value="__other__">أخرى…</option>
      </select>
      {mode === "other" && (
        <input className="input" type="text" value={value}
          placeholder="اكتب اسم البرنامج"
          onChange={e => onChange(e.target.value)}
          style={{ marginTop: 8 }} />
      )}
    </div>
  );
}

// ── FilterDropdown ────────────────────────────────────────────────────────────
function FilterDropdown({ label, value, options, onChange }) {
  if (!options.length) return null;
  return (
    <label style={{ color: "rgba(255,255,255,.8)", fontSize: 13, display: "flex",
                    alignItems: "center", gap: 8 }}>
      {label}:
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ background: "#0f2035", color: "#e8f0fe",
                 border: "1px solid rgba(255,255,255,.25)", borderRadius: 8,
                 padding: "7px 32px 7px 12px", fontFamily: "inherit", fontSize: 13,
                 cursor: "pointer", outline: "none",
                 appearance: "none", WebkitAppearance: "none",
                 backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='7' viewBox='0 0 10 7'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%231abc9c' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E\")",
                 backgroundRepeat: "no-repeat", backgroundPosition: "left 10px center" }}>
        <option value="" style={{ background: "#0f2035", color: "#e8f0fe" }}>— الكل —</option>
        {options.map(o => <option key={o} value={o} style={{ background: "#0f2035", color: "#e8f0fe" }}>{o}</option>)}
      </select>
    </label>
  );
}

// ── DataPreviewTable ──────────────────────────────────────────────────────────
function DataPreviewTable({ headers, rows, removed, onToggleRow,
                            deptIdx, degreeIdx, filters, onFilterChange }) {
  const deptOptions  = useMemo(() => uniqueCol(rows, deptIdx),   [rows, deptIdx]);
  const degreeOpts   = useMemo(() => uniqueCol(rows, degreeIdx), [rows, degreeIdx]);

  const visibleRows = rows.map((r, i) => ({ row: r, idx: i }))
    .filter(({ idx })  => !removed.has(idx))
    .filter(({ row }) => !filters.dept   || String(row[deptIdx])   === filters.dept)
    .filter(({ row }) => !filters.degree || String(row[degreeIdx]) === filters.degree);

  const cellTh = { padding: "8px 10px", fontSize: 11, fontWeight: 700, textAlign: "right",
                   borderBottom: "1px solid rgba(255,255,255,.15)", whiteSpace: "nowrap" };
  const cellTd = { padding: "6px 10px", textAlign: "right", color: "rgba(255,255,255,.85)",
                   verticalAlign: "top" };

  return (
    <div>
      <div style={{ display: "flex", gap: 12, marginBottom: 12, flexWrap: "wrap",
                    alignItems: "center" }}>
        <FilterDropdown label="القسم" value={filters.dept}
          options={deptOptions}
          onChange={v => onFilterChange({ ...filters, dept: v })} />
        <FilterDropdown label="الدرجة / المسمى الوظيفي" value={filters.degree}
          options={degreeOpts}
          onChange={v => onFilterChange({ ...filters, degree: v })} />
        <button className="btn btn-ghost btn-sm"
          onClick={() => onFilterChange({ dept: "", degree: "" })}>
          ↩ إزالة الفلاتر
        </button>
      </div>

      <div style={{ maxHeight: 500, overflow: "auto",
                    border: "1px solid rgba(255,255,255,.15)", borderRadius: 8,
                    background: "rgba(0,0,0,.2)" }}>
        <table style={{ width: "100%", fontSize: 12, color: "#fff",
                        borderCollapse: "collapse", direction: "rtl" }}>
          <thead style={{ position: "sticky", top: 0, background: "#1a3a5c", zIndex: 1 }}>
            <tr>
              <th style={{ ...cellTh, textAlign: "center", width: 60 }}>إجراء</th>
              {headers.map((h, i) => <th key={i} style={cellTh}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map(({ row, idx }) => (
              <tr key={idx} style={{ borderTop: "1px solid rgba(255,255,255,.05)" }}>
                <td style={{ ...cellTd, textAlign: "center" }}>
                  <button className="btn btn-danger btn-sm" title="حذف هذا الصف"
                    style={{ padding: "4px 10px", fontSize: 14 }}
                    onClick={() => onToggleRow(idx)}>🗑️</button>
                </td>
                {headers.map((_, ci) => (
                  <td key={ci} style={cellTd}>
                    {row[ci] == null ? "" : String(row[ci])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ color: "rgba(255,255,255,.5)", fontSize: 12, marginTop: 8,
                    textAlign: "center" }}>
        عدد الصفوف المعروضة: {visibleRows.length} من إجمالي {rows.length}
        {" "}(تم حذف {removed.size})
      </div>
    </div>
  );
}

// ── CollapsibleSection ────────────────────────────────────────────────────────
function CollapsibleSection({ title, defaultOpen = true, children, badge }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 18 }}>
      <button
        onClick={() => setOpen(o => !o)}
        aria-expanded={open}
        style={{
          width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "16px 22px", background: "transparent", border: "none", cursor: "pointer",
          fontFamily: "'Cairo',sans-serif",
        }}
      >
        <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ color: "#fff", fontSize: 16, fontWeight: 900 }}>{title}</span>
          {badge}
        </span>
        <span style={{
          color: "#1abc9c", fontSize: 15, transition: "transform .2s",
          display: "inline-block", transform: open ? "rotate(180deg)" : "rotate(0deg)",
        }} aria-hidden="true">▾</span>
      </button>
      {open && <div style={{ padding: "0 22px 22px" }}>{children}</div>}
    </div>
  );
}

// ── QuickStatsRow (compact live-analysis KPIs, used in Step 3 while trimming rows) ──
function QuickStatsRow({ result }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 14 }}>
      {[
        { n: result.n,                     l: "عدد المستجيبين", icon: "👥" },
        { n: result.axes.length,           l: "عدد المحاور",    icon: "📋" },
        { n: result.totalQuestions,        l: "عدد الأسئلة",    icon: "❓" },
        { n: `${result.overallAgreePct}%`, l: "نسبة الموافقة",  icon: "✅" },
      ].map((s, i) => (
        <div key={i} className="stat-card">
          <div style={{ fontSize: 26, marginBottom: 6 }}>{s.icon}</div>
          <div style={{ color: "#1abc9c", fontSize: 28, fontWeight: 900 }}>{s.n}</div>
          <div style={{ color: "rgba(255,255,255,.55)", fontSize: 11, marginTop: 4 }}>{s.l}</div>
        </div>
      ))}
    </div>
  );
}

// ── ResultsPreview ────────────────────────────────────────────────────────────
function ResultsPreview({ result }) {
  const is5 = result.scaleType === "likert-5";
  return (
    <div>
      <div style={{ marginBottom: 24 }}>
        <QuickStatsRow result={result} />
      </div>

      {Object.keys(result.byDegree).length > 0 && (
        <div className="card" style={{ padding: 22, marginBottom: 18 }}>
          <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 12, borderBottom: "1px solid rgba(255,255,255,.1)", paddingBottom: 8 }}>
            توزيع المشاركين
          </div>
          <div className="table-scroll">
            <table className="mini-table">
              <thead><tr><th>العدد</th><th>النسبة</th><th style={{ textAlign: "right" }}>الدرجة / الوظيفة</th></tr></thead>
              <tbody>
                {Object.entries(result.byDegree).sort((a, b) => b[1] - a[1]).map(([k, v], i) => (
                  <tr key={i}>
                    <td>{v}</td>
                    <td style={{ color: "#1abc9c", fontWeight: 700 }}>{(v / result.n * 100).toFixed(1)}%</td>
                    <td className="rtl-td" style={{ color: "#e8f0fe" }}>{k}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="card" style={{ padding: 22 }}>
        <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 12, borderBottom: "1px solid rgba(255,255,255,.1)", paddingBottom: 8 }}>
          ملخص المحاور
        </div>
        <div className="table-scroll">
          <table className="mini-table">
            <thead>
              <tr>
                <th style={{ width: 36 }}>م</th>
                <th style={{ textAlign: "right" }}>المحور</th>
                {is5 && <th>المتوسط</th>}
                <th>نسبة الموافقة</th>
                <th>الاتجاه</th>
              </tr>
            </thead>
            <tbody>
              {result.axes.map((ax, i) => (
                <tr key={i}>
                  <td style={{ color: "#1abc9c", fontWeight: 700 }}>{i + 1}</td>
                  <td className="rtl-td" style={{ fontSize: 12 }}>{ax.name}</td>
                  {is5 && <td style={{ fontWeight: 700, color: "#d6eaf8" }}>{ax.axisMean}</td>}
                  <td style={{ fontWeight: 700, color: "#1abc9c" }}>{ax.axisAgreePct}%</td>
                  <td>
                    <span className="badge" style={{ background: `${dirColor(ax.tier)}22`, color: dirColor(ax.tier) }}>
                      {ax.direction}
                    </span>
                  </td>
                </tr>
              ))}
              <tr>
                <td colSpan={is5 ? 2 : 2} style={{ textAlign: "right", fontWeight: 900, color: "#fff", background: "rgba(26,188,156,.18)", padding: "10px" }}>
                  الإجمالي
                </td>
                {is5 && <td style={{ fontWeight: 900, color: "#1abc9c", background: "rgba(26,188,156,.18)" }}>{result.overallMean}</td>}
                <td style={{ fontWeight: 900, color: "#1abc9c", fontSize: 15, background: "rgba(26,188,156,.18)" }}>
                  {result.overallAgreePct}%
                </td>
                <td style={{ background: "rgba(26,188,156,.18)" }}>
                  <span className="badge" style={{ background: "rgba(13,110,58,.35)", color: "#1abc9c" }}>
                    {result.overallDirection}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── ComparisonPreview ─────────────────────────────────────────────────────────
function ComparisonPreview({ comparison }) {
  const { axes, overall } = comparison;
  const years = overall.map(o => o.year);
  return (
    <div className="card" style={{ padding: 24 }}>
      <div style={{ color: "#fff", fontSize: 16, fontWeight: 700, marginBottom: 16 }}>جدول المقارنة</div>
      <div style={{ overflowX: "auto" }}>
        <table className="mini-table">
          <thead>
            <tr>
              <th style={{ width: 40 }}>م</th>
              <th style={{ textAlign: "right" }}>المحور</th>
              {years.map((y, i) => <th key={i}>{y}</th>)}
              <th>الاتجاه</th>
            </tr>
          </thead>
          <tbody>
            {axes.map((ax, ai) => (
              <tr key={ai}>
                <td style={{ color: "#1abc9c", fontWeight: 700 }}>{ai + 1}</td>
                <td className="rtl-td" style={{ fontSize: 12 }}>{ax.name}</td>
                {ax.years.map((y, yi) => (
                  <td key={yi} style={{ fontWeight: 700, color: "#d6eaf8" }}>
                    {y.agreePct !== null ? `${y.agreePct}%` : "—"}
                  </td>
                ))}
                <td>
                  <span className="badge" style={{
                    background: ax.trend === "تحسن" ? "rgba(13,110,58,.3)" : ax.trend === "تراجع" ? "rgba(146,43,33,.3)" : "rgba(120,66,18,.2)",
                    color: ax.trend === "تحسن" ? "#1abc9c" : ax.trend === "تراجع" ? "#e74c3c" : "#f0b27a"
                  }}>{ax.trend}</span>
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={2} style={{ textAlign: "right", fontWeight: 900, color: "#fff", background: "rgba(26,188,156,.15)", padding: "9px 10px" }}>الإجمالي</td>
              {overall.map((o, i) => (
                <td key={i} style={{ fontWeight: 900, color: "#1abc9c", fontSize: 14, background: "rgba(26,188,156,.15)" }}>
                  {o.agreePct}%
                </td>
              ))}
              <td style={{ background: "rgba(26,188,156,.15)" }}></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

const LOCAL_STEPS = ["قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "حساب المؤشرات", "تجهيز العرض"];
const DRIVE_STEPS = ["تحميل الملف من Drive", "قراءة البيانات", "تنظيف البيانات", "كشف نوع الاستبيان", "تجهيز العرض"];

const HUB_ICONS = {
  home: <><path d="M3 11 12 3l9 8"/><path d="M5 10v10h14V10M9 20v-6h6v6"/></>,
  analytics: <><path d="M4 19V9m6 10V5m6 14v-7m4 7H2"/><path d="m4 7 6-4 6 5 4-3"/></>,
  surveys: <><rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 3.5h6M9 8h6M9 12h6M9 16h4"/></>,
  semester: <><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01"/></>,
  dashboard: <><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></>,
  courses: <><path d="m3 6 9-4 9 4-9 4-9-4Z"/><path d="M7 8.2v5.3c0 1.7 2.2 3 5 3s5-1.3 5-3V8.2M21 6v7"/></>,
  settings: <><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .34 1.88l.06.06-2.83 2.83-.06-.06A1.7 1.7 0 0 0 15 19.4a1.7 1.7 0 0 0-1 .6 1.7 1.7 0 0 0-.4 1.1V21h-4v-.1A1.7 1.7 0 0 0 8.6 19.4a1.7 1.7 0 0 0-1.88.34l-.06.06-2.83-2.83.06-.06A1.7 1.7 0 0 0 4.6 15a1.7 1.7 0 0 0-.6-1 1.7 1.7 0 0 0-1.1-.4H3v-4h.1A1.7 1.7 0 0 0 4.6 8.6a1.7 1.7 0 0 0-.34-1.88l-.06-.06 2.83-2.83.06.06A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-.6 1.7 1.7 0 0 0 .4-1.1V3h4v.1A1.7 1.7 0 0 0 15.4 4a1.7 1.7 0 0 0 1.88-.34l.06-.06 2.83 2.83-.06.06A1.7 1.7 0 0 0 19.4 9c.17.36.5.75 1 .97.35.16.73.24 1.1.23h.1v4h-.1A1.7 1.7 0 0 0 19.4 15Z"/></>,
  help: <><circle cx="12" cy="12" r="9"/><path d="M9.8 9a2.3 2.3 0 1 1 3.4 2c-.8.45-1.2.9-1.2 2M12 17h.01"/></>,
};

function HubIcon({ name }) {
  return <svg viewBox="0 0 24 24" aria-hidden="true">{HUB_ICONS[name]}</svg>;
}

function QualityHub({ onOpen }) {
  const toolCards = [
    {
      id: "analytics", icon: "analytics", color: "#5eead4",
      title: "تحليل نتائج موجودة", tag: "Excel أو Google Drive",
      description: "ابدأ بملف واحد أو مجموعة ملفات، ودع النظام يتعرّف على نوع الاستبيان ثم أنشئ تقرير Word أو PDF.",
    },
    {
      id: "semester", icon: "semester", color: "#60a5fa",
      title: "إنشاء استبيانات الفصل الدراسي", tag: "Google Forms",
      description: "اختر القوالب والسنة والفصل، وأنشئ كل استبيانات الفصل منظمةً تلقائياً على Google Drive.",
    },
    {
      id: "courses", icon: "courses", color: "#a78bfa",
      title: "تقييم المقررات", tag: "دورة تقييم متكاملة",
      description: "جهّز بيانات المقررات، تابع نسب المشاركة، قسّم ملفات التقييم وراجع التوصيات من مساحة واحدة.",
    },
  ];

  return (
    <section className="hub" aria-label="الصفحة الرئيسية لوحدة ضمان الجودة">
      <div className="hub-hero">
        <div style={{ position: "relative", zIndex: 1 }}>
          <div className="hub-eyebrow"><span>◆</span> ابدأ من المهمة، لا من الأداة</div>
          <h1 className="hub-title">ماذا تريد أن تنجز اليوم؟</h1>
          <p className="hub-subtitle">
            اختر مهمة واحدة وسنقودك إلى الخطوة التالية. يمكنك تحليل نتائج موجودة، إنشاء
            استبيانات جديدة، أو إدارة دورة تقييم المقررات.
          </p>
        </div>
      </div>

      <div className="hub-section-head">
        <div><h2>أدوات وحدة الجودة</h2><p>اختر المهمة التي تريد إنجازها الآن</p></div>
        <span className="hub-eyebrow">{toolCards.length} مهام رئيسية</span>
      </div>
      <div className="tool-grid">
        {toolCards.map(tool => (
          <button key={tool.id} className={`tool-card ${tool.featured ? "featured" : ""}`}
            style={{ "--tool-color": tool.color }} onClick={() => onOpen(tool.id)}>
            <span className="tool-card-icon"><HubIcon name={tool.icon} /></span>
            <h3>{tool.title}</h3>
            <p>{tool.description}</p>
            <span className="tool-card-foot"><span className="tool-tag">{tool.tag}</span><span className="tool-arrow">←</span></span>
          </button>
        ))}
      </div>

      <div className="hub-note">
        <div className="hub-note-copy"><span className="hub-note-icon">✦</span><span><b style={{ color: "#fff" }}>مسار عمل مقترح:</b> أنشئ الاستبيان، اجمع الاستجابات، ثم استخدم أداة التحليل لإصدار التقرير النهائي.</span></div>
        <button className="btn btn-ghost btn-sm" onClick={() => onOpen("tutorial")}>عرض دليل الاستخدام</button>
      </div>
    </section>
  );
}

const SIDEBAR_ITEMS = [
  { id: "home", icon: "home", label: "الرئيسية" },
  { id: "analytics", icon: "analytics", label: "تحليل الاستبيانات" },
  { id: "semester", icon: "semester", label: "استبيانات الفصل الدراسي" },
  { id: "courses", icon: "courses", label: "تقييم المقررات" },
];

function ToolSidebar({ open, active, onNavigate, onTutorial, onToggle }) {
  return (
    <aside className={`tool-sidebar ${open ? "open" : "collapsed"}`} aria-label="التنقل بين أدوات الجودة">
      <div className="sidebar-brand">
        <span className="sidebar-brand-mark"><HubIcon name="home" /></span>
        <span className="sidebar-brand-copy">بوابة وحدة ضمان الجودة<small>ERU · QUALITY HUB</small></span>
        <button type="button" className="sidebar-toggle" onClick={onToggle} aria-label={open ? "طي القائمة الجانبية" : "فتح القائمة الجانبية"} title={open ? "طي القائمة" : "فتح القائمة"}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round"><path d="M4 6h16M4 12h16M4 18h16"/></svg>
        </button>
      </div>
      <div className="sidebar-heading">مساحات العمل</div>
      <nav className="sidebar-nav">
        {SIDEBAR_ITEMS.map(item => (
          <button key={item.id} type="button" className={`sidebar-item ${active === item.id ? "active" : ""}`} onClick={() => onNavigate(item.id)} title={!open ? item.label : undefined}>
            <span className="sidebar-icon"><HubIcon name={item.icon} /></span>
            <span className="sidebar-label">{item.label}</span>
          </button>
        ))}
      </nav>
      <div className="sidebar-separator" />
      <div className="sidebar-heading">أدوات إضافية</div>
      <nav className="sidebar-nav">
        <button type="button" className={`sidebar-item ${active === "surveys" ? "active" : ""}`} onClick={() => onNavigate("surveys")} title={!open ? "تصميم الاستبيانات" : undefined}>
          <span className="sidebar-icon"><HubIcon name="surveys" /></span><span className="sidebar-label">تصميم الاستبيانات</span>
        </button>
      </nav>
      <div className="sidebar-bottom sidebar-nav">
        <button type="button" className="sidebar-item" onClick={onTutorial} title={!open ? "دليل الاستخدام" : undefined}>
          <span className="sidebar-icon"><HubIcon name="help" /></span><span className="sidebar-label">دليل الاستخدام</span>
        </button>
        <button type="button" className={`sidebar-item ${active === "settings" ? "active" : ""}`} onClick={() => onNavigate("settings")} title={!open ? "الإعدادات" : undefined}>
          <span className="sidebar-icon"><HubIcon name="settings" /></span><span className="sidebar-label">الإعدادات</span>
        </button>
      </div>
    </aside>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep]         = useState(1);
  const [surveyType, setSurveyType] = useState("faculty");
  const [mode, setMode]         = useState("annual");
  const [meta, setMeta]         = useState({ year: "2024-2025", program: "", preparedBy: "", reviewer: "" });
  const [processing, setProcessing] = useState(false);
  const [pdfGenerating, setPdfGenerating] = useState(false);
  const [pdfCancelRequested, setPdfCancelRequested] = useState(false);
  const [pdfProgress, setPdfProgress] = useState({ current: 0, total: 0 });
  const [error, setError]       = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const [showSurveyManagement, setShowSurveyManagement] = useState(false);
  const [showSemesterSurveys, setShowSemesterSurveys] = useState(true);
  const [showCourseEval, setShowCourseEval] = useState(false);
  const [showHub, setShowHub] = useState(false);
  const [showTutorial, setShowTutorial] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => typeof window === "undefined" || window.innerWidth > 760);
  const [semesterInitialTab, setSemesterInitialTab] = useState("departments");
  const [settings, setSettings]         = useState(loadSettings);
  const [aiSettings, setAiSettings]     = useState(loadAiSettings);
  const [pdfIncludeCharts, setPdfIncludeCharts] = useState(() => loadSettings().includePdfCharts !== false);

  const [singleFile, setSingleFile]     = useState(null);
  const [singleResult, setSingleResult] = useState(null);
  const [showEnhancedView, setShowEnhancedView] = useState(false);
  const singleRef = useRef();
  const [dragging, setDragging]         = useState(false);

  // File-detection state (step 0 → 1)
  const [detectedAutoType, setDetectedAutoType] = useState(null);
  const [rawRows, setRawRows]                   = useState(null);

  // Processing animation overlay
  const [procSteps, setProcSteps] = useState(null);
  const [procFile,  setProcFile]  = useState("");

  const runStepAnim = useCallback((labels, startAt = 0, msPerStep = 480, onDone) => {
    const build = (cur) => labels.map((label, i) => ({
      label,
      status: i < cur ? "done" : i === cur ? "active" : "pending",
    }));
    setProcSteps(build(startAt));
    let cur = startAt;
    const tick = () => {
      cur++;
      if (cur >= labels.length) {
        setProcSteps(labels.map(label => ({ label, status: "done" })));
        setTimeout(() => { setProcSteps(null); onDone(); }, 380);
        return;
      }
      setProcSteps(build(cur));
      setTimeout(tick, msPerStep);
    };
    setTimeout(tick, msPerStep);
  }, []);

  // Batch mode (multiple files selected in step 0)
  const [batchMode,  setBatchMode]  = useState(false);
  const [batchFiles, setBatchFiles] = useState([]);
  const [batchYear,  setBatchYear]  = useState("2024-2025");

  // Upload source tab
  const [uploadTab,         setUploadTab]         = useState("upload");
  const [driveToken,        setDriveToken]        = useState(() => getStoredToken(DRIVE_TOKEN_KEY));
  const [driveConnecting,   setDriveConnecting]   = useState(false);
  const [driveConnection,   setDriveConnection]   = useState(() => getStoredToken(DRIVE_TOKEN_KEY) ? "checking" : "disconnected");
  const [driveAccount,      setDriveAccount]      = useState("");
  const [driveFiles,        setDriveFiles]        = useState([]);
  const [driveSearch,       setDriveSearch]       = useState("");
  const [driveLoading,      setDriveLoading]      = useState(false);
  const [driveSelected,     setDriveSelected]     = useState(null);
  const [driveSelectedIds,  setDriveSelectedIds]  = useState(() => new Set());
  const [driveProcessing,   setDriveProcessing]   = useState(false);
  const [driveBatchState,   setDriveBatchState]   = useState(null);
  const driveBatchCancelRef = useRef(false);
  const pdfCancelRef = useRef(false);
  const [driveFilterType,   setDriveFilterType]   = useState("");
  const [driveFilterYear,   setDriveFilterYear]   = useState("");
  const [driveFilterProgram,setDriveFilterProgram]= useState("");
  const [driveViewMode,     setDriveViewMode]     = useState("list"); // "list" | "dashboard" | "semester"
  const tokenClientRef = useRef(null);
  const explicitDriveRef = useRef(false);
  const triedSilentDriveRef = useRef(false);

  // Annual-mode preview state (Step 3)
  const [singleHeaders, setSingleHeaders]     = useState(null);
  const [singleAllRows, setSingleAllRows]     = useState(null);
  const [singleSchema, setSingleSchema]       = useState(null);
  const [singleMetaCols, setSingleMetaCols]   = useState(null);
  const [singleRemoved, setSingleRemoved]     = useState(new Set());
  const [singleFilters, setSingleFilters]     = useState({ dept: "", degree: "" });

  const slotCount = mode === "compare2" ? 2 : 3;
  const defaultSlots = () => [
    { year: "2022-2023", fileName: "", result: null, _file: null, _type: null, _program: null },
    { year: "2023-2024", fileName: "", result: null, _file: null, _type: null, _program: null },
    { year: "2024-2025", fileName: "", result: null, _file: null, _type: null, _program: null },
  ];
  const [slots, setSlots]       = useState(defaultSlots());
  const [comparison, setComparison] = useState(null);

  const schema = allSchemas()[surveyType];

  const processBuffer = useCallback((buf, filename) => {
    const rows = readExcel(buf);
    const type = detectSurveyType(filename, rows[0]) ?? surveyType;
    const s = allSchemas()[type] ?? schema;
    return analyze(rows, s);
  }, [surveyType, schema]);

  const fetchDriveFiles = useCallback(async (token) => {
    setDriveLoading(true);
    try {
      const q = DRIVE_FILE_MIMES.map(m => `mimeType='${m}'`).join(" or ");
      const params = new URLSearchParams({
        q: `(${q}) and trashed=false`,
        fields: "files(id,name,mimeType,modifiedTime,parents,webViewLink)",
        orderBy: "modifiedTime desc",
        pageSize: "1000",
      });
      const res = await fetch(`https://www.googleapis.com/drive/v3/files?${params}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      setDriveFiles(data.files || []);
    } catch (e) {
      setError("خطأ في تحميل ملفات Drive: " + e.message);
    } finally {
      setDriveLoading(false);
    }
  }, []);

  const ensureDriveTokenClient = useCallback(async () => {
    await loadGisScript();
    if (!tokenClientRef.current) {
      tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: DRIVE_SCOPE,
        prompt: "select_account",
        callback: async (resp) => {
          if (resp.error) {
            // A silent background attempt failing is normal (no prior session/consent) —
            // only surface an error banner for an explicit, user-clicked connect.
            if (explicitDriveRef.current) setError("فشل الاتصال بـ Google Drive: " + resp.error);
            setDriveConnecting(false);
            setDriveConnection("disconnected");
            return;
          }
          setDriveToken(resp.access_token);
          setDriveConnecting(false);
          setDriveConnection("connected");
          saveStoredToken(DRIVE_TOKEN_KEY, resp.access_token, resp.expires_in);
          await fetchDriveFiles(resp.access_token);
        },
      });
    }
    return tokenClientRef.current;
  }, [fetchDriveFiles]);

  const connectDrive = useCallback(async () => {
    if (!GOOGLE_CLIENT_ID) {
      setError("تعذر الاتصال: VITE_GOOGLE_CLIENT_ID غير مضبوط في إعدادات المشروع.");
      setDriveConnection("disconnected");
      return;
    }
    setDriveConnecting(true);
    setDriveConnection("connecting");
    try {
      const client = await ensureDriveTokenClient();
      explicitDriveRef.current = true;
      client.requestAccessToken({ prompt: "select_account" });
    } catch (e) {
      setDriveConnecting(false);
      setDriveConnection("disconnected");
      setError("فشل الاتصال بـ Google Drive: " + e.message);
    }
  }, [ensureDriveTokenClient]);

  const disconnectDrive = useCallback(() => {
    setDriveToken(null);
    setDriveConnection("disconnected");
    setDriveAccount("");
    clearStoredToken(DRIVE_TOKEN_KEY);
    setDriveFiles([]);
    setDriveSelected(null);
    setDriveSelectedIds(new Set());
    setDriveViewMode("list");
  }, []);

  useEffect(() => {
    if (!driveToken) { setDriveConnection("disconnected"); setDriveAccount(""); return; }
    let cancelled = false;
    setDriveConnection("checking");
    fetch("https://www.googleapis.com/drive/v3/about?fields=user(displayName,emailAddress)", {
      headers: { Authorization: `Bearer ${driveToken}` },
    }).then(async res => {
      if (!res.ok) throw new Error(String(res.status));
      const data = await res.json();
      if (!cancelled) {
        setDriveAccount(data.user?.emailAddress || data.user?.displayName || "");
        setDriveConnection("connected");
      }
    }).catch(() => {
      if (cancelled) return;
      setDriveToken(null);
      clearStoredToken(DRIVE_TOKEN_KEY);
      setDriveFiles([]);
      setDriveConnection("disconnected");
      setDriveAccount("");
    });
    return () => { cancelled = true; };
  }, [driveToken]);

  // Token persisted from a previous visit? Great, no Google call needed at all. Otherwise,
  // once on mount, try a silent refresh (no popup) before ever showing the connect button.
  useEffect(() => {
    if (!GOOGLE_CLIENT_ID || driveToken || triedSilentDriveRef.current) return;
    triedSilentDriveRef.current = true;
    (async () => {
      try {
        const client = await ensureDriveTokenClient();
        explicitDriveRef.current = false;
        client.requestAccessToken({ prompt: "" });
      } catch { /* ignore — falls through to the manual connect button */ }
    })();
  }, [driveToken, ensureDriveTokenClient]);

  // A token restored from localStorage skips connectDrive() entirely (and the file-list
  // fetch that normally happens in its callback) — fetch once the Drive tab is actually open.
  useEffect(() => {
    if (uploadTab === "drive" && driveToken && driveFiles.length === 0 && !driveLoading) {
      fetchDriveFiles(driveToken);
    }
  }, [uploadTab, driveToken, driveFiles.length, driveLoading, fetchDriveFiles]);

  const handleDriveFileSelect = useCallback(async (file) => {
    if (!driveToken || !file) return;
    setDriveSelected(file);
    setDriveProcessing(true);
    setError("");
    setProcFile(file.name);
    // Show step 0 (download) as active immediately
    setProcSteps(DRIVE_STEPS.map((label, i) => ({ label, status: i === 0 ? "active" : "pending" })));
    try {
      const { buffer: buf, filename } = await downloadDriveBuffer(file, driveToken);

      setSingleFile({ name: filename });
      const rows = readExcel(buf);
      setRawRows(rows);
      const detected = detectSurveyType(filename, rows[0]) ?? null;
      setDetectedAutoType(detected);
      setSurveyType(detected ?? "faculty");
      const prog = detectProgramFromFilename(filename);
      if (prog) setMeta(m => ({ ...m, program: prog }));

      // Animate remaining steps (download already done = step 0 done)
      runStepAnim(DRIVE_STEPS, 1, 400, () => {
        setDriveProcessing(false);
        setStep(2);
      });
    } catch (err) {
      setProcSteps(null);
      setDriveProcessing(false);
      setError(err.message ?? "خطأ في تحميل الملف من Drive");
    }
  }, [driveToken, runStepAnim]);

  const cancelDriveBatch = useCallback(() => {
    driveBatchCancelRef.current = true;
    setDriveBatchState(prev => prev ? {
      ...prev,
      cancelRequested: true,
      items: prev.items.map(item => item.status === "active"
        ? { ...item, stage: "جارٍ إيقاف إنشاء التقرير…" }
        : item),
    } : prev);
  }, []);

  const handleDriveBatchReports = useCallback(async (files, options = {}) => {
    const selectedFiles = (files ?? []).filter(Boolean);
    if (!driveToken || !selectedFiles.length || driveProcessing) return;
    const delivery = options.delivery === "download" ? "download" : "upload";
    const requiredMeta = options.reportMeta ?? meta;
    const preparedBy = String(requiredMeta.preparedBy ?? "").trim();
    const reviewer = String(requiredMeta.reviewer ?? "").trim();
    if (!preparedBy || !reviewer) {
      setError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء أي تقرير.");
      return;
    }
    driveBatchCancelRef.current = false;
    setDriveProcessing(true);
    setError("");
    setDriveBatchState({
      done: false,
      delivery,
      current: 0,
      total: selectedFiles.length,
      items: selectedFiles.map(file => ({ id: file.id, name: file.name, stage: "في الانتظار", status: "pending" })),
    });
    const updateItem = (id, patch) => setDriveBatchState(prev => prev ? ({
      ...prev,
      items: prev.items.map(item => item.id === id ? { ...item, ...patch } : item),
    }) : prev);
    const reports = [];

    for (let i = 0; i < selectedFiles.length; i++) {
      if (driveBatchCancelRef.current) break;
      const file = selectedFiles[i];
      setDriveBatchState(prev => prev ? { ...prev, current: i + 1 } : prev);
      try {
        updateItem(file.id, { status: "active", stage: "تحميل الاستجابات" });
        const { buffer, filename } = await downloadDriveBuffer(file, driveToken);
        if (driveBatchCancelRef.current) break;
        updateItem(file.id, { stage: "تحليل البيانات" });
        const rows = readExcel(buffer);
        const type = detectSurveyType(filename, rows[0]);
        const targetSchema = allSchemas()[type];
        if (!targetSchema) throw new Error("تعذر تحديد نوع الاستبيان تلقائيًا");
        const result = analyze(rows, targetSchema);
        updateItem(file.id, { stage: "تجهيز صفحات تقرير PDF" });
        const reportMeta = {
          year: requiredMeta.year || detectYearFromFilename(filename) || meta.year,
          program: detectProgramFromFilename(filename) || meta.program,
          preparedBy,
          reviewer,
        };
        const builtPdf = await buildBrandedReportPdf(
          result, reportMeta, settings,
          (page, total) => updateItem(file.id, { stage: `إنشاء PDF — صفحة ${page} من ${total}` }),
          { shouldCancel: () => driveBatchCancelRef.current }
        );
        if (driveBatchCancelRef.current) break;
        const reportName = safePdfReportName(filename);
        if (delivery === "upload") {
          updateItem(file.id, { stage: "رفع التقرير إلى Drive" });
          const uploaded = await uploadReportNextToSource(builtPdf.blob, reportName, file, driveToken);
          updateItem(file.id, { status: "done", stage: "تم رفع تقرير PDF", report: uploaded, localBlob: builtPdf.blob, reportName });
        } else {
          reports.push({ blob: builtPdf.blob, filename: reportName });
          updateItem(file.id, { status: "done", stage: "تم إنشاء تقرير PDF", localBlob: builtPdf.blob, reportName });
        }
      } catch (err) {
        if (err?.name === "AbortError" && driveBatchCancelRef.current) break;
        updateItem(file.id, { status: "error", stage: err.message || "فشلت المعالجة" });
      }
    }
    const cancelled = driveBatchCancelRef.current;
    setDriveBatchState(prev => prev ? { ...prev, done: true, cancelled } : prev);
    if (delivery === "download" && !cancelled) await downloadReports(reports, requiredMeta.year || meta.year);
    setDriveProcessing(false);
  }, [driveToken, driveProcessing, meta, settings]);

  const handleSemesterFormsBatch = useCallback(async (forms, token, options = {}) => {
    const selectedForms = (forms ?? []).filter(Boolean);
    if (!token || !selectedForms.length || driveProcessing) return;
    const delivery = options.delivery === "download" ? "download" : "upload";
    const requiredMeta = options.reportMeta ?? {};
    const preparedBy = String(requiredMeta.preparedBy ?? "").trim();
    const reviewer = String(requiredMeta.reviewer ?? "").trim();
    if (!preparedBy || !reviewer) {
      throw new Error("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء أي تقرير.");
    }
    driveBatchCancelRef.current = false;
    setDriveProcessing(true);
    setError("");
    setDriveBatchState({ done: false, delivery, current: 0, total: selectedForms.length, items: selectedForms.map(f => ({ id: f.id, name: f.name, stage: "في الانتظار", status: "pending" })) });
    const updateItem = (id, patch) => setDriveBatchState(prev => prev ? ({ ...prev, items: prev.items.map(item => item.id === id ? { ...item, ...patch } : item) }) : prev);
    const reports = [];
    for (let i = 0; i < selectedForms.length; i++) {
      if (driveBatchCancelRef.current) break;
      const file = selectedForms[i];
      setDriveBatchState(prev => prev ? { ...prev, current: i + 1 } : prev);
      try {
        updateItem(file.id, { status: "active", stage: "قراءة ردود Google Form" });
        const [form, responses] = await Promise.all([getForm(token, file.id), listAllResponses(token, file.id)]);
        if (driveBatchCancelRef.current) break;
        const rows = responsesToRows(form, responses);
        updateItem(file.id, { stage: "تحليل البيانات" });
        const type = detectSurveyType(file.name, rows[0]);
        const targetSchema = allSchemas()[type];
        if (!targetSchema) throw new Error("تعذر تحديد نوع الاستبيان تلقائيًا");
        const result = analyze(rows, targetSchema);
        updateItem(file.id, { stage: "تجهيز صفحات تقرير PDF" });
        const reportMeta = {
          year: requiredMeta.year || detectYearFromFilename(file.name) || meta.year,
          program: departmentFromSurveyName(file.name) || detectProgramFromFilename(file.name) || meta.program,
          preparedBy,
          reviewer,
        };
        const builtPdf = await buildBrandedReportPdf(
          result, reportMeta, settings,
          (page, total) => updateItem(file.id, { stage: `إنشاء PDF — صفحة ${page} من ${total}` }),
          { shouldCancel: () => driveBatchCancelRef.current }
        );
        if (driveBatchCancelRef.current) break;
        const reportName = safePdfReportName(file.name);
        if (delivery === "upload") {
          updateItem(file.id, { stage: "رفع التقرير بجوار النموذج" });
          const uploaded = await uploadReportNextToSource(builtPdf.blob, reportName, { parents: [file.parentId] }, token);
          updateItem(file.id, { status: "done", stage: "تم رفع تقرير PDF", report: uploaded, localBlob: builtPdf.blob, reportName });
        } else {
          reports.push({ blob: builtPdf.blob, filename: reportName });
          updateItem(file.id, { status: "done", stage: "تم إنشاء تقرير PDF", localBlob: builtPdf.blob, reportName });
        }
      } catch (err) {
        if (err?.name === "AbortError" && driveBatchCancelRef.current) break;
        updateItem(file.id, { status: "error", stage: err.message || "فشلت المعالجة" });
      }
    }
    const cancelled = driveBatchCancelRef.current;
    setDriveBatchState(prev => prev ? { ...prev, done: true, cancelled } : prev);
    if (delivery === "download" && !cancelled) await downloadReports(reports, requiredMeta.year || meta.year);
    setDriveProcessing(false);
  }, [driveProcessing, meta, settings]);

  // Same tail as handleDriveFileSelect, but for a survey picked via SemesterFormPicker —
  // rows are already fetched (Forms API responses converted to Excel-row shape) by the
  // time this runs, so "step 0: تحميل الملف من Drive" is marked done immediately.
  const handleDriveFormSelect = useCallback((rows, filename, department) => {
    setDriveProcessing(true);
    setError("");
    setProcFile(filename);
    setProcSteps(DRIVE_STEPS.map((label, i) => ({ label, status: i === 0 ? "done" : "pending" })));
    try {
      setSingleFile({ name: filename });
      setRawRows(rows);
      const detected = detectSurveyType(filename, rows[0]) ?? null;
      setDetectedAutoType(detected);
      setSurveyType(detected ?? "faculty");
      if (department) setMeta(m => ({ ...m, program: department }));
      runStepAnim(DRIVE_STEPS, 1, 400, () => {
        setDriveProcessing(false);
        setStep(2);
      });
    } catch (err) {
      setProcSteps(null);
      setDriveProcessing(false);
      setError(err.message ?? "خطأ في معالجة الاستبيان");
    }
  }, [runStepAnim]);

  // Step 0: read file headers → auto-detect → advance to step 1
  const handleFileSelected = useCallback((file) => {
    setDriveSelected(null);
    setSingleFile(file);
    setError("");
    setProcFile(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = readExcel(e.target.result);
        setRawRows(rows);
        const detected = detectSurveyType(file.name, rows[0]) ?? null;
        setDetectedAutoType(detected);
        setSurveyType(detected ?? "faculty");
        const prog = detectProgramFromFilename(file.name);
        if (prog) setMeta(m => ({ ...m, program: prog }));
        // Animate all 5 steps then advance
        runStepAnim(LOCAL_STEPS, 0, 520, () => setStep(2));
      } catch (err) {
        setProcSteps(null);
        setError(err.message ?? "تعذّر قراءة الملف. تأكد أنه ملف Excel صحيح.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, [runStepAnim]);

  // Step 1: user confirms the detected type → prepare data → step 2
  const handleValidationConfirm = useCallback(() => {
    if (!rawRows || !surveyType) return;
    setProcessing(true); setError("");
    try {
      const s = allSchemas()[surveyType];
      const { headers, dataRows, metaCols } = prepareData(rawRows, s);
      setSingleHeaders(headers);
      setSingleAllRows(dataRows);
      setSingleSchema(s);
      setSingleMetaCols(metaCols);
      setSingleRemoved(new Set());
      setSingleFilters({ dept: "", degree: "" });
      // Auto-fill preparedBy from settings responsible person (if empty)
      const responsible = settings.surveyResponsible?.[surveyType];
      if (responsible) setMeta(m => ({ ...m, preparedBy: m.preparedBy || responsible }));
      setStep(3);
    } catch (err) {
      setError(err.message ?? "خطأ في تحضير البيانات.");
    } finally {
      setProcessing(false);
    }
  }, [rawRows, surveyType, settings]);

  // Step 3 (comparison): process all files (slot 0 = singleFile from step 0)
  const handleProcessSingle = useCallback(() => {
    if (!singleFile) return;
    setProcessing(true); setError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = readExcel(e.target.result);
        const type = detectSurveyType(singleFile.name, rows[0]) ?? surveyType;
        const s = allSchemas()[type] ?? schema;
        const { headers, dataRows, metaCols } = prepareData(rows, s);
        setSingleHeaders(headers);
        setSingleAllRows(dataRows);
        setSingleSchema(s);
        setSingleMetaCols(metaCols);
        setSingleRemoved(new Set());
        setSingleFilters({ dept: "", degree: "" });
        // Auto-detect program from filename — only if user hasn't already filled it
        const detected = detectProgramFromFilename(singleFile.name);
        if (detected) setMeta(m => ({ ...m, program: detected }));
        setStep(3);
      } catch (err) {
        setError(err.message ?? "حدث خطأ في تحليل الملف.");
      } finally { setProcessing(false); }
    };
    reader.readAsArrayBuffer(singleFile);
  }, [singleFile, surveyType, schema]);

  const handleProcessMulti = useCallback(() => {
    const activeSlots = slots.slice(0, slotCount);
    if (!activeSlots.every(s => s._file)) { setError("ارفع ملفاً لكل سنة."); return; }

    // Validate unique years
    const years = activeSlots.map(s => (s.year ?? "").trim());
    if (new Set(years).size !== activeSlots.length) {
      setError("كل ملف يجب أن يكون لسنة دراسية مختلفة — تحقق من السنوات المدخلة.");
      return;
    }

    // Validate same survey type (when all slots have a detected type)
    const slotTypes = activeSlots.map(s => s._type).filter(Boolean);
    if (slotTypes.length === activeSlots.length) {
      const uniqueTypes = new Set(slotTypes);
      if (uniqueTypes.size > 1) {
        setError(`الملفات لأنواع استبيانات مختلفة: ${[...uniqueTypes].map(t => allSchemas()[t]?.label ?? t).join(" / ")} — تأكد أن جميع الملفات لنفس نوع الاستبيان.`);
        return;
      }
    }

    setProcessing(true); setError("");
    let pending = activeSlots.length;
    const results = Array(activeSlots.length).fill(null);
    let firstDetectedType = null;
    const firstDetectedProgram = activeSlots.find(s => s._program)?._program ?? null;

    activeSlots.forEach((slot, idx) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const rows = readExcel(e.target.result);
          const type = detectSurveyType(slot._file.name, rows[0]) ?? surveyType;
          if (!firstDetectedType) firstDetectedType = type;
          const s = allSchemas()[type] ?? schema;
          const result = analyze(rows, s);
          results[idx] = { year: slot.year, result };
        } catch (err) {
          setError(err.message ?? "خطأ في ملف " + slot._file.name);
        } finally {
          if (--pending === 0) {
            setProcessing(false);
            if (firstDetectedType) setSurveyType(firstDetectedType);
            if (firstDetectedProgram) setMeta(m => ({ ...m, program: m.program || firstDetectedProgram }));
            const validResults = results.filter(Boolean);
            if (validResults.length > 0) {
              const cmp = buildComparison(validResults);
              setComparison({ ...cmp, slots: results });
              setStep(4);
            }
          }
        }
      };
      reader.readAsArrayBuffer(slot._file);
    });
  }, [slots, slotCount, surveyType, schema]);

  const handleDropSingle = (e) => {
    e.preventDefault(); setDragging(false);
    const files = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
    if (!files.length) return;
    if (files.length === 1) { handleFileSelected(files[0]); }
    else                    { setBatchFiles(files); setBatchMode(true); }
  };

  const downloadAnnual = async () => {
    if (!String(meta.preparedBy ?? "").trim() || !String(meta.reviewer ?? "").trim()) {
      setError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء التقرير.");
      return;
    }
    setProcessing(true);
    try {
      const blob = await buildAnnualDocx(singleResult, { ...meta }, settings);
      const dept = meta.program ? `_${meta.program}` : "";
      downloadBlob(blob, `تقرير_${schema.label}${dept}_${meta.year}.docx`);
    } finally { setProcessing(false); }
  };

  const downloadAnnualPdf = async () => {
    if (!String(meta.preparedBy ?? "").trim() || !String(meta.reviewer ?? "").trim()) {
      setError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء التقرير.");
      return;
    }
    pdfCancelRef.current = false;
    setPdfCancelRequested(false);
    setPdfGenerating(true);
    setPdfProgress({ current: 0, total: 0 });
    try {
      const builtPdf = await buildBrandedReportPdf(
        singleResult, { ...meta }, { ...settings, includePdfCharts: pdfIncludeCharts },
        (current, total) => setPdfProgress({ current, total }),
        { shouldCancel: () => pdfCancelRef.current }
      );
      if (!pdfCancelRef.current) downloadBlob(builtPdf.blob, builtPdf.filename || safePdfReportName(singleFile?.name));
    } catch (err) {
      if (err?.name !== "AbortError") setError(`تعذّر إنشاء ملف PDF: ${err.message}`);
    } finally { setPdfGenerating(false); setPdfCancelRequested(false); }
  };

  const downloadComparison = async () => {
    if (!String(meta.preparedBy ?? "").trim() || !String(meta.reviewer ?? "").trim()) {
      setError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء التقرير.");
      return;
    }
    setProcessing(true);
    try {
      const blob = await buildComparisonDocx(comparison, { ...meta }, settings);
      const dept = meta.program ? `_${meta.program}` : "";
      const years = comparison.slots?.map(s => s.year).filter(Boolean).join("_و_") ?? meta.year;
      downloadBlob(blob, `مقارنة_${schema.label}${dept}_${years}.docx`);
    } finally { setProcessing(false); }
  };

  const reset = () => {
    setStep(1); setSingleFile(null); setSingleResult(null); setComparison(null);
    setSlots(defaultSlots()); setError(""); setProcessing(false); setShowSettings(false);
    setShowSemesterSurveys(false);
    setSingleHeaders(null); setSingleAllRows(null); setSingleSchema(null);
    setSingleMetaCols(null); setSingleRemoved(new Set());
    setSingleFilters({ dept: "", degree: "" });
    setDetectedAutoType(null); setRawRows(null);
    setBatchMode(false); setBatchFiles([]);
    setUploadTab("upload"); setDriveSelected(null); setDriveViewMode("list");
    setDriveFilterType(""); setDriveFilterYear(""); setDriveFilterProgram("");
  };

  const isAnnual = mode === "annual";
  const canProceed = isAnnual ? !!singleFile : slots.slice(0, slotCount).every(s => s._file);

  const openHubTool = (tool) => {
    if (tool === "tutorial") { setShowTutorial(true); return; }
    if (tool === "home") { goToHub(); return; }
    const isSemesterTool = tool === "semester" || tool === "semester-create" || tool === "semester-quick-analysis" || tool === "semester-dashboard";
    if (tool === "semester") setSemesterInitialTab("home");
    if (tool === "semester-create") setSemesterInitialTab("generate");
    if (tool === "semester-quick-analysis") setSemesterInitialTab("quick-analysis");
    if (tool === "semester-dashboard") setSemesterInitialTab("dashboard");
    setShowHub(false);
    setShowSurveyManagement(tool === "surveys");
    setShowSemesterSurveys(isSemesterTool);
    setShowCourseEval(tool === "courses");
    setShowSettings(tool === "settings");
    if (tool === "analytics") reset();
    if (window.innerWidth <= 760) setSidebarOpen(false);
  };

  const uploadAnnualPdf = async () => {
    if (!String(meta.preparedBy ?? "").trim() || !String(meta.reviewer ?? "").trim()) {
      setError("يجب إدخال اسم مُعدّ التحليل واسم مراجع التحليل قبل إنشاء التقرير.");
      return;
    }
    if (!driveToken) {
      setError("يرجى ربط Google Drive أولاً لرفع التقرير.");
      return;
    }
    pdfCancelRef.current = false;
    setPdfCancelRequested(false);
    setPdfGenerating(true);
    setPdfProgress({ current: 0, total: 0 });
    try {
      const builtPdf = await buildBrandedReportPdf(
        singleResult, { ...meta }, { ...settings, includePdfCharts: pdfIncludeCharts },
        (current, total) => setPdfProgress({ current, total }),
        { shouldCancel: () => pdfCancelRef.current }
      );
      if (pdfCancelRef.current) return;
      await uploadReportNextToSource(builtPdf.blob, builtPdf.filename || safePdfReportName(singleFile?.name), driveSelected ?? {}, driveToken);
      setError("");
    } catch (err) {
      if (err?.name !== "AbortError") setError(`تعذّر رفع ملف PDF: ${err.message}`);
    } finally { setPdfGenerating(false); setPdfCancelRequested(false); }
  };

  const openSemesterAnalysis = () => {
    openHubTool("analytics");
    setMode("annual");
    setStep(1);
    setUploadTab("drive");
    setDriveViewMode("semester");
  };

  const goToHub = () => {
    setShowSurveyManagement(false);
    setShowSemesterSurveys(false);
    setShowCourseEval(false);
    setShowSettings(false);
    setShowHub(true);
  };

  const activeTool = showHub
    ? "home"
    : showSurveyManagement
      ? "surveys"
    : showSemesterSurveys
        ? "semester"
        : showCourseEval
          ? "courses"
          : showSettings
            ? "settings"
            : "analytics";
  const activeToolTitle = activeTool === "settings"
    ? "الإعدادات"
    : activeTool === "surveys"
      ? "تصميم الاستبيانات"
      : SIDEBAR_ITEMS.find(item => item.id === activeTool)?.label ?? "بوابة وحدة ضمان الجودة";
  const showGoogleConnection = !showHub && (activeTool === "analytics" || activeTool.startsWith("semester"));
  const showAiChat = showHub || (activeTool === "analytics" && step === 5);

  return (
    <div className="quality-app" style={{
      minHeight: "100vh",
      background: "linear-gradient(145deg,#091a2d 0%,#102f4b 58%,#0b352f 100%)",
      fontFamily: "'Cairo',sans-serif", direction: "rtl",
    }}>
      <style>{CSS}</style>

      {/* Header */}
      <header className={`app-header quality-header ${sidebarOpen ? "sidebar-open" : "sidebar-closed"}`} style={{ padding: "14px 24px", borderBottom: "1px solid rgba(255,255,255,.08)", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11, color: "#fff", fontSize: 14, fontWeight: 900 }}>
          <button type="button" className="sidebar-toggle mobile-nav-toggle" onClick={() => setSidebarOpen(v => !v)} aria-label={sidebarOpen ? "طي القائمة الجانبية" : "فتح القائمة الجانبية"} title={sidebarOpen ? "طي القائمة" : "فتح القائمة"}>
            <svg width="21" height="21" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round"><path d="M4 6h16M4 12h16M4 18h16"/></svg>
          </button>
          <span>{activeToolTitle}</span>
        </div>
        {showGoogleConnection && (driveConnection === "connected" ? (
          <button type="button" className="google-global-btn connected" onClick={connectDrive} title="متصل لكل أدوات Drive وForms — اضغط لتحديث الاتصال أو تغيير الحساب"><GoogleDriveIcon size={20} /><span className="google-status-dot" /><span className="google-global-label">{driveAccount || "Google Drive متصل"}</span></button>
        ) : (
          <button type="button" className="google-global-btn disconnected" disabled={driveConnecting || driveConnection === "checking"} onClick={connectDrive}>
            <GoogleDriveIcon size={20} />
            <span className="google-global-label">{driveConnection === "checking" ? <>جاري التحقق<span className="loader-dots"><i/><i/><i/></span></> : driveConnecting ? <>جاري الاتصال<span className="loader-dots"><i/><i/><i/></span></> : "ربط Google Drive"}</span>
          </button>
        ))}
      </header>

      <div className={`app-shell ${sidebarOpen ? "sidebar-open" : "sidebar-closed"}`}>
        <ToolSidebar open={sidebarOpen} active={activeTool} onNavigate={openHubTool} onTutorial={() => setShowTutorial(true)} onToggle={() => setSidebarOpen(v => !v)} />
        <div className="app-content">
          <div className="app-main" style={{ maxWidth: 1400, margin: "0 auto", padding: "28px 32px" }}>

        {/* ── Survey Management view (new, independent of the existing wizard/engine) ── */}
        {showHub ? (
          <QualityHub onOpen={openHubTool} />
        ) : showSurveyManagement ? (
          <SurveyManagement />
        ) : showSemesterSurveys ? (
          <SemesterSurveys key={semesterInitialTab} initialTab={semesterInitialTab} onOpenAnalysis={openSemesterAnalysis} onAnalyzeForms={handleSemesterFormsBatch} onSectionChange={setSemesterInitialTab} googleAuth={{ token: driveConnection === "connected" ? driveToken : null, connecting: driveConnecting || driveConnection === "checking", connect: connectDrive, authError: error }} />
        ) : showCourseEval ? (
          <CourseEvaluationHub />
        ) : showSettings ? (
          <div className="card" style={{ padding: 32 }}>
            <SettingsPanel
              settings={settings} onChange={setSettings}
              aiSettings={aiSettings} onAiChange={setAiSettings}
            />
            <div style={{ textAlign: "center", marginTop: 24 }}>
              <button className="btn btn-primary" onClick={() => setShowSettings(false)}>
                ✓ حفظ والعودة
              </button>
            </div>
          </div>
        ) : batchMode ? (
          <BatchProcessor
            files={batchFiles}
            year={batchYear}
            onYearChange={setBatchYear}
            settings={settings}
            reportMeta={meta}
            onReportMetaChange={setMeta}
            onBack={() => { setBatchMode(false); setBatchFiles([]); setError(""); }}
          />
        ) : (
          <>
            {step > 0 && <StepBar step={step} />}

            {/* ══ STEP 0 — Choose report mode ══ */}
            {step === 0 && (
              <div>
                <ReportModePicker value={mode} onChange={setMode} />
                <div style={{ display: "flex", justifyContent: "center", marginTop: 24 }}>
                  <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 36px" }}
                    onClick={() => { setError(""); setStep(1); }}>
                    التالي ←
                  </button>
                </div>
              </div>
            )}

            {/* ══ STEP 1 — Upload file(s) ══ */}
            {step === 1 && isAnnual && (
              <div className="card upload-workspace">
                <WorkflowHeader step="2" icon="upload" title="رفع ملف الاستبيان" description="اختر من الجهاز أو Drive، وسيتم التعرف على نوعه تلقائيًا." />
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 18, padding: "10px 13px", borderRadius: 12, background: "rgba(26,188,156,.07)", border: "1px solid rgba(94,234,212,.15)" }}>
                  <span style={{ color: "rgba(255,255,255,.78)", fontSize: 13 }}>النطاق الحالي: <b style={{ color: "#8ff3df" }}>تقرير سنة واحدة</b></span>
                  <button type="button" className="btn btn-ghost btn-sm" onClick={() => setStep(0)}>تغيير إلى مقارنة</button>
                </div>
                {!(uploadTab === "drive" && driveToken) && (<>
                  <div style={{ color: "rgba(255,255,255,.45)", fontSize: 13, marginBottom: 24 }}>
                    يُكتشف نوع الاستبيان تلقائياً من الأعمدة · يدعم .xlsx / .xls / .csv
                  </div>
                </>)}

                {/* Source tabs */}
                <div className="upload-source-grid">
                  <button className={`upload-source-card ${uploadTab === "upload" ? "active" : ""}`} onClick={() => setUploadTab("upload")}>
                    <span className="upload-source-icon"><StepIcon name="upload" /></span>
                    <span><strong>رفع ملف من الجهاز</strong><small>Excel أو CSV من جهازك، مع دعم اختيار عدة ملفات</small></span>
                  </button>
                  <button className={`upload-source-card ${uploadTab === "drive" ? "active" : ""}`} onClick={() => setUploadTab("drive")}>
                    <span className="upload-source-icon"><GoogleDriveIcon size={24} /></span>
                    <span><strong>Google Drive</strong><small>ابحث واختر الاستبيانات المحفوظة على حسابك</small></span>
                  </button>
                </div>

                {/* ── Upload tab ── */}
                {uploadTab === "upload" && (
                  <div
                    className={`upload-zone upload-drop-modern ${dragging ? "drag" : ""}`}
                    onDragOver={e => { e.preventDefault(); setDragging(true); }}
                    onDragLeave={() => setDragging(false)}
                    onDrop={handleDropSingle}
                    onClick={() => singleRef.current.click()}
                  >
                    <div className="upload-drop-icon"><StepIcon name="upload" /></div>
                    <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>
                      اسحب الملفات هنا أو اضغط للاختيار
                    </div>
                    <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>
                      يدعم Excel وCSV، ويمكن اختيار عدة ملفات للتحليل الجماعي
                    </div>
                    <input ref={singleRef} type="file" accept=".xlsx,.xls,.csv" multiple
                      style={{ display: "none" }}
                      onChange={e => {
                        const files = [...e.target.files];
                        if (!files.length) return;
                        if (files.length === 1) { handleFileSelected(files[0]); }
                        else                    { setBatchFiles(files); setBatchMode(true); }
                      }} />
                  </div>
                )}

                {/* ── Google Drive tab ── */}
                {uploadTab === "drive" && (
                  <div style={{ maxWidth: driveToken ? "100%" : 520, margin: driveToken ? "0" : "0 auto", textAlign: "right" }}>
                    {!GOOGLE_CLIENT_ID ? (
                      /* No Client ID configured */
                      <div style={{
                        background: "rgba(243,156,18,.07)", border: "1px solid rgba(243,156,18,.3)",
                        borderRadius: 14, padding: "22px 24px", fontSize: 13, color: "#f1c40f", lineHeight: 1.9,
                      }}>
                        ⚠️ لم يتم إعداد Google Client ID بعد. اتبع الخطوات:<br/>
                        <span style={{ color: "rgba(255,255,255,.6)", fontSize: 12 }}>
                          ١. إنشاء مشروع في Google Cloud Console<br/>
                          ٢. تفعيل Google Drive API<br/>
                          ٣. إنشاء OAuth 2.0 Client ID (نوع: Web application)<br/>
                          ٤. إضافة <code style={{ color: "#1abc9c" }}>http://localhost:3000</code> في Authorized JavaScript Origins<br/>
                          ٥. نسخ الـ Client ID وإضافته في ملف <code style={{ color: "#1abc9c" }}>.env.local</code>:<br/>
                          <code style={{ color: "#1abc9c", fontSize: 11 }}>VITE_GOOGLE_CLIENT_ID=your_client_id_here</code><br/>
                          ٦. إعادة تشغيل الـ dev server
                        </span>
                      </div>
                    ) : !driveToken ? (
                      /* Not connected */
                      <div style={{ textAlign: "center", padding: "28px 0" }}>
                        <div style={{ width: 68, height: 68, display: "grid", placeItems: "center", margin: "0 auto 14px", borderRadius: 20, background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)" }}><GoogleDriveIcon size={38} /></div>
                        <div style={{ color: "rgba(255,255,255,.5)", fontSize: 13, marginBottom: 22 }}>
                          اربط حسابك على Google Drive لاختيار ملف الاستبيان مباشرة
                        </div>
                        <button onClick={connectDrive} style={{
                          display: "inline-flex", alignItems: "center", gap: 10,
                          padding: "13px 28px", borderRadius: 50, border: "1px solid #dadce0",
                          background: "#fff", color: "#3c4043", cursor: "pointer",
                          fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 15,
                          boxShadow: "0 2px 8px rgba(0,0,0,.2)", transition: "all .2s",
                        }}
                          onMouseOver={e => e.currentTarget.style.boxShadow = "0 4px 16px rgba(0,0,0,.3)"}
                          onMouseOut={e  => e.currentTarget.style.boxShadow = "0 2px 8px rgba(0,0,0,.2)"}
                        >
                          <GoogleDriveIcon size={23} />
                          ربط Google Drive
                        </button>
                      </div>
                    ) : driveLoading ? (
                      /* Loading files */
                      <div style={{ textAlign: "center", padding: "36px 0", color: "rgba(255,255,255,.5)" }}>
                        <div className="loader-visual" style={{ width: 58, height: 58 }}><span className="loader-ring"/><span className="loader-ring inner"/><span style={{ width: 20, height: 20, color: "#d8fff7" }}><StepIcon name="upload"/></span></div>
                        جاري قراءة ملفات Drive<span className="loader-dots"><i/><i/><i/></span>
                      </div>
                    ) : (
                      /* File list */
                      <div>
                        {/* Connected bar */}
                        <div className="drive-control-bar">
                          <span style={{ color: "#dce8f7", fontWeight: 800, fontSize: 12, display: "inline-flex", alignItems: "center", gap: 8 }}><GoogleDriveIcon size={21} /> Google Drive متصل <span className="google-status-dot" aria-hidden="true" /></span>
                          <div style={{ display: "flex", gap: 8, alignItems: "center", maxWidth: "100%" }}>
                            {/* View toggle */}
                            <div className="drive-view-switch">
                              {[
                                { id: "list", label: "الملفات" },
                                { id: "dashboard", label: "لوحة التحكم" },
                                { id: "semester", label: "استبيانات الفصل" },
                              ].map(v => (
                                <button key={v.id} className={`drive-view-btn ${driveViewMode === v.id ? "active" : ""}`} onClick={() => setDriveViewMode(v.id)}>{v.label}</button>
                              ))}
                            </div>
                            <button onClick={() => fetchDriveFiles(driveToken)} style={{
                              background: "rgba(255,255,255,.08)", border: "1px solid rgba(255,255,255,.15)",
                              borderRadius: 8, padding: "5px 11px", color: "rgba(255,255,255,.6)",
                              cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif",
                            }}>🔄</button>
                            <button onClick={disconnectDrive} style={{
                              background: "rgba(231,76,60,.1)", border: "1px solid rgba(231,76,60,.25)",
                              borderRadius: 8, padding: "5px 11px", color: "#ff6b6b",
                              cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif",
                            }}>قطع</button>
                          </div>
                        </div>

                        {/* Dashboard / Semester-survey / flat-list views */}
                        {driveViewMode === "dashboard" ? (
                          <DriveDashboard
                            files={driveFiles}
                            token={driveToken}
                            onSelectFile={(f) => { setDriveViewMode("list"); handleDriveFileSelect(f); }}
                          />
                        ) : driveViewMode === "semester" ? (
                          <SemesterFormPicker onFormSelected={handleDriveFormSelect} onAnalyzeForms={handleSemesterFormsBatch} googleAuth={{ token: driveConnection === "connected" ? driveToken : null, connecting: driveConnecting || driveConnection === "checking", connect: connectDrive, authError: error }} />
                        ) : (<>
                        {/* Filters row */}
                        {(() => {
                          const yearSet = new Set();
                          driveFiles.forEach(f => { const y = detectYearFromFilename(f.name); if (y) yearSet.add(y); });
                          const availableYears = [...yearSet].sort().reverse();
                          const selStyle = (active) => ({
                            flex: 1, minWidth: 0, background: "#0d1f33",
                            color: active ? "#1abc9c" : "rgba(255,255,255,.5)",
                            border: `1px solid ${active ? "rgba(26,188,156,.5)" : "rgba(255,255,255,.18)"}`,
                            borderRadius: 9, padding: "8px 10px",
                            fontFamily: "'Cairo',sans-serif", fontSize: 12,
                            cursor: "pointer", outline: "none", appearance: "none",
                          });
                          const hasFilter = driveFilterType || driveFilterYear || driveFilterProgram;
                          return (
                            <div className="drive-filter-grid">
                              <select value={driveFilterType}
                                onChange={e => { setDriveFilterType(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterType)}>
                                <option value="">كل الاستبيانات</option>
                                {Object.values(allSchemas()).map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                              </select>
                              <select value={driveFilterYear}
                                onChange={e => { setDriveFilterYear(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterYear)}>
                                <option value="">كل السنوات</option>
                                {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
                              </select>
                              <select value={driveFilterProgram}
                                onChange={e => { setDriveFilterProgram(e.target.value); setDriveSelected(null); }}
                                style={selStyle(!!driveFilterProgram)}>
                                <option value="">كل البرامج</option>
                                {PROGRAMS.map(p => <option key={p} value={p}>{p}</option>)}
                              </select>
                              {hasFilter && (
                                <button
                                  onClick={() => { setDriveFilterType(""); setDriveFilterYear(""); setDriveFilterProgram(""); setDriveSelected(null); }}
                                  style={{ padding: "7px 11px", background: "rgba(231,76,60,.12)", border: "1px solid rgba(231,76,60,.25)", borderRadius: 9, color: "#ff8a80", cursor: "pointer", fontSize: 12, fontFamily: "'Cairo',sans-serif", whiteSpace: "nowrap" }}>
                                  ↩ مسح
                                </button>
                              )}
                            </div>
                          );
                        })()}

                        {/* Search */}
                        <input
                          value={driveSearch}
                          onChange={e => { setDriveSearch(e.target.value); setDriveSelected(null); }}
                          placeholder="ابحث باسم الملف أو البرنامج..."
                          className="drive-search-field"
                          style={{
                            direction: "rtl",
                          }}
                        />

                        {/* Files */}
                        {(() => {
                          const filtered = driveFiles.filter(f => {
                            if (driveSearch && !f.name.toLowerCase().includes(driveSearch.toLowerCase())) return false;
                            if (driveFilterType && detectTypeHintFromFilename(f.name) !== driveFilterType) return false;
                            if (driveFilterYear && detectYearFromFilename(f.name) !== driveFilterYear) return false;
                            if (driveFilterProgram && detectProgramFromFilename(f.name) !== driveFilterProgram) return false;
                            return true;
                          });
                          const hasAnyFilter = driveSearch || driveFilterType || driveFilterYear || driveFilterProgram;
                          const selectedFiles = filtered.filter(f => driveSelectedIds.has(f.id));
                          return (
                            <>
                              {!!filtered.length && (
                                <div className="drive-batch-toolbar">
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                                    <span style={{ color: "#b9fff2", fontSize: 12, fontWeight: 800 }}>{selectedFiles.length} محدد</span>
                                    <button className="btn btn-ghost btn-sm" onClick={() => setDriveSelectedIds(prev => {
                                      const next = new Set(prev); filtered.forEach(f => next.add(f.id)); return next;
                                    })}>تحديد الظاهر</button>
                                    {!!driveSelectedIds.size && <button className="btn btn-ghost btn-sm" onClick={() => setDriveSelectedIds(new Set())}>إلغاء التحديد</button>}
                                  </div>
                                  <div style={{ display: "flex", gap: 8, flex: "1 1 360px", flexWrap: "wrap" }}>
                                    <input value={meta.preparedBy ?? ""} onChange={e => setMeta(m => ({ ...m, preparedBy: e.target.value }))} placeholder="مُعدّ التحليل *" aria-label="مُعدّ التحليل" style={{ flex: "1 1 160px", minWidth: 0, padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(255,255,255,.14)", background: "rgba(4,18,34,.45)", color: "#fff", fontFamily: "inherit", fontSize: 11 }} />
                                    <input value={meta.reviewer ?? ""} onChange={e => setMeta(m => ({ ...m, reviewer: e.target.value }))} placeholder="مراجع التحليل *" aria-label="مراجع التحليل" style={{ flex: "1 1 160px", minWidth: 0, padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(255,255,255,.14)", background: "rgba(4,18,34,.45)", color: "#fff", fontFamily: "inherit", fontSize: 11 }} />
                                  </div>
                                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                                    <button className="btn btn-ghost btn-sm" disabled={!selectedFiles.length || driveProcessing}
                                      style={{ opacity: selectedFiles.length && !driveProcessing ? 1 : .45 }}
                                      onClick={() => handleDriveBatchReports(selectedFiles, { delivery: "download", reportMeta: meta })}>
                                      تحليل وتحميل المحدد ({selectedFiles.length})
                                    </button>
                                    <button className="btn btn-primary btn-sm" disabled={!selectedFiles.length || driveProcessing}
                                      style={{ opacity: selectedFiles.length && !driveProcessing ? 1 : .45 }}
                                      onClick={() => handleDriveBatchReports(selectedFiles, { delivery: "upload", reportMeta: meta })}>
                                      تحليل ورفع المحدد ({selectedFiles.length})
                                    </button>
                                    <button className="btn btn-ghost btn-sm" disabled={driveProcessing}
                                      onClick={() => handleDriveBatchReports(filtered, { delivery: "download", reportMeta: meta })}>
                                      تحليل وتحميل الكل
                                    </button>
                                    <button className="btn btn-primary btn-sm" disabled={driveProcessing}
                                      onClick={() => handleDriveBatchReports(filtered, { delivery: "upload", reportMeta: meta })}>
                                      تحليل ورفع الكل
                                    </button>
                                  </div>
                                </div>
                              )}
                              {hasAnyFilter && (
                                <div style={{ color: "rgba(255,255,255,.35)", fontSize: 11, marginBottom: 6, textAlign: "right" }}>
                                  {filtered.length} من {driveFiles.length} ملف
                                </div>
                              )}
                              {!filtered.length ? (
                                <div style={{ textAlign: "center", color: "rgba(255,255,255,.3)", padding: 24, fontSize: 13 }}>
                                  {driveFiles.length ? "لا توجد ملفات تطابق الفلاتر" : "لا توجد ملفات Excel أو CSV في Drive"}
                                </div>
                              ) : (
                                <div className="drive-file-grid">
                                  {filtered.map((f, index) => {
                                    const isChecked = driveSelectedIds.has(f.id);
                                    const prog = detectProgramFromFilename(f.name);
                                    const yr = detectYearFromFilename(f.name);
                                    const type = detectTypeHintFromFilename(f.name);
                                    const tags = [type && allSchemas()[type]?.label, yr, prog].filter(Boolean);
                                    const toggleFile = () => {
                                      setDriveSelected(f);
                                      setDriveSelectedIds(prev => {
                                        const next = new Set(prev);
                                        if (next.has(f.id)) next.delete(f.id); else next.add(f.id);
                                        return next;
                                      });
                                    };
                                    return (
                                      <article key={f.id} className={`drive-file-card ${isChecked ? "selected" : ""}`}
                                        role="button" tabIndex={0} aria-pressed={isChecked}
                                        style={{ animationDelay: `${Math.min(index * 25, 250)}ms` }}
                                        onClick={toggleFile}
                                        onKeyDown={e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleFile(); } }}>
                                        <span className="drive-file-select" aria-hidden="true">{isChecked ? "✓" : "+"}</span>
                                        <span className="drive-file-kind" aria-hidden="true">
                                          <svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M6 2h9l3 3v17H6z"/><path d="M9 11h6M9 15h6M9 7h3"/></svg>
                                        </span>
                                        <div className="drive-file-name">{f.name}</div>
                                        {!!tags.length && <div className="drive-file-meta">{tags.map((tag, i) => <span className="drive-file-tag" key={`${tag}-${i}`}>{tag}</span>)}</div>}
                                        <div className="drive-file-foot">
                                          <span className="drive-file-date">{f.modifiedTime?.slice(0, 10) ?? ""}</span>
                                          <button className="drive-file-analyze" onClick={e => { e.stopPropagation(); handleDriveFileSelect(f); }}>تحليل الملف</button>
                                        </div>
                                      </article>
                                    );
                                  })}
                                </div>
                              )}
                            </>
                          );
                        })()}

                        </>)}
                      </div>
                    )}
                  </div>
                )}

                <InlineNotice text={error} style={{ marginTop: 20 }} />

                <div style={{ textAlign: "center", marginTop: 20 }}>
                  <button className="btn btn-ghost btn-sm" onClick={() => { setError(""); setStep(0); }}>→ السابق</button>
                </div>
              </div>
            )}

            {/* Comparison: upload all slots in step 1 */}
            {step === 1 && !isAnnual && (
              <div className="card" style={{ padding: 36 }}>
                <div style={{ color: "#fff", fontSize: 20, fontWeight: 900, textAlign: "center", marginBottom: 8 }}>
                  رفع ملفات الاستبيانات
                </div>
                <div style={{ color: "rgba(255,255,255,.4)", fontSize: 13, textAlign: "center", marginBottom: 24 }}>
                  {mode === "compare2" ? "حدد ملف لكل سنة من السنتين" : "حدد ملف لكل سنة من السنوات الثلاث"}
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16 }}>
                  {slots.slice(0, slotCount).map((slot, idx) => (
                    <FileSlot
                      key={idx}
                      slot={slot}
                      label={`العام الأكاديمي ${idx + 1}`}
                      onChange={updated => setSlots(prev => prev.map((s, i) => i === idx ? updated : s))}
                      driveToken={driveToken}
                      driveFiles={driveFiles}
                      driveLoading={driveLoading}
                      onConnectDrive={connectDrive}
                    />
                  ))}
                </div>

                {/* Inline conflict warnings */}
                {(() => {
                  const active = slots.slice(0, slotCount).filter(s => s._file);
                  if (active.length < 2) return null;
                  const warnings = [];
                  const yrs = active.map(s => (s.year ?? "").trim()).filter(Boolean);
                  if (new Set(yrs).size < yrs.length)
                    warnings.push("⚠ بعض الملفات لها نفس السنة الدراسية — يجب أن تكون السنوات مختلفة.");
                  const types = active.map(s => s._type).filter(Boolean);
                  if (types.length === active.length && new Set(types).size > 1)
                    warnings.push(`⚠ أنواع استبيانات مختلفة: ${[...new Set(types)].map(t => allSchemas()[t]?.label).join(" / ")} — يجب أن تكون الملفات لنفس نوع الاستبيان.`);
                  const progs = active.map(s => s._program).filter(Boolean);
                  if (progs.length > 1 && new Set(progs).size > 1)
                    warnings.push(`⚠ برامج مختلفة محتملة: ${[...new Set(progs)].join(" / ")} — تأكد أن الملفات لنفس البرنامج.`);
                  if (!warnings.length) return null;
                  return (
                    <div style={{ background: "rgba(243,156,18,.08)", border: "1px solid rgba(243,156,18,.3)", borderRadius: 10, padding: "12px 16px", marginTop: 16 }}>
                      {warnings.map((w, i) => (
                        <div key={i} style={{ color: "#ffd54f", fontSize: 12, fontWeight: 700, lineHeight: 1.7 }}>{w}</div>
                      ))}
                    </div>
                  );
                })()}

                <InlineNotice text={error} style={{ marginTop: 16 }} />

                <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 28 }}>
                  <button className="btn btn-ghost" onClick={() => { setError(""); setStep(0); }}>→ السابق</button>
                  <button
                    className="btn btn-primary"
                    disabled={!slots.slice(0, slotCount).every(s => s._file) || processing}
                    style={{ opacity: (!slots.slice(0, slotCount).every(s => s._file) || processing) ? 0.5 : 1 }}
                    onClick={handleProcessMulti}
                  >
                    {processing
                      ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحليل…</>
                      : "تحليل ←"}
                  </button>
                </div>
              </div>
            )}

            {/* ══ STEP 2 — Validation (annual only) ══ */}
            {step === 2 && rawRows && (() => {
              const s = allSchemas()[surveyType];
              const totalQ = s?.axes.reduce((acc, ax) => acc + ax.questions.length, 0) ?? 0;
              const dataRowCount = rawRows.length - 1;
              const colHeaders = rawRows[0] ?? [];
              const dataRows = rawRows.slice(1);
              const blankRowCount = dataRows.filter(
                row => row.every(cell => cell == null || String(cell).trim() === "")
              ).length;
              return (
                <div className="card" style={{ padding: 36 }}>
                  <WorkflowHeader step="3" icon="shield" title="التحقق من نوع الاستبيان" description="راجع النتيجة المختصرة وافتح التفاصيل فقط إذا احتجت تعديلًا." />

                  {/* Detected schema card */}
                  <div style={{
                    background: detectedAutoType ? "rgba(26,188,156,.10)" : "rgba(255,193,7,.07)",
                    border: `1.5px solid ${detectedAutoType ? "rgba(26,188,156,.4)" : "rgba(255,193,7,.4)"}`,
                    borderRadius: 16, padding: "20px 24px", marginBottom: 24,
                    display: "flex", alignItems: "center", gap: 20,
                  }}>
                    <div style={{ fontSize: 52, lineHeight: 1 }}>{s?.icon ?? "📋"}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{
                        display: "inline-block", padding: "3px 12px", borderRadius: 20, fontSize: 11,
                        fontWeight: 700, marginBottom: 8,
                        background: detectedAutoType ? "rgba(26,188,156,.25)" : "rgba(255,193,7,.2)",
                        color: detectedAutoType ? "#1abc9c" : "#ffd54f",
                      }}>
                        {detectedAutoType ? "✓ تم الاكتشاف التلقائي" : "⚠ لم يتم الاكتشاف — تم اختيار الافتراضي"}
                      </div>
                      <div style={{ color: "#fff", fontSize: 20, fontWeight: 900, marginBottom: 4 }}>
                        {s?.label}
                      </div>
                      <div style={{ color: "rgba(255,255,255,.55)", fontSize: 13, display: "flex", gap: 16, flexWrap: "wrap" }}>
                        <span>{s?.scale.type === "likert-5" ? "مقياس 5 درجات" : "مقياس 3 درجات"}</span>
                        <span>·</span>
                        <span>{s?.axes.length} محور</span>
                        <span>·</span>
                        <span>{totalQ} سؤال</span>
                      </div>
                    </div>
                  </div>

                  {/* File Health Summary */}
                  <div style={{ color: "rgba(255,255,255,.55)", fontSize: 12, fontWeight: 700,
                                marginBottom: 10, letterSpacing: .3 }}>
                    ملخص حالة الملف
                  </div>
                  <div style={{
                    display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 24,
                  }}>
                    {[
                      { label: "الملف", value: singleFile?.name ?? "—" },
                      { label: "عدد الاستجابات", value: `${dataRowCount} صف` },
                      { label: "عدد الأعمدة", value: `${colHeaders.length} عمود` },
                      { label: "صفوف فارغة", value: `${blankRowCount} صف`, warn: blankRowCount > 0 },
                    ].map((item, i) => (
                      <div key={i} className={`stat-chip ${item.warn ? "warn" : ""}`}>
                        <div className="stat-chip-value" style={{
                          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                        }}>
                          {item.value}
                        </div>
                        <div className="stat-chip-label">{item.label}</div>
                      </div>
                    ))}
                  </div>
                  {blankRowCount > 0 && (
                    <div style={{
                      background: "rgba(255,193,7,.08)", border: "1px solid rgba(255,193,7,.3)",
                      borderRadius: 12, padding: "10px 16px", marginBottom: 24,
                      color: "#ffd54f", fontSize: 12.5,
                    }}>
                      ⚠ تم رصد {blankRowCount} صف فارغ بالكامل — لن يتم استبعاده تلقائياً وسيُحتسب ضمن عدد المشاركين. يُفضّل حذفه يدوياً من خطوة "معاينة البيانات" التالية.
                    </div>
                  )}

                  <CollapsibleSection title="تفاصيل الأعمدة المكتشفة" defaultOpen={false}
                    badge={<span style={{ color: "rgba(255,255,255,.4)", fontSize: 11 }}>{colHeaders.length} عمود</span>}>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                      {colHeaders.slice(0, 9).map((h, i) => (
                        <span key={i} style={{
                          background: "rgba(255,255,255,.07)", border: "1px solid rgba(255,255,255,.1)",
                          borderRadius: 8, padding: "4px 10px", fontSize: 11.5,
                          color: "rgba(255,255,255,.75)",
                        }}>{String(h)}</span>
                      ))}
                      {colHeaders.length > 9 && (
                        <span style={{
                          background: "rgba(26,188,156,.12)", border: "1px solid rgba(26,188,156,.25)",
                          borderRadius: 8, padding: "4px 10px", fontSize: 11.5, color: "#1abc9c",
                        }}>+{colHeaders.length - 9} أخرى</span>
                      )}
                    </div>
                  </CollapsibleSection>

                  <CollapsibleSection title="تغيير نوع الاستبيان يدويًا" defaultOpen={false}
                    badge={<span style={{ color: "rgba(255,255,255,.4)", fontSize: 11 }}>عند الحاجة فقط</span>}>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 13, fontWeight: 700,
                                  marginBottom: 12 }}>
                      النوع غير صحيح؟ اختر يدوياً:
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      {Object.values(allSchemas()).map(sc => (
                        <button key={sc.id}
                          onClick={() => setSurveyType(sc.id)}
                          style={{
                            padding: "8px 18px", borderRadius: 10, cursor: "pointer",
                            fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 13,
                            transition: "all .2s",
                            border: `2px solid ${surveyType === sc.id ? "#1abc9c" : "rgba(255,255,255,.15)"}`,
                            background: surveyType === sc.id ? "rgba(26,188,156,.15)" : "rgba(255,255,255,.04)",
                            color: surveyType === sc.id ? "#1abc9c" : "rgba(255,255,255,.6)",
                          }}>
                          {sc.icon ?? "📋"} {sc.label}
                        </button>
                      ))}
                    </div>
                  </CollapsibleSection>

                  <InlineNotice text={error} style={{ marginBottom: 16 }} />

                  <div style={{ display: "flex", justifyContent: "center", gap: 12 }}>
                    <button className="btn btn-ghost" onClick={() => {
                      setStep(1); setRawRows(null); setSingleFile(null);
                      setDetectedAutoType(null); setError("");
                    }}>
                      → إعادة الرفع
                    </button>
                    <button className="btn btn-primary" style={{ fontSize: 16, padding: "13px 36px" }}
                      disabled={processing}
                      onClick={handleValidationConfirm}>
                      {processing
                        ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري التحضير…</>
                        : "✓ تأكيد ومتابعة ←"}
                    </button>
                  </div>
                </div>
              );
            })()}

            {/* ══ STEP 3 — Data Preview (annual) ══ */}
            {step === 3 && isAnnual && singleAllRows && singleSchema && (() => {
              const filteredRows = computeFilteredRows(
                singleAllRows, singleRemoved, singleMetaCols, singleFilters
              );
              const liveResult = filteredRows.length
                ? analyzeRows(singleHeaders, filteredRows, singleSchema, singleMetaCols)
                : null;
              return (
                <div>
                  {/* Live analysis dashboard — stays at the top for a quick sanity check
                      while trimming rows below, with Next/Previous right underneath it so
                      the user never has to scroll past the row table to move on. */}
                  <div className="card" style={{ padding: 28, marginBottom: 18 }}>
                    <WorkflowHeader step="4" icon="table" title="معاينة البيانات" description="راجع الملخص، وافتح جدول الصفوف فقط للحذف أو التصفية." />
                    {liveResult ? (
                      <QuickStatsRow result={liveResult} />
                    ) : (
                      <InlineNotice style={{ justifyContent: "center", padding: 20 }}>
                        لا توجد صفوف بعد تطبيق الفلاتر — أزل بعض الفلاتر للمتابعة.
                      </InlineNotice>
                    )}

                    <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 26 }}>
                      <button className="btn btn-ghost" onClick={() => setStep(2)}>→ السابق</button>
                      <button className="btn btn-primary"
                        disabled={!liveResult}
                        style={{ opacity: liveResult ? 1 : 0.5 }}
                        onClick={() => { setSingleResult(liveResult); setStep(4); }}>
                        متابعة ←
                      </button>
                    </div>
                  </div>

                  {/* Row-level data preview — collapsible, sits below the dashboard */}
                  <CollapsibleSection
                    title="🔍 معاينة البيانات"
                    defaultOpen={false}
                    badge={
                      <span style={{ color: "rgba(255,255,255,.4)", fontSize: 12, fontWeight: 400 }}>
                        راجع البيانات واحذف الصفوف غير المرغوبة
                      </span>
                    }
                  >
                    <DataPreviewTable
                      headers={singleHeaders}
                      rows={singleAllRows}
                      removed={singleRemoved}
                      onToggleRow={idx => setSingleRemoved(prev => {
                        const next = new Set(prev);
                        if (next.has(idx)) next.delete(idx); else next.add(idx);
                        return next;
                      })}
                      deptIdx={singleMetaCols?.deptIdx}
                      degreeIdx={singleMetaCols?.degreeIdx}
                      filters={singleFilters}
                      onFilterChange={setSingleFilters}
                    />
                  </CollapsibleSection>
                </div>
              );
            })()}

            {/* ══ STEP 4 — Metadata ══ */}
            {step === 4 && (
              <div>
                <MetadataForm meta={meta} onChange={setMeta} />
                <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 24 }}>
                  <button className="btn btn-ghost" onClick={() => setStep(isAnnual ? 3 : 1)}>→ السابق</button>
                  <button className="btn btn-primary" onClick={() => setStep(5)}>عرض النتائج ←</button>
                </div>
              </div>
            )}

            {/* ══ STEP 5 — Results ══ */}
            {step === 5 && (
              <div>
                <WorkflowHeader step="6" icon="chart" title="النتائج والتصدير" description="راجع الملخص ثم نزّل Word أو PDF أو افتح المعاينة المحسّنة." />
                {isAnnual && singleResult && (
                  <div>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 14, marginBottom: 18, textAlign: "center" }}>
                      <span style={{ color: "#1abc9c", fontWeight: 700 }}>{schema.label}</span>
                      {meta.program ? ` — ${meta.program}` : ""}
                      {" | "}<span style={{ color: "#d6eaf8" }}>{meta.year}</span>
                    </div>
                    <CollapsibleSection title="📊 ملخص النتائج" defaultOpen={true}>
                      <ResultsPreview result={singleResult} />
                    </CollapsibleSection>
                  </div>
                )}

                {!isAnnual && comparison && (
                  <div>
                    <div style={{ color: "rgba(255,255,255,.6)", fontSize: 14, marginBottom: 18, textAlign: "center" }}>
                      <span style={{ color: "#1abc9c", fontWeight: 700 }}>{schema.label}</span>
                      {meta.program ? ` — ${meta.program}` : ""}
                      {" | مقارنة "}
                      {comparison.slots.map(s => s.year).join(" / ")}
                    </div>
                    <CollapsibleSection title="📊 ملخص النتائج" defaultOpen={true}>
                      <ComparisonPreview comparison={comparison} />
                    </CollapsibleSection>
                  </div>
                )}

                <div style={{ textAlign: "center", padding: "28px 0", display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "center" }}>
                    {isAnnual ? (
                      <button className="btn btn-blue" style={{ fontSize: 16, padding: "14px 36px" }}
                        disabled={processing} onClick={downloadAnnual}>
                        {processing
                          ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="white" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري الإنشاء…</>
                          : "📄 تحميل التقرير (Word)"}
                      </button>
                    ) : (
                      <button className="btn btn-blue" style={{ fontSize: 16, padding: "14px 36px" }}
                        disabled={processing} onClick={downloadComparison}>
                        📄 تحميل تقرير المقارنة (Word)
                      </button>
                    )}
                    {isAnnual && singleResult && (
                      <>
                        <button className="btn btn-ghost" style={{ fontSize: 16, padding: "14px 36px" }}
                          disabled={pdfGenerating} onClick={downloadAnnualPdf}>
                          {pdfGenerating
                            ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="#e8f0fe" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري إنشاء PDF…</>
                            : "🖨️ تحليل وتحميل PDF"}
                        </button>
                        <button className="btn btn-primary" style={{ fontSize: 16, padding: "14px 36px" }}
                          disabled={pdfGenerating} onClick={uploadAnnualPdf}>
                          {pdfGenerating
                            ? <><svg className="spin" width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6" fill="none" stroke="currentColor" strokeWidth="2" strokeDasharray="28 10"/></svg> جاري إنشاء PDF…</>
                            : "☁️ تحليل ورفع PDF"}
                        </button>
                      </>
                    )}
                    {isAnnual && singleResult && (
                      <button
                        className="btn btn-ghost"
                        style={{ fontSize: 16, padding: "14px 36px", background: "rgba(30,58,138,.25)", borderColor: "rgba(37,99,235,.45)", color: "#93c5fd" }}
                        onClick={() => setShowEnhancedView(true)}
                      >
                        ✨ معاينة محسّنة
                      </button>
                    )}
                  </div>
                  {isAnnual && singleResult && (
                    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "rgba(255,255,255,.6)", cursor: "pointer" }}>
                      <input
                        type="checkbox"
                        checked={pdfIncludeCharts}
                        onChange={e => setPdfIncludeCharts(e.target.checked)}
                        style={{ width: 15, height: 15, cursor: "pointer" }}
                      />
                      تضمين الرسوم البيانية في تقرير الـ PDF (تعطيلها يقلل عدد الصفحات)
                    </label>
                  )}
                  <div style={{ color: "rgba(255,255,255,.35)", fontSize: 12 }}>
                    Word: تقرير كامل مع جداول ومحاور • PDF: تقرير مصمَّم بنفس الهوية البصرية جاهز للطباعة
                  </div>
                </div>
              </div>
            )}
          </>
        )}
          </div>
        </div>
      </div>

      {/* ── Tutorial overlay ── */}
      {showTutorial && <TutorialOverlay onClose={() => setShowTutorial(false)} />}

      {/* ── Processing steps overlay ── */}
      {procSteps && <ProcessingOverlay steps={procSteps} filename={procFile} />}

      {/* ── Drive multi-report progress ── */}
      {driveBatchState && <DriveBatchOverlay state={driveBatchState} onClose={() => setDriveBatchState(null)} onCancel={cancelDriveBatch} />}

      {/* ── Branded PDF generation overlay ── */}
      {pdfGenerating && <LoadingOverlay message="جاري إنشاء تقرير PDF…" progress={pdfProgress} cancelling={pdfCancelRequested}
        onCancel={() => { pdfCancelRef.current = true; setPdfCancelRequested(true); }} />}

      {/* ── Enhanced Report View overlay (isolated, zero impact on existing flow) ── */}
      {showEnhancedView && singleResult && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 9000,
          background: "#f1f5f9",
          overflowY: "auto",
        }}>
          {/* Close bar */}
          <div style={{
            position: "sticky", top: 0, zIndex: 9001,
            background: "#1e3a8a",
            padding: "10px 24px",
            display: "flex", alignItems: "center", justifyContent: "space-between",
            boxShadow: "0 2px 12px rgba(0,0,0,.25)",
          }}>
            <span style={{ color: "#fff", fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 15 }}>
              ✨ معاينة التقرير المحسّن
            </span>
            <button
              onClick={() => setShowEnhancedView(false)}
              style={{
                background: "rgba(255,255,255,.15)", border: "1px solid rgba(255,255,255,.3)",
                color: "#fff", borderRadius: 8, padding: "7px 20px",
                fontFamily: "'Cairo',sans-serif", fontWeight: 700, fontSize: 14, cursor: "pointer",
              }}
            >
              ✕ إغلاق
            </button>
          </div>
          <EnhancedReportView result={singleResult} meta={meta} settings={settings} />
        </div>
      )}

      {/* AI Chat — floating panel, always visible */}
      {showAiChat && <AiChat
        currentResult={singleResult}
        aiSettings={aiSettings}
        docSettings={settings}
        onAnalysisComplete={(result, type) => {
          setSingleResult(result);
          if (type) setSurveyType(type);
        }}
      />}
    </div>
  );
}
