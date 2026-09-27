import { useState } from "react";
import CourseTemplateTool from "./CourseTemplateTool.jsx";
import SurveyParticipationTool from "./SurveyParticipationTool.jsx";
import CourseSplitter from "./CourseSplitter.jsx";
import PdfRecommendationReviewer from "./PdfRecommendationReviewer.jsx";
import { QualityIcon } from "./UiElements.jsx";

const TABS = [
  { key: "guide", label: "نظرة عامة", hint: "ابدأ من هنا", icon: "compass" },
  { key: "template", label: "بيانات المقررات", hint: "القالب والتوزيع", icon: "sheet" },
  { key: "participation", label: "أداء الاستبيانات", hint: "مراجعة المشاركة", icon: "chart" },
  { key: "splitter", label: "تقسيم التقييم", hint: "ملفات المقررات", icon: "split" },
  { key: "recs", label: "مراجعة التوصيات", hint: "تقارير PDF", icon: "document" },
];

function HubIcon({ name, size = 20 }) {
  const iconName = name === "document" ? "report" : name === "arrow" ? "arrowForward" : name;
  return <QualityIcon name={iconName} size={size} />;
}

export default function CourseEvaluationHub() {
  const [activeTab, setActiveTab] = useState("guide");
  const active = TABS.find(tab => tab.key === activeTab);
  return (
    <div className="course-hub">
      <style>{COURSE_HUB_CSS}</style>
      <section className="course-hub-hero">
        <div className="course-hub-copy">
          <span className="course-hub-kicker">مركز العمليات الأكاديمية</span>
          <h1>دورة تقييم المقررات</h1>
          <p>مسار واحد من تجهيز بيانات المقررات ومتابعة المشاركة، حتى تقسيم النتائج ومراجعة التوصيات.</p>
        </div>
        <div className="course-hub-visual" aria-hidden="true"><HubIcon name="route" size={27}/><strong>4</strong><small>أدوات مترابطة</small></div>
      </section>

      <nav className="course-hub-nav" aria-label="أدوات تقييم المقررات">
        {TABS.map(tab => <button key={tab.key} className={`course-hub-tab ${activeTab === tab.key ? "active" : ""}`} onClick={() => setActiveTab(tab.key)} aria-current={activeTab === tab.key ? "page" : undefined}>
          <span className="course-hub-tab-icon"><HubIcon name={tab.icon}/></span><span><strong>{tab.label}</strong><small>{tab.hint}</small></span>
        </button>)}
      </nav>

      {activeTab !== "guide" && <div className="course-tool-heading">
        <button onClick={() => setActiveTab("guide")}><HubIcon name="arrow" size={16}/> العودة للمسار</button>
        <div><span>الأداة الحالية</span><strong>{active?.label}</strong></div>
      </div>}

      <main className="course-hub-content">
        {activeTab === "guide" && <GuideView onJump={setActiveTab} />}
        {activeTab === "template" && <CourseTemplateTool />}
        {activeTab === "participation" && <SurveyParticipationTool />}
        {activeTab === "splitter" && <CourseSplitter />}
        {activeTab === "recs" && <PdfRecommendationReviewer />}
      </main>
    </div>
  );
}

const STEPS = [
  { n: 1, tab: "template", icon: "sheet", title: "جهّز بيانات المقررات", summary: "أنشئ القالب ووزّع المراجعين وتأكد من اكتمال البيانات.", body: "حمّل القالب الفارغ أو ارفع نسخة موجودة، ثم راجع توزيع الأقسام وعبء العمل والبيانات الناقصة قبل إرسال الاستبيان.", cta: "فتح بيانات المقررات" },
  { n: 2, tab: "participation", icon: "chart", title: "راجع مستوى المشاركة", summary: "اعرف المقررات المكتملة والمتعثرة قبل اعتماد النتائج.", body: "ارفع تقرير أداء الاستبيانات ليتم تصنيف المقررات تلقائياً حسب عدد المشاركين، مع إمكانية تصدير تقرير Excel لكل فئة.", cta: "فتح أداء الاستبيانات" },
  { n: 3, tab: "splitter", icon: "split", title: "قسّم ملف التقييم", summary: "حوّل ملف UMIS المجمّع إلى ملف مستقل لكل مقرر.", body: "ارفع الملف المجمّع وقائمة المواد أو توزيع الأقسام عند الحاجة، وستحصل على ملفات منظمة مع توضيح المقررات الناقصة.", cta: "فتح تقسيم التقييم" },
  { n: 4, tab: "recs", icon: "document", title: "راجع التوصيات", summary: "راجع توصيات تقارير PDF وحدد المهم منها بسرعة.", body: "اعرض صفحة التوصية من كل تقرير، حدّد التقارير التي تحتوي على توصية فعلية، ثم صدّر القائمة إلى Excel.", cta: "فتح مراجعة التوصيات" },
];

function GuideView({ onJump }) {
  return <div className="course-guide">
    <div className="course-guide-heading"><div><span>مسار العمل المقترح</span><h2>أنجز دورة التقييم خطوة بخطوة</h2><p>يمكنك الدخول لأي أداة مباشرة، أو اتباع الترتيب لضمان اكتمال الدورة.</p></div><span className="course-guide-count">4 خطوات</span></div>
    <div className="course-step-grid">{STEPS.map((step, index) => <article className="course-step-card" key={step.n} style={{ animationDelay: `${index * 60}ms` }}>
      <div className="course-step-top"><span className="course-step-icon"><HubIcon name={step.icon}/></span><span className="course-step-number">{String(step.n).padStart(2, "0")}</span></div>
      <h3>{step.title}</h3><p>{step.summary}</p>
      <details><summary>تفاصيل الخطوة</summary><div>{step.body}</div></details>
      <button className="course-step-action" onClick={() => onJump(step.tab)}>{step.cta}<HubIcon name="arrow" size={16}/></button>
    </article>)}</div>
    <div className="course-guide-note"><span><HubIcon name="route"/></span><div><strong>بعد الانتهاء</strong><p>استخدم الملفات الناتجة داخل أداة تحليل الاستبيانات لإصدار التقرير النهائي.</p></div></div>
  </div>;
}

const COURSE_HUB_CSS = `
.course-hub{direction:rtl;color:#eef7ff;max-width:1260px;margin:0 auto;padding-bottom:28px}.course-hub-hero{position:relative;overflow:hidden;display:flex;align-items:center;justify-content:space-between;gap:28px;min-height:185px;padding:30px 34px;margin-bottom:16px;border:1px solid rgba(94,234,212,.16);border-radius:24px;background:linear-gradient(120deg,rgba(14,69,83,.76),rgba(15,48,76,.7));box-shadow:0 18px 45px rgba(2,12,27,.15)}.course-hub-hero:after{content:'';position:absolute;left:-70px;top:-110px;width:280px;height:280px;border-radius:50%;background:rgba(59,130,246,.1)}.course-hub-copy{position:relative;z-index:1}.course-hub-kicker{display:inline-flex;padding:5px 11px;border-radius:999px;color:#80efda;background:rgba(26,188,156,.1);border:1px solid rgba(94,234,212,.18);font-size:10px;font-weight:800}.course-hub-copy h1{margin:13px 0 7px;font-size:clamp(24px,3vw,37px);line-height:1.25}.course-hub-copy p{max-width:700px;margin:0;color:rgba(225,241,251,.58);font-size:12.5px;line-height:1.9}.course-hub-visual{position:relative;z-index:1;flex:0 0 145px;min-height:112px;display:grid;place-items:center;padding:16px;border-radius:20px;background:rgba(4,22,37,.3);border:1px solid rgba(255,255,255,.09)}.course-hub-visual svg{color:#70e8d2}.course-hub-visual strong{font-size:28px;line-height:1}.course-hub-visual small{color:rgba(255,255,255,.43);font-size:9px}
.course-hub-nav{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:7px;padding:8px;margin-bottom:18px;border-radius:18px;background:rgba(255,255,255,.035);border:1px solid rgba(255,255,255,.08)}.course-hub-tab{min-width:0;display:flex;align-items:center;gap:9px;padding:10px;border:1px solid transparent;border-radius:12px;background:transparent;color:rgba(238,247,255,.48);font-family:inherit;text-align:right;cursor:pointer;transition:.2s}.course-hub-tab:hover{color:#fff;background:rgba(255,255,255,.045)}.course-hub-tab.active{color:#eafffa;background:linear-gradient(135deg,rgba(26,188,156,.18),rgba(40,116,166,.1));border-color:rgba(94,234,212,.23)}.course-hub-tab-icon{flex:0 0 32px;width:32px;height:32px;display:grid;place-items:center;border-radius:9px;background:rgba(255,255,255,.055)}.course-hub-tab.active .course-hub-tab-icon{color:#72ead4;background:rgba(26,188,156,.13)}.course-hub-tab strong,.course-hub-tab small{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.course-hub-tab strong{font-size:10.5px}.course-hub-tab small{margin-top:2px;color:rgba(255,255,255,.3);font-size:8px}
.course-tool-heading{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:11px 14px;margin-bottom:14px;border-radius:14px;background:rgba(26,188,156,.055);border:1px solid rgba(94,234,212,.12)}.course-tool-heading div span,.course-tool-heading div strong{display:block}.course-tool-heading div span{color:#65d8c4;font-size:8px}.course-tool-heading div strong{font-size:13px}.course-tool-heading button{display:flex;align-items:center;gap:7px;border:0;background:transparent;color:rgba(255,255,255,.48);font-family:inherit;font-size:10px;cursor:pointer}.course-tool-heading button:hover{color:#8cf2dc}
.course-guide-heading{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin:4px 2px 16px}.course-guide-heading span{color:#69dfca;font-size:9px;font-weight:800}.course-guide-heading h2{font-size:19px;margin:5px 0}.course-guide-heading p{margin:0;color:rgba(255,255,255,.4);font-size:10px}.course-guide-heading .course-guide-count{padding:6px 10px;border-radius:999px;background:rgba(26,188,156,.08);border:1px solid rgba(94,234,212,.15);white-space:nowrap}.course-step-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.course-step-card{min-height:245px;display:flex;flex-direction:column;padding:19px;border:1px solid rgba(255,255,255,.085);border-radius:19px;background:linear-gradient(145deg,rgba(255,255,255,.055),rgba(255,255,255,.025));animation:fadeInUp .35s ease both;transition:.2s}.course-step-card:hover{transform:translateY(-3px);border-color:rgba(94,234,212,.2);box-shadow:0 16px 34px rgba(2,12,27,.16)}.course-step-top{display:flex;align-items:center;justify-content:space-between}.course-step-icon{width:40px;height:40px;display:grid;place-items:center;border-radius:12px;color:#78ead5;background:rgba(26,188,156,.1);border:1px solid rgba(94,234,212,.14)}.course-step-number{color:rgba(255,255,255,.16);font-size:22px;font-weight:900}.course-step-card h3{margin:15px 0 5px;font-size:14.5px}.course-step-card>p{margin:0 0 10px;color:rgba(233,246,255,.48);font-size:10.5px;line-height:1.8}.course-step-card details{font-size:9.5px;color:rgba(255,255,255,.42)}.course-step-card summary{width:max-content;color:#6edeca;cursor:pointer;font-weight:700}.course-step-card details div{margin-top:7px;padding:9px 10px;border-radius:9px;background:rgba(2,16,29,.24);line-height:1.8}.course-step-action{display:flex;align-items:center;justify-content:space-between;width:100%;margin-top:auto;padding:10px 12px;border-radius:10px;border:1px solid rgba(94,234,212,.16);background:rgba(26,188,156,.08);color:#bafff1;font-family:inherit;font-size:10px;font-weight:800;cursor:pointer;transition:.18s}.course-step-action:hover{background:rgba(26,188,156,.17);border-color:rgba(94,234,212,.3)}.course-guide-note{display:flex;align-items:center;gap:12px;margin-top:13px;padding:14px 16px;border-radius:15px;background:rgba(59,130,246,.06);border:1px solid rgba(96,165,250,.12)}.course-guide-note>span{width:35px;height:35px;display:grid;place-items:center;flex:0 0 35px;border-radius:10px;color:#82c9ff;background:rgba(59,130,246,.1)}.course-guide-note strong{font-size:11px}.course-guide-note p{margin:2px 0 0;color:rgba(255,255,255,.4);font-size:9.5px}
@media(max-width:900px){.course-hub-nav{grid-template-columns:repeat(3,minmax(0,1fr))}}@media(max-width:640px){.course-hub{padding-bottom:16px}.course-hub-hero{min-height:auto;padding:21px;border-radius:19px}.course-hub-copy p{font-size:11px}.course-hub-visual{display:none}.course-hub-nav{display:flex;overflow-x:auto;scroll-snap-type:x mandatory;padding:6px;margin-inline:-2px;scrollbar-width:none}.course-hub-nav::-webkit-scrollbar{display:none}.course-hub-tab{flex:0 0 145px;scroll-snap-align:start}.course-hub-tab small{display:none}.course-step-grid{grid-template-columns:1fr}.course-step-card{min-height:225px;padding:17px}.course-guide-heading{align-items:flex-end}.course-guide-heading h2{font-size:17px}.course-guide-heading p{display:none}.course-tool-heading{position:sticky;top:76px;z-index:9;backdrop-filter:blur(15px);background:rgba(10,37,54,.92)}}@media(max-width:380px){.course-hub-hero{padding:18px 15px}.course-hub-copy h1{font-size:22px}.course-guide-count{display:none}.course-hub-tab{flex-basis:132px}}
`;
