import { useState } from "react";
import CourseTemplateTool from "./CourseTemplateTool.jsx";
import SurveyParticipationTool from "./SurveyParticipationTool.jsx";
import CourseSplitter from "./CourseSplitter.jsx";
import PdfRecommendationReviewer from "./PdfRecommendationReviewer.jsx";

// ── Course Evaluation Hub ────────────────────────────────────────────────
// Groups every tool used to run a course-evaluation cycle in one place:
// the course/reviewer assignment template, the survey-participation report
// analyzer, the UMIS merged-file splitter (unchanged, just relocated here
// from its own top-level tab), and the PDF recommendation reviewer — plus a
// usage guide tying them together in the order they're actually used.

const TABS = [
  { key: "guide", label: "📖 دليل الاستخدام" },
  { key: "template", label: "📋 قالب بيانات المقررات" },
  { key: "participation", label: "📊 أداء الاستبيانات" },
  { key: "splitter", label: "🧩 تقسيم التقييم" },
  { key: "recs", label: "📄 مراجعة التوصيات" },
];

export default function CourseEvaluationHub() {
  const [activeTab, setActiveTab] = useState("guide");

  return (
    <div>
      <div style={{ marginBottom: 18 }}>
        <div style={{
          display: "inline-block", fontSize: 11, color: "#1abc9c",
          background: "rgba(26,188,156,.12)", border: "1px solid rgba(26,188,156,.35)",
          padding: "3px 12px", borderRadius: 999, marginBottom: 10, letterSpacing: .3,
        }}>ERU · تقييم المقررات</div>
        <div style={{ color: "#fff", fontSize: 22, fontWeight: 900 }}>📚 تقييم المقررات</div>
        <div style={{ color: "rgba(255,255,255,.55)", fontSize: 13, marginTop: 4 }}>
          كل أدوات تقييم المقررات في مكان واحد — من تجهيز بيانات المراجعين لحد تصدير التوصيات.
        </div>
      </div>

      <div className="card" style={{ padding: 10, marginBottom: 22, display: "flex", gap: 8, flexWrap: "wrap" }}>
        {TABS.map(t => (
          <button
            key={t.key}
            className={`btn btn-sm ${activeTab === t.key ? "btn-primary" : "btn-ghost"}`}
            onClick={() => setActiveTab(t.key)}
          >
            {t.label}
          </button>
        ))}
      </div>

      {activeTab === "guide" && <GuideView onJump={setActiveTab} />}
      {activeTab === "template" && <CourseTemplateTool />}
      {activeTab === "participation" && <SurveyParticipationTool />}
      {activeTab === "splitter" && <CourseSplitter />}
      {activeTab === "recs" && <PdfRecommendationReviewer />}
    </div>
  );
}

const STEPS = [
  {
    n: 1, tab: "template", title: "جهّز قالب بيانات المقررات",
    body: "حمّل القالب الفارغ (اسم المقرر، كود المقرر، عضو هيئة التدريس، عضو الهيئة المعاونة، القسم العلمي، هل يوجد لاب، القائم بالمراجعة) واملأه لكل الأقسام — أو ارفع نسخة موجودة وعدّل فيها. استخدم الـ insights (توزيع الأقسام، عبء العمل، البيانات الناقصة) للتأكد إن كل مقرر ليه مراجع ومعاون قبل النزول للاستبيان.",
    cta: "افتح تاب «قالب بيانات المقررات»",
  },
  {
    n: 2, tab: "participation", title: "راجع أداء الاستبيان بعد قفله",
    body: "حمّل تقرير أداء الاستبيانات من السيستم (COURSE_CODE، COURSE_DESCR_EN، NoOfVotes) وارفعه هنا. هتلاقي المواد اتقسّمت أوتوماتيك لـ: طبيعية، مشكوك في انتظامها (أقل من 10 مصوتين ومحتاجة تأكد يدوي)، لم يتم التقييم، ولا يوجد طلاب مسجلين — مع تقرير Excel جاهز لكل فئة.",
    cta: "افتح تاب «أداء الاستبيانات»",
  },
  {
    n: 3, tab: "splitter", title: "قسّم ملف تقييم المقررات المجمّع",
    body: "ملف UMIS بيرجع كل المواد مجمّعة في شيت واحد. ارفعه في أداة التقسيم مع قائمة المواد الأساسية (اختياري) وقائمة توزيع الأقسام (اختياري) — هتاخد ملف Excel مستقل لكل مادة، مقارن بالقائمة الأساسية عشان تعرف أي مادة لسه ناقصة، ومقسّم حسب القسم العلمي.",
    cta: "افتح تاب «تقسيم التقييم»",
  },
  {
    n: 4, tab: "recs", title: "راجع توصيات المراجعين",
    body: "بعد ما المراجعين يرفعوا تقاريرهم كـ PDF لكل مادة، اجمعهم في مجلد واحد وافتحه هنا — هيتم عرض آخر صفحة من كل ملف (مكان التوصية غالبًا) كصورة مصغّرة، حدّد اللي فيها توصية فعلية بضغطة واحدة، وصدّر القائمة كـ Excel بسرعة بدل ما تفتح كل ملف لوحده.",
    cta: "افتح تاب «مراجعة التوصيات»",
  },
  {
    n: 5, tab: null, title: "كمّل التحليل والتقارير النهائية",
    body: "استخدم نتائج الخطوات فوق (القالب، تقرير الأداء، الملفات المقسّمة، قائمة التوصيات) مع باقي أدوات النظام — تحليل الاستبيانات وتوليد تقارير Word — لإتمام دورة تقييم المقررات بالكامل.",
    cta: null,
  },
];

function GuideView({ onJump }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {STEPS.map(s => (
        <div key={s.n} className="card" style={{ padding: 22 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
            <span style={{
              width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,.08)",
              display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "#1abc9c",
            }}>{s.n}</span>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>{s.title}</span>
          </div>
          <p style={{ color: "rgba(255,255,255,.55)", fontSize: 13, lineHeight: 1.8, margin: "0 0 12px" }}>{s.body}</p>
          {s.cta && (
            <button className="btn btn-ghost btn-sm" onClick={() => onJump(s.tab)}>{s.cta} ←</button>
          )}
        </div>
      ))}
    </div>
  );
}
