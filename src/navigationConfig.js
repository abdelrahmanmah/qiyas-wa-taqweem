const HOME_TOOL = { id: "home", icon: "home", label: "الرئيسية", tag: "لوحة البداية", navGroup: "main" };

export const HOME_TOOL_CARDS = [
  {
    id: "summary", icon: "dashboard", color: "#fbbf24", featured: true, requiresGoogle: true, navGroup: "main",
    title: "ملخص الاستبيانات", label: "ملخص الاستبيانات", tag: "السنة والفصل",
    description: "اعرف بسرعة ما تم تحليله وما لم يُحلل بعد، مع عدد المشاركين في كل استبيان.",
  },
  {
    id: "analytics", icon: "analytics", color: "#5eead4", requiresGoogle: true, navGroup: "main",
    title: "تحليل نتائج موجودة", label: "تحليل الاستبيانات", tag: "Excel أو Google Drive",
    description: "ابدأ بملف واحد أو مجموعة ملفات، ودع النظام يتعرف على نوع الاستبيان ثم أنشئ تقرير Word أو PDF.",
  },
  {
    id: "semester", icon: "calendar", color: "#60a5fa", requiresGoogle: true, navGroup: "main",
    title: "إنشاء استبيانات الفصل الدراسي", label: "استبيانات الفصل الدراسي", tag: "Google Forms",
    description: "اختر القوالب والسنة والفصل، وأنشئ كل استبيانات الفصل منظمة تلقائياً على Google Drive.",
  },
  {
    id: "drive", icon: "drive", color: "#34d399", requiresGoogle: true, navGroup: "main",
    title: "مكتبة Drive", label: "مكتبة Drive", tag: "تقارير واستبيانات",
    description: "تصفح تقارير الجودة والاستبيانات من مكان واحد، مع عدادات فورية وبحث وفلاتر دقيقة ووصول سريع للملفات.",
  },
  {
    id: "courses", icon: "cap", color: "#a78bfa", navGroup: "evaluation",
    title: "تقييم المقررات", label: "تقييم المقررات", tag: "دورة تقييم متكاملة",
    description: "جهّز بيانات المقررات، تابع نسب المشاركة، قسّم ملفات التقييم وراجع التوصيات من مساحة واحدة.",
  },
  {
    id: "exam-paper", icon: "exam", color: "#fb7185", requiresGoogle: true, navGroup: "evaluation",
    title: "تقييم الورقة الامتحانية", label: "تقييم الورقة الامتحانية", tag: "الشكل والاستيفاء",
    description: "حمّل قالب المقررات، قيّم استيفاء شكل الورقة ومخرجات التعلم، ثم أنشئ التقرير أو ارفعه على Drive.",
  },
];

export const WORKSPACE_NAV_ITEMS = [
  HOME_TOOL,
  ...HOME_TOOL_CARDS.filter(tool => tool.navGroup === "main"),
];

export const EVALUATION_NAV_ITEMS = HOME_TOOL_CARDS.filter(tool => tool.navGroup === "evaluation");

export const AUXILIARY_NAV_ITEMS = [
  { id: "surveys", icon: "surveys", label: "تصميم الاستبيانات", tag: "قوالب وأسئلة", navGroup: "management" },
  { id: "settings", icon: "settings", label: "الإعدادات", tag: "الهوية والتقارير", navGroup: "system" },
];

export const SEMESTER_TAB_BY_TOOL = {
  semester: "home",
  "semester-create": "generate",
  "semester-dashboard": "dashboard",
};

const TOOL_BY_ID = Object.fromEntries([
  HOME_TOOL,
  ...HOME_TOOL_CARDS,
  ...AUXILIARY_NAV_ITEMS,
].map(tool => [tool.id, tool]));

export function isSemesterTool(tool) {
  return Object.prototype.hasOwnProperty.call(SEMESTER_TAB_BY_TOOL, tool);
}

export function canonicalToolId(tool) {
  return isSemesterTool(tool) ? "semester" : tool;
}

export function getToolLabel(tool) {
  return TOOL_BY_ID[canonicalToolId(tool)]?.label ?? "بوابة لجنة القياس والتقويم";
}

export function getToolTag(tool) {
  return TOOL_BY_ID[canonicalToolId(tool)]?.tag ?? "بوابة لجنة القياس والتقويم";
}

export function needsGoogleConnection(tool) {
  return Boolean(TOOL_BY_ID[canonicalToolId(tool)]?.requiresGoogle);
}
