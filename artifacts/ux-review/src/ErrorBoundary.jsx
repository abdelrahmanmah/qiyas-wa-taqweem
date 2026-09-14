import React from "react";

export default class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error("Application render failed", error, info);
  }

  render() {
    if (!this.state.error) return this.props.children;

    return (
      <main dir="rtl" style={{ minHeight: "100vh", display: "grid", placeItems: "center", padding: 24, background: "linear-gradient(145deg,#091a2d,#0b352f)", fontFamily: "Cairo, sans-serif", color: "#fff" }}>
        <section role="alert" style={{ width: "min(520px,100%)", padding: 28, borderRadius: 20, background: "rgba(255,255,255,.08)", border: "1px solid rgba(255,255,255,.14)", textAlign: "right" }}>
          <div aria-hidden="true" style={{ width: 48, height: 48, display: "grid", placeItems: "center", borderRadius: 14, marginBottom: 16, color: "#ffd6d2", background: "rgba(231,76,60,.16)", fontSize: 24 }}>!</div>
          <h1 style={{ margin: "0 0 8px", fontSize: 22 }}>تعذّر عرض الصفحة</h1>
          <p style={{ margin: "0 0 20px", color: "rgba(255,255,255,.72)", lineHeight: 1.8, fontSize: 14 }}>حدث خطأ غير متوقع. أعد تحميل الصفحة، وإذا استمر الخطأ تواصل مع مسؤول النظام.</p>
          <button type="button" onClick={() => window.location.reload()} style={{ minHeight: 44, padding: "10px 22px", border: 0, borderRadius: 999, cursor: "pointer", color: "#fff", background: "linear-gradient(135deg,#1abc9c,#16a085)", fontFamily: "inherit", fontWeight: 800 }}>إعادة تحميل الصفحة</button>
        </section>
      </main>
    );
  }
}
