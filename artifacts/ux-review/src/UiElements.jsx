export function GoogleDriveIcon({ size = 22, className = "" }) {
  return (
    <svg
      className={className}
      width={size}
      height={size}
      viewBox="0 0 48 48"
      aria-hidden="true"
      focusable="false"
      style={{ display: "block", flex: `0 0 ${size}px` }}
    >
      <path fill="#FFC107" d="M17 6h14l14 24H31L17 6Z" />
      <path fill="#1976D2" d="M10 42 3 30h28l7 12H10Z" />
      <path fill="#4CAF50" d="M3 30 17 6l7 12-7 12H3Z" />
    </svg>
  );
}

export function InlineNotice({ children, text, className = "", style }) {
  const content = children ?? text;
  if (!content) return null;

  return (
    <div
      className={`inline-notice ${className}`.trim()}
      role="alert"
      style={{
        display: "flex",
        alignItems: "flex-start",
        gap: 10,
        padding: "11px 14px",
        borderRadius: 11,
        color: "#dce8f7",
        background: "rgba(96,165,250,.08)",
        border: "1px solid rgba(147,197,253,.18)",
        fontSize: 12.5,
        fontWeight: 600,
        lineHeight: 1.75,
        textAlign: "right",
        ...style,
      }}
    >
      <span
        aria-hidden="true"
        style={{
          width: 20,
          height: 20,
          flex: "0 0 20px",
          display: "grid",
          placeItems: "center",
          marginTop: 1,
          borderRadius: 7,
          color: "#bfdbfe",
          background: "rgba(96,165,250,.14)",
          fontSize: 12,
          fontWeight: 900,
        }}
      >
        !
      </span>
      <span style={{ minWidth: 0 }}>{content}</span>
    </div>
  );
}
