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

const QUALITY_ICON_PATHS = {
  home: <><path d="M3 11 12 3l9 8"/><path d="M5 10v10h14V10M9 20v-6h6v6"/></>,
  settings: <><circle cx="12" cy="12" r="3"/><path d="M19 15.2a1.8 1.8 0 0 0 .4 2l.1.1-2.2 2.2-.1-.1a1.8 1.8 0 0 0-2-.4 1.8 1.8 0 0 0-1.1 1.7v.2H11v-.2A1.8 1.8 0 0 0 9.8 19a1.8 1.8 0 0 0-2 .4l-.1.1-2.2-2.2.1-.1a1.8 1.8 0 0 0 .4-2A1.8 1.8 0 0 0 4.3 14H4v-3h.3A1.8 1.8 0 0 0 6 9.8a1.8 1.8 0 0 0-.4-2l-.1-.1 2.2-2.2.1.1a1.8 1.8 0 0 0 2 .4A1.8 1.8 0 0 0 11 4.3V4h3v.3A1.8 1.8 0 0 0 15.2 6a1.8 1.8 0 0 0 2-.4l.1-.1 2.2 2.2-.1.1a1.8 1.8 0 0 0-.4 2 1.8 1.8 0 0 0 1.7 1.2h.3v3h-.3a1.8 1.8 0 0 0-1.7 1.2Z"/></>,
  footer: <><path d="M5 3h14v18H5z"/><path d="M8 7h8M8 11h8M8 17h8"/></>,
  report: <><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5M9 12h6M9 16h6"/></>,
  file: <><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5"/></>,
  print: <><path d="M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v8H6z"/></>,
  palette: <><path d="M12 3a9 9 0 0 0 0 18h1.5a2 2 0 0 0 0-4H12a1.5 1.5 0 0 1 0-3h2a7 7 0 0 0-2-11Z"/><circle cx="7.5" cy="10" r="1"/><circle cx="10" cy="6.5" r="1"/><circle cx="15" cy="7.5" r="1"/></>,
  sheet: <><path d="M6 2h9l3 3v17H6z"/><path d="M9 10h6M9 14h6M9 18h4"/></>,
  upload: <><path d="M12 16V4m-5 5 5-5 5 5"/><path d="M5 20h14"/></>,
  download: <><path d="M12 3v12m-5-5 5 5 5-5"/><path d="M5 21h14"/></>,
  plus: <path d="M12 5v14M5 12h14"/>,
  chart: <><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></>,
  user: <><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></>,
  users: <><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.9M16 3.2a4 4 0 0 1 0 7.6"/></>,
  warning: <><path d="M10.3 3.7 2.2 18a2 2 0 0 0 1.7 3h16.2a2 2 0 0 0 1.7-3L13.7 3.7a2 2 0 0 0-3.4 0Z"/><path d="M12 9v4M12 17h.01"/></>,
  search: <><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></>,
  trash: <><path d="M4 7h16M9 7V4h6v3M7 7l1 14h8l1-14M10 11v6M14 11v6"/></>,
  copy: <><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/></>,
  edit: <><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z"/></>,
  folder: <path d="M3 6h7l2 2h9v11H3z"/>,
  package: <><path d="m3 7 9-4 9 4-9 4-9-4Z"/><path d="M3 7v10l9 4 9-4V7M12 11v10"/></>,
  split: <><path d="M8 4H4v4M4 4l6 6M16 20h4v-4M20 20l-6-6M16 4h4v4M20 4l-5 5M8 20H4v-4M4 20l5-5"/></>,
  check: <><circle cx="12" cy="12" r="9"/><path d="m8 12 2.7 2.7L16.5 9"/></>,
  close: <path d="m7 7 10 10M17 7 7 17"/>,
  link: <><path d="M10 13a5 5 0 0 0 7.5.5l2-2a5 5 0 0 0-7-7l-1.1 1.1"/><path d="M14 11a5 5 0 0 0-7.5-.5l-2 2a5 5 0 0 0 7 7l1.1-1.1"/></>,
  dashboard: <><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></>,
  grid: <><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></>,
  list: <><path d="M9 6h12M9 12h12M9 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/></>,
  open: <><path d="M14 3h7v7M10 14 21 3"/><path d="M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/></>,
  clock: <><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></>,
  table: <><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11M15 9v11"/></>,
  filter: <path d="M4 5h16l-6 7v6l-4 2v-8Z"/>,
  refresh: <><path d="M20 7v5h-5"/><path d="M4 17v-5h5"/><path d="M6.1 8.5A7 7 0 0 1 18.8 7L20 12M4 12l1.2 5A7 7 0 0 0 17.9 15.5"/></>,
  info: <><circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7h.01"/></>,
  eye: <><path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z"/><circle cx="12" cy="12" r="2.5"/></>,
  arrowBack: <><path d="M19 12H5M11 18l-6-6 6-6"/></>,
  sparkles: <><path d="m12 3 1.2 3.8L17 8l-3.8 1.2L12 13l-1.2-3.8L7 8l3.8-1.2L12 3ZM19 15l.7 2.3L22 18l-2.3.7L19 21l-.7-2.3L16 18l2.3-.7L19 15ZM5 13l.7 2.3L8 16l-2.3.7L5 19l-.7-2.3L2 16l2.3-.7L5 13Z"/></>,
};

export function QualityIcon({ name, size = 20, className = "", strokeWidth = 1.8 }) {
  return (
    <svg className={className} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
      {QUALITY_ICON_PATHS[name] ?? QUALITY_ICON_PATHS.info}
    </svg>
  );
}

export function QualityPageHeader({ icon = "sparkles", eyebrow, title, description, actions, className = "" }) {
  return (
    <header className={`qa-page-header ${className}`.trim()}>
      <span className="qa-page-header-icon"><QualityIcon name={icon} size={24} /></span>
      <div className="qa-page-header-copy">
        {eyebrow && <span className="qa-page-eyebrow">{eyebrow}</span>}
        <h1>{title}</h1>
        {description && <p>{description}</p>}
      </div>
      {actions && <div className="qa-page-header-actions">{actions}</div>}
    </header>
  );
}

export function QualitySectionTitle({ icon = "sparkles", title, description, actions, className = "" }) {
  return (
    <div className={`qa-section-title ${className}`.trim()}>
      <span className="qa-section-title-icon"><QualityIcon name={icon} size={18} /></span>
      <div><h2>{title}</h2>{description && <p>{description}</p>}</div>
      {actions && <div className="qa-section-title-actions">{actions}</div>}
    </div>
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
