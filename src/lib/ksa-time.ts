export const KSA_TIME_ZONE = "Asia/Riyadh";

export function parseGraphDate(value: string): Date | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const hasZone = /Z$/i.test(trimmed) || /[+-]\d{2}:\d{2}$/.test(trimmed);
  const normalized = hasZone ? trimmed : `${trimmed}Z`;
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function getKsaDateKey(value: string | Date): string {
  const date = typeof value === "string" ? parseGraphDate(value) : value;
  if (!date) {
    return "unknown";
  }

  return new Intl.DateTimeFormat("en-CA", {
    timeZone: KSA_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

export function formatKsaDate(value: string | Date): string {
  const date = typeof value === "string" ? parseGraphDate(value) : value;
  if (!date) {
    return "Unknown date";
  }

  const formatted = new Intl.DateTimeFormat("en-GB", {
    timeZone: KSA_TIME_ZONE,
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(date);

  return `${formatted} KSA`;
}

export function formatKsaDateTime(value: string | Date): string {
  const date = typeof value === "string" ? parseGraphDate(value) : value;
  if (!date) {
    return "Unknown date";
  }

  const formatted = new Intl.DateTimeFormat("en-GB", {
    timeZone: KSA_TIME_ZONE,
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  }).format(date);

  return `${formatted} KSA`;
}

export function getKsaDateRangeBounds(daysBack = 30): {
  min: string;
  max: string;
} {
  const today = new Date();
  const minDate = new Date(today);
  minDate.setUTCDate(minDate.getUTCDate() - daysBack);

  return {
    min: getKsaDateKey(minDate),
    max: getKsaDateKey(today),
  };
}

export function formatKsaDateKeyLabel(dateKey: string): string {
  if (!dateKey || dateKey === "unknown") {
    return "Unknown date";
  }

  return formatKsaDate(`${dateKey}T12:00:00+03:00`);
}
