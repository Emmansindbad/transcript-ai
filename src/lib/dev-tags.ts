export const DEV_TAGS = [
  "None",
  "#backend",
  "#frontend",
  "#data",
  "#devops",
  "#api",
  "#ui",
  "#mobile",
  "#security",
  "#performance",
  "#docs",
  "#testing",
  "#a11y",
] as const;

export type DevTag = (typeof DEV_TAGS)[number];

const ALLOWED = new Map<string, DevTag>(
  DEV_TAGS.map((tag) => [tag.toLowerCase(), tag]),
);

export function normalizeDevTag(value: string | null | undefined): DevTag {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) {
    return "None";
  }

  const direct = ALLOWED.get(trimmed.toLowerCase());
  if (direct) {
    return direct;
  }

  const hashed = ALLOWED.get(`#${trimmed.toLowerCase()}`);
  return hashed ?? "None";
}
