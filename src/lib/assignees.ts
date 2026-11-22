import type { TenantUser } from "@/types/meeting";

function normalizeName(value: string): string {
  return value
    .toLowerCase()
    .replace(/[._@]/g, " ")
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function nameTokens(value: string): string[] {
  return normalizeName(value)
    .split(" ")
    .filter((token) => token.length > 1);
}

function directoryRank(user: TenantUser): number {
  const upn = user.userPrincipalName.toLowerCase();
  let rank = 0;
  if (!upn.includes("#ext#")) {
    rank += 2;
  }
  if (upn.endsWith("@sindbad.tech")) {
    rank += 1;
  }
  return rank;
}

function preferUser(current: TenantUser | null, candidate: TenantUser): TenantUser {
  if (!current || directoryRank(candidate) > directoryRank(current)) {
    return candidate;
  }
  return current;
}

function toMatch(user: TenantUser): { displayName: string; userPrincipalName: string } {
  return {
    displayName: user.displayName,
    userPrincipalName: user.userPrincipalName,
  };
}

export function matchAssignee(
  pointPerson: string,
  users: TenantUser[],
): { displayName: string; userPrincipalName: string } | null {
  const raw = pointPerson.trim();
  if (!raw || /^unassigned$/i.test(raw) || /^tbd$/i.test(raw) || /^none$/i.test(raw)) {
    return null;
  }

  const needle = normalizeName(raw);
  if (!needle) {
    return null;
  }

  const needleSignificant = nameTokens(raw).join(" ");
  let exact: TenantUser | null = null;
  for (const user of users) {
    const displayName = normalizeName(user.displayName);
    const localPart = normalizeName(user.userPrincipalName.split("@")[0] ?? "");
    const displaySignificant = nameTokens(user.displayName).join(" ");
    if (
      displayName === needle ||
      localPart === needle ||
      (needleSignificant.length > 0 && displaySignificant === needleSignificant)
    ) {
      exact = preferUser(exact, user);
    }
  }
  if (exact) {
    return toMatch(exact);
  }

  const tokens = nameTokens(raw);
  if (tokens.length === 0) {
    return null;
  }

  let best: TenantUser | null = null;
  let bestScore = 0;

  for (const user of users) {
    const haystackTokens = new Set(
      nameTokens(`${user.displayName} ${user.userPrincipalName.split("@")[0] ?? ""}`),
    );
    const matched = tokens.filter((token) => haystackTokens.has(token)).length;
    const score = matched / tokens.length;
    if (score > bestScore) {
      bestScore = score;
      best = user;
    } else if (
      score === bestScore &&
      score > 0 &&
      best &&
      directoryRank(user) > directoryRank(best)
    ) {
      best = user;
    }
  }

  if (best && bestScore >= 0.66) {
    return toMatch(best);
  }

  return null;
}

export function directoryEmail(user: TenantUser): string {
  const mail = user.mail?.trim() ?? "";
  if (mail.includes("@")) {
    return mail;
  }
  return user.userPrincipalName.trim();
}

function personFromMatch(
  matched: { displayName: string; userPrincipalName: string },
  users: TenantUser[],
): { displayName: string; email: string } {
  const user = users.find(
    (candidate) =>
      candidate.userPrincipalName.toLowerCase() === matched.userPrincipalName.toLowerCase(),
  );
  if (!user) {
    return { displayName: matched.displayName, email: matched.userPrincipalName };
  }
  return { displayName: user.displayName, email: directoryEmail(user) };
}

export function resolveDirectoryAssignee(
  pointPerson: string,
  storedUpn: string | undefined,
  users: TenantUser[],
): { displayName: string; email: string } | null {
  const matched = matchAssignee(pointPerson, users);
  if (matched) {
    return personFromMatch(matched, users);
  }

  const typed = pointPerson.trim();
  if (typed && !/^(unassigned|tbd|none)$/i.test(typed)) {
    return null;
  }

  const stored = storedUpn?.trim() ?? "";
  if (!stored.includes("@")) {
    return null;
  }
  const known = users.find(
    (user) =>
      user.userPrincipalName.toLowerCase() === stored.toLowerCase() ||
      user.mail?.toLowerCase() === stored.toLowerCase(),
  );
  return {
    displayName: known?.displayName ?? stored,
    email: known ? directoryEmail(known) : stored,
  };
}

export function assigneeFieldsForName(
  pointPerson: string,
  users: TenantUser[],
): { pointPerson: string; assigneeUpn: string; assigneeDisplayName: string } {
  const resolved = resolveDirectoryAssignee(pointPerson, undefined, users);
  return {
    pointPerson,
    assigneeUpn: resolved?.email ?? "",
    assigneeDisplayName: resolved?.displayName ?? "",
  };
}

export function applyAssigneeMatches<T extends { pointPerson: string }>(
  items: T[],
  users: TenantUser[],
): Array<T & { assigneeUpn?: string; assigneeDisplayName?: string }> {
  return items.map((item) => {
    const matched = matchAssignee(item.pointPerson, users);
    if (!matched) {
      return item;
    }

    return {
      ...item,
      pointPerson: matched.displayName,
      assigneeDisplayName: matched.displayName,
      assigneeUpn: matched.userPrincipalName,
    };
  });
}
