import "server-only";

import type { NextRequest } from "next/server";

function isLoopbackAddress(value: string): boolean {
  const trimmed = value.trim().toLowerCase();
  return (
    trimmed === "127.0.0.1" ||
    trimmed === "::1" ||
    trimmed === "0:0:0:0:0:0:0:1" ||
    trimmed === "::ffff:127.0.0.1" ||
    trimmed.startsWith("127.")
  );
}

export function isLocalRequest(request: NextRequest): boolean {
  const host = request.headers.get("host") ?? "";
  const hostname = host.split("]")[0]?.replace("[", "").split(":")[0] ?? "";
  if (hostname === "localhost" || isLoopbackAddress(hostname)) {
    const forwarded = request.headers.get("x-forwarded-for");
    if (forwarded) {
      const first = forwarded.split(",")[0]?.trim() ?? "";
      if (first && !isLoopbackAddress(first) && first.toLowerCase() !== "localhost") {
        return false;
      }
    }
    return true;
  }
  return false;
}
