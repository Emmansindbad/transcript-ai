import { exec } from "node:child_process";
import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      {
        error: "Shutdown is only allowed from this PC.",
        code: "FORBIDDEN",
      },
      { status: 403 },
    );
  }

  const pid = process.pid;
  setTimeout(() => {
    if (process.platform === "win32") {
      exec(`taskkill /PID ${pid} /T /F`);
      return;
    }
    process.exit(0);
  }, 400);

  return NextResponse.json({
    ok: true,
    message: "Analyzer is shutting down on this PC.",
  });
}
