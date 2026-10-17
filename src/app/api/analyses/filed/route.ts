import { NextRequest, NextResponse } from "next/server";
import { markItemsFiled } from "@/lib/store";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

interface FiledBody {
  filingKeys?: string[];
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  try {
    const body = (await request.json()) as FiledBody;
    const keys = await markItemsFiled(body.filingKeys ?? []);
    return NextResponse.json({ filedKeys: keys });
  } catch (error) {
    return NextResponse.json(
      {
        error: getErrorMessage(error),
        code: getErrorCode(error),
      },
      { status: getErrorStatusCode(error) },
    );
  }
}
