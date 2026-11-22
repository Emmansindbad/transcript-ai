import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem, SavedAnalysis } from "@/types/meeting";
import { retainCrucialSourceLinks } from "@/lib/chat-links";
import { applyFiledState, getSavedAnalysis, saveAnalysis } from "@/lib/store";
import { listFiledKeys } from "@/lib/store";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export async function GET(request: NextRequest): Promise<NextResponse> {
  try {
    const meetingKey = request.nextUrl.searchParams.get("meetingKey")?.trim();
    if (!meetingKey) {
      return NextResponse.json(
        { error: "meetingKey is required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const saved = await getSavedAnalysis(meetingKey);
    if (!saved) {
      return NextResponse.json({ analysis: null });
    }

    const filedKeys = await listFiledKeys();
    return NextResponse.json({
      analysis: {
        ...saved,
        actionItems: applyFiledState(saved.actionItems, filedKeys).map((item) => ({
          ...item,
          description: retainCrucialSourceLinks(item.description),
        })),
      } satisfies SavedAnalysis,
    });
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

interface SaveAnalysisBody {
  meetingKey?: string;
  analysis?: string;
  actionItems?: MeetingActionItem[];
  modelUsed?: string;
  videoUsed?: boolean;
  videoNotice?: string;
}

export async function PUT(request: NextRequest): Promise<NextResponse> {
  try {
    const body = (await request.json()) as SaveAnalysisBody;
    if (!body.meetingKey?.trim() || !body.actionItems) {
      return NextResponse.json(
        { error: "meetingKey and actionItems are required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const saved = await saveAnalysis({
      meetingKey: body.meetingKey.trim(),
      analysis: body.analysis ?? "",
      actionItems: body.actionItems.map((item) => ({
        ...item,
        description: retainCrucialSourceLinks(item.description),
      })),
      modelUsed: body.modelUsed,
      videoUsed: body.videoUsed,
      videoNotice: body.videoNotice,
      savedAt: new Date().toISOString(),
    });

    return NextResponse.json({ analysis: saved });
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
