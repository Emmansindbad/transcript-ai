import { NextResponse } from "next/server";
import { getSummaryEmailRecipient, getTargetUserId } from "@/lib/env";
import { listTenantUsers } from "@/lib/m365";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export async function GET(): Promise<NextResponse> {
  try {
    const users = await listTenantUsers();
    const targetUserId = getTargetUserId().toLowerCase();
    const defaultUser =
      users.find(
        (user) =>
          user.userPrincipalName.toLowerCase() === targetUserId ||
          user.id.toLowerCase() === targetUserId,
      ) ?? null;

    return NextResponse.json({
      users,
      defaultUserId: defaultUser?.userPrincipalName ?? getTargetUserId(),
      summaryEmailRecipient: getSummaryEmailRecipient(),
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
