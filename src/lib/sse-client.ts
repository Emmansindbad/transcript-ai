export async function readSseEvents(
  response: Response,
  onEvent: (event: Record<string, unknown>) => void,
): Promise<void> {
  if (!response.body) {
    throw new Error("The analysis response did not include a stream.");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }

    buffer += decoder.decode(value, { stream: true });
    const parts = buffer.split("\n\n");
    buffer = parts.pop() ?? "";

    for (const part of parts) {
      const dataLine = part
        .split("\n")
        .find((line) => line.startsWith("data: "));
      if (!dataLine) {
        continue;
      }
      try {
        const parsed = JSON.parse(dataLine.slice(6)) as Record<string, unknown>;
        onEvent(parsed);
      } catch {
        // Ignore malformed chunks; the next event may complete the JSON.
      }
    }
  }
}
