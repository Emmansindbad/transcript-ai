import { readFileSync } from "node:fs";
import { GoogleGenAI } from "@google/genai";

const env = Object.fromEntries(
  readFileSync(".env.local", "utf8")
    .split("\n")
    .filter((line) => line.includes("=") && !line.startsWith("#"))
    .map((line) => {
      const index = line.indexOf("=");
      return [line.slice(0, index), line.slice(index + 1)];
    }),
);

const apiKey = env.GEMINI_API_KEY;
console.log("Key prefix:", apiKey?.slice(0, 6) + "...");

try {
  const ai = new GoogleGenAI({ apiKey });
  const response = await ai.models.generateContent({
    model: "gemini-2.5-pro",
    contents: "Say hello in one sentence.",
    config: {
      systemInstruction: "You are a helpful assistant.",
    },
  });
  console.log("SUCCESS:", response.text?.slice(0, 100));
} catch (error) {
  console.error("ERROR:", error);
  if (error instanceof Error && "cause" in error) {
    console.error("CAUSE:", error.cause);
  }
}

// Also try REST directly
try {
  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=${apiKey}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: "You are helpful." }] },
        contents: [{ parts: [{ text: "Say hello." }] }],
      }),
    },
  );
  const body = await res.text();
  console.log("REST status:", res.status);
  console.log("REST body:", body.slice(0, 300));
} catch (error) {
  console.error("REST fetch failed:", error);
}
