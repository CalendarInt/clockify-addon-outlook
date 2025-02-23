import { NextResponse } from "next/server";

export async function POST(request: Request, response: Response) {
  return NextResponse.json({ response: "installed" });
}
