import { createClient } from "@/lib/server";
import axios from "axios";
import { addHours, formatISO, parse } from "date-fns";
import { NextResponse } from "next/server";

export async function POST(request: Request, response: Response) {
  const body = await request.json();
  const supabase = await createClient();
  console.log(body, "assignment published");

  let scopedUser = null;

  const user = await supabase
    .from("users")
    .select()
    .eq("id", body.userId as string);
  if (!user.data) {
    return;
  }
  scopedUser = user.data[0];

  if (!scopedUser?.provider?.azure?.connected) {
    console.log("disconnected");
    return NextResponse.json("disconnected");
  }

  if (user.data && user.data[0].provider?.azure?.sync?.outlookScheduledTime?.value) {
    const start = formatISO(
      parse(body.startTime ?? "00:00", "HH:mm", new Date(body.period.start))
    );

    const end = formatISO(addHours(start, body.hoursPerDay));

    try {
      let response = await axios.post(
        `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events`,
        {
          subject: body.note ?? "No title",
          body: {
            contentType: "text",
            content: body.id
          },
          start: {
            dateTime: start,
            timeZone: "UTC"
          },
          end: {
            dateTime: end,
            timeZone: "UTC"
          },
          categories: ["Clockify"]
        },
        {
          headers: {
            Authorization: `Bearer ${scopedUser.provider.azure.access_token}`,
            "Content-Type": "application/json"
          }
        }
      );

      console.log(response.data);
      return NextResponse.json(response.data);
    } catch (error) {
      console.error("Error creating event:", error);
      return NextResponse.json({ error: "Failed to create event" }, { status: 500 });
    }
  }

  return NextResponse.json("installed");
}
