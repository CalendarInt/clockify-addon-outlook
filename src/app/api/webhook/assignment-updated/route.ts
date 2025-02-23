import { createClient } from "@/lib/server";
import axios from "axios";
import { addHours, formatISO, parse } from "date-fns";
import { NextResponse } from "next/server";

export async function POST(request: Request) {
  const body = await request.json();
  const supabase = await createClient();
  console.log(body, "assignment updated");

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

    // Get all events with this assignment ID
    let response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events`,
      {
        headers: {
          Authorization: `Bearer ${scopedUser.provider.azure.access_token}`,
          "Content-Type": "application/json",
        },
        params: {
          $filter: `body/content eq '${body.id}'`
        }
      }
    );

    if (response.data.value.length > 0) {
      try {
        // Update the event
        let updateResponse = await axios.patch(
          `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events/${response.data.value[0].id}`,
          {
            start: {
              dateTime: start,
              timeZone: "UTC"
            },
            end: {
              dateTime: end,
              timeZone: "UTC"
            }
          },
          {
            headers: {
              Authorization: `Bearer ${scopedUser.provider.azure.access_token}`,
              "Content-Type": "application/json"
            }
          }
        );
        console.log(updateResponse.data, "Event updated");
        return NextResponse.json(updateResponse.data);
      } catch (error) {
        console.error("Error updating event:", error);
        return NextResponse.json({ error: "Failed to update event" }, { status: 500 });
      }
    }
  }

  return NextResponse.json("installed");
}
