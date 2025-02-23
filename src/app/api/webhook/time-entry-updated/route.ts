import { createClient } from "@/lib/server";
import axios, { AxiosError } from "axios";
import { NextResponse } from "next/server";

async function refreshAccessToken(refreshToken: string) {
  try {
    const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    const params = new URLSearchParams({
      client_id: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
      scope: 'offline_access Calendars.ReadWrite'
    });

    const response = await axios.post(tokenEndpoint, params, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000"
      }
    });

    console.log('Token refresh response:', response.data);

    return {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token
    };
  } catch (error) {
    console.error('Error refreshing token:', error);
    throw error;
  }
}

export async function POST(request: Request, response: Response) {
  const body = await request.json();
  const supabase = await createClient();
  console.log(body, "time entry updated");

  try {
    let scopedUser = null;

    const user = await supabase
      .from("users")
      .select()
      .eq("id", body.userId as string);
    
    if (!user.data) {
      console.error("User not found:", body.userId);
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }
    scopedUser = user.data[0];

    if (!scopedUser?.provider?.azure?.connected) {
      console.log("Azure calendar not connected for user:", body.userId);
      return NextResponse.json({ error: "Azure calendar not connected" }, { status: 400 });
    }

    if (!scopedUser.provider.azure.calendarId) {
      console.error("Calendar ID not found for user:", body.userId);
      return NextResponse.json({ error: "Calendar ID not found" }, { status: 400 });
    }

    // Get the existing event with this time entry ID
    try {
      // Convert dates to handle timezone consistently
      const startTime = new Date(body.timeInterval.start);
      const endTime = new Date(body.timeInterval.end);
      
      // Get events within the time range
      let response = await axios.get(
        `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events`,
        {
          headers: {
            Authorization: `Bearer ${scopedUser.provider.azure.access_token}`,
            "Content-Type": "application/json",
            Prefer: 'outlook.timezone="UTC"'
          },
          params: {
            $select: "id,subject,body,start,end",
            $filter: `SingleValueExtendedProperties/any(ep: ep/id eq 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name clockifyId' and ep/value eq '${body.id}')`,
          }
        }
      );

      console.log("Event response:", response.data);

      // Find the event with matching ID in body content
      const matchingEvent = response.data.value.find(
        (event: any) => event.body?.content?.includes(body.id)
      );

      if (!matchingEvent) {
        console.log("No matching event found for time entry:", body.id);
        return NextResponse.json({ error: "Event not found" }, { status: 404 });
      }

      // Update the existing event
      try {
        let updateResponse = await axios.patch(
          `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events/${matchingEvent.id}`,
          {
            start: {
              dateTime: body.timeInterval.start,
              timeZone: "UTC"
            },
            end: {
              dateTime: body.timeInterval.end,
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

        console.log("Event updated successfully:", updateResponse.data.id);
        return NextResponse.json(updateResponse.data);
      } catch (updateError) {
        if (updateError instanceof AxiosError) {
          console.error("Failed to update event:", {
            status: updateError.response?.status,
            data: updateError.response?.data,
            error: updateError.message
          });
          
          if (updateError.response?.status === 401) {
            return NextResponse.json({ error: "Authentication failed" }, { status: 401 });
          }
          
          return NextResponse.json(
            { error: "Failed to update event", details: updateError.response?.data },
            { status: updateError.response?.status || 500 }
          );
        }
        throw updateError;
      }
    } catch (searchError) {
      if (searchError instanceof AxiosError) {
        console.error("Failed to search for event:", {
          status: searchError.response?.status,
          data: searchError.response?.data,
          error: searchError.message
        });
        
        return NextResponse.json(
          { error: "Failed to search for event", details: searchError.response?.data },
          { status: searchError.response?.status || 500 }
        );
      }
      throw searchError;
    }
  } catch (error) {
    console.error("Unexpected error:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
