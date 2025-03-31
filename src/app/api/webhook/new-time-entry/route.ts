import { createClient } from "@/lib/server";
import axios from "axios";
import { NextResponse } from "next/server";

async function refreshAccessToken(refreshToken: string) {
  try {
    const tokenEndpoint =
      "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    const params = new URLSearchParams({
      client_id: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
      scope: "offline_access Calendars.ReadWrite",
    });

    const response = await axios.post(tokenEndpoint, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Origin: process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000",
      },
    });

    console.log("Token refresh response:", response.data);

    return {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token,
    };
  } catch (error) {
    console.error("Error refreshing token:", error);
    throw error;
  }
}

export async function POST(request: Request) {
  const body = await request.json();
  const supabase = await createClient();

  console.log(body, "new time entry");

  try {
    let scopedUser = null;

    const user = await supabase
      .from("users")
      .select()
      .eq("id", body.userId as string);

    if (!user.data) {
      return NextResponse.json({ error: "User not found" }, { status: 404 });
    }
    scopedUser = user.data[0];

    if (!scopedUser?.provider?.azure?.connected) {
      return NextResponse.json(
        { error: "Azure not connected" },
        { status: 400 }
      );
    }

    try {
      // Refresh token first
      const tokens = await refreshAccessToken(
        scopedUser.provider.azure.refresh_token
      );

      // Update tokens in database
      await supabase
        .from("users")
        .update({
          provider: {
            ...scopedUser.provider,
            azure: {
              ...scopedUser.provider.azure,
              access_token: tokens.access_token,
              refresh_token: tokens.refresh_token,
            },
          },
        })
        .eq("id", body.userId);

      const startTime = new Date(body.timeInterval.start);
      const endTime = new Date(body.timeInterval.end);

      const client = body.project?.clientName
        ? `${body.project?.clientName} : `
        : "";
      const project = body.project?.name ?? "";
      const task = body.task?.name ? ` : ${body.task?.name}` : "";
      const description = body.description ? ` - ${body.description}` : "";

      let calendarResponse = await axios.post(
        `https://graph.microsoft.com/v1.0/me/calendars/${scopedUser.provider.azure.calendarId}/events`,
        {
          subject: client + project + task + description,
          body: {
            contentType: "text",
            content: body.id,
          },
          start: {
            dateTime: startTime.toISOString(),
            timeZone: "UTC",
          },
          end: {
            dateTime: endTime.toISOString(),
            timeZone: "UTC",
          },
          categories: ["Blue category"],
          singleValueExtendedProperties: [
            {
              id: "String {66f5a359-4659-4830-9070-00040ec6ac6e} Name clockifyId",
              value: body.id,
            },
          ],
        },
        {
          headers: {
            Authorization: `Bearer ${tokens.access_token}`,
            "Content-Type": "application/json",
            Prefer: 'outlook.timezone="UTC"',
          },
        }
      );

      // console.log('Calendar response:', calendarResponse.data);
      return NextResponse.json(calendarResponse.data);
    } catch (error: any) {
      console.error("API Error details:", {
        status: error.response?.status,
        data: error.response?.data,
        headers: error.response?.headers,
      });

      if (error.response?.status === 401) {
        console.log("Token expired, disconnecting Azure...");

        await supabase
          .from("users")
          .update({
            provider: {
              ...scopedUser.provider,
              azure: {
                ...scopedUser.provider.azure,
                connected: false,
              },
            },
          })
          .eq("id", body.userId);

        return NextResponse.json(
          {
            error: "Token expired, please reconnect Azure",
            details:
              error.response?.data?.error?.message ||
              "No error details available",
          },
          { status: 401 }
        );
      }
      throw error;
    }
  } catch (error: any) {
    console.error("Unexpected error:", error.response?.data || error);
    return NextResponse.json(
      {
        error: "Failed to create calendar event",
        details: error.message,
      },
      { status: 500 }
    );
  }
}
