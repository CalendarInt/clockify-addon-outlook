import { createClient } from "@/lib/client";
import axiosInstance from "@/app/proba/axiosInterceptorInstance";
import { QueryClient } from "@tanstack/react-query";
import { ClockifyToken } from "@/lib/models/clockify-token";
import {
  subMonths,
  addYears,
  addHours,
  parse,
  formatISO,
  addMonths,
} from "date-fns";

const getBaseUrl = (workspaceId: string) => {
  console.log(`Current URL: ${window.location.href}`);
  return process.env.NODE_ENV === "development" || window.location.href.includes('developer')
    ? `https://developer.clockify.me/report/v1/workspaces/${workspaceId}/reports/detailed`
    : `https://reports.api.clockify.me/v1/workspaces/${workspaceId}/reports/detailed`;
};

export const fetchOutlookCalendars = async (
  jwt: ClockifyToken,
  queryClient: QueryClient
) => {
  let scopedUser = queryClient.getQueryData(["user"]) as any;
  const supabase = createClient();

  try {
    // Get Microsoft Graph access token
    const account = scopedUser.provider?.azure?.access_token;
    if (!account) {
      throw new Error("No Azure account found");
    }

    // Fetch user's calendars
    const response = await axiosInstance.get(
      "https://graph.microsoft.com/v1.0/me/calendars",
      {
        headers: {
          Authorization: `Bearer ${account}`,
          "Content-Type": "application/json",
        },
      }
    );

    // Check if Clockify calendar exists
    let clockifyCalendar = response.data.value.find(
      (calendar: any) => calendar.name === "Clockify Addon Calendar"
    );

    // If no Clockify calendar exists, create one
    if (!clockifyCalendar) {
      const newCalendar = await axiosInstance.post(
        "https://graph.microsoft.com/v1.0/me/calendars",
        {
          name: "Clockify Addon Calendar",
        },
        {
          headers: {
            Authorization: `Bearer ${account}`,
            "Content-Type": "application/json",
          },
        }
      );
      clockifyCalendar = newCalendar.data;
    }

    // Update user in database with calendar ID
    let updatedUser = await supabase
      .from("users")
      .update({
        provider: {
          ...scopedUser.provider,
          azure: {
            ...scopedUser.provider.azure,
            calendarId: clockifyCalendar.id,
          },
        },
      })
      .eq("id", jwt.user)
      .select("*");

    if (updatedUser?.data) {
      queryClient.setQueryData(["user"], updatedUser.data[0]);
    }

    return clockifyCalendar;
  } catch (error) {
    console.error("Error in fetchOutlookCalendars:", error);
    throw error;
  }
};

export const syncWithOutlookCalendar = async (
  timeEntries: any[],
  queryClient: QueryClient
) => {
  console.log("syncWithOutlookCalendar");
  let scopedUser = queryClient.getQueryData(["user"]) as any;
  const accessToken = scopedUser.provider?.azure?.access_token;
  const calendarId = scopedUser.provider?.azure?.calendarId;

  if (!accessToken || !calendarId) {
    throw new Error("Missing Azure credentials or calendar ID");
  }

  console.log(timeEntries);

  // Prepare batch requests
  const batchRequests = timeEntries.map((entry, index) => ({
    id: (index + 1).toString(),
    method: "POST",
    url: `/me/calendars/${calendarId}/events`,
    body: {
      subject: entry.description,
      body: {
        contentType: "text",
        content: entry._id
      },
      start: {
        dateTime: entry.timeInterval.start,
        timeZone: "UTC"
      },
      end: {
        dateTime: entry.timeInterval.end,
        timeZone: "UTC"
      },
      categories: ["Blue category"],
      singleValueExtendedProperties: [
        {
          id: "String {66f5a359-4659-4830-9070-00040ec6ac6e} Name clockifyId",
          value: entry._id,
        },
      ],
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: 'outlook.timezone="UTC"',
    },
  }));

  // Split requests into chunks of 20 (Microsoft Graph batch limit)
  const batchSize = 20;
  const batches = [];
  for (let i = 0; i < batchRequests.length; i += batchSize) {
    batches.push(batchRequests.slice(i, i + batchSize));
  }

  console.log(batches);

  try {
    // Process each batch
    for (const batchChunk of batches) {
      await axiosInstance.post(
        'https://graph.microsoft.com/v1.0/$batch',
        {
          requests: batchChunk
        },
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );
      
      // Add a small delay between batches to avoid rate limits
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
  } catch (error) {
    console.error("Error in batch operation:", error);
    throw error;
  }
};

export const timeEntriesSyncMutation = async (
  jwt: ClockifyToken,
  authToken: string,
  queryClient: QueryClient,
  controlValue: boolean,
  calendar: string,
  type: string
) => {
  let scopedUser = queryClient.getQueryData(["user"]) as any;

  if (!controlValue) {
    await updateFormStateInDatabase(
      scopedUser,
      type,
      controlValue,
      jwt,
      queryClient
    );
    return [];
  }

  try {
    // Fetch time entries from Clockify
    console.log({
      dateRangeEnd: addMonths(new Date(), 1),
      dateRangeStart: subMonths(new Date(), 1),
      detailedFilter: {},
      amountShown: "HIDE_AMOUNT",
      users: {
        ids: [jwt.user],
      },
    });

    const detailedReport = await axiosInstance.post(
      getBaseUrl(jwt.workspaceId),
      {
        dateRangeEnd: addMonths(new Date(), 1),
        dateRangeStart: subMonths(new Date(), 1),
        detailedFilter: {},
        amountShown: "HIDE_AMOUNT",
        users: {
          ids: [jwt.user],
        },
      },
      {
        headers: {
          "x-addon-token": authToken,
        },
      }
    );

    console.log(detailedReport.data);

    const timeEntries = detailedReport.data.timeentries.filter(
      (timeEntry: any) => {
        const client = timeEntry?.clientName
          ? `${timeEntry?.clientName} : `
          : "";
        const project = timeEntry?.projectName ?? "";
        const task = timeEntry?.taskName ? ` : ${timeEntry.taskName}` : "";
        const description = timeEntry?.description
          ? ` - ${timeEntry.description}`
          : "";
        timeEntry.description = `${client}${project}${task}${description}`;
        return timeEntry.type === "REGULAR";
      }
    );

    if (timeEntries.length > 0) {
      await syncWithOutlookCalendar(timeEntries, queryClient);
    }
  } catch (error) {
    throw error;
  }

  await updateFormStateInDatabase(
    scopedUser,
    type,
    controlValue,
    jwt,
    queryClient
  );
};

export const timeOffSyncMutation = async (
  jwt: ClockifyToken,
  queryClient: QueryClient,
  controlValue: boolean,
  calendar: string,
  type: string
) => {
  let scopedUser = queryClient.getQueryData(["user"]) as any;

  updateFormStateInDatabase(scopedUser, type, controlValue, jwt, queryClient);
};

async function updateFormStateInDatabase(
  scopedUser: any,
  type: string,
  formControlValue: any,
  jwt: ClockifyToken,
  queryClient: QueryClient
) {
  const supabase = createClient();

  let updatedUser = await supabase
    .from("users")
    .update({
      provider: {
        ...scopedUser.provider,
        azure: {
          ...scopedUser.provider.azure,
          sync: {
            ...scopedUser.provider.azure?.sync,
            [type]: {
              value: formControlValue,
              initialized: true,
            },
          },
        },
      },
    })
    .eq("id", jwt.user)
    .select("*");

  if (updatedUser?.data) {
    queryClient.setQueryData(["user"], updatedUser.data[0]);
  }
}

export async function disconnectOutlookCalendar(
  jwt: ClockifyToken,
  scopedUser: any,
  queryClient: any
) {
  const supabase = createClient();
  let updatedUser = await supabase
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
    .eq("id", jwt.user)
    .select("*");

  if (updatedUser?.data) {
    queryClient.setQueryData(["user"], updatedUser.data[0]);
  }
}
