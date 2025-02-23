import { NextResponse } from "next/server";
const baseUrl =
  process.env.NODE_ENV === "development"
    ? "https://herring-endless-firmly.ngrok-free.app"
    : "https://clockify-addon-outlook.vercel.app/";

export async function GET() {
  const manifest = {
    key: "OutlookCalendarIntegration",
    name: "Outlook Calendar Sync",
    iconPath: "/tab_icon.svg",
    description: "Transfer time entries and approved time offs to Outlook Calendar",
    baseUrl: baseUrl,
    lifecycle: [
      {
        type: "INSTALLED",
        path: "api/lifecycle/installed",
      },
      {
        type: "DELETED",
        path: "/lifecycle/uninstalled",
      },
      {
        type: "SETTINGS_UPDATED",
        path: "/lifecycle/settings-updated",
      },
    ],
    webhooks: [
      {
        event: "NEW_TIME_ENTRY",
        path: "/api/webhook/new-time-entry",
      },
      {
        event: "TIME_OFF_REQUEST_APPROVED",
        path: "/api/webhook/time-off-request-approved",
      },
      {
        event: "TIMER_STOPPED",
        path: "/api/webhook/timer-stopped",
      },
      {
        event: "TIME_ENTRY_UPDATED",
        path: "/api/webhook/time-entry-updated",
      },
      {
        event: "TIME_OFF_REQUESTED",
        path: "/api/webhook/time-off-requested",
      },
      // {
      //   event: "ASSIGNMENT_PUBLISHED",
      //   path: "/api/webhook/assignment-published",
      // },
      // {
      //   event: "ASSIGNMENT_UPDATED",
      //   path: "/api/webhook/assignment-updated",
      // },
    ],
    components: [
      {
        type: "sidebar",
        accessLevel: "EVERYONE",
        path: "/",
        label: "Outlook Calendar Sync",
        iconPath: "tab_icon.svg",
      },
    ],
    minimalSubscriptionPlan: "FREE",
    scopes: [
      "CLIENT_READ",
      "PROJECT_READ",
      "TASK_READ",
      "TIME_ENTRY_READ",
      "USER_READ",
      "WORKSPACE_READ",
      "APPROVAL_READ",
      "REPORTS_READ",
      "TIME_OFF_READ",
    ],
  };

  return NextResponse.json(manifest, {
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET",
    },
  });
}
