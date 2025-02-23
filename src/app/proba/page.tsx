"use client";
import { useEffect, useState } from "react";
import { useSearchParams } from "next/navigation";
import { Switch } from "@/components/ui/switch";
import { jwtDecode } from "jwt-decode";
import { PublicClientApplication } from "@azure/msal-browser";
import { createClient } from "@/lib/client";
import { Button } from "@/components/ui/button";
import {
  Form,
  FormControl,
  FormField,
  FormItem,
  FormLabel,
} from "@/components/ui/form";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import Loading from "@/components/loading";
import {
  fetchOutlookCalendars,
  timeEntriesSyncMutation,
  disconnectOutlookCalendar,
} from "../outlookQueris";

async function fetchUser(jwt: any) {
  const supabase = createClient();
  const existingUser = await supabase.from("users").select().eq("id", jwt.user);

  return existingUser.data?.[0];
}

export default function OutlookTest() {
  const queryClient = useQueryClient();
  const searchParams = useSearchParams();
  let jwt: any;

  if (searchParams?.get("auth_token")) {
    jwt = jwtDecode(searchParams?.get("auth_token") ?? "");
  }

  const formSchema = z.object({
    outlookTimeEntry: z.boolean().default(false),
    outlookTimeOff: z.boolean().default(false),
    outlookScheduledTime: z.boolean().default(false),
  });

  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: {
      outlookTimeEntry: false,
      outlookTimeOff: false,
      outlookScheduledTime: false,
    },
  });

  const {
    data: supabaseUser,
    refetch: refetchSupabaseUser,
    isFetching: isLoadingSupabaseUser,
  } = useQuery({
    queryKey: ["user"],
    queryFn: () => fetchUser(jwt),
    staleTime: Infinity,
    refetchInterval: false,
  });

  useEffect(() => {
    if (!supabaseUser) {
      return;
    }
    form.reset({
      outlookTimeEntry:
        !!supabaseUser?.provider?.azure?.sync?.outlookTimeEntry?.value,
      outlookTimeOff:
        !!supabaseUser?.provider?.azure?.sync?.outlookTimeOff?.value,
      outlookScheduledTime:
        !!supabaseUser?.provider?.azure?.sync?.outlookScheduledTime?.value,
    });
  }, [supabaseUser]);

  let pca: PublicClientApplication;
  pca = new PublicClientApplication({
    auth: {
      clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || "",
      authority: "https://login.microsoftonline.com/common",
      redirectUri:
        process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000",
    },
  });

  const {
    data: outlookCalendars,
    refetch: refetchOutlookCalendars,
    isFetching: isLoadingOutlookCalendars,
  } = useQuery({
    queryKey: ["calendars"],
    queryFn: () => fetchOutlookCalendars(jwt, queryClient),
    staleTime: Infinity,
    refetchInterval: false,
    enabled: false,
  });

  const {
    mutate: timeEntriesSync,
    isPending: isTimeEntriesSyncMutationPending,
  } = useMutation({
    mutationFn: ({ calendar, type }: { calendar: string; type: string }) =>
      timeEntriesSyncMutation(
        jwt,
        searchParams.get("auth_token") as string,
        queryClient,
        (form.getValues() as any)[type],
        calendar,
        type
      ),
    onSuccess: async (codeResponse) => {},
    onError: (error) => {
      form.setValue("outlookTimeEntry", !form.getValues().outlookTimeEntry);
    },
  });

  async function outlookLogin() {
    try {
      await pca.initialize();
      const loginResponse = await pca.loginPopup({
        scopes: ["User.Read", "Calendars.ReadWrite"],
      });

      const supabase = createClient();
      const jwt = jwtDecode(searchParams.get("auth_token") as string) as any;

      const existingUser = await supabase
        .from("users")
        .select()
        .eq("id", jwt.user as string);

      if (existingUser?.data) {
        let updatedUser = await supabase
          .from("users")
          .update({
            provider: {
              ...existingUser.data[0].provider,
              azure: {
                access_token: loginResponse.accessToken,
                connected: true,
              },
            },
          })
          .eq("id", jwt.user as string)
          .select("*");
      }

      refetchSupabaseUser();
      refetchOutlookCalendars();
    } catch (error) {
      console.error("Login error:", error);
    }
  }

  const disconnect = async () => {
    let scopedUser = queryClient.getQueryData(["user"]) as any;
    await disconnectOutlookCalendar(jwt, scopedUser, queryClient);
  };

  const loadingArray = [
    isLoadingOutlookCalendars,
    isLoadingSupabaseUser,
    isTimeEntriesSyncMutationPending,
  ];

  console.log(supabaseUser);

  return (
    <div className="flex flex-col items-start p-28 gap-12">
      {loadingArray.some((el) => el) ? (
        <Loading />
      ) : (
        <div className="flex flex-col gap-12 w-full">
          <div className="flex flex-col gap-4 items-center">
            <div className="flex flex-row gap-4 items-center self-start">
              Outlook Calendar
              {supabaseUser && !supabaseUser?.provider?.azure?.connected ? (
                <Button onClick={outlookLogin} className="bg-blue-600">
                  Connect
                </Button>
              ) : (
                <Button onClick={disconnect} className="bg-green-600">
                  Disconnect
                </Button>
              )}
            </div>

            <Form {...form}>
              <form className="w-full space-y-6">
                <div className="flex flex-row gap-5 items-center">
                  <div className="w-1/5">
                    <FormField
                      control={form.control}
                      name="outlookTimeEntry"
                      render={({ field }) => (
                        <FormItem className="flex flex-row items-center justify-between rounded-lg border p-3 shadow-sm">
                          <div className="space-y-0.5">
                            <FormLabel>Time entry</FormLabel>
                          </div>
                          <FormControl>
                            <Switch
                              checked={field.value}
                              onCheckedChange={(e) => {
                                field.onChange(e);
                                timeEntriesSync({
                                  calendar: "Outlook",
                                  type: field.name,
                                });
                              }}
                            />
                          </FormControl>
                        </FormItem>
                      )}
                    />
                  </div>
                  <h1>
                    Time entries will be displayed in your Outlook Calendar
                  </h1>
                </div>
              </form>
            </Form>
          </div>
        </div>
      )}
    </div>
  );
}
