"use client";
import { useEffect, useState } from "react";
import { useSearchParams } from "next/navigation";
import { Switch } from "@/components/ui/switch";
import { jwtDecode } from "jwt-decode";
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
  timeOffSyncMutation,
} from "./outlookQueris";
import axios from "axios";

async function generateCodeChallenge() {
  // Generate random string
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  const codeVerifier = Array.from(array, byte => byte.toString(16).padStart(2, '0')).join('');

  // Create code challenge
  const encoder = new TextEncoder();
  const data = encoder.encode(codeVerifier);
  const digest = await crypto.subtle.digest('SHA-256', data);
  const base64Digest = btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

  return { codeVerifier, codeChallenge: base64Digest };
}

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

  const { mutate: timeOffSync, isPending: isTimeOffSyncPending } = useMutation({
    mutationFn: ({ calendar, type }: { calendar: string; type: string }) =>
      timeOffSyncMutation(
        jwt,
        queryClient,
        (form.getValues() as any)[type],
        calendar,
        type
      ),
    onSuccess: async (codeResponse) => {},
    onError: (error) => {
      form.setValue("outlookTimeOff", !form.getValues().outlookTimeOff);
    },
  });

  const outlookLogin = async () => {
    try {
      // Generate PKCE values
      const { codeVerifier, codeChallenge } = await generateCodeChallenge();
      
      // Store code verifier for later use
      localStorage.setItem('azure_code_verifier', codeVerifier);

      // Step 1: Get Authorization Code
      const authEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
      const params = new URLSearchParams({
        client_id: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
        response_type: 'code',
        redirect_uri: process.env.NEXT_PUBLIC_REDIRECT_URI!,
        scope: 'offline_access Calendars.ReadWrite',
        state: crypto.randomUUID(),
        code_challenge: codeChallenge,
        code_challenge_method: 'S256'
      });

      // Store state in localStorage to verify when we return
      localStorage.setItem('azure_auth_state', params.get('state')!);

      // Open Microsoft login in a popup
      const width = 600;
      const height = 600;
      const left = window.screenX + (window.outerWidth - width) / 2;
      const top = window.screenY + (window.outerHeight - height) / 2;
      
      const authWindow = window.open(
        `${authEndpoint}?${params.toString()}`,
        'Azure Login',
        `width=${width},height=${height},left=${left},top=${top}`
      );

      // Poll the popup for the redirect
      const pollTimer = setInterval(() => {
        try {
          if (authWindow?.closed) {
            clearInterval(pollTimer);
            return;
          }

          const currentUrl = authWindow?.location.href;
          if (currentUrl?.includes('code=')) {
            clearInterval(pollTimer);
            authWindow?.close();

            const urlParams = new URLSearchParams(new URL(currentUrl).search);
            const code = urlParams.get('code');
            const state = urlParams.get('state');
            const storedState = localStorage.getItem('azure_auth_state');

            // Clear stored state
            localStorage.removeItem('azure_auth_state');

            if (!code || !state || state !== storedState) {
              console.error("Invalid authorization response");
              return;
            }

            // Exchange code for tokens
            handleAuthCode(code);
          }
        } catch (e) {
          // Ignore cross-origin errors while polling
        }
      }, 100);

    } catch (error) {
      console.error("Login error:", error);
    }
  };
  
  const handleAuthCode = async (code: string) => {
    try {
      // Get stored code verifier
      const codeVerifier = localStorage.getItem('azure_code_verifier');
      if (!codeVerifier) {
        throw new Error('No code verifier found');
      }

      // Clear stored code verifier
      localStorage.removeItem('azure_code_verifier');

      // Step 2: Exchange code for tokens
      const tokenResponse = await axios.post('/api/auth/azure/token', { 
        code,
        code_verifier: codeVerifier
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
                access_token: tokenResponse.data.access_token,
                refresh_token: tokenResponse.data.refresh_token,
                calendarId: existingUser.data[0].provider?.azure?.calendarId,
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
      console.error("Token exchange error:", error);
    }
  };

  useEffect(() => {
    const handleAuthCallback = async () => {
      const params = new URLSearchParams(window.location.search);
      const code = params.get('code');
      const state = params.get('state');
      const storedState = localStorage.getItem('azure_auth_state');

      localStorage.removeItem('azure_auth_state');

      if (!code || !state || state !== storedState) {
        console.error("Invalid authorization response");
        return;
      }

      try {
        const tokenResponse = await axios.post('/api/auth/azure/token', { code });
        
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
                  access_token: tokenResponse.data.access_token,
                  refresh_token: tokenResponse.data.refresh_token,
                  calendarId: existingUser.data[0].provider?.azure?.calendarId,
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
        console.error("Token exchange error:", error);
      }
    };

    if (window.location.search.includes('code=')) {
      handleAuthCallback();
    }
  }, [searchParams]);

  const disconnect = async () => {
    let scopedUser = queryClient.getQueryData(["user"]) as any;
    await disconnectOutlookCalendar(jwt, scopedUser, queryClient);
  };

  const connect = async () => {
    const supabase = createClient();
    let scopedUser = queryClient.getQueryData(["user"]) as any;

    let updatedUser = await supabase
      .from("users")
      .update({
        provider: {
          ...scopedUser.provider,
          azure: {
            access_token: scopedUser.provider.azure.access_token,
            sync: scopedUser.provider.azure.sync,
            calendarId: scopedUser.provider.azure.calendarId,
            connected: true,
          },
        },
      })
      .eq("id", jwt.user)
      .select("*");

    if (updatedUser?.data) {
      queryClient.setQueryData(["user"], updatedUser.data[0]);
    }
  };

  const loadingArray = [
    isLoadingOutlookCalendars,
    isLoadingSupabaseUser,
    isTimeEntriesSyncMutationPending,
    isTimeOffSyncPending,
  ];

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
                <Button
                  onClick={
                    supabaseUser?.provider?.azure?.access_token
                      ? connect
                      : outlookLogin
                  }
                  className="bg-blue-600"
                >
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
                <div className="flex flex-row gap-5 items-center">
                  <div className="w-1/5">
                    <FormField
                      control={form.control}
                      name="outlookTimeOff"
                      render={({ field }) => (
                        <FormItem className="flex flex-row items-center justify-between rounded-lg border p-3 shadow-sm">
                          <div className="space-y-0.5">
                            <FormLabel>Time off</FormLabel>
                          </div>
                          <FormControl>
                            <Switch
                              checked={field.value}
                              onCheckedChange={(e) => {
                                field.onChange(e);
                                timeOffSync({
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
                  <h1>Time off will be displayed in your Outlook Calendar</h1>
                </div>
              </form>
            </Form>
          </div>
        </div>
      )}
    </div>
  );
}
