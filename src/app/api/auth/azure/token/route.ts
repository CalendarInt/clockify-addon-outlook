import { NextResponse } from "next/server";
import axios from "axios";

export async function POST(request: Request) {
  try {
    const { code, code_verifier } = await request.json();

    if (!process.env.NEXT_PUBLIC_AZURE_CLIENT_ID) {
      console.error("Missing Azure client ID in environment variables");
      return NextResponse.json(
        { error: "Server configuration error" }, 
        { status: 500 }
      );
    }

    // Exchange authorization code for tokens
    const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    const params = new URLSearchParams({
      client_id: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID,
      code: code,
      code_verifier: code_verifier,
      redirect_uri: process.env.NEXT_PUBLIC_REDIRECT_URI!,
      grant_type: "authorization_code",
      scope: "offline_access Calendars.ReadWrite"
    });

    console.log('Token request params:', {
      client_id: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID,
      redirect_uri: process.env.NEXT_PUBLIC_REDIRECT_URI,
    });

    const response = await axios.post(tokenEndpoint, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000"
      }
    });

    console.log('Token response:', response.data);

    return NextResponse.json({
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token,
      expires_in: response.data.expires_in
    });

  } catch (error: any) {
    console.error("Token exchange error details:", {
      status: error.response?.status,
      data: error.response?.data,
      message: error.message
    });
    
    return NextResponse.json(
      { 
        error: "Failed to exchange code for tokens",
        details: error.response?.data || error.message
      }, 
      { status: error.response?.status || 500 }
    );
  }
} 