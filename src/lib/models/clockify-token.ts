export interface ClockifyToken {
  addonId: string;
  backendUrl: string;
  exp: number;
  iat: number;
  iss: string;
  jti: string;
  language: string;
  locationsUrl: string;
  ptoUrl: string;
  reportsUrl: string;
  screenshotsUrl: string;
  sub: string;
  type: string;
  user: string;
  workspaceId: string;
  workspaceRole: string;
}
