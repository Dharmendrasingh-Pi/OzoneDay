export interface IOzoneDayProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  spHttpClient: SPHttpClient;
  listWebUrl: string;
}
import { SPHttpClient } from "@microsoft/sp-http";
