import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEditFormProps {
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
