import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IFormRequestProps {
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
