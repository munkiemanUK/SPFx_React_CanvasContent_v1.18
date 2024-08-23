import {SPHttpClient} from '@microsoft/sp-http';

export interface ICanvasContentProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteUrl: string;
  context: any;
  spHttpClient: SPHttpClient;
  numGroups : number;
  useList : string;  
}
