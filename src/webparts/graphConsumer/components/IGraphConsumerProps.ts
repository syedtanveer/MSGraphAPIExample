import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGraphConsumerProps {
  context: WebPartContext;
}

export interface IGraphConsumerState {
  users: Array<IUserItem>;
  searchFor: string;
}

export interface IUserItem {
  displayName: string;
  mail: string;
  userPrincipalName: string;
}
