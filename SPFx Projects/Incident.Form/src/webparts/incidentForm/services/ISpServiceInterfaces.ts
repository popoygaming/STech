export interface ISharePointListItemResponse {
  Id: number;
}

export interface IIncidentListItemPayload {
  [fieldTitle: string]: string | undefined;
}

export interface IChoiceFieldResponse {
  Choices?: string[] | { results: string[] };
}

export interface ISharePointFieldResponse {
  Title: string;
  InternalName: string;
}

export interface ISharePointFieldCollectionResponse {
  value?: ISharePointFieldResponse[];
  d?: {
    results?: ISharePointFieldResponse[];
  };
}
