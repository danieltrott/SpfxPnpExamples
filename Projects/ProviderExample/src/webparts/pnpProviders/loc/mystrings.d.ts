declare interface IPnpProvidersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'PnpProvidersWebPartStrings' {
  const strings: IPnpProvidersWebPartStrings;
  export = strings;
}
