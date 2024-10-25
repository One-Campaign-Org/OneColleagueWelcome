declare interface IColleagueWelcomeWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PrefixFieldLabel: string;
  CustomCssLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'ColleagueWelcomeWebPartStrings' {
  const strings: IColleagueWelcomeWebPartStrings;
  export = strings;
}
