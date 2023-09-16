declare interface IDesignerWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    ImageSourceFieldLabel: string;
    ShowDesignerLabel: string;
    uploadToDocLibraryLabel: string;
    AppLocalEnvironmentSharePoint: string;
    AppLocalEnvironmentTeams: string;
    AppLocalEnvironmentOffice: string;
    AppLocalEnvironmentOutlook: string;
    AppSharePointEnvironment: string;
    AppTeamsTabEnvironment: string;
    AppOfficeEnvironment: string;
    AppOutlookEnvironment: string;
}

declare module "DesignerWebPartStrings" {
    const strings: IDesignerWebPartStrings;
    export = strings;
}
