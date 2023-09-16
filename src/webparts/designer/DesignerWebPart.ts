import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { Guid, Version } from "@microsoft/sp-core-library";
import {
    IPropertyPaneConfiguration,
    PropertyPaneCheckbox,
    PropertyPaneHorizontalRule,
    PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";

import {
    DesignerMini,
    DesignerMiniDoneEventData,
    IDesignerMiniTheme,
    IDesignerMiniView,
} from "@designer/mini";
import { EmbeddedDesignerApp, ForwardedAppConfig } from "@designerapp/embedded";
import { SPFI, SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as strings from "DesignerWebPartStrings";
import Designer from "./components/Designer";
import { IDesignerProps } from "./components/IDesignerProps";

export interface IDesignerWebPartProps {
    imageSrc: string;
    showDesigner: boolean;
    uploadToDocLibrary: boolean;
}

enum DesignerWebPartProperties {
    imageSrc = "imageSrc",
    showDesigner = "showDesigner",
    uploadToDocLibrary = "uploadToDocLibrary",
}

export default class DesignerWebPart extends BaseClientSideWebPart<IDesignerWebPartProps> {
    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = "";
    private _sp: SPFI;
    private _miniApp: DesignerMini;
    private _fullApp: EmbeddedDesignerApp;

    public render(): void {
        const element: React.ReactElement<IDesignerProps> = React.createElement(
            Designer,
            {
                imageSrc: this.properties.imageSrc,
                showDesigner: this.properties.showDesigner,
                isDarkTheme: this._isDarkTheme,
                environmentMessage: this._environmentMessage,
                hasTeamsContext: !!this.context.sdks.microsoftTeams,
                userDisplayName: this.context.pageContext.user.displayName,
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onInit(): Promise<void> {
        this._sp = spfi().using(SPFx(this.context));
        return this._getEnvironmentMessage().then((message) => {
            this._environmentMessage = message;
        });
    }

    private _getEnvironmentMessage(): Promise<string> {
        if (!!this.context.sdks.microsoftTeams) {
            // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app
                .getContext()
                .then((context) => {
                    let environmentMessage: string = "";
                    switch (context.app.host.name) {
                        case "Office": // running in Office
                            environmentMessage = this.context
                                .isServedFromLocalhost
                                ? strings.AppLocalEnvironmentOffice
                                : strings.AppOfficeEnvironment;
                            break;
                        case "Outlook": // running in Outlook
                            environmentMessage = this.context
                                .isServedFromLocalhost
                                ? strings.AppLocalEnvironmentOutlook
                                : strings.AppOutlookEnvironment;
                            break;
                        case "Teams": // running in Teams
                            environmentMessage = this.context
                                .isServedFromLocalhost
                                ? strings.AppLocalEnvironmentTeams
                                : strings.AppTeamsTabEnvironment;
                            break;
                        default:
                            throw new Error("Unknown host");
                    }

                    return environmentMessage;
                });
        }

        return Promise.resolve(
            this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentSharePoint
                : strings.AppSharePointEnvironment
        );
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        this._isDarkTheme = !!currentTheme.isInverted;
        const { semanticColors } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty(
                "--bodyText",
                semanticColors.bodyText || null
            );
            this.domElement.style.setProperty(
                "--link",
                semanticColors.link || null
            );
            this.domElement.style.setProperty(
                "--linkHovered",
                semanticColors.linkHovered || null
            );
        }
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0");
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneCheckbox("showDesigner", {
                                    text: strings.ShowDesignerLabel,
                                    checked: false,
                                }),
                                PropertyPaneCheckbox("uploadToDocLibrary", {
                                    text: strings.uploadToDocLibraryLabel,
                                    checked: false,
                                }),
                                PropertyPaneHorizontalRule(),
                                PropertyPaneTextField("imageSrc", {
                                    label: strings.ImageSourceFieldLabel,
                                }),
                                PropertyPaneHorizontalRule(),
                            ],
                        },
                    ],
                },
            ],
        };
    }

    // Update the property pane field value after the user selects a new image
    protected onPropertyPaneFieldChanged(
        propertyPath: string,
        oldValue: any,
        newValue: any
    ): void {
        if (propertyPath === "imageSrc" && newValue) {
            this.updatePropertyPane(
                DesignerWebPartProperties.imageSrc,
                newValue
            );
        }
        if (propertyPath === "showDesigner" && newValue) {
            if (newValue === true) {
                this.updatePropertyPane(DesignerWebPartProperties.imageSrc, "");
                this.createMiniApp();
                this.toggleImageVisibility(false);
            } else {
                this.closeMiniApp();
            }
        }
    }

    // Update the property pane field value

    private updatePropertyPane = (
        propertyPath: DesignerWebPartProperties,
        newValue: string
    ): void => {
        switch (propertyPath) {
            case DesignerWebPartProperties.imageSrc:
                this.properties.imageSrc = newValue;
                break;
            case DesignerWebPartProperties.showDesigner:
                this.properties.showDesigner = newValue === "true";
                break;
            case DesignerWebPartProperties.uploadToDocLibrary:
                this.properties.uploadToDocLibrary = newValue === "true";
                break;
        }
        this.context.propertyPane.refresh(); // Trigger a property pane refresh
        this.render();
    };

    // Upload the file to the Site Assets library
    private async uploadFileToLibrary(
        byteArray: Uint8Array,
        fileName: string
    ): Promise<string | undefined> {
        try {
            console.log("Starting upload...");

            // Create the Doc Lib if it doesn't exist
            const listEnsureResult = await this._sp.web.lists.ensure(
                "Designer Images",
                "Designer Images",
                101
            );

            // check if the list was created, or if it already existed:
            if (listEnsureResult.created) {
                console.log("Designer Images is created!");
            } else {
                console.log("Designer Images already exists!");
            }

            // get the folder where the file will be uploaded
            const folder =
                this._sp.web.getFolderByServerRelativePath("Designer Images");

            // create a new file in the folder
            const file = await folder.files.addUsingPath(
                encodeURI(fileName),
                byteArray,
                {
                    Overwrite: false,
                }
            );

            console.log(
                `File uploaded successfully. URL: ${file.data.ServerRelativeUrl}`
            );

            return file.data.ServerRelativeUrl;
        } catch (error) {
            console.error("Error uploading file:", error);
        }
    }

    // Handle the done event from Designer apps
    private doneHandler = async (
        data: DesignerMiniDoneEventData
    ): Promise<void> => {
        try {
            const media = data.media;
            let imageSource = "";
            if (this.properties.uploadToDocLibrary) {
                // Upload the file to the Site Assets library
                const fileExtension = data.mimeType.split("/")[1];
                const fileName = `DesignerExport_${Guid.newGuid().toString()}.${fileExtension}`;
                console.log(`Uploading file: ${fileName} to Style Library`);
                const result = await this.uploadFileToLibrary(media, fileName);

                if (result) {
                    imageSource = result;
                } else {
                    console.log("No Image URL received after upload.");
                }
            } else {
                // Convert the Uint8Array to a base64 string
                const base64String = Buffer.from(media).toString("base64");
                imageSource = `data:${data.mimeType};base64,${base64String}`;
            }

            // Update the property pane value
            this.updatePropertyPane(
                DesignerWebPartProperties.imageSrc,
                imageSource
            );

            // Show the image
            this.toggleImageVisibility(true);
        } catch (error) {
            console.log("Error while handling the done event:", error);
        }
    };

    private createFullApp = (data: ForwardedAppConfig): void => {
        const designerFullContainer = document.getElementById(
            "full-container"
        ) as HTMLDivElement;
        designerFullContainer.style.display = "block";
        this._fullApp = new EmbeddedDesignerApp({
            baseAppURL: new URL("https://designer.microsoft.com"),
            clientName: "MiniTestApp",
            clientId: "uuid-1234",
            container: designerFullContainer,
            sessionId: "uuid-1234",
            correlationId: "uuid-1234",
            forwardedConfig: data,
            insertMode: true,
            disableEmbeddedCSPEnforcement: true,
            enableSandboxing: true,
            suggestionsConfig: { scenario: "TestApp-Form" },
            useEnterpriseTOULink: true,
            platform: "Web",
            hostBuildVersion: "0.0.0",
            hostEnvironment: "TestApp-HostEnv",
            hostScenario: "TestApp-HostScenario",
        });
        const closeApp = (): void => {
            this._fullApp.destroy();
            designerFullContainer.style.display = "none";

            // Close the mini app
            this.closeMiniApp();
        };
        this._fullApp.on("done", async (data: DesignerMiniDoneEventData) => {
            await this.doneHandler(data);
            closeApp();
        });
        this._fullApp.on("cancel", () => {
            closeApp();
        });
        this._fullApp.initialize().catch((error: any) => {
            console.error(error);
        });
    };

    private createMiniApp = (): void => {
        const designerMiniContainer = document.getElementById(
            "mini-container"
        ) as HTMLDivElement;
        designerMiniContainer.style.display = "block";
        this._miniApp = new DesignerMini({
            // refer to our API docs on the DesignerMini Class to adjust the settings here.
            miniURL: new URL(
                "https://cdn.designerapp.osi.office.net/mini-app/index.html"
            ),
            serviceBaseURL: new URL("https://designerapp.officeapps.live.com"),
            // give your container element a width and height, the iframe will fill it.
            container: designerMiniContainer,
            telemetryConfig: {
                audienceGroup: "Other",
                clientId: "uuid-1234",
                clientName: "MiniTestApp",
                hostBuildVersion: "0.0.0",
                hostEnvironment: "TestApp-HostEnv",
                isMicrosoftInternal: true,
                optionalDiagnosticsAllowed: true,
                sessionId: "uuid-1234",
                correlationId: "uuid-1234",
                platform: "Web",
            },
            disableEmbeddedCSPEnforcement: true,
            enableSandboxing: true,
            insertMode: true,
            hideSizeSelector: false,
            preferredImageOutputFormat: "jpg",
            useEnterpriseTOULink: true,
            viewConfig: {
                view: IDesignerMiniView.Pane,
                theme: IDesignerMiniTheme.Auto,
            },
        });

        // Initialize with optional design suggestions
        this._miniApp.initialize().catch((error: any) => {
            console.error(error);
        });

        this._miniApp.on("done", async (data: DesignerMiniDoneEventData) => {
            await this.doneHandler(data);
            this.closeMiniApp();
        });

        this._miniApp.on("customize", (data: ForwardedAppConfig) => {
            this.createFullApp(data);
        });
    };

    private closeMiniApp = (): void => {
        const designerMiniContainer = document.getElementById(
            "mini-container"
        ) as HTMLDivElement;
        this._miniApp.destroy();
        this.updatePropertyPane(
            DesignerWebPartProperties.showDesigner,
            "false"
        );
        designerMiniContainer.style.display = "none";
    };

    private toggleImageVisibility = (show: boolean): void => {
        // Show the image
        const imageContainer = document.getElementById(
            "designer-image"
        ) as HTMLDivElement;
        imageContainer.style.visibility = show ? "visible" : "hidden";
    };
}
