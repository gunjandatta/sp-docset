import { ContextInfo, Helper, List, SPTypes } from "gd-sprest-bs";

/** List Name */
export const ListName = "Doc Set Demo";

/**
 * Configuration
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: ListName,
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                ContentTypesEnabled: true
            },
            ContentTypes: [
                {
                    Name: "Dashboard Item",
                    Description: "Displayed in the ribbon new item text.",
                    ParentName: "Document Set",
                    onCreated: () => {
                        // Get the list document set home page
                        List(ListName).RootFolder().Folders("Forms").Folders("Document Set").Files("docsethomepage.aspx").execute(page => {
                            // Add the dashboard webpart
                            Helper.WebPart.addWebPartToPage(page.ServerRelativeUrl, {
                                title: "Dashboard",
                                description: "Displays the custom dashboard.",
                                chromeType: "None",
                                index: 0,
                                zone: "WebPartZone_Top",
                                content: '<div id="dashboard"></div>'
                            });

                            // Add the reference to the library and initialize script
                            // Add the dashboard webpart
                            Helper.WebPart.addWebPartToPage(page.ServerRelativeUrl, {
                                title: "Library Reference",
                                description: "Loads the dashboard library and initializes the solution.",
                                chromeType: "None",
                                index: 0,
                                zone: "WebPartZone_Bottom",
                                content: [
                                    '<script type="text/javascript" src="' + ContextInfo.webServerRelativeUrl + '/siteassets/sp-docset.js"></script>',
                                    '<script type="text/javascript">',
                                    '// Wait for the library to be loaded',
                                    'SP.SOD.executeOrDelayUntilScriptLoaded(function() {',
                                    '\t// Render the dashboard',
                                    '\tDocSetDemo.Dashboard(document.querySelector("#dashboard"));',
                                    '}, "docset-demo");',
                                    '</script>'
                                ].join('\n')
                            });
                        });
                    }
                }
            ]
        }
    ]
});