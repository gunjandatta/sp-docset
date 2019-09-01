import { Components, Helper, ContextInfo } from "gd-sprest-bs";
import { DataSource } from "./ds";

/**
 * Dashboard
 */
export const Dashboard = (el: HTMLElement) => {
    // Render a jumbotron
    Components.Jumbotron({
        el,
        title: "Document Set Dashboard",
        content: [
            '<h3>Welcome</h3>',
            '<p>You can customize this dashboard to display whatever you want</p>'
        ].join('\n')
    });

    // Display a loading dialog
    Helper.SP.ModalDialog.showWaitScreenWithNoClose("Loading the Information").then(dlg => {
        // Load the item information
        DataSource.getItemInfo().then(
            // Success
            info => {
                // Render cards
                Components.CardGroup({
                    el,
                    isDeck: true,
                    cards: [
                        {
                            body: [
                                {
                                    title: "Item Information",
                                    onRender: (el, card) => {
                                        // Render a list form
                                        Components.ListForm.renderDisplayForm({
                                            el,
                                            info
                                        });
                                    }
                                }
                            ]
                        },
                        {
                            body: [
                                {
                                    title: "Menu",
                                    onRender: (el, card) => {
                                        // Render buttons
                                        Components.ButtonGroup({
                                            el,
                                            isVertical: true,
                                            buttons: [
                                                {
                                                    text: "View Item",
                                                    type: Components.ButtonTypes.Primary,
                                                    onClick: () => {
                                                        // Display the item form in a modal dialog
                                                        Helper.SP.ModalDialog.showModalDialog({
                                                            title: "View Item",
                                                            url: ContextInfo.listUrl + "/Forms/DispForm.aspx?ID=" + info.item.Id
                                                        });
                                                    }
                                                },
                                                {
                                                    text: "Edit Item",
                                                    type: Components.ButtonTypes.Secondary,
                                                    onClick: () => {
                                                        // Display the item form in a modal dialog
                                                        Helper.SP.ModalDialog.showModalDialog({
                                                            title: "Edit Item",
                                                            url: ContextInfo.listUrl + "/Forms/EditForm.aspx?ID=" + info.item.Id,
                                                            dialogReturnValueCallback: result => {
                                                                // See if an update occurred
                                                                if (result == 1) {
                                                                    // Reload the page
                                                                    document.location.reload();
                                                                }
                                                            }
                                                        });
                                                    }
                                                }
                                            ]
                                        });
                                    }
                                }
                            ]
                        }
                    ]
                });

                // Close the dialog
                dlg.close();
            },
            // Error
            () => {
                // Render an error
                Components.Alert({
                    el,
                    header: "Error Getting Data",
                    content: "Please <a href='#' onclick='javascript:document.location.href = document.location.href'>click here</a> to refresh the page",
                    type: Components.AlertTypes.Danger
                });

                // Close the dialog
                dlg.close();
            }
        );
    });
}