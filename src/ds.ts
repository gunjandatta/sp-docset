import { Helper, ContextInfo, Web } from "gd-sprest-bs";

/**
 * Data Source
 */
export const DataSource = {
    /**
     * Gets the item id from the query string value "ID"
     */
    getItemId: (): PromiseLike<number> => {
        // Return a promise
        return new Promise((resolve, reject) => {
            let folderUrl = null;

            // Get the item id from the querystring
            let qs = document.location.search.split('?');
            qs = qs.length > 0 ? qs[1].split('&') : [];
            for (let i = 0; i < qs.length; i++) {
                let qsItem = qs[i].split('=');
                let key = qsItem[0];
                let value = qsItem[1];

                // See if this is the "id" key
                if (key == "ID") {
                    // Resolve the promise
                    resolve(parseInt(value));
                    return;
                }
                // Else, see if this is the root folder url
                else if (key == "RootFolder") {
                    // Set the folder url
                    folderUrl = value;
                }
            }

            // See if the folder url exists
            if (folderUrl) {
                // Get the folder's associate list item
                Web().getFolderByServerRelativeUrl().ListItemAllFields().execute(item => {
                    // Resolve the promise
                    resolve(item.Id);
                }, reject);
            }

            // Reject the promise
            reject("Unable to find item information in querystring.");
        });
    },

    /**
     * Gets the item information
     */
    getItemInfo: (): PromiseLike<Helper.IListFormResult> => {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the item id
            DataSource.getItemId().then(itemId => {
                // Get the item information
                Helper.ListForm.create({
                    listName: ContextInfo.listTitle,
                    itemId,
                    fields: ["Title", "DocumentSetDescription"]
                }).then(resolve, reject);
            }, reject);
        });
    }
}