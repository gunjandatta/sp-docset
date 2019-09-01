import { Helper, ContextInfo } from "gd-sprest-bs";

/**
 * Data Source
 */
export const DataSource = {
    /**
     * Gets the item id from the query string value "ID"
     */
    getItemId: (): number => {
        let itemId = 0;

        // Get the item id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 0 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');

            // See if this is the "id" key
            if (qsItem[0] == "ID") {
                // Set the item id
                itemId = parseInt(qsItem[1]);
                break;
            }
        }

        // Return the item id
        return itemId;
    },

    /**
     * Gets the item information
     */
    getItemInfo: (): PromiseLike<Helper.IListFormResult> => {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the item id
            let itemId = DataSource.getItemId();
            if (itemId) {
                // Get the item information
                Helper.ListForm.create({
                    listName: ContextInfo.listTitle,
                    itemId,
                    fields: ["Title", "DocumentSetDescription"]
                }).then(resolve, reject);
            } else {
                // Reject the promise
                reject();
            }
        });
    }
}