/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import { getSP } from "../../../pnpjs-config";

class TransmittalService {

    private sp!: SPFI;
    private context!: FormCustomizerContext;

    // public async init(context: FormCustomizerContext): Promise<SPFI> {
    //     this.sp = getSP(context);
    //     this.context = context;

    //     return this.sp;
    // }
    public async init(context: FormCustomizerContext) {
        if (!!context) {
            this.sp = getSP(context);
            this.context = context;
        }
        return this.sp;
    }
    public async getBaseLineContentTypeId() {
        if ('list' in this.context) {
            const cType = await this.sp.web.lists.getByTitle(this.context.list.title)();
            console.log(cType);


        }
    }
    public async createListItem<T extends Record<string, unknown>>(listTitle: string, object: T) {
        const i = await this.sp.web.lists.getByTitle(listTitle).items.add(object);
        return i;
    }
    public async updateListItem(listTitle: string, object: any, itemId: number) {
        const i = await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update(object);
        return i;
    }

    public async getUserNameById(userId: number): Promise<string> {
        try {
            const user = await this.sp.web.siteUsers
                .getById(userId)();

            return user.Title;
        } catch (error) {
            console.error("Error getting user:", error);
            return "";
        }
    }

    

    // public async checkForDuplicates(projectReference: string, listTitle: string): Promise<string | undefined> {
    //     try {
    //         // Construct CAML query to check for duplicate ProjectReference
    //         let filterQuery = `ProjectReference eq '${projectReference}'`;
    //         //   if (itemId) {
    //         //     filterQuery += ` and Id ne ${itemId}`;
    //         //   }

    //         // Query the SharePoint list
    //         const items = await this.sp.web.lists.getByTitle(listTitle).items
    //             .filter(filterQuery)
    //             .select("Id")();

    //         // If items are found, return error message for duplicate
    //         if (items.length > 0) {
    //             return "A transmittal with this Reference already exists.";
    //         }
    //         return undefined;
    //     } catch (error) {
    //         console.error("Error checking for duplicates:", error);
    //         return "Error validating reference. Please try again.";
    //     }
    // }

}
export default new TransmittalService();