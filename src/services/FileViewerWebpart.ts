import { ClientsideWebpart, ColumnControl } from "@pnp/sp/clientside-pages";
import "@pnp/sp/webs";

// we create a class to wrap our functionality in a reusable way
export class FileViewerWebpart {// extends ClientsideWebpart {
    private json;
    constructor(control: ClientsideWebpart) {
        //super((<any>control).json);
        this.json = (<any>control).json;
    }
    public get EmbeddedFileType(): string {
        return this.EmbeddedFileUrl.toLocaleLowerCase().split("?")[0].split("#")[0].split('.').pop();
    }
    public get EmbeddedFileUrl(): string {
        return this.json.webPartData?.properties?.file || "";
    }
    public get IsFileViewer() {
        return this.json.webPartId == "b7dd04e1-19ce-4b24-9132-b60a1c2b910d";
    }

}