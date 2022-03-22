import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { IInsightItem } from './IInsightItem';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IWebAddResult } from "@pnp/sp/webs"; import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import "@pnp/sp/webs";
import "@pnp/sp/profiles";
import { Web } from "@pnp/sp/webs";
import { ClientsideWebpart, IClientsidePage } from '@pnp/sp/clientside-pages';
import { AssignFrom } from "@pnp/core";

import { handleError } from './ErrorHandler';
import { getSP } from './pnpjsconfig';
import { FileViewerWebpart } from './FileViewerWebpart';


export class InsightService {
    constructor(private _context: ApplicationCustomizerContext) {

    }

    public async getInsightItme(): Promise<IInsightItem> {
        let ret: IInsightItem = {
            jobTitle: "",
            department: "",
            location: "",
            userId: "",
            employeeType: "",
            pageUrl: "",
            likeCount: 0,
            videoFileUrl: "",
            VideoFileViews: 0
        };
        const sp = getSP();
        const result = await sp.web();
        const currentPageUrl = new URL(window.location.href);
        ret.pageUrl = currentPageUrl.origin + currentPageUrl.pathname;
        // get current page like info
        const currentPageFile = await sp.web.getFileByServerRelativePath(currentPageUrl.pathname);
        const currentPageListItem = await currentPageFile.getItem();
        const likeInfo = await currentPageListItem.getLikedByInformation();
        ret.likeCount = likeInfo.likeCount;

        // Find File Viewer web part. 
        const page: IClientsidePage = await sp.web.loadClientsidePage(currentPageUrl.pathname);
        const ctrlFileViewer = page.findControl<ClientsideWebpart>((c: ClientsideWebpart) => {
            let ret = false;
            const wp = new FileViewerWebpart(c);
            ret = wp.IsFileViewer;
            return ret;
        });

        if (ctrlFileViewer) {
            const wp = new FileViewerWebpart(ctrlFileViewer);
            if (wp.EmbeddedFileType == "mp4") {
                ret.videoFileUrl = wp.EmbeddedFileUrl;
                const videoUrl = new URL(wp.EmbeddedFileUrl);
                const videoSCUrl = `${videoUrl.origin}/${videoUrl.pathname.split("/")[1]}/${videoUrl.pathname.split("/")[2]}`;
                const videoSP = spfi(videoSCUrl).using(AssignFrom(sp.web));
                const videoFile = await videoSP.web.getFileByUrl(videoUrl.href);
                const videoFileListItem = await videoFile.getItem();
                const apiUrl = (videoFileListItem as any)._url;

                const graphUrl = `https://graph.microsoft.com/v1.0/sites/${videoUrl.hostname}:/${videoUrl.pathname.split("/")[1]}/${videoUrl.pathname.split("/")[2]}:/lists/${apiUrl.split("'")[1]}/items/${videoFileListItem["ID"]}/driveItem/analytics/alltime`;
                console.log("video file list item:", videoFileListItem);

                const msGraphClient = await this._context.msGraphClientFactory.getClient();
                const viewsResult = await msGraphClient.api(graphUrl).get();
                ret.VideoFileViews = +viewsResult?.access?.actionCount;
            }
        }

        // get user profile. 
        const profile = await sp.profiles.myProperties();
        profile.UserProfileProperties.forEach(prop => {
            if (prop.Key == "Department") ret.department = prop.Value;
            if (prop.Key == "Title") ret.jobTitle = prop.Value;
            if (prop.Key == "Office") ret.location = prop.Value;
            if (prop.Key == "msOnline-ObjectId") ret.userId = prop.Value;
        });

        return ret;

    }
}