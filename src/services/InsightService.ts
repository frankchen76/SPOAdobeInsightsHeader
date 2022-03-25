import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { IInsightItem } from './IInsightItem';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IWebAddResult } from "@pnp/sp/webs"; import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import "@pnp/sp/webs";
import "@pnp/sp/profiles";
import "@pnp/sp/folders";
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
        let sp: SPFI = getSP();
        const result = await sp.web();
        const currentPageUrlObj = new URL(window.location.href);
        let currentPageUrl = currentPageUrlObj.origin + currentPageUrlObj.pathname;

        if (currentPageUrl.toLocaleLowerCase().indexOf(".aspx") == -1) {
            // home page for the site. 
            const currentSP = getSP();
            sp = spfi(currentPageUrl).using(AssignFrom(currentSP.web));
            // get home page
            const rootFolder = await sp.web.rootFolder();
            currentPageUrl = `${currentPageUrl}/${rootFolder.WelcomePage}`;
        }

        ret.pageUrl = currentPageUrl;
        // get current page like info
        //const currentPageFile = await sp.web.getFileByServerRelativePath(currentPageUrlObj.pathname);
        const currentPageFile = await sp.web.getFileByUrl(currentPageUrl);
        const currentPageListItem = await currentPageFile.getItem();
        const likeInfo = await currentPageListItem.getLikedByInformation();
        ret.likeCount = likeInfo.likeCount;

        // Find File Viewer web part. 
        const page: IClientsidePage = await sp.web.loadClientsidePage(new URL(currentPageUrl).pathname);
        const ctrlFileViewer = page.findControl<ClientsideWebpart>((c: ClientsideWebpart) => {
            let ret = false;
            const wp = new FileViewerWebpart(c);
            // TODO: add additional check if multiple FileViewers are included.  // && wp.EmbeddedFileType=="mp4"
            ret = wp.IsFileViewer;
            return ret;
        });

        if (ctrlFileViewer) {
            const wp = new FileViewerWebpart(ctrlFileViewer);
            // TODO: add additional file extensions
            if (wp.EmbeddedFileType == "mp4") {
                ret.videoFileUrl = wp.EmbeddedFileUrl;
                const videoUrlObj = new URL(wp.EmbeddedFileUrl);
                const videoSCUrl = `${videoUrlObj.origin}/${videoUrlObj.pathname.split("/")[1]}/${videoUrlObj.pathname.split("/")[2]}`;
                const videoSP = spfi(videoSCUrl).using(AssignFrom(sp.web));
                const videoFile = await videoSP.web.getFileByUrl(videoUrlObj.href);
                const videoFileListItem = await videoFile.getItem();
                const videoParentInfo = await videoFileListItem.getParentInfos();
                const apiUrl = (videoFileListItem as any)._url;

                // apiUrl.split("'")[1]
                const graphUrl = `https://graph.microsoft.com/v1.0/sites/${videoUrlObj.hostname}:/${videoUrlObj.pathname.split("/")[1]}/${videoUrlObj.pathname.split("/")[2]}:/lists/${videoParentInfo.ParentList.Id}/items/${videoFileListItem["ID"]}/driveItem/analytics/alltime`;
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