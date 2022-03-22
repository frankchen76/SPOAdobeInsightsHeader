import * as React from 'react';
import { useState } from "react";
import * as ReactDom from 'react-dom';
import { Callout, IconButton, Label, Stack, TextField } from "office-ui-fabric-react";
import { Dialog } from '@microsoft/sp-dialog';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IWebAddResult } from "@pnp/sp/webs"; import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { handleError } from '../../../services/ErrorHandler';
import { getSP } from '../../../services/pnpjsconfig';
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import { ClientsideWebpart, IClientsidePage } from '@pnp/sp/clientside-pages';
import { AssignFrom } from "@pnp/core";
import { FileViewerWebpart } from '../../../services/FileViewerWebpart';
import { InsightService } from '../../../services/InsightService';
import { IInsightItem } from '../../../services/IInsightItem';

export interface IInsightComponentProps {
    title: string;
    context: ApplicationCustomizerContext;
    item: IInsightItem;
};
export const InsightComponent = (props: IInsightComponentProps) => {
    const [isCalloutVisible, setIsCalloutVisible] = useState<Boolean>(false);
    const _onDismissHandler = (ev?: any) => {
        setIsCalloutVisible(!isCalloutVisible);
    };
    const _onSearchHandler = async (): Promise<void> => {
        setIsCalloutVisible(!isCalloutVisible);
        // const service = new InsightService(props.context);
        // const item = await service.getInsightItme();
        // console.log("InsightItem", item);
        // const sp = getSP();
        // const result = await sp.web();
        // Dialog.alert(result.Title);
        // const currentPageUrl = new URL(window.location.href);
        // const page: IClientsidePage = await sp.web.loadClientsidePage(currentPageUrl.pathname);
        // // Find File Viewer web part. 
        // //const ctrlFileViewer = page.findControlById("62dd52a3-8b2d-4584-9f2c-e26db4d7a125");
        // const ctrlFileViewer = page.findControl<ClientsideWebpart>((c: ClientsideWebpart) => {
        //     let ret = false;
        //     const wp = new FileViewerWebpart(c);
        //     ret = wp.IsFileViewer;
        //     return ret;
        // });

        // // ctrlFileViewer.
        // if (ctrlFileViewer) {
        //     const wp = new FileViewerWebpart(ctrlFileViewer);
        //     const videoUrl = new URL(wp.EmbeddedFileUrl);
        //     const videoSCUrl = `${videoUrl.origin}/${videoUrl.pathname.split("/")[1]}/${videoUrl.pathname.split("/")[2]}`;
        //     const videoSP = spfi(videoSCUrl).using(AssignFrom(sp.web));
        //     const videoFile = await videoSP.web.getFileByUrl(videoUrl.href);
        //     const videoFileListItem = await videoFile.getItem();
        //     const apiUrl = (videoFileListItem as any)._url;

        //     //const videoList = await (await videoFileListItem.list()).Id;

        //     //const videoList = await videoSP.web.lists.getByTitle(videoUrl.pathname.split("/")[3]).select("Id", "Title")();

        //     //const videoList = await videoFileListItem.list;

        //     const graphUrl = `https://graph.microsoft.com/v1.0/sites/${videoUrl.hostname}:/${videoUrl.pathname.split("/")[1]}/${videoUrl.pathname.split("/")[2]}:/lists/${apiUrl.split("'")[1]}/items/${videoFileListItem["ID"]}/driveItem/analytics/alltime`;
        //     //console.log("video file list item:", videoFileListItem);

        //     const msGraphClient = await props.context.msGraphClientFactory.getClient();
        //     const viewsResult = await msGraphClient.api(graphUrl).get();
        //     const views = viewsResult?.access?.actionCount;

        //     console.log(wp.EmbeddedFileUrl);
        // }
        // console.log(currentPageUrl);


    };
    return (
        <Stack>
            <Stack tokens={{ childrenGap: 2 }} horizontal>
                <IconButton
                    id="header-btn"
                    iconProps={{ iconName: 'Info' }}
                    title="Search"
                    ariaLabel="Search"
                    onClick={_onSearchHandler} />
                {isCalloutVisible && <Callout
                    gapSpace={0}
                    target={`#header-btn`}
                    onDismiss={_onDismissHandler}
                    setInitialFocus
                >
                    <pre >{props.item ? JSON.stringify(props.item, null, 4) : "No data"}</pre>
                </Callout>}
                <Label>{props.title}</Label>
            </Stack>
        </Stack>
    );
};

