import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log, SPEventArgs } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AdobeInsightsApplicationCustomizerStrings';
import { InsightComponent } from './components/InsightsComponent';
import {
    Logger,
    ConsoleListener,
    LogLevel,
    PnPLogging
} from "@pnp/logging";
import { getSP } from '../../services/pnpjsconfig';
import { IInsightItem } from '../../services/IInsightItem';
import { InsightService } from '../../services/InsightService';


const LOG_SOURCE: string = 'AdobeInsightsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAdobeInsightsApplicationCustomizerProperties {
    // This is an example; replace with your own property
    testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AdobeInsightsApplicationCustomizer
    extends BaseApplicationCustomizer<IAdobeInsightsApplicationCustomizerProperties> {
    private _topPlaceholder: PlaceholderContent | undefined;

    public async onInit(): Promise<void> {
        Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

        let message: string = this.properties.testMessage;
        if (!message) {
            message = '(No properties were provided.)';
        }

        //Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

        //const graph = graphfi().using(SPFx(this.context));

        Logger.subscribe(ConsoleListener(LOG_SOURCE, { color: '#0b6a0b', warning: 'magenta' }));
        Logger.activeLogLevel = LogLevel.Info;

        getSP(this.context);

        this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
        this.context.application.navigatedEvent.add(this, (args: SPEventArgs) => {
            //console.log(`navigatedEvent was called.`, args);
            setTimeout(async () => {
                console.log("starting...");
                const service = new InsightService(this.context);
                const item = await service.getInsightItme();
                this._renderPlaceHolders(item);
                console.log("updated.");
            }, 50);

        });

        return Promise.resolve();
    }

    private _onDispose(): void {
        console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
    }

    private _renderPlaceHolders(item?: IInsightItem): void {
        console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
        console.log(
            "Available placeholders: ",
            this.context.placeholderProvider.placeholderNames
                .map(name => PlaceholderName[name])
                .join(", ")
        );
        console.log("item:", item);

        if (!this._topPlaceholder) {
            this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
                PlaceholderName.Top,
                { onDispose: this._onDispose }
            );
        }
        if (this._topPlaceholder) {
            const element = React.createElement(
                InsightComponent,
                {
                    title: item && item.pageUrl != undefined ? "loaded" : "loading",
                    context: this.context,
                    item: item
                }
            );
            // const element = React.createElement(
            //     SPFxHeader,
            //     {
            //         text: "(Top property was not defined.)"
            //     }
            // );

            ReactDom.render(element, this._topPlaceholder.domElement);// as React.Component<IHeaderProps, React.ComponentState, any>;

        } else {
            console.error("The expected placeholder (Top) was not found.");
        }

    }

}
