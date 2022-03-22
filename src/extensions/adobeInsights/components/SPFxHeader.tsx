import * as React from 'react';
import styles from './SPFxHeader.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IconButton, Label, Stack, TextField } from 'office-ui-fabric-react';
import { spfi, SPFI } from '@pnp/sp';
import { getSP } from '../../../services/pnpjsconfig';

export interface ISPFxHeaderProps {
    text: string;
}
export interface ISPFxHeaderState {
    searchKeyword: string;
}

export class SPFxHeader extends React.Component<ISPFxHeaderProps, ISPFxHeaderState> {
    private _selectedLocation: string;
    private LOG_SOURCE = "🅿PnPjsExample";
    private LIBRARY_NAME = "Documents";
    private _sp: SPFI;
    constructor(props: ISPFxHeaderProps) {
        super(props);
        // this.state = {
        //   searchKeyword: ""
        // };
        this._sp = getSP();
    }

    private _onSearchHandler = async (): Promise<void> => {
        //window.location.href = "https://www.bing.com/search?q=";// + this.state.searchKeyword;
        const result = await spfi(this._sp).web();
        console.log("result: ", result);

    }
    private _onKeywordChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        // this.setState({ searchKeyword: newValue });
    }

    public render(): React.ReactElement<ISPFxHeaderProps> {
        return (
            <Stack>
                <Stack tokens={{ childrenGap: 2 }} horizontal>
                    <Label>SPFxHeader: </Label>
                    <TextField placeholder="Please enter search keyword here" width={300} onChange={this._onKeywordChanged} />
                    <IconButton iconProps={{ iconName: 'Search' }} title="Search" ariaLabel="Search" onClick={this._onSearchHandler} />
                </Stack>
            </Stack>

        );
    }
}
