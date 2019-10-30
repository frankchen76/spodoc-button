import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { PrimaryButton, autobind, Modal } from 'office-ui-fabric-react';

import * as strings from 'CallAzureFunctionCommandSetStrings';
import { HttpClientConfiguration, HttpClient } from '@microsoft/sp-http';
import { setup as pnpSetup } from "@pnp/common";
import { sp } from "@pnp/sp";
import ProgressDialog from './ProgressDialog';
import { ISetting } from './ISetting';
import { ISettingListItem } from './ISettingListItem';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICallAzureFunctionCommandSetProperties {
    // This is an example; replace with your own properties
    sampleTextOne: string;
    sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CallAzureFunctionCommandSet';

export default class CallAzureFunctionCommandSet extends BaseListViewCommandSet<ICallAzureFunctionCommandSetProperties> {
    private SETTING_LISTNAME = "Settings";
    private _commandSettings: ISetting[];
    @override
    public onInit(): Promise<void> {
        return super.onInit().then(_ => {
            Log.info(LOG_SOURCE, 'Initialized CallAzureFunctionCommandSet');
            pnpSetup({
                spfxContext: this.context
            });

            return this._loadCommandSettings().then(result => { this._commandSettings = result; });
        });
    }

    @autobind
    private _loadCommandSettings(): Promise<ISetting[]> {
        return sp.web.lists.getByTitle(this.SETTING_LISTNAME).items.get().then((items: ISettingListItem[]): ISetting[] => {
            let ret = new Array<ISetting>();
            items.map(item => {
                ret.push({
                    title: item.Title,
                    setting: JSON.parse(item.Value)
                });
            });
            return ret;
        });
    }
    @autobind
    private _initCommand(cmd: Command, setting: ISetting, event: IListViewCommandSetListViewUpdatedParameters) {
        const needToShow = setting.setting.displayLists.indexOf(this.context.pageContext.list.title) != -1;
        cmd.visible = needToShow;
        cmd.title = setting.setting.title;
        cmd.iconImageUrl = setting.setting.iconImage;
    }

    @override
    public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
        if (this._commandSettings != null && this._commandSettings.length > 0) {
            this._commandSettings.forEach(setting => {
                const cmd: Command = this.tryGetCommand(setting.title);
                if (cmd) {
                    this._initCommand(cmd, setting, event);
                }
            });
        }
    }

    @autobind
    private _runAzureFunctionHandler(cmdTitle: string) {
        if (this._commandSettings != null && this._commandSettings.length > 0) {
            const res = this._commandSettings.filter(setting => { return setting.title == cmdTitle; });
            if (res != null && res.length == 1 && res[0].setting.apiUrl != null) {
                const setting = res[0];
                const dlg: ProgressDialog = new ProgressDialog();
                //dlg.title="Status";
                dlg.message = "running Azure Function";
                dlg.show();
                this.context.httpClient.get(setting.setting.apiUrl, HttpClient.configurations.v1).then(result => {
                    return dlg.close();
                }).then(_ => {
                    Dialog.alert("The Azure Function was completed successfully");
                });
            } else {
                Dialog.alert("Please add 'AzureFunctionUrl' to 'Settings' list");
            }
        }
    }

    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
        switch (event.itemId) {
            case 'COMMAND_1':
                this._runAzureFunctionHandler("COMMAND_1");
                break;
            case 'COMMAND_2':
                Dialog.alert(`${this.properties.sampleTextTwo}`);
                break;
            default:
                throw new Error('Unknown command');
        }
    }
}
