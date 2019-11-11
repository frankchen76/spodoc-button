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
import { HttpClientConfiguration, HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { setup as pnpSetup } from "@pnp/common";
import { sp } from "@pnp/sp";
import ProgressDialog from './ProgressDialog';
import { ISetting } from './ISetting';
import { ISettingListItem } from './ISettingListItem';
import { IDurableFunctionResult } from './IDurableFunctionResult';
import { IDurableFunctionCustomStatus } from './IDurableFunctionCustomStatus';
import { IAzureFunctionMessage } from './IAzureFunctionMessage';

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
        } else {
            //hide the command buttone when there is settings found.
            const cmd: Command = this.tryGetCommand("COMMAND_1");
            if (cmd) {
                cmd.visible = false;
            }
        }
    }

    private _getCurrentFolder(): string {
        let ret = "";
        const qs1 = new URLSearchParams(window.location.href);
        if (qs1.has("id")) {
            ret = decodeURIComponent(qs1.get("id"));
        } else {
            ret = decodeURIComponent(this.context.pageContext.list.serverRelativeUrl);
        }
        return ret;
    }

    @autobind
    private _runAzureFunctionHandler(cmdTitle: string) {
        if (this._commandSettings != null && this._commandSettings.length > 0) {
            const res = this._commandSettings.filter(setting => { return setting.title == cmdTitle; });
            if (res != null && res.length == 1 && res[0].setting.apiUrl != null) {
                const setting = res[0];
                const dlg: ProgressDialog = new ProgressDialog();
                //dlg.title="Status";
                dlg.message = "Running Azure Function";
                dlg.show();

                const azureFunctionMessage: IAzureFunctionMessage = {
                    siteUrl: this.context.pageContext.web.absoluteUrl,
                    listTitle: this.context.pageContext.list.title,
                    listId: this.context.pageContext.list.id.toString(),
                    currentFolder: this._getCurrentFolder()
                };
                let requestHeader = new Headers();
                requestHeader.append("Content-Type", "application/json");
                requestHeader.append('Cache-Control', 'no-cache');
                const options: IHttpClientOptions = {
                    body: JSON.stringify(azureFunctionMessage),
                    headers: requestHeader
                };
                this.context.httpClient.post(setting.setting.apiUrl, HttpClient.configurations.v1, options).then(result => {
                    switch (result.status) {
                        case 200:
                            dlg.close().then(_ => {
                                Dialog.alert("The Azure Function was completed successfully").then(() => {
                                    if (setting.setting.refreshPage) {
                                        window.location.reload(true);
                                    }
                                });
                            });
                            break;
                        case 202:
                            result.json().then((apiResult: IDurableFunctionResult) => {
                                if (apiResult.statusQueryGetUri) {
                                    const timerId = setInterval(_ => {
                                        console.log(`Access ${apiResult.statusQueryGetUri}...`);
                                        this.context.httpClient.get(apiResult.statusQueryGetUri, HttpClient.configurations.v1).then(statusResult => {
                                            console.log(`status: ${statusResult.status}`);
                                            if (statusResult.status == 200) {
                                                clearInterval(timerId);
                                                dlg.close().then(() => {
                                                    Dialog.alert("The Azure Function was completed successfully").then(() => {
                                                        if (setting.setting.refreshPage) {
                                                            window.location.reload(true);
                                                        }
                                                    });
                                                });
                                            } else {
                                                statusResult.json().then((customStatus: IDurableFunctionCustomStatus) => {
                                                    if (customStatus.customStatus) {
                                                        dlg.updateStatus(`${customStatus.customStatus.Process}% ${customStatus.customStatus.Message}`);
                                                    }
                                                });

                                            }
                                        });
                                    }, 500);
                                }
                            });
                            break;
                        default:
                            dlg.close().then(_ => {
                                Dialog.alert(`The Azure Function was completed failed. status: ${result.status}`);
                            });
                            break;
                    }
                }).catch(error => {
                    dlg.close().then(() => {
                        Dialog.alert(`Calling Azure Function was failed. error: ${error}`);
                    });
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
