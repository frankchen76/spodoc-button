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
import {sp} from "@pnp/sp";
import ProgressDialog from './ProgressDialog';

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
  private SETTING_LISTNAME="Settings";
  @override
  public onInit(): Promise<void> {
    return super.onInit().then(_=>{
      Log.info(LOG_SOURCE, 'Initialized CallAzureFunctionCommandSet');
      pnpSetup({
        spfxContext: this.context
      });
    });
    // Log.info(LOG_SOURCE, 'Initialized CallAzureFunctionCommandSet');
    // return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @autobind
  private _runAzureFunctionHandler(){
    const dlg: ProgressDialog = new ProgressDialog();
    //dlg.title="Status";
    dlg.message="running Azure Function";
    dlg.show();
    sp.web.lists.getByTitle(this.SETTING_LISTNAME).items.get().then((items:any[])=>{
      let azureFuncationUrl = items[0]["Value"];
      if(azureFuncationUrl!=null && azureFuncationUrl!=""){
        return this.context.httpClient.get(azureFuncationUrl,HttpClient.configurations.v1);
      }else{
        Dialog.alert("Please add 'AzureFunctionUrl' to 'Settings' list");
      }
    }).then(result=>{
      return dlg.close();
    }).then(()=>{
      Dialog.alert("successful");
    })
    .catch(err=>{
      Dialog.alert(err);
    });
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_RunAzureFunction':
          // const dialog: ColorPickerDialog = new ColorPickerDialog();
          // dialog.message = 'Pick a color:';
          // // Use 'EEEEEE' as the default color for first usage
          // dialog.colorCode = '#EEEEEE';
          // dialog.show().then(() => {
          //   //this._colorCode = dialog.colorCode;
          //   Dialog.alert(`Picked color: ${dialog.colorCode}`);
          // });
          //const dlg: ProgressDialog = new ProgressDialog();
          //dlg.show();
      this._runAzureFunctionHandler();

        // Dialog.alert(`${this.properties.sampleTextOne}`);
        //this.context.pageContext.
        // this.context.httpClient.get("https://ezcode-testfunction1.azurewebsites.net/api/HttpTriggerCSharp1?code=7bo6F86FMGPgddQnQo9Heea1uB0wxTH/BaWZ28fQMm333HqeLGXEWw==",HttpClient.configurations.v1)
        // .then(result=>{
        //   //alert("done");
        //   Dialog.alert(`Done. result: ${result}`);
        // });
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
