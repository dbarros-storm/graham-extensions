import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import ShareDocumentsDialog from './components/Dialog/ShareDocumentsDialog'

import * as strings from 'ShareMultiDocumentsCommandSetStrings';
import { Dialog } from '@microsoft/sp-dialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IShareMultiDocumentsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ShareMultiDocumentsCommandSet';

export default class ShareMultiDocumentsCommandSet extends BaseListViewCommandSet<IShareMultiDocumentsCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ShareMultiDocumentsCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('Share_Documents');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'Share_Documents':
      const dialog: ShareDocumentsDialog = new ShareDocumentsDialog();
      dialog.message = 'Select a document:';
      dialog.spHttpClient = this.context.spHttpClient;
      dialog.siteUrl = this.context.pageContext.web.absoluteUrl;
      // Use 'EEEEEE' as the default color for first usage
      //dialog.fileName = this._fileName || '';
      dialog.show().then(() => {
        //this._fileName = dialog.fileName;
        
      });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
