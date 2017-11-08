import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import ReviseDocumentDialog from './ReviseDocumentDialog';

import * as strings from 'ReviseDocumentCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IReviseDocumentCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ReviseDocumentCommandSet';

export default class ReviseDocumentCommandSet extends BaseListViewCommandSet<IReviseDocumentCommandSetProperties> {
  private _fileName: string;
  
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ReviseDocumentCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('Revise_Document');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'Revise_Document':
        const dialog: ReviseDocumentDialog = new ReviseDocumentDialog();
        dialog.message = 'Select a document:';
        // Use 'EEEEEE' as the default color for first usage
        dialog.fileName = this._fileName || '';
        dialog.show().then(() => {
          this._fileName = dialog.fileName;
          Dialog.alert(`Picked color: ${dialog.fileName}`);
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
