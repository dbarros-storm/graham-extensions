import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import {
  autobind,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';


interface IReviseDocumentsDialogContentProps {
    message: string;
    close: () => void;
    submit: (file: string) => void;
  }

  class RevideDocumentsDialogContent extends React.Component<IReviseDocumentsDialogContentProps, {}> {
    private _pickedFile: string;
  
    constructor(props) {
      super(props);
    }
  
    public render(): JSX.Element {
      return <DialogContent
        title='Revise Document'
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >

      <input id="document" className="document-path" type="file" value={this._pickedFile} data-custom-file-change={this._onFileSelected} />

        <DialogFooter>
          <Button text='Cancel' title='Cancel' onClick={this.props.close} />
          <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._pickedFile); }} />
        </DialogFooter>
      </DialogContent>;
    }
  
    @autobind
    private _onFileSelected(file: string, event: Event): void {
      this._pickedFile = file;
    }
}

export default class ReviseDocumentDialog extends BaseDialog {
    public message: string;
    public fileName: string;
  
    public render(): void {
      ReactDOM.render(<RevideDocumentsDialogContent
        close={ this.close }
        message={ this.message }
        submit={ this._submit }
      />, this.domElement);
    }
  
    public getConfig(): IDialogConfiguration {
      return {
        isBlocking: false
      };
    }
  
    @autobind
    private _submit(file: string): void {
      var fileUploadInput = (this.domElement.getElementsByClassName("document-path"))[0].nodeValue;
      
    //   if (typeof (fileUploadInput) !== 'undefined' && fileUploadInput !== '') {

    //       var itemId = getParameterByName("ID", "");
    //       //alert(itemId);

    //       var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('Fit%20Out')/items(" + itemId + ")";

    //       getListItemFile(_spPageContextInfo.webAbsoluteUrl, itemId).done(function(targetListItem){uploadFile(targetListItem);});
    // }
  }

    private nextChar(c) {
        //return String.fromCharCode(((c.charCodeAt(0) + 1 - 65) % 25) + 65);
        //alert(String.fromCharCode(c.charCodeAt(0) + 1));
        return String.fromCharCode(c.charCodeAt(0) + 1);
    }
  }

