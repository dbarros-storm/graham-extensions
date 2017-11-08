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
import OfficeUiFabricPeoplePicker from '../PeoplePicker/OfficeUiFabricPeoplePicker';
import { IOfficeUiFabricPeoplePickerProps } from '../PeoplePicker/IOfficeUiFabricPeoplePickerProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


interface IShareDocumentsDialogContentProps {
    message: string;
    spHttpClient: SPHttpClient;
    siteUrl: string;
    close: () => void;
    submit: (file: string) => void;
  }

  class ShareDocumentsDialogContent extends React.Component<IShareDocumentsDialogContentProps, {}> {
    private _pickedFile: string;
  
    constructor(props) {
      super(props);
    }
  
    public render(): JSX.Element {
      let currentPicker: JSX.Element | undefined = undefined;

      currentPicker = this.PeoplePicker();

      return <DialogContent
        title='Revise Document'
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >

      {currentPicker}

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

    
  private PeoplePicker(){
    const element: React.ReactElement<IOfficeUiFabricPeoplePickerProps> = React.createElement(
      OfficeUiFabricPeoplePicker,
      {
        description: "",
        spHttpClient: this.props.spHttpClient,
        siteUrl: this.props.siteUrl,
        typePicker: "Normal",
        principalTypeUser: true,
        principalTypeSharePointGroup: true,
        principalTypeSecurityGroup: false,
        principalTypeDistributionList: false,
        numberOfItems: 10
      }
    );

    return element;
  }
}

export default class ShareDocumentsDialog extends BaseDialog {
    public message: string;
    public spHttpClient: SPHttpClient;
    public siteUrl: string;
    public fileName: string;
  
    public render(): void {
      ReactDOM.render(<ShareDocumentsDialogContent
        spHttpClient={this.spHttpClient}
        siteUrl={this.siteUrl}
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

