/* tslint:disable */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
/* tslint:enable */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  CompactPeoplePicker,
  IBasePickerSuggestionsProps,
  IBasePicker,
  ListPeoplePicker,
  NormalPeoplePicker,
  ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { people, mru } from './PeoplePickerData';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Promise } from 'es6-promise';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
//import './PeoplePicker.Types.Example.scss';
import {
  Button,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';

export interface IPeoplePickerExampleState {
  currentPicker?: number | string;
  delayResults?: boolean;
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems?: IPersonaProps[];
}

interface IReviseDocumentsDialogContentProps {
  message: string;
  close: () => void;
  submit: (file: string) => void;
}

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'People Picker Suggestions available',
  //suggestionsContainerAriaLabel: 'Suggested contacts'
};

const limitedSearchAdditionalProps: IBasePickerSuggestionsProps = {
  searchForMoreText: 'Load all Results',
  resultsMaximumNumber: 10,
  searchingText: 'Searching...'
};

const limitedSearchSuggestionProps: IBasePickerSuggestionsProps = assign(limitedSearchAdditionalProps, suggestionProps);

export class PeoplePickerTypesExample extends BaseComponent<any, IPeoplePickerExampleState> {
  private _picker: IBasePicker<IPersonaProps>;
  private _peopleList;

  constructor() {
    super();
    
    people.forEach((persona: IPersonaProps) => {
      let target: IPersonaWithMenu = {};

      assign(target, persona);
      this._peopleList.push(target);
    });

    //this._searchPeople("", this._peopleList);

    this.state = {
      currentPicker: 1,
      delayResults: false,
      peopleList: this._peopleList,
      mostRecentlyUsed: mru,
      currentSelectedItems: []
    };
  }

  public render() {
    let currentPicker: JSX.Element | undefined = undefined;

    currentPicker = this._renderLimitedSearch();

    // switch (this.state.currentPicker) {
    //   case 1:
    //     currentPicker = this._renderNormalPicker();
    //     break;
    //   case 2:
    //     currentPicker = this._renderCompactPicker();
    //     break;
    //   case 3:
    //     currentPicker = this._renderListPicker();
    //     break;
    //   case 4:
    //     currentPicker = this._renderPreselectedItemsPicker();
    //     break;
    //   case 5:
    //     currentPicker = this._renderLimitedSearch();
    //     break;
    //   case 6:
    //     currentPicker = this._renderProcessSelectionPicker();
    //   case 7:
    //     currentPicker = this._renderControlledPicker();
    //     break;
    //   default:
    // }

    return (
      <div>
        <DialogContent
        title='Revise Document'
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        { currentPicker }
        <DialogFooter>
          <Button text='Cancel' title='Cancel' onClick={this.props.close} />
          <PrimaryButton text='OK' title='OK'  />
        </DialogFooter>
      </DialogContent>;
      </div>
    );
  }

  private _getTextFromItem(persona: IPersonaProps): string {
    return persona.primaryText as string;
  }
  
  private _renderListPicker() {
    return (
      <ListPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        onEmptyInputFocus={ this._returnMostRecentlyUsed }
        getTextFromItem={ this._getTextFromItem }
        className={ 'ms-PeoplePicker' }
        pickerSuggestionsProps={ suggestionProps }
        key={ 'list' }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        onValidateInput={ this._validateInput }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
      />
    );
  }

  private _renderNormalPicker() {
    return (
      <NormalPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        onEmptyInputFocus={ this._returnMostRecentlyUsed }
        getTextFromItem={ this._getTextFromItem }
        pickerSuggestionsProps={ suggestionProps }
        className={ 'ms-PeoplePicker' }
        key={ 'normal' }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        onValidateInput={ this._validateInput }
        removeButtonAriaLabel={ 'Remove' }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
        //onInputChange={ this._onInputChange }
      />
    );
  }

  private _renderCompactPicker() {
    return (
      <CompactPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        onEmptyInputFocus={ this._returnMostRecentlyUsed }
        getTextFromItem={ this._getTextFromItem }
        pickerSuggestionsProps={ suggestionProps }
        className={ 'ms-PeoplePicker' }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        onValidateInput={ this._validateInput }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
      />
    );
  }

  private _renderPreselectedItemsPicker() {
    return (
      <CompactPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        onEmptyInputFocus={ this._returnMostRecentlyUsed }
        getTextFromItem={ this._getTextFromItem }
        className={ 'ms-PeoplePicker' }
        defaultSelectedItems={ people.splice(0, 3) }
        key={ 'list' }
        pickerSuggestionsProps={ suggestionProps }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        onValidateInput={ this._validateInput }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
      />
    );
  }

  private _renderLimitedSearch() {
    limitedSearchSuggestionProps.resultsFooter = this._renderFooterText;

    return (
      <CompactPeoplePicker
        onResolveSuggestions={ this._onFilterChangedWithLimit }
        onEmptyInputFocus={ this._returnMostRecentlyUsedWithLimit }
        getTextFromItem={ this._getTextFromItem }
        className={ 'ms-PeoplePicker' }
        onGetMoreResults={ this._onFilterChanged }
        pickerSuggestionsProps={ limitedSearchSuggestionProps }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
      />
    );
  }

  private _renderProcessSelectionPicker() {
    return (
      <NormalPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        onEmptyInputFocus={ this._returnMostRecentlyUsed }
        getTextFromItem={ this._getTextFromItem }
        pickerSuggestionsProps={ suggestionProps }
        className={ 'ms-PeoplePicker' }
        onRemoveSuggestion={ this._onRemoveSuggestion }
        onValidateInput={ this._validateInput }
        removeButtonAriaLabel={ 'Remove' }
        onItemSelected={ this._onItemSelected }
        inputProps={ {
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
          'aria-label': 'People Picker'
        } }
        componentRef={ this._resolveRef('_picker') }
      />
    );
  }

  private _renderControlledPicker() {
    let controlledItems = [];
    for (let i = 0; i < 5; i++) {
      let item = this.state.peopleList[i];
      if (this.state.currentSelectedItems!.indexOf(item) === -1) {
        controlledItems.push(this.state.peopleList[i]);
      }
    }
    return (
      <div>
        <NormalPeoplePicker
          onResolveSuggestions={ this._onFilterChanged }
          getTextFromItem={ this._getTextFromItem }
          pickerSuggestionsProps={ suggestionProps }
          className={ 'ms-PeoplePicker' }
          key={ 'controlled' }
          selectedItems={ this.state.currentSelectedItems }
          onChange={ this._onItemsChange }
          inputProps={ {
            onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
            onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called')
          } }
          componentRef={ this._resolveRef('_picker') }
        />
        <label> Click to Add a person </label>
        { controlledItems.map(item => <div>
          <DefaultButton
            className='controlledPickerButton'
            // tslint:disable-next-line:jsx-no-lambda
            onClick={ () => {
              this.setState({
                currentSelectedItems: this.state.currentSelectedItems!.concat([item])
              });
            } }
          >
            <Persona { ...item} />
          </DefaultButton>
        </div>) }
      </div>
    );
  }

  @autobind
  private _onItemsChange(items: any[]) {
    this.setState({
      currentSelectedItems: items
    });
  }

  @autobind
  private _onSetFocusButtonClicked() {
    if (this._picker) {
      this._picker.focus();
    }
  }

  @autobind
  private _renderFooterText(): JSX.Element {
    return <div>No additional results</div>;
  }

  @autobind
  private _onRemoveSuggestion(item: IPersonaProps): void {
    let { peopleList, mostRecentlyUsed: mruState } = this.state;
    let indexPeopleList: number = peopleList.indexOf(item);
    let indexMostRecentlyUsed: number = mruState.indexOf(item);

    if (indexPeopleList >= 0) {
      let newPeople: IPersonaProps[] = peopleList.slice(0, indexPeopleList).concat(peopleList.slice(indexPeopleList + 1));
      this.setState({ peopleList: newPeople });
    }

    if (indexMostRecentlyUsed >= 0) {
      let newSuggestedPeople: IPersonaProps[] = mruState.slice(0, indexMostRecentlyUsed).concat(mruState.slice(indexMostRecentlyUsed + 1));
      this.setState({ mostRecentlyUsed: newSuggestedPeople });
    }
  }

  @autobind
  private _onItemSelected(item: IPersonaProps) {
    const processedItem = assign({}, item);
    processedItem.primaryText = `${item.primaryText} (selected)`;
    return new Promise<IPersonaProps>((resolve, reject) => setTimeout(() => resolve(processedItem), 250));
  }

  @autobind
  private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = this._filterPersonasByText(filterText);

      filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
      filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
      return this._filterPromise(filteredPersonas);
    } else {
      return [];
    }
  }

  @autobind
  private _returnMostRecentlyUsed(currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    let { mostRecentlyUsed } = this.state;
    mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
    return this._filterPromise(mostRecentlyUsed);
  }

  @autobind
  private _returnMostRecentlyUsedWithLimit(currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    let { mostRecentlyUsed } = this.state;
    mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
    mostRecentlyUsed = mostRecentlyUsed.splice(0, 3);
    return this._filterPromise(mostRecentlyUsed);
  }

  @autobind
  private _onFilterChangedWithLimit(filterText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    return this._onFilterChanged(filterText, currentPersonas, 3);
  }

  private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    if (this.state.delayResults) {
      return this._convertResultsToPromise(personasToReturn);
    } else {
      return personasToReturn;
    }
  }

  private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
  }

  private _filterPersonasByText(filterText: string): IPersonaProps[] {
    return this.state.peopleList.filter(item => this._doesTextStartWith(item.primaryText as string, filterText));
  }

  private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
  }

  private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
    return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
  }

  @autobind
  private _toggleDelayResultsChange(toggleState: boolean) {
    this.setState({ delayResults: toggleState });
  }

  @autobind
  private _dropDownSelected(option: IDropdownOption) {
    this.setState({ currentPicker: option.key });
  }

  @autobind
  private _validateInput(input: string) {
    if (input.indexOf('@') !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

  private searchPeopleFromMock(): IPersonaProps[] {
    return this._peopleList = [
      {
        imageUrl: './images/persona-female.png',
        imageInitials: 'PV',
        primaryText: 'Annie Lindqvist',
        secondaryText: 'Designer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      },
      {
        imageUrl: './images/persona-male.png',
        imageInitials: 'AR',
        primaryText: 'Aaron Reid',
        secondaryText: 'Designer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      },
      {
        imageUrl: './images/persona-male.png',
        imageInitials: 'AL',
        primaryText: 'Alex Lundberg',
        secondaryText: 'Software Developer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      },
      {
        imageUrl: './images/persona-male.png',
        imageInitials: 'RK',
        primaryText: 'Roko Kolar',
        secondaryText: 'Financial Analyst',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      },
    ];
  }

    private _searchPeople(terms: string, results: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    //return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));

      const userRequestUrl: string = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
      let principalType: number = 0;
      if (this.props.principalTypeUser === true) {
        principalType += 1;
      }
      if (this.props.principalTypeSharePointGroup === true) {
        principalType += 8;
      }
      if (this.props.principalTypeSecurityGroup === true) {
        principalType += 4;
      }
      if (this.props.principalTypeDistributionList === true) {
        principalType += 2;
      }
      const data = {
        'queryParams': {
          'AllowEmailAddresses': true,
          'AllowMultipleEntities': false,
          'AllUrlZones': false,
          'MaximumEntitySuggestions': this.props.numberOfItems,
          'PrincipalSource': 15,
          // PrincipalType controls the type of entities that are returned in the results.
          // Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
          // These values can be combined (example: 13 is security + SP groups + users)
          'PrincipalType': principalType,
          'QueryString': terms
        }
      };

      return new Promise<IPersonaProps[]>((resolve, reject) =>
        this.props.spHttpClient.post(userRequestUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json',
              "content-type": "application/json"
            },
            body: JSON.stringify(data)
          })
          .then((response: SPHttpClientResponse) => {
            return response.json();
          })
          .then((response: any): void => {
            let relevantResults: any = JSON.parse(response.value);
            let resultCount: number = relevantResults.length;
            let people = [];
            let persona: IPersonaProps = {};
            if (resultCount > 0) {
              for (var index = 0; index < resultCount; index++) {
                var p = relevantResults[index];
                let account = p.Key.substr(p.Key.lastIndexOf('|') + 1);

                persona.primaryText = p.DisplayText;
                persona.imageUrl = `/_layouts/15/userphoto.aspx?size=S&accountname=${account}`;
                persona.imageShouldFadeIn = true;
                persona.secondaryText = p.EntityData.Title;
                people.push(persona);
              }
            }
            resolve(people);
          }, (error: any): void => {
            reject(this._peopleList = []);
          })
        );
  }

  /**
   * Takes in the picker input and modifies it in whichever way
   * the caller wants, i.e. parsing entries copied from Outlook (sample
   * input: "Aaron Reid <aaron>").
   *
   * @param input The text entered into the picker.
   */
  private _onInputChange(input: string): string {
    const outlookRegEx = /<.*>/g;
    const emailAddress = outlookRegEx.exec(input);

    if (emailAddress && emailAddress[0]) {
      return emailAddress[0].substring(1, emailAddress[0].length - 1);
    }

    return input;
  }
}


export default class ReviseDocumentDialog extends BaseDialog {
  public message: string;
  public fileName: string;

  public render(): void {
    ReactDOM.render(<PeoplePickerTypesExample
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
    this.fileName = file;
    this.close();
  }
}