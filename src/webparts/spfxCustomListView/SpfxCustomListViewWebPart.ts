import * as React from 'react';
import * as ReactDom from 'react-dom';
import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel,
  IPropertyPaneField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdownOptionType
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxCustomListViewWebPartStrings';
import SpfxCustomListView from './components/SpfxCustomListView';
import { ISpfxCustomListViewProps } from './components/ISpfxCustomListViewProps';
import { sp } from "@pnp/sp";
import { PnPService } from './services/PnPService';
import { mergeOptions } from '@pnp/common';
import { PropertyPaneDescription, ShowTitle } from 'SpfxCustomListViewWebPartStrings';
//component to reference external CSS from SP CDN 
import { SPComponentLoader } from '@microsoft/sp-loader';
import { FilterURLParameter } from './Entities/FilterURLParameter';
import { FilterSegment } from './Entities/FilterSegment';
import * as _ from 'lodash';
export interface ISpfxCustomListViewWebPartProps {
  description: string;
  title: string;
  titleURL: string;
  titleColor: string;
  enableExpandCollapse: boolean;
  expandCollapseDefaultState: string;
  // DataSource
  siteName: IPropertyPaneDropdownOption;
  listName: string;
  fieldName: string;
  sortName: string;
  ascending: boolean;
  fieldValuesName: string;
  columnsSelected: Array<string>;
  columnDisplayName:  Array<{key:string, displayName:string, type:string}>;
  fieldValuesSortName: string;

  filters: Array<any>;
  filtersExternal: Array<any>;

  //BCS Connection
  externalColumnsSelected: Array<string>;
  externalColumnDisplayName:  Array<{key:string, displayName:string, type:string}>;
  externalListName: string;
  externalFieldName: string;
  displayNameRename: string;
  externalFieldValuesName: string;
  columnAsHTML: string;

  // Variable Source
  // fieldQueryUrlVariableName: string;
  // fieldQueryUrlName: string;
  fieldQueryUrlSegmentName: string;
  fieldSegmentIndex: number;
  // variableSource: { variableName: string; value: string };
  variableSourceParameters: Array<{ variableName: string; value: string }>;
  variableSourceParametersRender: Array<FilterURLParameter>;
  variableSourceSegments: Array<{ variableSegmentName: string; value: string }>;
  variableSourceSegmentsRender: Array<FilterSegment>;
  dynamicPropertiesFilters: any;
  // Display Options
  fieldPager: boolean;
  fieldPageLimit: number;
  fieldExportToExcel: boolean;
  fieldPrint: boolean;
  fieldTotal: boolean;
  // Layout
  layoutSelection: boolean;
  fieldHeader: string;
  fieldBody: string;
  fieldFooter: string;
  fieldCSS: string;
  fieldJavascript: string;
  fieldShowTitle: string;
}

export default class SpfxCustomListViewWebPart extends BaseClientSideWebPart<ISpfxCustomListViewWebPartProps> {
  private sites: IPropertyPaneDropdownOption[];
  private sitesDropdownDisabled: boolean = true;
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private fields: IPropertyPaneDropdownOption[];
  private fieldsDropdownDisabled: boolean = true;
  private _externalLists: IPropertyPaneDropdownOption[];
  private _externalFields: IPropertyPaneDropdownOption[];
  private _externalListsDropdownDisabled: boolean = true;
  private _externalFieldsDropdownDisabled: boolean = true;
  private _externalColumnsSelected: Array<string> = new Array();
  private _externalColumnsDisplayNames: Array<{key:string, displayName:string, type:string}> = new Array();
  private _pnpService;
  private columnsSelected: Array<string> = new Array();
  private _columnsDisplayNames: Array<{key:string, displayName:string, type: string}> = new Array();
  private filtersSelected: Array<string> = new Array();
  // private filtersProperties: Array<filterPropertiesValue> = new Array();
  private layoutSelectionRender: Array<IPropertyPaneField<any>> = new Array();
  private _filterRender: Array<any> = new Array();
  private _filterExternalRender: Array<any> = new Array();
  private idDynamic: number = 0; // this is to have ids dynamic for some propertyPaneFields
  private idDynamicExternal: number = 0; // this is to have ids dynamic for some propertyPaneFields have duplicated  for filter in external list need to be reviewed
  private fieldsButtonDisabled: boolean = true;
  private _externalFieldsButtonDisabled: boolean = true;
  private _variableSourceParameters: Array<{variableName: string; value: string}> = new Array();
  private _variableSourceParametersRender: Array<FilterURLParameter> = new Array();
  private _variableSourceSegments: Array<{variableSegmentName: string; value: string}> = new Array();
  private _variableSourceSegmentsRender: Array<FilterSegment> = new Array();
  private _numberPageLimit: IPropertyPaneDropdownOption[] = new Array<IPropertyPaneDropdownOption>();

  public onInit(): Promise<void> {
    SPComponentLoader.loadCss('/Style%20Library/mw.portal/css/mw-portal.css');

    // return super.onInit().then(_ => {
    //   // other init code may be present
    //   this._pnpService = new PnPService(this.context);
    //   sp.setup({
    //     spfxContext: this.context
    //   });
    // });

    //set the default web part properties
    return new Promise<void>((resolve, _reject) => {
      if (this.properties.expandCollapseDefaultState === undefined) {
        this.properties.expandCollapseDefaultState = 'Open';
      }
      if (this.properties.title === undefined) {
        this.properties.title = 'Title';
      }
      if (this.properties.titleColor === undefined) {
        this.properties.titleColor = 'Blue';
      }
      if (this.properties.enableExpandCollapse === undefined) {
        this.properties.enableExpandCollapse = false;
      }
      resolve(undefined);

      //return onInit for componentloader
      return super.onInit().then(_ => {
        // other init code may be present
        this._pnpService = new PnPService(this.context);
        sp.setup({
          spfxContext: this.context
        });
      });;
    });
  }

  public render(): void {
    const element: React.ReactElement<ISpfxCustomListViewProps> = React.createElement(
      SpfxCustomListView,
      {
        description: this.properties.description,
        title: this.properties.title,
        titleURL: this.properties.titleURL,
        titleColor: this.properties.titleColor,
        enableExpandCollapse: this.properties.enableExpandCollapse,
        expandCollapseDefaultState: this.properties.expandCollapseDefaultState,
        // DataSource
        siteName: this.properties.siteName,
        listName: this.properties.listName,
        fieldName: this.properties.fieldName,
        sortName: this.properties.sortName,
        ascending: this.properties.ascending,
        fieldValuesName: this.properties.fieldValuesName,
        columnsSelected: this.properties.columnsSelected,
        columnDisplayName: this.properties.columnDisplayName,
        fieldValuesSortName: this.properties.fieldValuesSortName,


        filters: this.properties.filters,
        context: this.context,

        // BCS Connection
        externalListName: this.properties.externalListName,
        externalFieldName: this.properties.externalFieldName,
        externalFieldValuesName: this.properties.externalFieldValuesName,
        externalColumnsSelected: this.properties.externalColumnsSelected,
        externalColumnDisplayName: this.properties.externalColumnDisplayName,
        displayNameRename: this.properties.displayNameRename,
        columnAsHTML: this.properties.columnAsHTML,


        // Variable Source
        variableSourceParameters: this.properties.variableSourceParameters,
        variableSourceParametersRender: this.properties.variableSourceParametersRender,
        variableSourceSegments: this.properties.variableSourceSegments,
        variableSourceSegmentsRender: this.properties.variableSourceSegmentsRender,
      
        dynamicPropertiesFilters: this.properties,

        // Display Options
        fieldPager: this.properties.fieldPager,
        fieldPageLimit: this.properties.fieldPageLimit,
        fieldExportToExcel: this.properties.fieldExportToExcel,
        fieldPrint: this.properties.fieldPrint,
        fieldTotal: this.properties.fieldTotal,
        // Layout
        layoutSelection: this.properties.layoutSelection,
        fieldHeader: this.properties.fieldHeader,
        fieldBody: this.properties.fieldBody,
        fieldFooter: this.properties.fieldFooter,
        fieldCSS: this.properties.fieldCSS,
        fieldJavascript: this.properties.fieldJavascript,
        fieldShowTitle: this.properties.fieldShowTitle
      }
    );

    ReactDom.render(element, this.domElement);
  }
//#region SP Sites, Lists and Columns Dropdowns
  private loadSites(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array();
      this._pnpService.getAllSites().then(sites => {
        options.push({ key: '', text: 'Select Site' });
        sites.forEach(async site => {
          let navigator = await this.navigatorSiteName(site.Url);
          options.push({ key: site.Url, text: navigator });
          options.sort(function (a, b) {
            if (a.text > b.text) {
              return 1;
            }
            if (a.text < b.text) {
              return -1;
            }
            // a must be equal to b
            return 0;
          });
        });
        resolve(options);
      });
    });
  }

  private async navigatorSiteName(siteUrl: string): Promise<string> {
    let buildingUrl = "";
    let segmentsComplete = siteUrl.toString().split('/');
    //remove two first elements protocol and empty value after the split
    let segments = segmentsComplete.slice(2, segmentsComplete.length);
    const urlValue = new URL(siteUrl);
    let protocol = urlValue.protocol;
    let hostname = urlValue.hostname;
    let navigator = "";

    for (let index = 0; index < segments.length; index++) {
      if (index == 0) {
        buildingUrl = protocol + "//" + hostname;
      }
      if (index == 1) {
        buildingUrl = protocol + "//" + hostname + "/" + segments[index];
      }
      if (index == 2) {
        buildingUrl = protocol + "//" + hostname + "/" + segments[index - 1] + "/" + segments[index] ;
      }

      await this._pnpService.getWebTitle(buildingUrl).then(title => {
        if (index == 0) {
          navigator = title;
        }
        if (index == 1)
        {
          navigator = navigator + ">" + title;
        }
        if (index == 2)
        {
          navigator = navigator + ">" + title;
        }
      });
    }
    return await navigator;
  }
  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.siteName) {
      // resolve to empty options since no site has been selected
      return Promise.resolve();
    }

    const wp: SpfxCustomListViewWebPart = this;
    let opts = new Map<IPropertyPaneDropdownOption, Array<IPropertyPaneDropdownOption>>();
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array();
      this._pnpService.getListsFromSite(this.properties.siteName).then(lists => {
        //this._columnsDisplayNames = [];
        options.push({ key: '', text: 'Select List' });
        lists.forEach(list => {
          options.push({ key: list.Title, text: list.Title });
          options.sort(function (a, b) {
            if (a.text > b.text) {
              return 1;
            }
            if (a.text < b.text) {
              return -1;
            }
            // a must be equal to b
            return 0;
          });
          //this._columnsDisplayNames.push({key: field.InternalName, displayName: field.Title });
        });
        opts.set(this.properties.siteName, options)
        resolve(opts.get(wp.properties.siteName));
      });
    });
  }
  private loadFields(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    const wp: SpfxCustomListViewWebPart = this;
    let opts = new Map<string, Array<IPropertyPaneDropdownOption>>();
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      let options: Array<IPropertyPaneDropdownOption> = new Array();
      this._pnpService.getFieldsFromList(this.properties.listName, this.properties.siteName).then(fields => {
        this._columnsDisplayNames = [];
        options.push({ key: '', text: 'Select column' });
        fields.forEach(field => {
          options.push({ key: field.InternalName, text: field.Title });
          options.sort(function (a, b) {
            if (a.text > b.text) {
              return 1;
            }
            if (a.text < b.text) {
              return -1;
            }
            // a must be equal to b
            return 0;
          });
          this._columnsDisplayNames.push({key: field.InternalName, displayName: field.Title, type: field["odata.type"]});
        });
        opts.set(this.properties.listName, options)
        resolve(opts.get(wp.properties.listName));
      });
    });
  }
  //#endregion

//#region SP List Multiple Columns
  private SelectedColumnsClick(oldVal: any): any {
    if (this.properties.fieldValuesName == undefined) this.properties.fieldValuesName = "";
    if (this.properties.fieldValuesName.length == 0) {
      this.columnsSelected = this.properties.fieldValuesName.split(",").slice(1);
    } else {
      this.columnsSelected = this.properties.fieldValuesName.split(",")
    }
    this.columnsSelected.push(this.properties.fieldName);
    this.properties.columnsSelected = this.columnsSelected;
    this.properties.fieldValuesName = this.columnsSelected.toString();

    this.properties.columnDisplayName = _.filter(this._columnsDisplayNames, (column) => _.includes(this.properties.columnsSelected, column.key));
    console.log(this.properties.fieldValuesName);
    console.log(this.properties.columnDisplayName);
  }

  private ParseToDisplayName(displayNames: Array<{ key: string, displayName: string }>) {
    if (displayNames)
    {
      let names: Array<string> = new Array<string>();
      displayNames.forEach(element => {
        if (element.displayName)
        {
          names.push(element.displayName);
        }
        else
        {
          names.push(element.key);
        }
        
      });
      return names.toString();
    }
  }

  private RemoveColumnsClick() {
    if (this.properties.fieldValuesName == undefined) this.properties.fieldValuesName = "";
    if (this.properties.fieldValuesName.length == 0) {
      this.columnsSelected = this.properties.fieldValuesName.split(",").slice(1);
    } else {
      this.columnsSelected = this.properties.fieldValuesName.split(",")
    }
    // Remove the string selected
    const index = this.columnsSelected.indexOf(this.properties.fieldName, 0);
    if (index > -1) {
      this.columnsSelected.splice(index, 1);
    }
    this.properties.fieldValuesName = this.columnsSelected.toString();
    this.properties.columnsSelected = this.columnsSelected;
    this.properties.columnDisplayName = _.filter(this._columnsDisplayNames, (column) => _.includes(this.properties.columnsSelected, column.key));
    console.log(this.properties.fieldValuesName);
    console.log(this.properties.columnDisplayName);
  }
  //#endregion

//#region BCS Connection Lists and Columns Dropdowns
private externalLoadLists(): Promise<IPropertyPaneDropdownOption[]> {
  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
    let options: Array<IPropertyPaneDropdownOption> = new Array();
    this._pnpService.getAllExternalLists().then(lists => {
      options.push({ key: '', text: 'Select List' });
      lists.forEach(list => {
        options.push({ key: list.Title, text: list.Title });
        options.sort(function (a, b) {
          if (a.text > b.text) {
            return 1;
          }
          if (a.text < b.text) {
            return -1;
          }
          // a must be equal to b
          return 0;
        });
      });
      resolve(options);
    });
  });
}
private externalLoadFields(): Promise<IPropertyPaneDropdownOption[]> {
  if (!this.properties.externalListName) {
    // resolve to empty options since no list has been selected
    return Promise.resolve();
  }

  const wp: SpfxCustomListViewWebPart = this;
  let opts = new Map<string, Array<IPropertyPaneDropdownOption>>();
  return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
    let options: Array<IPropertyPaneDropdownOption> = new Array();
    this._pnpService.getExternalFieldsFromList(this.properties.externalListName).then(fields => {
      this._externalColumnsDisplayNames = [];
      options.push({ key: '', text: 'Select column' });
      fields.forEach(field => {
        options.push({ key: field.InternalName, text: field.Title });
        options.sort(function (a, b) {
          if (a.text > b.text) {
            return 1;
          }
          if (a.text < b.text) {
            return -1;
          }
          // a must be equal to b
          return 0;
        });
         this._externalColumnsDisplayNames.push({key: field.InternalName, displayName: field.Title, type: field["odata.type"] });
      });
      opts.set(this.properties.externalListName, options)
      resolve(opts.get(wp.properties.externalListName));
    });
  });
}
//#endregion

//#region BCS Connection Multiple Colummns
  private ExternalSelectedColumnsClick(oldVal: any): any {
    if (this.properties.externalFieldValuesName == undefined) this.properties.externalFieldValuesName = "";
    if (this.properties.externalFieldValuesName.length == 0) {
      this._externalColumnsSelected = this.properties.externalFieldValuesName.split(",").slice(1);
    } else {
      this._externalColumnsSelected = this.properties.externalFieldValuesName.split(",")
      // this._externalColumnsDisplayNames = this.properties.externalColumnDisplayName;
    }

    this._externalColumnsSelected.push(this.properties.externalFieldName);

    this.properties.externalColumnsSelected = this._externalColumnsSelected;
    this.properties.externalFieldValuesName = this._externalColumnsSelected.toString();

    let temp = new Array<{ key: string, displayName: string, type: string }>();

    temp = _.filter(this._externalColumnsDisplayNames, (column) => _.includes(this.properties.externalColumnsSelected, column.key));
    let newValue;
    temp.forEach((column, index) => {
      if (column.key == this.properties.externalFieldName) {
        if (this.properties.displayNameRename) {
          newValue = { key: column.key, displayName: this.properties.displayNameRename, type: column.type }
        }
        else {
          newValue = { key: column.key, displayName: this.properties.externalFieldName, type: column.type }
        }
        if (this.properties.externalColumnDisplayName) {
          this.properties.externalColumnDisplayName.push(newValue);
        }
        else {
          this.properties.externalColumnDisplayName = new Array<{ key: string, displayName: string, type: string }>();
          this.properties.externalColumnDisplayName.push(newValue);
        }

      }
    })

    console.log(this.properties.externalFieldValuesName);
    console.log(this.properties.externalColumnDisplayName);
    this.properties.displayNameRename = "";
  }

  private ExternalRemoveColumnsClick() {
    // Remove the string selected
    let externalColumns: Array<string> = new Array();
    let columnToRemove = this.properties.externalColumnDisplayName.filter(obj => obj.key === this.properties.externalFieldName)[0];
    const index = this.properties.externalColumnDisplayName.indexOf(columnToRemove, 0);
    if (index > -1) {
      this.properties.externalColumnDisplayName.splice(index, 1);
    }
    this.properties.externalColumnDisplayName.forEach(obj => externalColumns.push(obj.displayName));
    this._externalColumnsSelected = externalColumns;
    this.properties.externalFieldValuesName = this._externalColumnsSelected.toString();
    this.properties.externalColumnsSelected = this._externalColumnsSelected;
  }
  //#endregion

  private SelectedSortByClick(oldVal: any): any {
    if (this.properties.fieldValuesSortName == undefined) this.properties.fieldValuesSortName = "";
    if (this.properties.fieldValuesSortName.length == 0) {
      this.filtersSelected = this.properties.fieldValuesSortName.split(",").slice(1);
    } else {
      this.filtersSelected = this.properties.fieldValuesSortName.split(",")
    }
    this.filtersSelected.push(this.properties.sortName);
    this.properties.fieldValuesSortName = this.filtersSelected.toString();
    console.log(this.properties.fieldValuesSortName);
  }
  private RemoveSortByClick() {
    if (this.properties.fieldValuesSortName == undefined) this.properties.fieldValuesSortName = "";
    if (this.properties.fieldValuesSortName.length == 0) {
      this.filtersSelected = this.properties.fieldValuesSortName.split(",").slice(1);
    } else {
      this.filtersSelected = this.properties.fieldValuesSortName.split(",")
    }
    // Remove the string selected
    const index = this.filtersSelected.indexOf(this.properties.sortName, 0);
    if (index > -1) {
      this.filtersSelected.splice(index, 1);
    }
    this.properties.fieldValuesSortName = this.filtersSelected.toString();
    console.log(this.properties.fieldValuesSortName);
  }

  private InitialLoadPropertyPaneFiledFilters(): void {
    if (this._filterRender.length == 0) {
      this._filterRender.push(PropertyPaneDropdown('filter0', {
        label: strings.FiltersFieldLabel,
        options: this.fields,
        selectedKey: '',
        disabled: this.fieldsDropdownDisabled
      }));
      this._filterRender.push(PropertyPaneDropdown('operator0', {
        label: "",
        selectedKey: "select",
        options: [
          { key: 'select', text: 'Select filter operator' },
          { key: 'eq', text: 'Equals' },
          { key: 'ne', text: 'Not Equals' },
          { key: 'gt', text: 'Greater than' },
          { key: 'ge', text: 'Greater than or Equals' },
          { key: 'lt', text: 'Less than' },
          { key: 'le', text: 'Less than or Equals' },
          { key: 'IsNotNull', text: 'Is Not Null' },
          { key: 'IsNull', text: 'Is Null' },
          { key: 'contains', text: 'Contains' },
          { key: 'beginsWith', text: 'Begins With' }
        ]
      }));
      this._filterRender.push(PropertyPaneTextField('valueFilter0', {
        placeholder: "Enter search words"
      }));
      this._filterRender.push(PropertyPaneDropdown('logicOperator0', {
        label: "",
        selectedKey: "",
        options: [
          { key: '', text: 'Select logic operator' },
          { key: ' and ', text: 'and' },
          { key: ' or ', text: 'or' }
        ]
      }));

      this._filterRender.push(PropertyPaneButton('AddFilter0',
        {
          text: "Add filter",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Add",
          onClick: this.AddNewFilterClick.bind(this),
          disabled: this.fieldsButtonDisabled
        }));
      // this.UpdatePropertiesDynamic(this.idDynamic);
    }
    this.properties.filters = this._filterRender;
  }

  private InitialLoadPropertyPaneFiledFiltersExternal(): void {
    if (this._filterExternalRender.length == 0) {
      this._filterExternalRender.push(PropertyPaneDropdown('filterexternal0', {
        label: strings.FiltersFieldLabel,
        options: this._externalFields,
        selectedKey: '',
        disabled: this._externalFieldsDropdownDisabled
      }));
      this._filterExternalRender.push(PropertyPaneDropdown('operatorexternal0', {
        label: "",
        selectedKey: "select",
        options: [
          { key: 'select', text: 'Select filter operator' },
          { key: 'eq', text: 'Equals' },
          { key: 'ne', text: 'Not Equals' },
          { key: 'gt', text: 'Greater than' },
          { key: 'ge', text: 'Greater than or Equals' },
          { key: 'lt', text: 'Less than' },
          { key: 'le', text: 'Less than or Equals' },
          { key: 'IsNotNull', text: 'Is Not Null' },
          { key: 'IsNull', text: 'Is Null' },
          { key: 'contains', text: 'Contains' },
          { key: 'beginsWith', text: 'Begins With' }
        ]
      }));
      this._filterExternalRender.push(PropertyPaneTextField('valueFilterexternal0', {
        placeholder: "Enter search words"
      }));
      this._filterExternalRender.push(PropertyPaneDropdown('logicOperatorexternal0', {
        label: "",
        selectedKey: "",
        options: [
          { key: '', text: 'Select logic operator' },
          { key: ' and ', text: 'and' },
          { key: ' or ', text: 'or' }
        ]
      }));

      this._filterExternalRender.push(PropertyPaneButton('AddFilterexternal0',
        {
          text: "Add filter",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Add",
          onClick: this.AddNewFilterExternalClick.bind(this),
          disabled: this._externalFieldsButtonDisabled
        }));
      // this.UpdatePropertiesDynamic(this.idDynamic);
    }
    this.properties.filtersExternal = this._filterExternalRender;
  }
 //#region Variable Source Segment 
  private AddFilterParamClick(): IPropertyPaneField<any>[] {
    let index = 0;
    if (this.properties.variableSourceParametersRender) 
    {
      index = this.properties.variableSourceParametersRender.length;
    }
    else
    {
      index = this._variableSourceParametersRender.length;
    }

    // clean property have values
    if (this.properties["fieldQueryUrlVariableName" + index] && this.properties["fieldQueryUrlName" + index]) {
      this.properties["fieldQueryUrlVariableName" + index] = "";
      this.properties["fieldQueryUrlName" + index] = "";
    }
    let initialFilterURLParameter: FilterURLParameter =
    {
      parameterNameControl: PropertyPaneTextField('fieldQueryUrlName' + index, {
        label: strings.fieldQueryUrlName + " " + index.toString()
      }),
      variableNameControl: PropertyPaneTextField('fieldQueryUrlVariableName' + index, {
        label: strings.fieldQueryUrlVariableName + " " + index.toString()
      }),
      trashControl: PropertyPaneButton('removeParamFilter' + index ,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.filterParameterToRemove.bind(this, index)  // this.filterParameterToRemove.bind(this, newId)
      })
    };
    // if (newId == 0)
    // {
    //   initialFilterURLParameter.trashControl = null;
    // }
    if (this.properties.variableSourceParametersRender != undefined) this._variableSourceParametersRender = this.properties.variableSourceParametersRender;
    this._variableSourceParametersRender.push(initialFilterURLParameter);

    // Reordering controls
    this.readjustmentControlsParameters();
    // to save filters after refresh page
    this.properties.variableSourceParametersRender = this._variableSourceParametersRender;
    return this.getFilterURLParameterRender();
  }

  private filterParameterToRemove(id): void {

    delete this._variableSourceParametersRender[id];
    delete this.properties.variableSourceParametersRender[id];

    // remove null values
    this._variableSourceParametersRender = _.without(this._variableSourceParametersRender, undefined, null);
    this.properties.variableSourceParametersRender = _.without(this.properties.variableSourceParametersRender, undefined, null);

    // clean values in properties
    this.properties["fieldQueryUrlVariableName" + id] = "";
    this.properties["fieldQueryUrlName" + id] = "";

    this.readjustmentControlsParameters();

    this.readjustmentPropertyValuesParameters(id)

    console.log(this._variableSourceParametersRender);
  }

  private readjustmentPropertyValuesParameters(id: number) {
    this._variableSourceParametersRender.forEach((filterParam, index) => {
      //readjustments properties
      let variableParam = this.properties["fieldQueryUrlVariableName" + index];
      let nameParam = this.properties["fieldQueryUrlName" + index];
      if ((!variableParam && !nameParam) || (variableParam && !nameParam) || (!variableParam && nameParam)) {
        // Starting readjustments values
        for (let index = id; index < this._variableSourceParametersRender.length; index++) {
          // save values of the next parameter
          let nextid = index + 1;
          let urlParamenter: Record<'variableName' | 'valueName', string> = {
            variableName: this.properties["fieldQueryUrlVariableName" + nextid],
            valueName: this.properties["fieldQueryUrlName" + nextid]
          }

          this.properties["fieldQueryUrlVariableName" + index] = urlParamenter.variableName;
          this.properties["fieldQueryUrlName" + index] = urlParamenter.valueName;
          // clean next property
          this.properties["fieldQueryUrlVariableName" + nextid] = "";
          this.properties["fieldQueryUrlName" + nextid] = "";
        }
      }
    })
    // to save filters after refresh page
    this.properties.variableSourceParametersRender = this._variableSourceParametersRender;
  }

  private readjustmentControlsParameters() {
    this._variableSourceParametersRender.forEach((filterParam, index) => {
      // readjustments controls
      this._variableSourceParametersRender[index].parameterNameControl.targetProperty = "fieldQueryUrlName" + index;
      this._variableSourceParametersRender[index].variableNameControl.targetProperty = "fieldQueryUrlVariableName" + index;
      this._variableSourceParametersRender[index].trashControl = PropertyPaneButton('removeParamFilter' + index,
        {
          text: "Remove",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Trash",
          onClick: this.filterParameterToRemove.bind(this, index)  // this.filterParameterToRemove.bind(this, newId)
        })
    });
  }
  private AddFilterSegmentClick(): IPropertyPaneField<any>[] 
  {
    let index = 0;
    if (this.properties.variableSourceSegmentsRender) 
    {
      index = this.properties.variableSourceSegmentsRender.length;
    }
    else
    {
      index = this._variableSourceSegmentsRender.length;
    }
        // clean property have values
        if (this.properties["fieldSegmentIndex" + index] && this.properties["fieldQueryUrlSegmentName" + index]) {
          this.properties["fieldSegmentIndex" + index] = 0;
          this.properties["fieldQueryUrlSegmentName" + index] = "";
        }
    let initialFilterSegment: FilterSegment =
    {
      segmentVariableNameControl: PropertyPaneTextField('fieldQueryUrlSegmentName' + index, {
        label: strings.fieldQueryUrlSegmentName + " " + index.toString()
      }),
      indexSegmentNameControl: PropertyPaneSlider('fieldSegmentIndex' + index, {
        label: strings.fieldSegmentIndex + " " + index.toString(),
        min: 1,
        max: new URL(window.location.href).toString().split('/').length - 2,
        value: 1,
        showValue: true,
        step: 1
      }),
      trashControl: PropertyPaneButton('removeSegmentFilter' + index ,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.filterSegmentToRemove.bind(this, index)  // this.filterSegmentToRemove.bind(this, newId)
      })
    };
    if (this.properties.variableSourceSegmentsRender != undefined) this._variableSourceSegmentsRender = this.properties.variableSourceSegmentsRender;
    this._variableSourceSegmentsRender.push(initialFilterSegment);

    // Reordering controls
    this.readjustmentControlsSegments();

    // to save filters after refresh page
    this.properties.variableSourceSegmentsRender = this._variableSourceSegmentsRender;
    return this.getFilterSegmentRender();
  }

  private filterSegmentToRemove(id): void {
    delete this._variableSourceSegmentsRender[id];
    delete this.properties.variableSourceSegmentsRender[id];

    // remove null values
    this._variableSourceSegmentsRender = _.without(this._variableSourceSegmentsRender, undefined, null);
    this.properties.variableSourceSegmentsRender = _.without(this.properties.variableSourceSegmentsRender, undefined, null);

    // clean values in properties
    this.properties["fieldSegmentIndex" + id] = 0;
    this.properties["fieldQueryUrlSegmentName" + id] = "";

    this.readjustmentControlsSegments();

    this.readjustmentPropertyValuesSegments(id)

    console.log(this._variableSourceSegmentsRender);
  }

  private readjustmentPropertyValuesSegments(id: number) {
    this._variableSourceSegmentsRender.forEach((filterParam, index) => {
      //readjustments properties
      let variableSegment = this.properties["fieldSegmentIndex" + index];
      // if (variableSegment == undefined) variableSegment = 0;
      let nameSegment = this.properties["fieldQueryUrlSegmentName" + index];
      if (!nameSegment) {
        // Starting readjustments values
        for (let index = id; index < this._variableSourceSegmentsRender.length; index++) {
          // save values of the next parameter
          let nextid = index + 1;
          let urlParamenter: Record<'variableName' | 'valueName', string> = {
            variableName: this.properties["fieldSegmentIndex" + nextid],
            valueName: this.properties["fieldQueryUrlSegmentName" + nextid]
          }

          this.properties["fieldSegmentIndex" + index] = urlParamenter.variableName;
          this.properties["fieldQueryUrlSegmentName" + index] = urlParamenter.valueName;
          // clean next property
          this.properties["fieldSegmentIndex" + nextid] = 0;
          this.properties["fieldQueryUrlSegmentName" + nextid] = "";
        }
      }
    })
    // to save filters after refresh page
    this.properties.variableSourceSegmentsRender = this._variableSourceSegmentsRender;
  }

  private readjustmentControlsSegments() {
    this._variableSourceSegmentsRender.forEach((filterParam, index) => {
      // readjustments controls
      this._variableSourceSegmentsRender[index].segmentVariableNameControl.targetProperty = "fieldQueryUrlSegmentName" + index;
      this._variableSourceSegmentsRender[index].indexSegmentNameControl.targetProperty = "fieldSegmentIndex" + index;
      this._variableSourceSegmentsRender[index].trashControl = PropertyPaneButton('removeSegmentFilter' + index,
        {
          text: "Remove",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Trash",
          onClick: this.filterSegmentToRemove.bind(this, index) 
        })
    });
  }
  //#endregion

  private AddNewFilterClick(): IPropertyPaneField<any>[] {
    this.idDynamic++;
    if (this._filterRender.length == 5) {
      // this._filterRender = [];
      //remove the last Add Button
      this._filterRender.pop();
      //add the remove button in the initial filters controls
      this._filterRender.push(PropertyPaneButton('removeFilter0',
        {
          text: "Remove",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Trash",
          onClick: this.RemoveFilterRuleClick.bind(this, 0) // pass 0 because is the first filter
        }));
      this.addDefaultControlsForFilter();
    }
    else {
      //remove the last Add Button
      this._filterRender.pop();
      this.addDefaultControlsForFilter();
    }
    // assing to the property
    this.properties.filters = this._filterRender;
    return this._filterRender;
  }

  private AddNewFilterExternalClick(): IPropertyPaneField<any>[] {
    this.idDynamicExternal++;
    if (this._filterExternalRender.length == 5) {
      // this._filterRender = [];
      //remove the last Add Button
      this._filterExternalRender.pop();
      //add the remove button in the initial filters controls
      this._filterExternalRender.push(PropertyPaneButton('removeFilterexternal0',
        {
          text: "Remove",
          buttonType: PropertyPaneButtonType.Command,
          icon: "Trash",
          onClick: this.RemoveFilterExternalRuleClick.bind(this, 0) // pass 0 because is the first filter
        }));
      this.addDefaultControlsForFilterExternal();
    }
    else {
      //remove the last Add Button
      this._filterExternalRender.pop();
      this.addDefaultControlsForFilterExternal();
    }
    // assing to the property
    this.properties.filtersExternal = this._filterExternalRender;
    return this._filterExternalRender;
  }

  private filterToRemove(id): Array<any> {
    let controlsFromFilterToRemove: Array<any> = new Array();
    controlsFromFilterToRemove.push(PropertyPaneDropdown('filter' + id, {
      label: strings.FiltersFieldLabel,
      options: this.fields,
      disabled: this.fieldsDropdownDisabled
    }));
    controlsFromFilterToRemove.push(PropertyPaneDropdown('operator' + id, {
      label: "",
      selectedKey: "select",
      options: [
        { key: 'select', text: 'Select filter operator' },
        { key: 'eq', text: 'Equals' },
        { key: 'ne', text: 'Not Equals' },
        { key: 'gt', text: 'Greater than' },
        { key: 'ge', text: 'Greater than or Equals' },
        { key: 'lt', text: 'Less than' },
        { key: 'le', text: 'Less than or Equals' },
        { key: 'IsNotNull', text: 'Is Not Null' },
        { key: 'IsNull', text: 'Is Null' },
        { key: 'contains', text: 'Contains' },
        { key: 'beginsWith', text: 'Begins With' }
      ]
    }));
    controlsFromFilterToRemove.push(PropertyPaneTextField('valueFilter' + id, {
      placeholder: "Enter search words"
    }));
    controlsFromFilterToRemove.push(PropertyPaneDropdown('logicOperator' + id, {
      label: "",
      selectedKey: "",
      options: [
        { key: '', text: 'Select logic operator' },
        { key: ' and ', text: 'and' },
        { key: ' or ', text: 'or' }
      ]
    }));

    controlsFromFilterToRemove.push(PropertyPaneButton('removeFilter' + id,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.RemoveFilterRuleClick.bind(this, this.idDynamic)
      }));

    return controlsFromFilterToRemove;
  }

  private filterExternalToRemove(id): Array<any> {
    let controlsFromFilterToRemove: Array<any> = new Array();
    controlsFromFilterToRemove.push(PropertyPaneDropdown('filterexternal' + id, {
      label: strings.FiltersFieldLabel,
      options: this._externalFields,
      disabled: this._externalFieldsDropdownDisabled
    }));
    controlsFromFilterToRemove.push(PropertyPaneDropdown('operatorexternal' + id, {
      label: "",
      selectedKey: "select",
      options: [
        { key: 'select', text: 'Select filter operator' },
        { key: 'eq', text: 'Equals' },
        { key: 'ne', text: 'Not Equals' },
        { key: 'gt', text: 'Greater than' },
        { key: 'ge', text: 'Greater than or Equals' },
        { key: 'lt', text: 'Less than' },
        { key: 'le', text: 'Less than or Equals' },
        { key: 'IsNotNull', text: 'Is Not Null' },
        { key: 'IsNull', text: 'Is Null' },
        { key: 'contains', text: 'Contains' },
        { key: 'beginsWith', text: 'Begins With' }
      ]
    }));
    controlsFromFilterToRemove.push(PropertyPaneTextField('valueFilterexternal' + id, {
      placeholder: "Enter search words"
    }));
    controlsFromFilterToRemove.push(PropertyPaneDropdown('logicOperatorexternal' + id, {
      label: "",
      selectedKey: "",
      options: [
        { key: '', text: 'Select logic operator' },
        { key: ' and ', text: 'and' },
        { key: ' or ', text: 'or' }
      ]
    }));

    controlsFromFilterToRemove.push(PropertyPaneButton('removeFilterexternal' + id,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.RemoveFilterExternalRuleClick.bind(this, this.idDynamicExternal)
      }));

    return controlsFromFilterToRemove;
  }

  private addDefaultControlsForFilter(): void {
    this._filterRender.push(PropertyPaneDropdown('filter' + this.idDynamic, {
      label: strings.FiltersFieldLabel,
      options: this.fields,
      selectedKey: '',
      disabled: this.fieldsDropdownDisabled
    }));
    this._filterRender.push(PropertyPaneDropdown('operator' + this.idDynamic, {
      label: "",
      selectedKey: "select",
      options: [
        { key: 'select', text: 'Select filter operator' },
        { key: 'eq', text: 'Equals' },
        { key: 'ne', text: 'Not Equals' },
        { key: 'gt', text: 'Greater than' },
        { key: 'ge', text: 'Greater than or Equals' },
        { key: 'lt', text: 'Less than' },
        { key: 'le', text: 'Less than or Equals' },
        { key: 'IsNotNull', text: 'Is Not Null' },
        { key: 'IsNull', text: 'Is Null' },
        { key: 'contains', text: 'Contains' },
        { key: 'beginsWith', text: 'Begins With' }
      ]
    }));
    this._filterRender.push(PropertyPaneTextField('valueFilter' + this.idDynamic, {
      placeholder: "Enter search words"
    }));
    this._filterRender.push(PropertyPaneDropdown('logicOperator' + this.idDynamic, {
      label: "",
      selectedKey: "",
      options: [
        { key: '', text: 'Select logic operator' },
        { key: ' and ', text: 'and' },
        { key: ' or ', text: 'or' }
      ]
    }));

    this._filterRender.push(PropertyPaneButton('removeFilter' + this.idDynamic,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.RemoveFilterRuleClick.bind(this, this.idDynamic)
      }));

    this._filterRender.push(PropertyPaneButton('AddFilter' + this.idDynamic,
      {
        text: "Add filter",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Add",
        onClick: this.AddNewFilterClick.bind(this),
        disabled: this.fieldsButtonDisabled
      }));
    // this.UpdatePropertiesDynamic(this.idDynamic);
  }

  private addDefaultControlsForFilterExternal(): void {
    this._filterExternalRender.push(PropertyPaneDropdown('filterexternal' + this.idDynamicExternal, {
      label: strings.FiltersFieldLabel,
      options: this._externalFields,
      selectedKey: '',
      disabled: this._externalFieldsDropdownDisabled
    }));
    this._filterExternalRender.push(PropertyPaneDropdown('operatorexternal' + this.idDynamicExternal, {
      label: "",
      selectedKey: "select",
      options: [
        { key: 'select', text: 'Select filter operator' },
        { key: 'eq', text: 'Equals' },
        { key: 'ne', text: 'Not Equals' },
        { key: 'gt', text: 'Greater than' },
        { key: 'ge', text: 'Greater than or Equals' },
        { key: 'lt', text: 'Less than' },
        { key: 'le', text: 'Less than or Equals' },
        { key: 'IsNotNull', text: 'Is Not Null' },
        { key: 'IsNull', text: 'Is Null' },
        { key: 'contains', text: 'Contains' },
        { key: 'beginsWith', text: 'Begins With' }
      ]
    }));
    this._filterExternalRender.push(PropertyPaneTextField('valueFilterexternal' + this.idDynamicExternal, {
      placeholder: "Enter search words"
    }));
    this._filterExternalRender.push(PropertyPaneDropdown('logicOperatorexternal' + this.idDynamicExternal, {
      label: "",
      selectedKey: "",
      options: [
        { key: '', text: 'Select logic operator' },
        { key: ' and ', text: 'and' },
        { key: ' or ', text: 'or' }
      ]
    }));

    this._filterExternalRender.push(PropertyPaneButton('removeFilterexternal' + this.idDynamicExternal,
      {
        text: "Remove",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Trash",
        onClick: this.RemoveFilterExternalRuleClick.bind(this, this.idDynamicExternal)
      }));

    this._filterExternalRender.push(PropertyPaneButton('AddFilterexternal' + this.idDynamicExternal,
      {
        text: "Add filter",
        buttonType: PropertyPaneButtonType.Command,
        icon: "Add",
        onClick: this.AddNewFilterExternalClick.bind(this),
        disabled: this._externalFieldsButtonDisabled
      }));
    // this.UpdatePropertiesDynamic(this.idDynamic);
  }

  private RemoveFilterRuleClick(idControl: any) {

    // enable add button when is one filter at first in the array
    if (this._filterRender.length == 6) {
      let controlTpUpdate = this._filterRender[5];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterRender[5] = controlTpUpdate;
    }
    // if we have more than one filter and the last one logicOperator have a value selected then enable add button, basically are filter in the middle of the array
    if (this._filterRender.length > 6 && this.properties[this._filterRender[this._filterRender.length - 3].targetProperty] != undefined) {
      let controlTpUpdate = this._filterRender[this._filterRender.length - 1];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterRender[this._filterRender.length - 1] = controlTpUpdate;
    }

    // remove the last filter that not have logic operator selected
    if (this._filterRender.length > 6 && this.properties[this._filterRender[this._filterRender.length - 3].targetProperty] == undefined && this._filterRender[this._filterRender.length - 1].targetProperty == "AddFilter" + idControl) {
      let controlTpUpdate = this._filterRender[this._filterRender.length - 1];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterRender[this._filterRender.length - 1] = controlTpUpdate;
    }
    // remove filters
    let filterControlsToRemove = this.filterToRemove(idControl)
    filterControlsToRemove.forEach(controlToRemove => {
      this._filterRender = this._filterRender.filter(obj => obj.targetProperty !== controlToRemove.targetProperty);

      // clean values added in dynamic properties
      filterControlsToRemove.forEach((element) => {
       this.properties[element.targetProperty] = undefined;
      });

    });
    // assing to the property
    this.properties.filters = this._filterRender;
  }

  private RemoveFilterExternalRuleClick(idControl: any) {

    // enable add button when is one filter at first in the array
    if (this._filterExternalRender.length == 6) {
      let controlTpUpdate = this._filterExternalRender[5];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterExternalRender[5] = controlTpUpdate;
    }
    // if we have more than one filter and the last one logicOperator have a value selected then enable add button, basically are filter in the middle of the array
    if (this._filterExternalRender.length > 6 && this.properties[this._filterExternalRender[this._filterExternalRender.length - 3].targetProperty] != undefined) {
      let controlTpUpdate = this._filterExternalRender[this._filterExternalRender.length - 1];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterExternalRender[this._filterExternalRender.length - 1] = controlTpUpdate;
    }

    // remove the last filter that not have logic operator selected
    if (this._filterExternalRender.length > 6 && this.properties[this._filterExternalRender[this._filterExternalRender.length - 3].targetProperty] == undefined && this._filterExternalRender[this._filterExternalRender.length - 1].targetProperty == "AddFilter" + idControl) {
      let controlTpUpdate = this._filterExternalRender[this._filterExternalRender.length - 1];

      controlTpUpdate.properties.disabled = false;
      // update record in the array
      this._filterExternalRender[this._filterExternalRender.length - 1] = controlTpUpdate;
    }
    // remove filters
    let filterControlsToRemove = this.filterExternalToRemove(idControl)
    filterControlsToRemove.forEach(controlToRemove => {
      this._filterExternalRender = this._filterExternalRender.filter(obj => obj.targetProperty !== controlToRemove.targetProperty);

      // clean values added in dynamic properties
      filterControlsToRemove.forEach((element) => {
       this.properties[element.targetProperty] = undefined;
      });

    });
    // assing to the property
    this.properties.filtersExternal = this._filterExternalRender;
  }

//TODO make this generic to use this in both kind of list
  private CleanDynamicFilters(limit: number) {
    for (let index = 0; index < limit; index++) {
      if (this.properties["filter" + index] != undefined
        && this.properties["operator" + index] != undefined
        && this.properties["valueFilter" + index] != undefined) {
        this.properties["filter" + index] = undefined;
        this.properties["operator" + index] = undefined;
        this.properties["valueFilter" + index] = undefined;
        this.properties["logicOperator" + index] = undefined;
      }
    }
  }

  private CleanDynamicFiltersExternal(limit: number) {
    for (let index = 0; index < limit; index++) {
      if (this.properties["filterexternal" + index] != undefined
        && this.properties["operatorexternal" + index] != undefined
        && this.properties["valueFilterexternal" + index] != undefined) {
        this.properties["filterexternal" + index] = undefined;
        this.properties["operatorexternal" + index] = undefined;
        this.properties["valueFilterexternal" + index] = undefined;
        this.properties["logicOperatorexternal" + index] = undefined;
      }
    }
  }
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // the following line is to allow add filter after select an logic operator if not selected you won't be able to add more filter
    if (propertyPath.lastIndexOf("logicOperator", 0) == 0) {
      // update array enabling add button
      this._filterRender.forEach(controlToUpdate => {
        // If start with and control selected is not empty and always oldvlue needs to be empty
        if (controlToUpdate.targetProperty.lastIndexOf("AddFilter", 0) == 0 && this.properties[propertyPath] != "" && oldValue == undefined) {
          // enable add button
          controlToUpdate.properties.disabled = false;
          // update record in the array
          this._filterRender[propertyPath] = controlToUpdate;
        }
      });

      if (this.properties[propertyPath] == "") {
        // Show Alert
        alert("Please select a logic operator. ")
        // rollback selection
        this.properties[propertyPath] = oldValue;
      }
    }

    if (propertyPath.lastIndexOf("logicOperatorexternal", 0) == 0) {
      // update array enabling add button
      this._filterExternalRender.forEach(controlToUpdate => {
        // If start with and control selected is not empty and always oldvlue needs to be empty
        if (controlToUpdate.targetProperty.lastIndexOf("AddFilter", 0) == 0 && this.properties[propertyPath] != "" && oldValue == undefined) {
          // enable add button
          controlToUpdate.properties.disabled = false;
          // update record in the array
          this._filterExternalRender[propertyPath] = controlToUpdate;
        }
      });

      if (this.properties[propertyPath] == "") {
        // Show Alert
        alert("Please select a logic operator. ")
        // rollback selection
        this.properties[propertyPath] = oldValue;
      }
    }

    if (propertyPath === 'siteName' &&
    newValue) {
    // push new list value
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // get previously selected item
    const previousItem: string = this.properties.listName;
    // reset selected item
    this.properties.listName = undefined;
    // this.properties.fieldValuesName = undefined;
    // this.properties.columnsSelected = undefined;
    // this.properties.columnDisplayName = undefined;
    // this.properties.fieldValuesSortName = undefined;
    // this.properties.fieldPager = false;
    // push new item value
    this.onPropertyPaneFieldChanged('listName', previousItem, this.properties.listName);
    // disable item selector until new items are loaded
    this.listsDropdownDisabled = true;
    // refresh the item selector control by repainting the property pane
    this.context.propertyPane.refresh();
    // communicate loading items
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
      .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
        // store items
        this.lists = itemOptions;
        // enable item selector
        this.listsDropdownDisabled = false;
        // Set initial filters
        this._filterRender = new Array<any>();
        // reset filters
        this.properties.filters = undefined;
        //Reset filters
        this.CleanDynamicFilters(20);

        // reset selected columns 
        this.properties.fieldValuesName = undefined;
        this.properties.fieldValuesSortName = undefined;

        // Load initial controls for filter
        this.InitialLoadPropertyPaneFiledFilters();
        // clear status indicator
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        // re-render the web part as clearing the loading indicator removes the web part body
        this.render();
        // refresh the item selector control by repainting the property pane
        this.context.propertyPane.refresh();
      });
  }

    if (propertyPath === 'listName' &&
      newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousItem: string = this.properties.fieldName;
      // reset selected item
      this.properties.fieldName = undefined;
      this.properties.fieldValuesName = undefined;
      this.properties.columnsSelected = undefined;
      this.properties.columnDisplayName = undefined;
      this.properties.fieldValuesSortName = undefined;
      this.properties.fieldPager = false;
      // push new item value
      this.onPropertyPaneFieldChanged('fieldName', previousItem, this.properties.fieldName);
      // disable item selector until new items are loaded
      this.fieldsDropdownDisabled = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'fields');

      this.loadFields()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.fields = itemOptions;
          // enable item selector
          this.fieldsDropdownDisabled = false;
          // Set initial filters
          this._filterRender = new Array<any>();
          // reset filters
          this.properties.filters = undefined;
          //Reset filters
          this.CleanDynamicFilters(20);

          // reset selected columns 
          this.properties.fieldValuesName = undefined;
          this.properties.fieldValuesSortName = undefined;

          // Load initial controls for filter
          this.InitialLoadPropertyPaneFiledFilters();
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    // else {
    //   super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // }

    if (propertyPath === 'externalListName' &&
      newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousItem: string = this.properties.externalFieldName;
      // reset selected item
      this.properties.externalFieldName = undefined;
      this.properties.externalFieldValuesName = undefined;
      this.properties.externalColumnsSelected = undefined;
      this.properties.externalColumnDisplayName = undefined;
      this.properties.columnAsHTML = undefined;
      // this.properties.fieldValuesSortName = undefined;
      this.properties.fieldPager = false;
      // push new item value
      this.onPropertyPaneFieldChanged('externalFieldName', previousItem, this.properties.externalFieldName);
      // disable item selector until new items are loaded
      this._externalFieldsDropdownDisabled = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, '_externalFields');

      this.externalLoadFields()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this._externalFields = itemOptions;
          // enable item selector
          this._externalFieldsDropdownDisabled = false;
          // // Set initial filters
           this._filterExternalRender = new Array<any>();
          // // reset filters
           this.properties.filtersExternal = undefined;
          // //Reset filters
          this.CleanDynamicFiltersExternal(20);

          // reset selected columns 
          this.properties.externalFieldName = undefined;
          // this.properties.fieldValuesSortName = undefined;

          // Load initial controls for filter
           this.InitialLoadPropertyPaneFiledFiltersExternal();
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    // else {
    //   super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // }

     if((propertyPath !== 'listName' && !newValue) || (propertyPath !== 'externalListName' && !newValue))
    // if(propertyPath !== 'listName' || propertyPath !== 'externalListName')
    {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;
    this._externalListsDropdownDisabled = !this._externalLists;
    // if (this.sites) {
    //   return;
    // }

    this.loadSites().then((siteOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.sites = siteOptions;
      this.sitesDropdownDisabled = false;

      this.context.propertyPane.refresh();
      return this.loadLists();;
    })
    .then((listOptions: IPropertyPaneDropdownOption[]): void => {
      this.lists = listOptions;
      //this.sitesDropdownDisabled = !this.properties.siteName;

      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    });

    if (this.lists) {
      return;
    }
    if (this._externalLists) {
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'options');
    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        // load initial filter after refresh
        if (this.properties.filters)
        { //this is to have the binding in add and remove button after refresh because was lost
          this.properties.filters.forEach((filter, index) => {

            if (filter.targetProperty.lastIndexOf("removeFilter", 0) === 0) 
            {
              let id: Number = Math.floor( index / 5 );
              this.properties.filters[index] = PropertyPaneButton(filter.targetProperty,
                {
                  text: "Remove",
                  buttonType: PropertyPaneButtonType.Command,
                  icon: "Trash",
                  onClick: this.RemoveFilterRuleClick.bind(this, id)
                });
            }  
            if (filter.targetProperty.lastIndexOf("AddFilter", 0) === 0) 
            {
              this.properties.filters[index] = PropertyPaneButton(filter.targetProperty,
              {
                text: "Add filter",
                buttonType: PropertyPaneButtonType.Command,
                icon: "Add",
                onClick: this.AddNewFilterClick.bind(this),
                disabled: this.fieldsButtonDisabled
              });
            }       
          });

          this._filterRender = this.properties.filters;

        }
          
        this.context.propertyPane.refresh();
        return this.loadFields();
      })
      .then((fieldOptions: IPropertyPaneDropdownOption[]): void => {
        this.fields = fieldOptions;
        this.fieldsDropdownDisabled = !this.properties.listName;

        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
      
// Load list and columns for BCS Connection
      this.externalLoadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this._externalLists = listOptions;
        this._externalListsDropdownDisabled = false;

                // load initial filter after refresh
                if (this.properties.filtersExternal)
                { //this is to have the binding in add and remove button after refresh because was lost
                  this.properties.filtersExternal.forEach((filter, index) => {
        
                    if (filter.targetProperty.lastIndexOf("removeFilterexternal", 0) === 0) 
                    {
                      let id: Number = Math.floor( index / 5 );
                      this.properties.filtersExternal[index] = PropertyPaneButton(filter.targetProperty,
                        {
                          text: "Remove",
                          buttonType: PropertyPaneButtonType.Command,
                          icon: "Trash",
                          onClick: this.RemoveFilterExternalRuleClick.bind(this, id)
                        });
                    }  
                    if (filter.targetProperty.lastIndexOf("AddFilterexternal", 0) === 0) 
                    {
                      this.properties.filtersExternal[index] = PropertyPaneButton(filter.targetProperty,
                      {
                        text: "Add filter",
                        buttonType: PropertyPaneButtonType.Command,
                        icon: "Add",
                        onClick: this.AddNewFilterExternalClick.bind(this),
                        disabled: this._externalFieldsButtonDisabled
                      });
                    }       
                  });
        
                  this._filterExternalRender = this.properties.filtersExternal;
        
                }

        this.context.propertyPane.refresh();
        return this.externalLoadFields();
      })
      .then((fieldOptions: IPropertyPaneDropdownOption[]): void => {
        this._externalFields = fieldOptions;
        this._externalFieldsDropdownDisabled = !this.properties.externalListName;

        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
    // Initial dropdown pafe per limit
    _.range(1, 101).forEach(number => {
      this._numberPageLimit.push({ key: number, text: number.toString() });
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }
  private getFieldPropsPaneLayout(layoutSelection: boolean): IPropertyPaneField<any>[] {
    if (layoutSelection == undefined || this.layoutSelectionRender == undefined) {
      this.properties.layoutSelection = false;
      var layoutDDL = PropertyPaneToggle('layoutSelection', {
        label: strings.LayoutSelectionFieldLabel,
        offText: 'Table View'
      })
      this.layoutSelectionRender.push(layoutDDL);
    }
    if (layoutSelection) {
      this.layoutSelectionRender = [];
      var layoutDDL = PropertyPaneToggle('layoutSelection', {
        label: strings.LayoutSelectionFieldLabel,
        checked: true,
        onText: 'HTML View'
      })
      this.layoutSelectionRender.push(layoutDDL);
      this.layoutSelectionRender.push(
        PropertyPaneTextField('fieldHeader', {
          multiline: true,
          rows: 4,
          value: this.properties.fieldHeader,
          placeholder: "Enter Header here",
          label: 'Header'
        }));
      this.layoutSelectionRender.push(PropertyPaneTextField('fieldBody', {
        multiline: true,
        rows: 4,
        value: this.properties.fieldBody,
        placeholder: "Enter Body text here",
        label: 'Body'
      }));
      this.layoutSelectionRender.push(PropertyPaneTextField('fieldFooter', {
        multiline: true,
        rows: 4,
        value: this.properties.fieldFooter,
        placeholder: "Enter Footer here",
        label: 'Footer'
      }));
      this.layoutSelectionRender.push(PropertyPaneTextField('fieldCSS', {
        multiline: true,
        rows: 4,
        value: this.properties.fieldCSS,
        placeholder: "Enter CSS here",
        label: 'CSS'
      }));
      this.layoutSelectionRender.push(PropertyPaneTextField('fieldJavascript', {
        multiline: true,
        rows: 4,
        value: this.properties.fieldJavascript,
        placeholder: "Enter JavaScript here",
        label: 'Javascript'
      }));
    }
    if (layoutSelection == false) {
      this.layoutSelectionRender = [];
      var layoutDDL = PropertyPaneToggle('layoutSelection', {
        label: strings.LayoutSelectionFieldLabel,
        checked: false,
        offText: 'Table View'
      })
      this.layoutSelectionRender.push(layoutDDL);
      this.layoutSelectionRender.push(PropertyPaneCheckbox('fieldShowTitle', {
        text: strings.ShowTitle
      }));
    }
    return this.layoutSelectionRender;
  }
// for rendering the controls
  private getFilterURLParameterRender(): IPropertyPaneField<any>[] {
    let controls = new Array<IPropertyPaneField<any>>();
    // if (this.variableSourceParametersRender.length == 0) {
    //   this.variableSourceParametersRender.push(this.initialFilterURLParameter);
    // }
    if (this._variableSourceParametersRender != undefined) {
      if (this._variableSourceParametersRender.length == 0 && this.properties.variableSourceParametersRender != undefined) 
      {
        let countForId: number = 0;
        this.properties.variableSourceParametersRender.forEach(element => {
          if(element!=null)
          {
            // if (element.trashControl != null) 
            // {
              let trashControl = PropertyPaneButton('removeParamFilter' + countForId ,
              {
                text: "Remove",
                buttonType: PropertyPaneButtonType.Command,
                icon: "Trash",
                onClick: this.filterParameterToRemove.bind(this, countForId)  // this.filterParameterToRemove.bind(this, newId)
              })
              // controls.push(trashControl);
              element.trashControl = trashControl; // add remove to bindng again
            //} 
            this._variableSourceParametersRender.push(element);
          } 
          countForId++;
        });
      }
      this._variableSourceParametersRender.forEach(filter => {
        if(filter!=null)
        {
          controls.push(filter.parameterNameControl);
          controls.push(filter.variableNameControl);
          controls.push(filter.trashControl);
        }
      });
    }
    return controls;
  }
  private getFilterSegmentRender(): IPropertyPaneField<any>[] {
    let controls = new Array<IPropertyPaneField<any>>();
    // if (this.variableSourceParametersRender.length == 0) {
    //   this.variableSourceParametersRender.push(this.initialFilterURLParameter);
    // }
    if (this._variableSourceSegmentsRender != undefined) {
      if (this._variableSourceSegmentsRender.length == 0 && this.properties.variableSourceSegmentsRender != undefined) 
      {
        let countForId: number = 0;
        this.properties.variableSourceSegmentsRender.forEach(element => {
          if(element!=null)
          {
            // if (element.trashControl != null) 
            // {
              let trashControl = PropertyPaneButton('removeSegmentFilter' + countForId ,
              {
                text: "Remove",
                buttonType: PropertyPaneButtonType.Command,
                icon: "Trash",
                onClick: this.filterSegmentToRemove.bind(this, countForId)  // this.filterParameterToRemove.bind(this, newId)
              })
              // controls.push(trashControl);
              element.trashControl = trashControl; // add remove to bindng again
            // } 
            this._variableSourceSegmentsRender.push(element);
          } 
          countForId++;
        });
      }
      this._variableSourceSegmentsRender.forEach(filter => {
        if(filter != null)
        {
          controls.push(filter.indexSegmentNameControl);
          controls.push(filter.segmentVariableNameControl);
          controls.push(filter.trashControl);
        }
      });
    }
    return controls;
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    //configure conditional properties    
    //wp default expand/collapse control variable
    let wpDefaultState: any = [];

    //conditional statement, if expnd/collapse is true, show toggle open or close
    if (this.properties.enableExpandCollapse) {
      wpDefaultState = PropertyPaneDropdown('expandCollapseDefaultState', {
        label: strings.WebPartDefaultStateLabel,
        options: [
          { key: 'Open', text: 'Open' },
          { key: 'Close', text: 'Close' }
        ]
      })
    }

    // #region ##########Variable source setting##########
    const url = new URL(window.location.href);
    const queryParms = new URLSearchParams(url.searchParams);
    // let countParameters = Array.from(url.searchParams).length; // this line doesn't works then use regex
    var matches = url.toString().match(/[a-z\d]+=[a-z\d]+/gi);
    var countParameters = matches ? matches.length : 0;

    this._variableSourceParameters = [];
    for (let index = 0; index <= countParameters; index++) 
    {
      this._variableSourceParameters.push({ variableName: "$$" + this.properties["fieldQueryUrlVariableName" + index] + "$$", value: queryParms.get(this.properties["fieldQueryUrlName" + index]) });
    }
    
    //get segment in the URL
    let segmentsComplete = url.toString().split('/');
    //remove two first elements protocol and empty value after the split
    let arrUrlSegments = segmentsComplete.slice(2, segmentsComplete.length);

    var countSegment = url.toString().split('/').length - 2;

    this._variableSourceSegments = [];
    for (let index = 0; index <= countSegment; index++) 
    {
      this._variableSourceSegments.push({ variableSegmentName: "$$" + this.properties["fieldQueryUrlSegmentName" + index] + "$$", value: arrUrlSegments[this.properties["fieldSegmentIndex" + index] - 1]});
    }

    // Save properties with variable source values
    this.properties.variableSourceParameters = this._variableSourceParameters;
    this.properties.variableSourceSegments = this._variableSourceSegments;

   //#endregion ##############################################
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.AppearanceGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('titleURL', {
                  label: strings.TitleURLFieldLabel,
                  description: "Enter the Title Url"
                }),
                PropertyPaneDropdown('titleColor', {
                  label: strings.TitleColorFieldLabel,
                  options: [
                    { key: 'Blue', text: 'Blue' },
                    { key: 'Green', text: 'Green' },
                    { key: 'Grey', text: 'Grey' },
                  ],
                  selectedKey: 'Blue'
                }),
                PropertyPaneToggle('enableExpandCollapse', {
                  key: 'Expand/Collapse',
                  label: strings.EnableExpandCollapse,
                  checked: false
                }),
                wpDefaultState //conditonal variable
              ]
            },
            {
              groupName: strings.DataSourceGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneDropdown('siteName', {
                  label: strings.SitesFieldLabel,
                  selectedKey: '',
                  options: this.sites,
                  disabled: this.sitesDropdownDisabled
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.ListsFieldLabel,
                  options: this.lists,
                  selectedKey: '',
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('fieldName', {
                  label: strings.ColumnsFieldLabel,
                  options: this.fields,
                  selectedKey: '',
                  disabled: this.fieldsDropdownDisabled
                }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Add Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Add",
                    onClick: this.SelectedColumnsClick.bind(this)
                  }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Remove Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Remove",
                    onClick: this.RemoveColumnsClick.bind(this)
                  }),
                PropertyPaneLabel('fieldValuesName', {
                  text: this.ParseToDisplayName(this.properties.columnDisplayName)
                }),
                PropertyPaneDropdown('sortName', {
                  label: strings.SortbyFieldLabel,
                  options: this.fields,
                  selectedKey: '',
                  disabled: this.fieldsDropdownDisabled
                }),
                PropertyPaneButton('ClickFilter',
                  {
                    text: "Add Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Add",
                    onClick: this.SelectedSortByClick.bind(this)
                  }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Remove Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Remove",
                    onClick: this.RemoveSortByClick.bind(this)
                  }),
                PropertyPaneLabel('fieldValuesFilterName', {
                  text: this.properties.fieldValuesSortName
                }),
                PropertyPaneCheckbox('ascending', {
                  text: strings.AscendingDescending
                }),
                PropertyPaneDropdown('groupByDDL', {
                  label: strings.GroupbyFieldLabel,
                  options: this.fields,
                  selectedKey: '',
                  disabled: this.fieldsDropdownDisabled
                })
              ].concat(this._filterRender)
            },
            {
              groupName: strings.BCSConnectionGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneDropdown('externalListName', {
                  label: strings.ExternalListsFieldLabel,
                  options: this._externalLists,
                  selectedKey: '',
                  disabled: this._externalListsDropdownDisabled
                }),
                PropertyPaneDropdown('externalFieldName', {
                  label: strings.ExternalColumnsFieldLabel,
                  options: this._externalFields,
                  selectedKey: '',
                  disabled: this._externalFieldsDropdownDisabled
                }),
                PropertyPaneTextField('displayNameRename', {
                  label: strings.DisplayNameFieldLabel,
                  description: "Enter the DisplayName for column selected"
                }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Add Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Add",
                    onClick: this.ExternalSelectedColumnsClick.bind(this)
                  }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Remove Field",
                    buttonType: PropertyPaneButtonType.Command,
                    icon: "Remove",
                    onClick: this.ExternalRemoveColumnsClick.bind(this)
                  }),
                PropertyPaneLabel('externalFieldValuesName', {
                  text: this.ParseToDisplayName(this.properties.externalColumnDisplayName)
                }),
                PropertyPaneTextField('columnAsHTML', {
                  label: strings.ColumnHTMLabel,
                  description: "Enter HTML Column"
                }),
              ].concat(this._filterExternalRender)
            },
            {
              groupName: strings.VariableSourceGroupName,
              isCollapsed: true,
              groupFields:
                [
                  PropertyPaneLabel('fieldQueryUrlParameter', {
                    text: strings.fieldQueryUrlParameter,
                  })].concat(this.getFilterURLParameterRender())
                     .concat(PropertyPaneButton('AddFilterParam',
                     {
                       text: "Add Parameter",
                       buttonType: PropertyPaneButtonType.Command,
                       icon: "Add",
                       onClick: this.AddFilterParamClick.bind(this),
                       disabled: false
                     }))
                     .concat(PropertyPaneLabel('fieldQueryUrlSegment', {
                       text: strings.fieldQueryUrlSegment,
                     }))
                     .concat(this.getFilterSegmentRender())  
                     .concat(PropertyPaneButton('AddFilterSegment',
                     {
                       text: "Add Segment",
                       buttonType: PropertyPaneButtonType.Command,
                       icon: "Add",
                       onClick: this.AddFilterSegmentClick.bind(this),
                       disabled: false
                     }))         
            },
            {
              groupName: strings.DisplayOptionsGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneCheckbox('fieldPager', {
                  text: strings.PagerFieldLabel
                }),
                // PropertyPaneSlider('fieldPageLimit', {
                //   label: strings.PageLimitFieldLabel,
                //   min: 1,
                //   max: 100,
                //   value: 1,
                //   showValue: true,
                //   step: 1
                // }),
                PropertyPaneDropdown('fieldPageLimit', {
                  label: strings.PageLimitFieldLabel,
                  options: this._numberPageLimit,
                  selectedKey: 100,
                  disabled: false
                }),
                PropertyPaneCheckbox('fieldExportToExcel', {
                  text: strings.ExportToExcelFieldLabel
                }),
                PropertyPaneCheckbox('fieldPrint', {
                  text: strings.PrintFieldLabel
                }),
                PropertyPaneCheckbox('fieldTotal', {
                  text: strings.TotalFieldLabel
                })
              ]
            },
            {
              groupName: strings.LayoutGroupName,
              isCollapsed: true,
              groupFields:
                this.getFieldPropsPaneLayout(this.properties.layoutSelection)
            }
          ]
        }
      ]
    };
  }
}
