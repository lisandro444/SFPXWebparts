import * as React from 'react';
import styles from './SpfxCustomListView.module.scss';
import { ISpfxCustomListViewProps } from './ISpfxCustomListViewProps';
import { PnPService } from '../services/PnPService';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISpfxCustomListViewState } from './ISpfxCustomListViewState';
import * as jQuery from 'jquery';
import { Guid } from "guid-typescript"; //import GUID generator
import  * as _ from 'lodash'; //import lodash utilities
// import { Link } from 'office-ui-fabric-react';
import { CSVLink, CSVDownload } from "react-csv";
import { FilterPropertiesValue } from '../Entities/FilterPropertiesValue';
import { FilterURLParameter } from '../Entities/FilterURLParameter';
import { FilterSegment } from '../Entities/FilterSegment';
import { JavaScript } from 'SpfxCustomListViewWebPartStrings';
import * as moment from 'moment';

export default class SpfxCustomListView extends React.Component<ISpfxCustomListViewProps, ISpfxCustomListViewState> {
  private _pnpService;
   private _titleColor: 'Blue';
  private _filtersValuesSelected: Array<FilterPropertiesValue> = new Array();
  private _filtersExternalValuesSelected: Array<FilterPropertiesValue> = new Array();
  private _paramtersVariableSource: Array<FilterURLParameter> = new Array();
  private _segmentsVariableSource: Array<FilterSegment> = new Array();
  private _itemsFiltered: Array<any> = new Array();
  // private _renderingBody: any;
  private _totalItems: number;
  public constructor(props) {
    super(props); //initiate access of wp properties
    this._pnpService = new PnPService(this.props.context);
     //generate unique guid for toggle function
     let uID = Guid.raw(); //generate GUID as a string
     uID = _.replace(uID, new RegExp('-','g'), ''); //lodash replace all instances of "-" character in guid
     let toggleID = "tglSpan" + uID;
     let contentContainerID = "Container" + uID;
     let rangePager: Record<'min' | 'max', number> = null;
    this.state = {
      url: this.props.context.pageContext.web.absoluteUrl,
      description: this.props.description,
      titleColor: this._titleColor,
      iconToggle: '',
      cssTitleColor: 'BlueBG',
      cssToggleWrapperDisplay: '',
      uniqueToggleID: toggleID,
      contentContainer: contentContainerID,
      // DataSource
      listName: this.props.listName,
      fieldName: this.props.fieldName,
      fieldValuesName: this.props.fieldValuesName,
      fieldValuesSortName: this.props.fieldValuesSortName,
      itemsresult: new Array(),
      itemsexternalresult: new Array(),
      filters: this.props.filters,
      totalListItems: 0,

      // Layout
      layoutSelection: this.props.layoutSelection,
      fieldHeader: this.props.fieldHeader,
      fieldBody: this.props.fieldBody,
      fieldFooter: this.props.fieldFooter,
      fieldCSS: this.props.fieldCSS,
      fieldJavascript: this.props.fieldJavascript,
      fieldShowTitle: this.props.fieldShowTitle
    };

    //Register event handlers
    this._toggleClick = this._toggleClick.bind(this);
    this._openToggle = this._openToggle.bind(this);
    this._closeToggle = this._closeToggle.bind(this);
    this.handleChangePager = this.handleChangePager.bind(this);

    // Rendering Variable Source values
    this.loadVariableSourceParameters(20);
    this.loadVariableSourceSegments(20);
    // Rendering Variable Source values
    this.loadFiltersValues(20);
    this.loadFiltersExternalValues(20);

    // if(this.props.fieldPager)
    // {
      // let value = jQuery("#selectPager")[0].value.split(" to ");;
      // let min = value[0];
      // let max = value[1];
      let min = 1;
      let max = this.props.fieldPageLimit;
      rangePager = {
        min: min,
        max: max
      }
    // }
    if (this.props.listName && this.props.fieldValuesName) {
      this._pnpService.getItemsFromSPList(this.props.siteName, this.props.listName, this.props.fieldValuesName, this._filtersValuesSelected, this.props.fieldValuesSortName, this.props.variableSourceParameters, this.props.variableSourceSegments, this.props.ascending, rangePager).then(items => {
        if (items) {
            this.setState({
              itemsresult: items,
              iconToggle: localStorage.getItem('myVisibleState')
            });
        }
      });
    }

    if (this.props.externalListName && this.props.externalFieldValuesName) {
        this._pnpService.getItemsFromBCSExternalList(this.props.externalListName, this.props.externalFieldValuesName, this._filtersExternalValuesSelected, this.props.variableSourceParameters, this.props.variableSourceSegments, rangePager).then(items => {
          if (items) {
            this.setState({
              itemsexternalresult: items
            });
          }
        });
    }

    if (this.props.listName) {
      this._pnpService.getTotalItems(this.props.siteName, this.props.listName).then(itemsLenght => {
        if (itemsLenght) {
          //this._totalItems = items.length
          this.setState({
            totalListItems: itemsLenght,
          });
        }
      });
    }
  }

  public componentDidMount() {

    //this.setState({iconToggle: localStorage.getItem('myVisibleState')});

  }

  public componentDidUpdate(prevProps, prevState) {
   //invoked immediately after updating occurs on props or state
  //compare previous and current props then take action

  //update local storage with current icon toggle state
    localStorage.setItem('myiconState', this.state.iconToggle);

       //close/open status of the toggle
       if(this.props.expandCollapseDefaultState !== prevProps.expandCollapseDefaultState){
        if(this.props.expandCollapseDefaultState == 'Close'){
          this._openToggle();
        }
        if(this.props.expandCollapseDefaultState == 'Open'){
          this._closeToggle();
        }

      }

      //if expand/collapse toggle is active
      if(this.props.enableExpandCollapse !== prevProps.enableExpandCollapse) {
        if(this.props.enableExpandCollapse == true){
          this.setState({
            cssToggleWrapperDisplay: 'block'
          });
          //localStorage.setItem('myVisibleState', this.state.cssToggleWrapperDisplay);

          if(this.props.expandCollapseDefaultState == 'Open'){
            this._closeToggle();
          }
          if(this.props.expandCollapseDefaultState == 'Close'){
            this._openToggle();
          }


        }else{

          //if expand/collapse toggle is inactive, hide toggle function, show hidden web part or hidden content
          this.setState({
            cssToggleWrapperDisplay: 'none'
          });


          //if Web Part Appearance Component is not empty, toggle web part container
          //else, toggle next web part
          let toggledWebPart;
          let wpContentDiv = jQuery("." +`${this.state.contentContainer}`);
          if(!jQuery(wpContentDiv).is(':empty')){
            //show/hide web part content
            toggledWebPart = jQuery("." +`${this.state.contentContainer}`);

          }else{

            //hide/show next web part
            //if modern version use ControlZone
            //if Publishing version use s4-wpcell-plain

            let modernWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.ControlZone').nextAll('.ControlZone:first');
            let publishingWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.s4-wpcell-plain').next('div');

            console.log(modernWPZone.length);
            console.log(publishingWPZone.length);

            if(modernWPZone.length > 0){
              toggledWebPart = modernWPZone;
            }else{
              toggledWebPart = publishingWPZone;
            }
          }
          let wpDisplay = true;
          jQuery(toggledWebPart).toggle(wpDisplay);


        }

      }

    // Rendering Variable Source values
    this.loadVariableSourceParameters(20);
    this.loadVariableSourceSegments(20);
    // Rendering Filters values
    this.loadFiltersValues(20);
    this.loadFiltersExternalValues(20);

    if (this.props.fieldValuesName !==  prevProps.fieldValuesName)
    {
      let rangePager: Record<'min' | 'max', number> = null;
      // if(this.props.fieldPager)
      // {
        let min = 1;
        let max = this.props.fieldPageLimit;
        //this._firstTime = false;
        rangePager = {
          min: min,
          max: max
        }
      //}
      this._pnpService.getItemsFromSPList(this.props.siteName, this.props.listName, prevProps.fieldValuesName,this._filtersValuesSelected, prevProps.fieldValuesSortName, this.props.variableSourceParameters, this.props.variableSourceSegments, this.props.ascending, rangePager).then(items => {
        if (items) {
          this.setState({
            itemsresult: items,
            totalListItems: items.length
          });
        }
      });
    }

    if (this.props.externalFieldValuesName !==  prevProps.externalFieldValuesName)
    {
      let rangePager: Record<'min' | 'max', number> = null;
      // if(this.props.fieldPager)
      // {
        let min = 1;
        let max = this.props.fieldPageLimit;
        //this._firstTime = false;
        rangePager = {
          min: min,
          max: max
        }
      // }
      if (this.props.externalListName) {
        this._pnpService.getItemsFromBCSExternalList(this.props.externalListName, this.props.externalFieldValuesName, this._filtersExternalValuesSelected, this.props.variableSourceParameters, this.props.variableSourceSegments, rangePager).then(items => {
          if (items) {
            this.setState({
              itemsexternalresult: items
            });
          }
        });
      }
    }

    if(this.props.fieldPageLimit !== prevProps.fieldPageLimit)
    {
      let rangePager: Record<'min' | 'max', number> = null;
      // if(this.props.fieldPager)
      // {
        let min = 1;
        let max = this.props.fieldPageLimit;
        //this._firstTime = false;
        rangePager = {
          min: min,
          max: max
        }
      // }
      this._pnpService.getItemsFromSPList(this.props.siteName, this.props.listName, prevProps.fieldValuesName,this._filtersValuesSelected, prevProps.fieldValuesSortName, this.props.variableSourceParameters, this.props.variableSourceSegments, this.props.ascending, rangePager).then(items => {
        if (items) {
          this.setState({
            itemsresult: items
          });
        }
      });

    }

  }

  public componentWillReceiveProps(props) {
    // Rendering Variable Source values
    this.loadVariableSourceParameters(20);
    this.loadVariableSourceSegments(20);
    // Rendering Filters values
    this.loadFiltersValues(20);
    this.loadFiltersExternalValues(20);

    let rangePager: Record<'min' | 'max', number> = null;
    // if (this.props.fieldPager) {
    let value;
    let max;
    if(jQuery("#selectPager")[0])
    {
      value = jQuery("#selectPager")[0].value.split(" to ");
    }
    else value = [1,100];
      let min = value[0];
      //let max = value[1];
      if(!this.props.fieldPager)
      {
        max = props.fieldPageLimit;
      }
      else
      {
        max = value[1];
      }

      // let min = 1;
      // let max = this.props.fieldPageLimit;
      rangePager = {
        min: min,
        max: max
      }
    // }
    if (this.props.listName) {
      this._pnpService.getItemsFromSPList(this.props.siteName, this.props.listName, this.props.fieldValuesName, this._filtersValuesSelected, props.fieldValuesSortName, this.props.variableSourceParameters, this.props.variableSourceSegments, props.ascending, rangePager, props.externalListName).then(items => {
        if (items) {
          this.setState({
            itemsresult: items
          });
        }
      });
    }
    if (this.props.externalListName) {
      this._pnpService.getItemsFromBCSExternalList(props.externalListName, this.props.externalFieldValuesName, this._filtersExternalValuesSelected, this.props.variableSourceParameters, this.props.variableSourceSegments, rangePager).then(items => {
        if (items) {
          this.setState({
            itemsexternalresult: items
          });
        }
      });
    }
  }

  public _openToggle() {
    //set icon to (+) and hide the next web part
    this.setState({
      iconToggle: 'icon-open'
    });

    //if Web Part Appearance Component is not empty, toggle web part container
    //else, toggle next web part
    let toggledWebPart;
    let wpContentDiv = jQuery("." +`${this.state.contentContainer}`);

    if(!jQuery(wpContentDiv).is(':empty')){
      //show/hide web part content
      toggledWebPart = jQuery("." +`${this.state.contentContainer}`);

    }else{

      //hide/show next web part
      //if modern version use ControlZone
      //if Publishing version use s4-wpcell-plain

      let modernWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.ControlZone').nextAll('.ControlZone:first');
      let publishingWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.s4-wpcell-plain').next('div');

      console.log(modernWPZone.length);
      console.log(publishingWPZone.length);

      if(modernWPZone.length > 0){
        toggledWebPart = modernWPZone;
      }else{
        toggledWebPart = publishingWPZone;
      }

    }
    let wpDisplay = false;
    jQuery(toggledWebPart).toggle(wpDisplay);
  }

  public _closeToggle(){
    //set icon to (x) and show the next web part
    this.setState({
      iconToggle: 'icon-close'
    });

    //if Web Part Appearance Component is not empty, toggle web part container
    //else, toggle next web part
    let toggledWebPart;
    let wpContentDiv = jQuery("." +`${this.state.contentContainer}`);
    if(!jQuery(wpContentDiv).is(':empty')){
      //web part content
      toggledWebPart = jQuery("." +`${this.state.contentContainer}`);

    }else{

      //hide/show next web part
      //if modern version use ControlZone
      //if Publishing version use s4-wpcell-plain

      let modernWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.ControlZone').nextAll('.ControlZone:first');
      let publishingWPZone = jQuery("." +`${this.state.uniqueToggleID}`).closest('.s4-wpcell-plain').next('div');

      console.log(modernWPZone.length);
      console.log(publishingWPZone.length);

      if(modernWPZone.length > 0){
        toggledWebPart = modernWPZone;
      }else{
        toggledWebPart = publishingWPZone;
      }

    }

    let wpDisplay = true;
    jQuery(toggledWebPart).toggle(wpDisplay);

  }

  public _toggleClick(togID: string){
    let toggleState = this.state.iconToggle;
    let toggleItem = togID;
    let toggledWebPart;

    //if Web Part Appearance Component is not empty, toggle web part container
    //else, toggle next web part
    let wpContentDiv = jQuery("." +`${this.state.contentContainer}`);
    if(!jQuery(wpContentDiv).is(':empty')){
      //show/hide web part content
      toggledWebPart = jQuery("." +`${this.state.contentContainer}`);
    }else{

      //hide/show next web part
      //if modern version use ControlZone
      //if Publishing version use s4-wpcell-plain

      let modernWPZone = jQuery("." +`${toggleItem}`).closest('.ControlZone').nextAll('.ControlZone:first');
      let publishingWPZone = jQuery("." +`${toggleItem}`).closest('.s4-wpcell-plain').next('div');

      console.log(modernWPZone.length);
      console.log(publishingWPZone.length);

      if(modernWPZone.length > 0){
        toggledWebPart = modernWPZone;
      }else{
        toggledWebPart = publishingWPZone;
      }


    }


    //toggle visibility and animate
    jQuery(toggledWebPart).toggle('fast', function() {});

    //toggle state to icon close (x)
    if(toggleState == 'icon-open'){
      this.setState({
        iconToggle: 'icon-close'
      });

    }else{ //toggle state to icon-open (+)
      this.setState({
        iconToggle: 'icon-open'
      });
    }


  }

//this method get the all values selected in the dataSource filters
  private loadFiltersValues(limit: number) {
    this._filtersValuesSelected = [];
    for (let index = 0; index < limit; index++) {
      if (this.props.dynamicPropertiesFilters['filter' + index] && this.props.dynamicPropertiesFilters['operator' + index] && this.props.dynamicPropertiesFilters['valueFilter' + index]) {
        this._filtersValuesSelected.push({
          filter: this.props.dynamicPropertiesFilters['filter' + index] ,
          operator: this.props.dynamicPropertiesFilters['operator' + index],
          valueFilter: this.props.dynamicPropertiesFilters['valueFilter' + index],
          logicOperator: this.props.dynamicPropertiesFilters['logicOperator' + index]
        });
      }
    }
  }

  private loadFiltersExternalValues(limit: number) {
    this._filtersExternalValuesSelected = [];
    for (let index = 0; index < limit; index++) {
      if (this.props.dynamicPropertiesFilters['filterexternal' + index] && this.props.dynamicPropertiesFilters['operatorexternal' + index] && this.props.dynamicPropertiesFilters['valueFilterexternal' + index]) {
        this._filtersExternalValuesSelected.push({
          filter: this.props.dynamicPropertiesFilters['filterexternal' + index] ,
          operator: this.props.dynamicPropertiesFilters['operatorexternal' + index],
          valueFilter: this.props.dynamicPropertiesFilters['valueFilterexternal' + index],
          logicOperator: this.props.dynamicPropertiesFilters['logicOperatorexternal' + index]
        });
      }
    }
  }

  //this method get the all values selected in the variable Source filters
  private loadVariableSourceParameters(limit: number) {
    this._paramtersVariableSource = [];
    for (let index = 0; index < limit; index++) {
      if (this.props.dynamicPropertiesFilters['fieldQueryUrlName' + index] && this.props.dynamicPropertiesFilters['fieldQueryUrlVariableName' + index]) {
        this._paramtersVariableSource.push({
          parameterNameControl : this.props.dynamicPropertiesFilters['fieldQueryUrlName' + index],
          variableNameControl: this.props.dynamicPropertiesFilters['fieldQueryUrlVariableName' + index]
        });
      }
    }
  }

    //this method get the all values selected in the variable Source filters
  private loadVariableSourceSegments(limit: number) {
    this._segmentsVariableSource = [];
    for (let index = 0; index < limit; index++) {
      if (this.props.dynamicPropertiesFilters['fieldSegmentIndex' + index] && this.props.dynamicPropertiesFilters['fieldQueryUrlSegmentName' + index]) {
        this._segmentsVariableSource.push({
          indexSegmentNameControl: this.props.dynamicPropertiesFilters['fieldSegmentIndex' + index],
          segmentVariableNameControl: this.props.dynamicPropertiesFilters['fieldQueryUrlSegmentName' + index]
        });
      }
    }
  }

  createSelectItemsPager() {
    let items = [];
    let maxOptions = 1;
    let maxResult = this.state.totalListItems;
    console.log("Total items: " + this.state.totalListItems);
    let multNumber = this.props.fieldPageLimit;
    if (multNumber > 0) {
      maxOptions = Math.ceil(maxResult / multNumber)
    }

    items.push(<option>1 to {multNumber}</option>);

    for (let i = 2; i <= maxOptions; i++) {
      if (multNumber <= maxResult) {
        if (i * multNumber > maxResult) {
          items.push(<option>{i * multNumber - multNumber} to {maxResult}</option>);
        }
        else {
          items.push(<option>{i * multNumber - multNumber} to {i * multNumber}</option>);
        }
      }
    }
    if (multNumber > maxResult) {
      items[0] = <option>1 to {maxResult}</option>;
    }

    return items;
  }

  handleChangePager(e) {
    let values = e.target.value.split(" to ");
    let min = values[0];
    let max = values[1];
    //this._firstTime = false;
    let rangePager: Record<'min' | 'max', number> = {
      min: min,
      max: max
    }

    if (this.props.listName && this.props.fieldValuesName) {
      this._pnpService.getItemsFromSPList(this.props.siteName, this.props.listName, this.props.fieldValuesName, this._filtersValuesSelected, this.props.fieldValuesSortName, this.props.variableSourceParameters, this.props.variableSourceSegments, this.props.ascending, rangePager).then(items => {
        if (items) {
          this._itemsFiltered = items
          this.setState({
            itemsresult: items
            // filters: this.props.filters,
            // fieldBody: this.props.fieldBody,
            // listName: this.props.listName
          });
        }
      });
    }

  }

  public render(): React.ReactElement<ISpfxCustomListViewProps> {
    let currentState = this.state;
    let currentProps = this.props;
    let renderBody: any;
    let header: any;
    let tableViewResults: any;
    let renderingColumnsResults: Array<any> = new Array<any>();
    let pagerDDL = <select id="selectPager" onChange={this.handleChangePager}>
      {this.createSelectItemsPager()}
    </select>


    let print = (): any => {
      var divContent = document.getElementById("divcontent").innerHTML;
      var print = document.getElementById("ifmcontentstoprint") as HTMLIFrameElement
      print.contentWindow.document.write(divContent);
      print.contentWindow.document.close();
      print.contentWindow.focus();
      print.contentWindow.print();
    }

    //#region Appearance Section
      //Note: all custom css is loaded from mw-portal.css via SPComponentLoader

      let color = this.props.titleColor;
      let expandable = this.props.enableExpandCollapse;

      //hide wp options header; show if tools are selected
      let wpToolsDisplay = 'none';
      if(this.props.fieldPager || this.props.fieldPrint || this.props.fieldExportToExcel){
        wpToolsDisplay = 'flex';
      }

      //expand collapse keep state on page refresh
      let toggleIcon: any;
      let toggleIconDisplay: any;
      if(this.props.expandCollapseDefaultState == 'Open'){
        toggleIcon = 'icon-close';
      }
      if(this.props.expandCollapseDefaultState == 'Close'){
        toggleIcon = 'icon-open';
      }
      if(this.props.enableExpandCollapse == true){
        toggleIconDisplay = 'block';
      }
      if(this.props.enableExpandCollapse == false){
        toggleIconDisplay = 'none';
      }

      //setup expand/collapse toggle wrapper include unique id from state
      let expandCollapseWrapper = <div className={`webPartMinRestore`} style={{display: `${toggleIconDisplay}`}}>
        <span className={`${this.state.uniqueToggleID} webPartMinRestore ${toggleIcon}`} onClick={() => this._toggleClick(this.state.uniqueToggleID)}>&nbsp;</span></div>;

      //setup Title UI and dynamically update color based on wp props,state and custom class ex:GreenBG
      //Hide Title bg if no title is entered
      let titleNoURLDisplay: any;
      if(!this.props.titleURL){
        titleNoURLDisplay = 'wpTitleNoURL';
      }

      let collapsableMenu: any;
      if(this.props.title){
        collapsableMenu = <div className={`webPartTitle ${this.props.titleColor}BG ${styles.wpTitleReact}` } role='none'>
        <span className={`spnTitle`}><a className={`${titleNoURLDisplay}`} href={this.props.titleURL} target='_blank'>{this.props.title}</a></span>{expandCollapseWrapper}</div>;
      }


      //setup container for web part content
      //This is required for the Web Part Appearance Component
      //For Web Part Title Only web part !!!Leave EMPTY
      let wpContentContainer = <div className={`${this.state.contentContainer} wpContent`}></div>

    //#endregion


    //#region HTML View
    let replaceValues = (body: string, item: any, columns: Array<{key:string, displayName:string, type:string}>): string => {
      let result = body;
      let parameters = this.props.variableSourceParameters.filter(x=>x.variableName != "$$undefined$$");
      let segments = this.props.variableSourceSegments.filter(x=>x.variableSegmentName != "$$undefined$$");
      let columnHTML = this.props.columnAsHTML;
      if (columns !== undefined && columns !== null && result) {
        columns.forEach(column => {
          if (column!== undefined && column.key == "clientname" && columnHTML) // replacement column as HTML
          {
            let linkName;
            if (parameters.length > 0) {
              parameters.forEach(param => {
                linkName = columnHTML.replace(param.variableName, param.value);
              });
            }
            if (segments.length > 0) {
              segments.forEach(segment => {
                linkName = columnHTML.replace(segment.variableSegmentName, segment.value);
              });
            }
            if (column.displayName) {
              linkName = linkName.replace("$$" + column.displayName + "$$", item[column.key]);
              result = result.replace("$$" + column.displayName + "$$", linkName);
            }
            else {
              linkName = linkName.replace("$$" + column.displayName + "$$", item[column.key]);
              result = result.replace("$$" + column.key + "$$", linkName)
            }
          }
          else {
            if (column !== undefined && column.displayName) {
              let itemValue = item[column.key];
              if (itemValue) {
                if (column.type == "SP.FieldDateTime") 
                {
                  itemValue = moment(itemValue).utc().format("MM/DD/YYYY")
                  result = result.replace("$$" + column.displayName + "$$", itemValue)
                }
                if (column.type == "SP.FieldUrl") 
                {
                  if(result.indexOf("$$" + column.displayName + "$$") >= 0)
                  {
                    let asLink;
                    if (itemValue.Description)
                    {
                      asLink = '<a href="' + itemValue.Url + '" target="_blank">' + itemValue.Description + '</a>'
                    }
                    else
                    {
                      asLink = '<a href="' + itemValue.Url + '" target="_blank">' + itemValue.Url + '</a>'
                    }
                    
                    result = result.replace("$$" + column.displayName + "$$", asLink)
                  }
                  if(result.indexOf(column.displayName + ".Url") >= 0)
                  {
                    let asLink = itemValue.Url;
                    result = result.replace("$$" + column.displayName + ".Url$$", asLink)
                  }
                  if(result.indexOf(column.displayName + ".Description") >= 0)
                  {
                    itemValue = itemValue.Description;
                    result = result.replace("$$" + column.displayName + ".Description$$", itemValue)
                  }
                }
                else
                {
                  result = result.replace("$$" + column.displayName + "$$", itemValue)
                }
              }
            }
            else {
              if (column !== undefined) {
                let itemValue = item[column.key];
                if (column.type == "SP.FieldDateTime") {
                  itemValue = moment(itemValue).utc().format("MM/DD/YYYY")
                }
                result = result.replace("$$" + column.key + "$$", itemValue)
              }
            }
          }
        });
      }
      if(result) {
        result = result.replace(/null/g, '')
          .replace(/undefined/g, '');
      }
      return result;
    };
    if (this.state.itemsresult && this.state.itemsexternalresult && this.state.itemsexternalresult.length == 0) {
      renderBody = this.state.itemsresult.map(function (item) {
        return (
          <div className={styles.row} dangerouslySetInnerHTML={{ __html: replaceValues(currentProps.fieldBody, item, currentProps.columnDisplayName) }} />
        );
      });
    }
    if (this.state.itemsexternalresult && this.state.itemsresult && this.state.itemsresult.length == 0) {
      renderBody = this.state.itemsexternalresult.map(function (item) {
        return (
          <div className={styles.row} dangerouslySetInnerHTML={{ __html: replaceValues(currentProps.fieldBody, item, currentProps.externalColumnDisplayName) }} />
        );
      });
    }
    if (this.state.itemsresult  && this.state.itemsexternalresult && currentProps.columnDisplayName) {
      renderBody = this.state.itemsresult.concat(this.state.itemsexternalresult).map(function (item) {
        return (
          <div className={styles.row} dangerouslySetInnerHTML={{ __html: replaceValues(currentProps.fieldBody, item, currentProps.columnDisplayName.concat(currentProps.externalColumnDisplayName)) }} />
        );
      });
    }
    //#endregion
    //#region Table View
    if (currentProps.fieldShowTitle) {
      if(currentProps.fieldValuesName && !currentProps.externalFieldValuesName)
      {
        header = currentProps.columnDisplayName.map(function (field) {
          if (field.displayName)
          {
            return (
              <th>{field.displayName}</th>
            );
          }
          else
          {
            return (
              <th>{field.key}</th>
            );
          }
        });
      }
      if(currentProps.externalFieldValuesName && !currentProps.fieldValuesName)
      {
        header = currentProps.externalColumnDisplayName.map(function (field) {
          if (field.displayName)
          {
            return (
              <th>{field.displayName}</th>
            );
          }
          else
          {
            return (
              <th>{field.key}</th>
            );
          }
        });
      }
      if(currentProps.fieldValuesName && currentProps.externalFieldValuesName)
      {
        header = currentProps.columnDisplayName.concat(currentProps.externalColumnDisplayName).map(function (field) {
          if (field.displayName)
          {
            return (
              <th>{field.displayName}</th>
            );
          }
          else
          {
            return (
              <th>{field.key}</th>
            );
          }
        });
      }
    }
    else {
      header = null;
    }

    let renderingColumns = (item: any): Array<any> => {
      renderingColumnsResults = [];
      if (this.props.fieldValuesName) {
        this.props.columnDisplayName.map(function (column) {
          let itemValue = item[column.key];
          if (itemValue) 
          {
              if (column.type == "SP.FieldText" || column.type == "SP.FieldMultiLineText") 
              {
                renderingColumnsResults.push(<td><div dangerouslySetInnerHTML={{ __html: itemValue }} /></td>);
              }
              if (column.type == "SP.FieldDateTime") 
              {
                itemValue = moment(itemValue).utc().format("MM/DD/YYYY")
              }
              if (column.type == "SP.FieldUrl") 
              {
                if (itemValue) 
                {
                  let asLink;
                  if (itemValue.Description) {
                    asLink = '<a href="' + itemValue.Url + '" target="_blank">' + itemValue.Description + '</a>'
                  }
                  else {
                    asLink = '<a href="' + itemValue.Url + '" target="_blank">' + itemValue.Url + '</a>'
                  }
                  renderingColumnsResults.push(<td><div dangerouslySetInnerHTML={{ __html: asLink }} /></td>);
                }
              }
              else
              {
                renderingColumnsResults.push(<td>{itemValue} </td>);
              }     
          }
          else renderingColumnsResults.push(<td></td>);
        });
      }
      if (this.props.externalFieldValuesName) {
        let parameters = this.props.variableSourceParameters.filter(x=>x.variableName != "$$undefined$$");
        let segments = this.props.variableSourceSegments.filter(x=>x.variableSegmentName != "$$undefined$$");
        let columnHTML = this.props.columnAsHTML;
        this.props.externalColumnDisplayName.map(function (column) {
          let itemValue = item[column.key];
          if (itemValue)
          {
            if (column.key == "clientname" && columnHTML) // replacement column as HTML
            {
              let result;
              if (parameters.length > 0) {
                parameters.forEach(param => {
                  result = columnHTML.replace(param.variableName, param.value);
                });
              }
              if (segments.length > 0) {
                segments.forEach(segment => {
                  result = columnHTML.replace(segment.variableSegmentName, segment.value);
                });
              }
              if(column.displayName)
              {
                result  = result.replace("$$" + column.displayName + "$$", itemValue);
              }
              else
              {
                result  = result.replace("$$" + column.key + "$$", itemValue);
              }

              renderingColumnsResults.push(<td><div dangerouslySetInnerHTML={{ __html: result }} /></td>);
            }
            else
            {
              if (column.type == "SP.FieldDateTime")
              {
                itemValue = moment(itemValue).utc().format("MM/DD/YYYY")
              }
              renderingColumnsResults.push(<td>{itemValue} </td>);
            }
          }
          else renderingColumnsResults.push(<td></td>);
        });
      }
      return renderingColumnsResults;
    }
    tableViewResults = this.state.itemsresult.concat(this.state.itemsexternalresult).map(function (item) {
      return (
        <tr>
          {renderingColumns(item)}
        </tr>
      );
    });
    //#endregion

    //Rendering
    if (currentProps.layoutSelection) {
      return (
        <div className={styles.spfxCustomListView}>
          {collapsableMenu}
          <style dangerouslySetInnerHTML={{ __html: currentProps.fieldCSS }}>
            </style>
            <span dangerouslySetInnerHTML={{ __html: currentProps.fieldJavascript }}>
            </span>
          <div className={`${this.state.contentContainer} wpContent`}>

            {/*web part tools wrapper*/}
            <div className={styles.wpToolsWrapper} style={{display: `${wpToolsDisplay}`}}>

              {/*pager*/}
              <div className={styles.pagerWrapper}>
                {currentProps.fieldPager && <div>{pagerDDL} of {this.state.totalListItems}</div>}
              </div>
              {/*export to excel*/}
              <div className={styles.exportWrapper}>
                {currentProps.fieldExportToExcel && <div>
                  <CSVLink className={styles.wpTextLink} data={this.state.itemsresult} filename={"htmlView.csv"}
                    onClick={() => { console.log("You Downloaded the Csv"); }}>
                    Export to Excel
                  </CSVLink>
                </div>}
              </div>
              {/*print*/}
              <div className={styles.printWrapper}>
                {currentProps.fieldPrint && <div><button className={styles.mwButton} onClick={print} value="Print">Print</button></div>}
              </div>
            </div>{/*end web part tools wrapper*/}
          </div>
          <div id="divcontent" className={`${this.state.contentContainer} wpContent`}>
            <div className={styles.container}>
              <div className={styles.Header} dangerouslySetInnerHTML={{ __html: currentProps.fieldHeader }} />
            </div>
            <div className={styles.container}>
              {renderBody}
            </div>
            <div className={styles.container}>
              <div className={styles.Footer} dangerouslySetInnerHTML={{ __html: currentProps.fieldFooter }} />
              {currentProps.fieldTotal && <div className={styles.container}>Total: {this.state.totalListItems}</div>}
            </div>
            <iframe id="ifmcontentstoprint" style={{ display: 'none' }}></iframe>
          </div>
        </div >
      );
    }
    else {
      return (
        <div className={styles.spfxCustomListView}>
          {collapsableMenu}
          <div className={`${this.state.contentContainer} wpContent`}>
          {/*web part tools wrapper*/}
          <div className={styles.wpToolsWrapper} style={{display: `${wpToolsDisplay}`}}>
            {/*pager*/}
            <div className={styles.pagerWrapper}>
              {currentProps.fieldPager && <div>{pagerDDL} of {this.state.totalListItems}</div>}
            </div>
            {/*export to excel*/}
            <div className={styles.exportWrapper}>
              {currentProps.fieldExportToExcel &&
                <CSVLink className={styles.wpTextLink} type="button" data={this.state.itemsresult} filename={"tableView.csv"}
                  onClick={() => { console.log("You Downloaded the Csv"); }}>
                  Export to Excel
                </CSVLink>}
            </div>
            {/*print*/}
            <div className={styles.printWrapper}>
              {currentProps.fieldPrint && <button className={styles.mwButton} onClick={print} value="Print">Print</button>}
            </div>
          </div>{/*end web part tools wrapper*/}

          <div id="divcontent">
            <table className={styles.wpTableWrapper}>
              <tr>
                {header}
              </tr>
              {tableViewResults}
            </table>
            {currentProps.fieldTotal && <div>Total: {this.state.totalListItems}</div>}
          </div>
          <iframe id="ifmcontentstoprint" style={{ display: 'none' }}></iframe>
          </div>
        </div>
      );
    }
  }
}
