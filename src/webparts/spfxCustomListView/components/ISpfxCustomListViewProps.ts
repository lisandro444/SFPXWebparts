import { IPropertyPaneDropdownOption, IPropertyPaneField, WebPartContext } from "@microsoft/sp-webpart-base";
import { FilterSegment } from "../../../../lib/webparts/spfxCustomListView/Entities/FilterSegment";
import { FilterURLParameter } from "../Entities/FilterURLParameter";
// import { filterPropertiesValue } from "../SpfxCustomListViewWebPart";
export interface ISpfxCustomListViewProps {
  // Appearence
  description: string;
  title: string;
  titleURL: string;
  titleColor: string;
  enableExpandCollapse: boolean;
  expandCollapseDefaultState: string; 
  // Data Source
  siteName:IPropertyPaneDropdownOption;
  listName: string;
  fieldName:string;
  sortName: string;
  ascending:boolean;
  fieldValuesName:string;
  columnsSelected: Array<string>;
  columnDisplayName:  Array<{key:string, displayName:string, type: string}>;
  fieldValuesSortName:string;


  
  filters: Array<any>;
  context: WebPartContext;

  //BCS Connection
  externalColumnsSelected: Array<string>;
  externalColumnDisplayName: Array<{ key: string, displayName: string, type: string}>;
  externalListName: string;
  externalFieldName: string;
  externalFieldValuesName: string;
  columnAsHTML: string;

  // Variable Source
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
