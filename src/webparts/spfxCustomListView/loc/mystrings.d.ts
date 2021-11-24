declare interface ISpfxCustomListViewWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  DataSourceGroupName:string;
  BCSConnectionGroupName: string;
  VariableSourceGroupName:string;
  DisplayOptionsGroupName: string;
  LayoutGroupName: string;
  TitleFieldLabel: string;
  TitleURLFieldLabel: string;
  TitleColorFieldLabel: string;  
  WebPartDefaultStateLabel: string;
  EnableExpandCollapse: string;

  // Data Source
  SitesFieldLabel: string;
  ListsFieldLabel: string;
  ColumnsFieldLabel: string;
  ColumnsFieldValueLabel: string;
  SortbyFieldLabel: string;
  AscendingDescending: string;
  GroupbyFieldLabel: string;
  FiltersFieldLabel: string;
  FiltersOperatorFieldLabel: string;
  FiltersValueFieldLabel: string;
  LogicOperatorFieldLabel: string;

  // BCS Connection
  ExternalListsFieldLabel: string;
  ExternalColumnsFieldLabel: string;
  DisplayNameFieldLabel: string;
  ColumnHTMLabel: string;


  // Variable Source
  fieldQueryUrlParameter: string;
  fieldQueryUrlVariableName: string;
  fieldQueryUrlName: string;
  fieldQueryUrlSegment: string;
  fieldQueryUrlSegmentName: string;
  fieldSegmentIndex: string;

  // Display Options
  PagerFieldLabel: string;
  PageLimitFieldLabel: string;
  ExportToExcelFieldLabel: string;
  PrintFieldLabel: string;
  TotalFieldLabel: string;

  //Layout
  LayoutSelectionFieldLabel: string;
  Header: string;
  Body: string;
  Footer: string;
  CSS: string;
  JavaScript: string;
  ShowTitle: string;
}

declare module 'SpfxCustomListViewWebPartStrings' {
  const strings: ISpfxCustomListViewWebPartStrings;
  export = strings;
}
