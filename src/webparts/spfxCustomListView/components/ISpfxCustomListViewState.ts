export interface ISpfxCustomListViewState {
    url:string;
    // Appearence
    description: string;
    titleColor: string;
    iconToggle: string;
    cssTitleColor: string;    
    cssToggleWrapperDisplay: string;
    uniqueToggleID: string;
    contentContainer: string;

    // Data Source
    listName: string;
    fieldName:string;
    fieldValuesName:string;
    fieldValuesSortName:string;
    itemsresult: Array<any>;
    itemsexternalresult: Array<any>;
    filters: Array<any>;
    totalListItems: number;
    // context: WebPartContext;
    // Layout
    layoutSelection: boolean;
    fieldHeader: string;
    fieldBody: string;
    fieldFooter: string;
    fieldCSS: string;
    fieldJavascript: string;
    fieldShowTitle: string;
}