import { IPropertyPaneField } from "@microsoft/sp-webpart-base";

export interface FilterURLParameter {
    variableNameControl: IPropertyPaneField<any>;
    parameterNameControl: IPropertyPaneField<any>;
    trashControl?:IPropertyPaneField<any>;
  }