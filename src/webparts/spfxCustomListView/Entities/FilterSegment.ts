import { IPropertyPaneField } from "@microsoft/sp-webpart-base";

export interface FilterSegment {
    segmentVariableNameControl: IPropertyPaneField<any>;
    indexSegmentNameControl: IPropertyPaneField<any>;
    trashControl?:IPropertyPaneField<any>;
  }