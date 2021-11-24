import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Web } from '@pnp/sp/webs';
import { ICamlQuery } from '@pnp/sp/lists';
import { IItems, Items, sp } from "@pnp/sp/presets/all";
import "@pnp/sp/search";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";
import { Dictionary, filter } from "lodash";
import { FilterPropertiesValue } from "../Entities/FilterPropertiesValue";
import * as moment from "moment";

export class PnPService {
    private _context;
    private _siteUrl;
    constructor(context: WebPartContext) {
        this._context = context;
        this._siteUrl = context.pageContext.web.absoluteUrl;
    }

    public async getAllSites(): Promise<any[]> {
        var results = [];
        console.log("Site URL: " + this._siteUrl);
        const urlValue = new URL(this._siteUrl);
        let protocol = urlValue.protocol;
        let hostname = urlValue.hostname;

        try {
            await sp.search({ Querytext: "path:" + protocol + "//" + hostname + "(contentclass=sts_site OR contentclass=sts_web)(-WebTemplate:SPSPERS AND -SiteTemplate:APPCATALOG)", SelectProperties: ["Title", "SiteId", "Path", "Url", "SiteName"], RowLimit: 500, TrimDuplicates: false })
                .then((data: SearchResults) => {

                    console.log(data.PrimarySearchResults);
                    results = data.PrimarySearchResults;
                });
            return results;
        }
        catch (error) {
            console.log("error in getAllSites: " + error);

        }
    }
    public async getItemsFiltered(listName): Promise<any[]> {
        try {
            let initialweb = sp.web;
            const caml: ICamlQuery = {
                ViewXml: "<View><Query><Where><IsNotNull><FieldRef Name='Title'/></IsNotNull></Where></Query></View>",
            };
            let items = await initialweb.lists.getByTitle(listName).getItemsByCAMLQuery(caml);
            console.log("CAML Query");
            console.log(items);
            return items;
        }
        catch (error) {
            console.log("Error in getItemsFiltered: " + error);
            return null;
        }
    }

    public async getFieldsFromList(listName, siteUrl): Promise<any[]> {
        try {
            let web = Web(siteUrl);
            let fields = await web.lists.getByTitle(listName).fields.get();;
            console.log(listName + " fields : ");
            console.log(fields);
            return fields;
        }
        catch (error) {
            console.log("Error in getFieldsFromList: " + error);
            return null;
        }
    }

    public async getListsFromSite(siteUrl): Promise<any[]> {
        try {
            let web = Web(siteUrl);
            let lists = await web.lists.get();
            return lists;
        }
        catch (error) {
            console.log("Error in getAllLists(): " + error);
            return null;
        }
    }

    public async getTotalItems(siteUrl, listName): Promise<number> {
        // let initialweb = sp.web;
        let initialweb = Web(siteUrl);
        return (await initialweb.lists.getByTitle(listName).items.get()).length;

    }

    public async getItemsFromSPList(siteUrl, listName, columns, filters: Array<FilterPropertiesValue>, orderBy: string, variableSourceParameters: Array<any>, variableSourceSegments: Array<any>, ascendingParam?: boolean, pager?: Record<'min' | 'max', number>): Promise<any[]> {
        try {
            //let initialweb = sp.web;
            let initialweb = Web(siteUrl);
            let items = initialweb.lists.getByTitle(listName).items;
            let query = "";
            let ascending = false;

            if (filters.length > 0) { //items = items.filter(`ID eq '1' or ID eq '2'`)
                filters.forEach(filter => {
                    if (filter.logicOperator) {
                        query = query.concat(filter.filter + " " + filter.operator + " '" + filter.valueFilter + "'" + filter.logicOperator);
                    }
                    else {
                        query = query.concat(filter.filter + " " + filter.operator + " '" + filter.valueFilter + "'");
                    }
                    //Variable Source Replacement
                    // Parameters
                    variableSourceParameters.forEach(variableSourceParam => {
                        // this condition is for some particular case when the variableName output in variable source is equal to some column in the list
                        if (variableSourceParam.variableName == "$$" + filter.filter + "$$")
                        {
                            query = query.replace(variableSourceParam.variableName,"x" + variableSourceParam.variableName);
                            console.log("You can't use Variable output Name with the same name as internal columns in the list, Please change the name");
                            return; // continue with the next
                        }
                        if (variableSourceParam.variableName == "$$" + filter.valueFilter + "$$") {
                            //filter.valueFilter = variableSource.value;
                            query = query.replace(filter.valueFilter, variableSourceParam.value)
                        }
                    });
                    // Segments
                    variableSourceSegments.forEach(variableSourceSegment => {
                        // this condition is for some particular case when the variableName output in variable source is equal to some column in the list
                        if (variableSourceSegment.variableName == "$$" + filter.filter + "$$") {
                            query = query.replace(variableSourceSegment.variableName, "x" + variableSourceSegment.variableName);
                            console.log("You can't use Variable output Name with the same name as internal columns in the list, Please change the Name of your Variable.");
                            return; // continue with the next
                        }
                        if (variableSourceSegment.variableSegmentName == "$$" + filter.valueFilter + "$$") {
                            //filter.valueFilter = variableSource.value;
                            query = query.replace(filter.valueFilter, variableSourceSegment.value)
                        }
                    });

                    // Replacement for Today compare
                    if(filter.valueFilter == "Today")
                    {
                        var date = new Date();
                        console.log(moment(date).utc().format("MM/DD/YYYY"));
                        query = query.replace(filter.valueFilter, moment(date).utc().format("MM/DD/YYYY"));
                    }
                });

                // if the string end with and/or and remove it
                if (this.confirmEnding(query, " and ")) query = query.slice(0, -5);
                if (this.confirmEnding(query, " or ")) query = query.slice(0, -4);

                items = items.filter(query);
            }
            if (orderBy) {
                if (ascendingParam) ascending = ascendingParam;
                orderBy.split(",").forEach(sort => {
                    items = items.orderBy(sort, ascending);
                });

                items = items.select(columns);
            }
            if (pager.max) {
                items = items.skip(pager.min - 1).top(pager.max - pager.min + 1);
            }
            return items.get();
        }
        catch (error) {
            console.log("Error in getItemsFromList: " + error);
            return null;
        }
    }

    public async getItemsFromBCSExternalList(externalListName: string, columns, filters: Array<FilterPropertiesValue>, variableSourceParameters: Array<any>, variableSourceSegments: Array<any>, pager?: Record<'min' | 'max', number>): Promise<any[]>
    {
        try {
            const urlValue = new URL(this._siteUrl);
            let protocol = urlValue.protocol;
            let hostname = urlValue.hostname;
            let web = Web(protocol + "//" + hostname + "/mwexternalsources/");
            let items;

            if(!pager)
            {
                items = web.lists.getByTitle(externalListName).items.top(20); // put limit if pager is not selected
            }
            else
            {
                items = web.lists.getByTitle(externalListName).items;
            }
            let query = "";
            let ascending = false;

            if (filters.length > 0) { //items = items.filter(`ID eq '1' or ID eq '2'`)
                filters.forEach(filter => {
                    if (filter.logicOperator) {
                      if(query.length > 0){
                        query = query.concat(" " + filter.logicOperator + " " +filter.filter + " " + filter.operator + " '" + filter.valueFilter + "'");
                      } else {
                        query = query.concat(filter.filter + " " + filter.operator + " '" + filter.valueFilter + "'" + filter.logicOperator);
                      }
                    }
                    else {
                        query = query.concat(filter.filter + " " + filter.operator + " '" + filter.valueFilter + "'");
                    }
                    //Variable Source Replacement
                    // Parameters
                    variableSourceParameters.forEach(variableSourceParam => {
                        // this condition is for some particular case when the variableName output in variable source is equal to some column in the list
                        if (variableSourceParam.variableName == "$$" + filter.filter + "$$")
                        {
                            query = query.replace(variableSourceParam.variableName,"x" + variableSourceParam.variableName);
                            console.log("You can't use Variable output Name with the same name as internal columns in the list, Please change the name");
                            return; // continue with the next
                        }
                        if (variableSourceParam.variableName == "$$" + filter.valueFilter + "$$") {
                            //filter.valueFilter = variableSource.value;
                            query = query.replace(filter.valueFilter, variableSourceParam.value)
                        }
                    });
                    // Segments
                    variableSourceSegments.forEach(variableSourceSegment => {
                        // this condition is for some particular case when the variableName output in variable source is equal to some column in the list
                        if (variableSourceSegment.variableName == "$$" + filter.filter + "$$")
                        {
                            query = query.replace(variableSourceSegment.variableName,"x" + variableSourceSegment.variableName);
                            console.log("You can't use Variable output Name with the same name as internal columns in the list, Please change the Name of your Variable.");
                            return; // continue with the next
                        }
                        if (variableSourceSegment.variableSegmentName == "$$" + filter.valueFilter + "$$") {
                            //filter.valueFilter = variableSource.value;
                            query = query.replace(filter.valueFilter, variableSourceSegment.value)
                        }
                    });

                    // Replacement for Today compare
                    if(filter.valueFilter == "Today")
                    {
                        var date = new Date();
                        console.log(moment(date).utc().format("MM/DD/YYYY"));
                        query = query.replace(filter.valueFilter, moment(date).utc().format("MM/DD/YYYY"));
                    }
                });

                // if the string end with and/or and remove it
                if (this.confirmEnding(query, " and ")) query = query.slice(0, -5);
                if (this.confirmEnding(query, " or ")) query = query.slice(0, -4);

                items = items.filter(query);
            }
            // if (orderBy) {
            //     if (ascendingParam) ascending = ascendingParam;
            //     orderBy.split(",").forEach(sort => {
            //         items = items.orderBy(sort, ascending);
            //     });

            //     items = items.select(columns);
            // }
            if (pager.max) {
                items = items.skip(pager.min - 1).top(pager.max - pager.min + 1);
            }
            return items.get();
        }
        catch (error) {
            console.log("Error in getItemsFromBCSExternalList: " + error);
            return null;
        }
    }

    public async getAllExternalLists(): Promise<any[]> {
        try {
            const urlValue = new URL(this._siteUrl);
            let protocol = urlValue.protocol;
            let hostname = urlValue.hostname;
            let web = Web(protocol + "//" + hostname + "/mwexternalsources/");
            let externalLists = await web.lists.filter('HasExternalDataSource eq true').get();
            return externalLists;

        }
        catch (error) {
            console.log("Error in getAllExternalLists(): " + error);
            return null;
        }
    }

    public async getExternalFieldsFromList(listName): Promise<any[]> {
        try {
            const urlValue = new URL(this._siteUrl);
            let protocol = urlValue.protocol;
            let hostname = urlValue.hostname;
            let web = Web(protocol + "//" + hostname + "/mwexternalsources/");
            let fields = await web.lists.getByTitle(listName).fields.get();;
            console.log(listName + " External Fields : ");
            console.log(fields);
            return fields;
        }
        catch (error) {
            console.log("Error in getExternalFieldsFromList: " + error);
            return null;
        }
    }

    public async getWebTitle(url): Promise<string>
    {

        try {
            let web = Web(url);
            let title = await (await web.get()).Title;
            return title;
        }
        catch (error) {
            console.log("Error in getWebTitle(): " + error);
            return null;
        }
    }

    //#region Helpers
  private confirmEnding(string, target) {
    if (string.substr(-target.length) === target) {
      return true;
    } else {
      return false;
    }
  }
    //#endregion
}
