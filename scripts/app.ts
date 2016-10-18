namespace SupportedActions {
    export const OpenItems = "OpenItems";
    export const OpenQuery = "OpenQuery";
}

namespace WellKnownQueries {
    export const AssignedToMe = "A2108D31-086C-4FB0-AFDA-097E4CC46DF4";
    export const UnsavedWorkItems = "B7A26A56-EA87-4C97-A504-3F028808BB9F";
    export const FollowedWorkItems = "202230E0-821E-401D-96D1-24A7202330D0";
    export const CreatedBy = "53FB153F-C52C-42F1-90B6-CA17FC3561A8";
    export const SearchResults = "2CBF5136-1AE5-4948-B59A-36F526D9AC73";
    export const CustomWiql = "08E20883-D56C-4461-88EB-CE77C0C7936D";
    export const RecycleBin = "2650C586-0DE4-4156-BA0E-14BCFB664CCA";
}

export var queryExclusionList = [
    WellKnownQueries.AssignedToMe,
    WellKnownQueries.UnsavedWorkItems,
    WellKnownQueries.FollowedWorkItems,
    WellKnownQueries.CreatedBy,
    WellKnownQueries.SearchResults,
    WellKnownQueries.CustomWiql,
    WellKnownQueries.RecycleBin];

export function isSupportedQueryId(queryId: string) {
    return queryId && queryExclusionList.indexOf(queryId.toUpperCase()) === -1;
}

export function generateUrl(action: string, collection: string, project: string, qid?: string, wids?: number[], columns?: string[]): string {
    let url = `tfs://ExcelRequirements/${action}?cn=${collection}&proj=${project}`;

    if (action === SupportedActions.OpenItems) {
        if (!wids) {
            throw new Error(`'wids' must be provided for '${SupportedActions.OpenItems}' action.`);
        }
        url += `&wid=${wids}`;

        if (columns && columns.length > 0) {
            url += `&columns=${columns}`;
        }
    }
    else if (action === SupportedActions.OpenQuery) {
        if (!qid) {
            throw new Error(`'qid' must be provided for '${SupportedActions.OpenQuery}' action.`);
        }
        url += `&qid=${qid}`;
    }
    else {
        throw new Error(`Unsupported action provided: ${action}`);
    }

    if (url.length > 2000) {
        throw new Error('Generated url is exceeds the maxlength, please reduce the number of work items you selected.');
    }

    return url;
}

export interface IQueryObject {
    id: string;
    isPublic: boolean;
    name: string;
    path: string;
    wiql: string;
}

export interface IActionContext {
    id?: number
    query?: IQueryObject;
    queryText?: string;
    ids?: number[];
    workItemIds?: number[];
    columns?: string[];
}

export var openQueryAction = {
    getMenuItems: (context: any) => {
        if (!context || !context.query || !context.query.wiql || !isSupportedQueryId(context.query.id)) {
            return null;
        }
        else {
            return [<IContributedMenuItem>{
                title: "Open in Excel",
                text: "Open in Excel",
                icon: "img/miniexcellogo.png",
                action: (actionContext: IActionContext) => {
                    if (actionContext && actionContext.query && actionContext.query.id) {
                        let qid = actionContext.query.id;
                        let context = VSS.getWebContext();
                        let collectionUri = context.collection.uri;
                        let projectName = context.project.name;

                        window.location.href = generateUrl(SupportedActions.OpenQuery, collectionUri, projectName, qid);
                    }
                }
            }];
        }
    }
};

export var openWorkItemsAction = {
    getMenuItems: (context: any) => {
        return [<IContributedMenuItem>{
            title: "Open in Excel",
            text: "Open in Excel",
            icon: "img/miniexcellogo.png",
            action: (actionContext: IActionContext) => {
                let wids = actionContext.ids || actionContext.workItemIds || (actionContext.id > 0 ? [actionContext.id] : null);
                let columns = actionContext.columns;
                let context = VSS.getWebContext();
                let collectionUri = context.collection.uri;
                let projectName = context.project.name;

                window.location.href = generateUrl(SupportedActions.OpenItems, collectionUri, projectName, null, wids, columns);
            }
        }];
    }
};

export var openQueryOnToolbarAction = {
    getMenuItems: (context: any) => {
        return [<IContributedMenuItem>{
            title: "Open in Excel",
            text: "Open in Excel",
            icon: "img/miniexcellogo.png",
            showText: true,
            action: (actionContext: IActionContext) => {
                if (actionContext && actionContext.query && actionContext.query.wiql && isSupportedQueryId(actionContext.query.id)) {
                    let qid = actionContext.query.id;
                    let context = VSS.getWebContext();
                    let collectionUri = context.collection.uri;
                    let projectName = context.project.name;

                    window.location.href = generateUrl(SupportedActions.OpenQuery, collectionUri, projectName, qid);
                }
                else {
                    alert("This operation is not supported to this query. This extension supports queries saved in My Queries and Shared Queries.");
                }
            }
        }];
    }
};
