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
    id?: number;            // From card
    workItemId?: number;    // From work item form
    query?: IQueryObject;
    queryText?: string;
    ids?: number[];
    workItemIds?: number[]; // From backlog/iteration (context menu) and query results (toolbar and context menu)
    columns?: string[];
}

export var openQueryAction = {
    execute: (actionContext: IActionContext) => {
        if (actionContext && actionContext.query && actionContext.query.id) {
            const qid = actionContext.query.id;
            const context = VSS.getWebContext();
            const collectionUri = context.collection.uri;
            const projectName = context.project.name;

            const url = generateUrl(SupportedActions.OpenQuery, collectionUri, projectName, qid);
            openUrl(url);
        }
    }
};

export var openWorkItemsAction = {
    execute: (actionContext: IActionContext) => {
        const wids = actionContext.ids ||
            actionContext.workItemIds ||
            (actionContext.workItemId > 0 ? [actionContext.workItemId] : null) ||
            (actionContext.id > 0 ? [actionContext.id] : null);
        const columns = actionContext.columns;
        const context = VSS.getWebContext();
        const collectionUri = context.collection.uri;
        const projectName = context.project.name;

        const url = generateUrl(SupportedActions.OpenItems, collectionUri, projectName, null, wids, columns);
        openUrl(url);
    }
};

export var openQueryOnToolbarAction = {
    execute: (actionContext: IActionContext) => {
        if (actionContext && actionContext.query && actionContext.query.wiql && isSupportedQueryId(actionContext.query.id)) {
            const qid = actionContext.query.id;
            const context = VSS.getWebContext();
            const collectionUri = context.collection.uri;
            const projectName = context.project.name;

            const url = generateUrl(SupportedActions.OpenQuery, collectionUri, projectName, qid);
            openUrl(url);
        }
        else {
            alert("Unable to perform operation. To use this extension, queries must be saved in My Queries or Shared Queries.");
        }
    }
};

function openUrl(url: string) {
    showNotification();
    VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: any) => {
        navigationService.navigate(url);
    });
}

function showNotification() {
    const extensionContext = VSS.getExtensionContext();
    let dialog: IExternalDialog;

    VSS.getService(VSS.ServiceIds.Dialog).then((hostDialogService: IHostDialogService) => {
        hostDialogService.openDialog(`${extensionContext.publisherId}.${extensionContext.extensionId}.notificationDialog`,
            {
                title: "We're opening this in Microsoft Excel...",
                width: 470,
                height: 250,
                modal: true,
                draggable: false,
                resizable: false,
                buttons: {
                    "ok": {
                        id: "ok",
                        text: "Dismiss",
                        click: () => {
                            dialog.close();
                        },
                        class: "cta",
                    }
                }
            },
            {
                close: () => dialog.close(),
            })
            .then(d => {
                dialog = d;
            });
    });
}
