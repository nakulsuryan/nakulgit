import { Text } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';



export class ListService {

    private spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this.spHttpClient = spHttpClient;
    }

    public getListsFromWeb(webUrl: string): Promise<Array<{url: string, title: string}>> {
        return new Promise<Array<{url: string, title: string}>>((resolve, reject) => {
            const endpoint = Text.format("{0}/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl&$filter=(IsPrivate eq false) and (IsCatalog eq false) and (Hidden eq false)&$expand=RootFolder", webUrl);
            this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    response.json().then((data: any) => {
                        const listTitles: Array<{url: string, title: string}> = data.value.map((list) => {
                                return {url: list.RootFolder.ServerRelativeUrl, title: list.Title};
                            });
                        resolve( listTitles.sort( (a, b) => a.title.localeCompare(b.title)) );
                    })
                    .catch((error) => { reject(error); });
                } else {
                    reject(response);
                }
            })
            .catch((error) => { reject(error); });
        });
    }
    ///_api/web/lists/getbytitle('<list title>')/fields?$filter=Hidden eq false and ReadOnlyField eq false
    public getColumnsFromList(webUrl : string, ListName : string): Promise<Array<{name: string, title: string}>> {
        return new Promise<Array<{name: string, title: string}>>((resolve, reject) => {
            const endpoint = Text.format("{0}/_api/web/lists/getbytitle('"+ListName+"')/fields?$select=Title,InternalName&$filter=ReadOnlyField%20eq%20false", webUrl);
            this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    response.json().then((data: any) => {
                        const columnTitles: Array<{name: string, title: string}> = data.value.map((field) => {
                                return {name: field.InternalName, title: field.Title};
                            });
                        resolve( columnTitles.sort( (a, b) => a.title.localeCompare(b.title)) );
                    })
                    .catch((error) => { reject(error); });
                } else {
                    reject(response);
                }
            })
            .catch((error) => { reject(error); });
        });
    }

    public getAllViewsOfList(webUrl : string, ListName : string): Promise<Array<{query: string, title: string}>> {
        return new Promise<Array<{query: string, title: string}>>((resolve, reject) => {
            const endpoint = Text.format("{0}/_api/web/lists/getbytitle('"+ListName+"')/views?$select=Title,ViewQuery", webUrl);
            this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    response.json().then((data: any) => {
                        const viewTitles: Array<{query: string, title: string}> = data.value.map((view) => {
                                return {query: view.ViewQuery, title: view.Title};
                            });
                        resolve( viewTitles.sort( (a, b) => a.title.localeCompare(b.title)) );
                    })
                    .catch((error) => { reject(error); });
                } else {
                    reject(response);
                }
            })
            .catch((error) => { reject(error); });
        });
    }

}
