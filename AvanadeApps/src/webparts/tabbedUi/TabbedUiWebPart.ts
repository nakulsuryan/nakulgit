import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType,DisplayMode  } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneButton
} from '@microsoft/sp-webpart-base';

import * as strings from 'TabbedUiWebPartStrings';
import TabbedUi from './components/TabbedUi';
import { ITabbedUiProps } from './components/ITabbedUiProps';
import { ListService } from '../../common/services/ListService';
import { SPHttpClient, ISPHttpClientOptions , SPHttpClientConfiguration,SPHttpClientResponse} from '@microsoft/sp-http';
export interface ITabbedUiWebPartProps {
  layout:string;
  titleHeader: string;
  listName: string;
  columnName:string;
  title:string;
  description:string;
  titleImageUrl:string;
  view:string;

}

  import { update, get } from '@microsoft/sp-lodash-subset';
  import { ControlMode } from '../../common/datatypes/ControlMode';
  import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';  
 

export default class TabbedUiWebPart extends BaseClientSideWebPart<ITabbedUiWebPartProps> {
  private dropDownOptions: IPropertyPaneDropdownOption[] =[]; 
  private lists: IPropertyPaneDropdownOption[];
  private columns: IPropertyPaneDropdownOption[];
  private views: IPropertyPaneDropdownOption[];
  private listService: ListService;
  private cachedLists = null;
  private cachedColumns = null;
  private cachedViews = null;
  protected onInit(): Promise<void> {
    return super.onInit().then( _ => {
      this.listService = new ListService(this.context.spHttpClient);
    });
  }


  private listsDropdownDisabled: boolean = true;
  private columnsDropdownDisabled: boolean = true;
  private buttonDisabled: boolean = false;
  public render(): void {
    const props = this.properties;
    const element: React.ReactElement<ITabbedUiProps > = React.createElement(
      TabbedUi,
      {
        siteUrl:this.context.pageContext.web.absoluteUrl,        
        view: this.properties.view,
        listName:this.properties.listName,
        columnName: this.properties.columnName,
        title:this.properties.title,//"Title",
        titleImageUrl: this.properties.titleImageUrl,
        description:this.properties.description,//"TabDescription",
        tabPublish:'TabPublish',
        layout:this.properties.layout,
        isEditMode: this.displayMode==DisplayMode.Edit,
        titleHeader:this.properties.titleHeader,
        setTitle: function(title:string){
        props.titleHeader=title;

        
        }

      }
    );

    ReactDom.render(element, this.domElement);
  }
 
  protected createList():void
  {
        if(this.lists.filter(x => x.text == 'TabbedUIList').length>0)
        {
          confirm("TabbedUIList is already present here: \n"+this.context.pageContext.web.absoluteUrl.toString()+"/Lists/TabbedUIList");
          this.buttonDisabled=true;
        }
        else
        { 
            confirm("TabbedUIList creation in progress, will notify you once it gets completed");
            this.buttonDisabled=true;
            var newUrl=this.context.pageContext.web.absoluteUrl.toString()+"/_api/web/lists";
            const body: string = JSON.stringify({ '__metadata': { 'type': 'SP.List' }, 'AllowContentTypes': true,
            'BaseTemplate': 100, 'ContentTypesEnabled': true, 'Description': 'enter values for TabbedUI', 'Title': 'TabbedUIList' });
            const opt: ISPHttpClientOptions = { headers: { 'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''}, body: body };
            this.context.spHttpClient.post(newUrl, SPHttpClient.configurations.v1, opt).then((response: SPHttpClientResponse) => {
              if(response.ok) {
                this.createRichTextColumn("TabDecription");        
              }
                })
                .catch((error) => { 
                    console.log('Error in creating list');
                    console.log(error); 
                }); 
              
          } 
    }
     
  
          protected createColumn(columnN,columnType):void
          {

            var newUrl=this.context.pageContext.web.absoluteUrl.toString()+"/_api/web/lists/getbytitle('TabbedUIList')/Fields";
            const body: string = JSON.stringify({ '__metadata': { 'type': 'SP.Field' }, 'Title':  columnN ,'FieldTypeKind': columnType});
            const opt: ISPHttpClientOptions = { headers: { 'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''}, body: body };
            this.context.spHttpClient.post(newUrl, SPHttpClient.configurations.v1, opt).then((response: SPHttpClientResponse) => {
              if(response.ok) {
                interface columnsUI{
                  columnTitle: string,
                  fieldType: number
                }
                var columnsAll: columnsUI[];
                var columnsAll=[
                  
                            {columnTitle:"TitleImageUrl",fieldType:2},
                            {columnTitle:"TabOrder",fieldType:1},
                            {columnTitle:"TabPublish",fieldType:8},
                            {columnTitle:"TabGroupName",fieldType:2}
                          ];
                for(var i=0;i<columnsAll.length;i++){
                  if(columnsAll[i].columnTitle==columnN && (i+1<columnsAll.length)){
                  
                        this.createColumn(columnsAll[i+1].columnTitle,columnsAll[i+1].fieldType);
                  }
                  
                }
                
                if(columnsAll[columnsAll.length-1].columnTitle==columnN){
                  this.addToDefaultView('TitleImageUrl');
                }
              }
                })
                .catch((error) => { 
                    console.log('Error in creating column');
                    console.log(error); 
                }); 
                  
                  
          }
        protected createRichTextColumn(columnN):void
        {   
         
                  var newUrl=this.context.pageContext.web.absoluteUrl.toString()+"/_api/web/lists/getbytitle('TabbedUIList')/Fields";
                  const body: string = JSON.stringify({ '__metadata': { 'type': 'SP.FieldMultiLineText' }, 'Title':'TabDecription', 'FieldTypeKind': 3,'SchemaXml':'<Field DisplayName=\"TabDecription\" Type=\"Note\" Required=\"FALSE\"  RichText=\"TRUE\" NumLines=\"6\" RestrictedMode=\"TRUE\" RichTextMode=\"FullHtml\" AppendOnly=\"FALSE\" />' });
                  const opt: ISPHttpClientOptions = { headers: { 'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=verbose',
                  'odata-version': ''}, body: body };
                  this.context.spHttpClient.post(newUrl, SPHttpClient.configurations.v1, opt) .then((response: SPHttpClientResponse) => 
                  {
                    if(response.ok) 
                    {
                      var newUrl=this.context.pageContext.web.absoluteUrl.toString()+"/_api/web/lists/getbytitle('TabbedUIList')/views/getbytitle('All Items')/ViewFields/AddViewField('"+columnN+"')";
                      const body: string = JSON.stringify({  'strField':columnN }
                    );
                      const opt: ISPHttpClientOptions = { headers: { 'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=verbose',
                      'odata-version': ''}, body: body };
                      this.context.spHttpClient.post(newUrl, SPHttpClient.configurations.v1, opt)
                      .then((response: SPHttpClientResponse) => 
                      {
                        if(response.ok) 
                        this.createColumn("TitleImageUrl",2);
                
                      })
                    }             
                  }).catch((error)=>
                  {
                    console.log("Error in creating rich text column");
                    console.log(error);
                  });
         

        }

           
    protected addToDefaultView(columnN):void
            {
              var newUrl=this.context.pageContext.web.absoluteUrl.toString()+"/_api/web/lists/getbytitle('TabbedUIList')/views/getbytitle('All Items')/ViewFields/AddViewField('"+columnN+"')";
              const body: string = JSON.stringify({  'strField':columnN });
              const opt: ISPHttpClientOptions = { headers: { 'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''}, body: body };
              this.context.spHttpClient.post(newUrl, SPHttpClient.configurations.v1, opt).then((response: SPHttpClientResponse) => {
                if(response.ok) 
                {
                  interface columnsUI{
                    columnTitle: string
                  }
                  var columnsAll: columnsUI[];
                  var columnsAll=[
                              
                              {columnTitle:"TitleImageUrl"},
                              {columnTitle:"TabOrder"},
                              {columnTitle:"TabPublish"},
                              {columnTitle:"TabGroupName"}
                            ];
                  for(var i=0;i<columnsAll.length;i++)
                  {
                    if(columnsAll[i].columnTitle==columnN && (i+1<columnsAll.length))
                    {
                      this.addToDefaultView(columnsAll[i+1].columnTitle);
                    } 
                  }
                  if((columnsAll[columnsAll.length-1].columnTitle==columnN))
                  {
                    confirm("TabbedUIList is created here: \n"+this.context.pageContext.web.absoluteUrl.toString()+"/Lists/TabbedUIList");
                    this.cachedLists=null;
                    this.loadLists()
                    .then((listOptions: IPropertyPaneDropdownOption[]):  Promise<IPropertyPaneDropdownOption[]> =>  {
                      this.lists = listOptions;
                      this.listsDropdownDisabled = false;
                      this.context.propertyPane.refresh();
                    this.buttonDisabled=true;
                    return this.loadColumns(this.properties.listName);
                  } );
                  } 
                }
                  })
                  .catch((error) => { 
                      console.log("Error in adding all created column to default view");
                      console.log(error); 
                  }); 
                    
                    
            }
  
    
              
  
 private loadLists(): Promise<IDropdownOption[]> {
  return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
    if (Environment.type === EnvironmentType.Local) {
      resolve( [{
        key: 'TabbedUIList',
        text: 'TabbedUIList'
      },
      {
        key: 'ITSFAQList',
        text: 'ITSFAQList'
        }] );
    } else if (Environment.type === EnvironmentType.SharePoint ||
              Environment.type === EnvironmentType.ClassicSharePoint) {
      try {
        if (!this.cachedLists) {
          return this.listService.getListsFromWeb(this.context.pageContext.web.absoluteUrl)
            .then( (lists) => {
              this.cachedLists = lists.map( (l) => ({ key: l.title, text: l.title } as IDropdownOption) );
              resolve( this.cachedLists );
            } );
        } else {
          // using cached lists if available to avoid loading spinner every time property pane is refreshed
          return resolve( this.cachedLists );
        }
      } 
      catch (error) 
      {
        console.log(error);
      }
    }
  });
}
private loadViews(ListName : string): Promise<IDropdownOption[]> {
  return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
    if (Environment.type === EnvironmentType.Local) {
        const views = {
          TabbedUIList: [
            {
              key: 'View1',
              text: 'View1'
            },
            {
              key: 'View2',
              text: 'View2'
            }
          ],
          ITSFAQList: [
            {
              key: 'View3',
              text: 'View3'
            },
            {
              key: 'View4',
              text: 'View4'
            }
          ]
        };
        resolve(views[this.properties.listName]);

    } else if (Environment.type === EnvironmentType.SharePoint ||
              Environment.type === EnvironmentType.ClassicSharePoint) {
      try {
        if (!this.cachedViews) {
          return this.listService.getAllViewsOfList(this.context.pageContext.web.absoluteUrl,ListName)
            .then( (views) => {
              this.cachedViews = views.map( (v) => ({ key: v.query, text: v.title } as IDropdownOption) );
              resolve( this.cachedViews );
            } );
        } else {
          // using cached views if available to avoid loading spinner every time property pane is refreshed
          return resolve( this.cachedViews);
        }
      }
      catch (error) 
      {
        console.log(error);
      }
    }
  });
}
private loadColumns(ListName : string): Promise<IDropdownOption[]> {
  return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
    if (Environment.type === EnvironmentType.Local) {
        const columns = {
          TabbedUIList: [
            {
              key: 'Title11',
              text: 'Title11'
            },
            {
              key: 'Title12',
              text: 'Title12'
            }
          ],
          ITSFAQList: [
            {
              key: 'Title13',
              text: 'Title13'
            },
            {
              key: 'Title114',
              text: 'Title14'
            }
          ]
        };
        resolve(columns[this.properties.listName]);

    } else if (Environment.type === EnvironmentType.SharePoint ||
              Environment.type === EnvironmentType.ClassicSharePoint) {
      try {
        if (!this.cachedColumns) {
          return this.listService.getColumnsFromList(this.context.pageContext.web.absoluteUrl,ListName)
            .then( (columns) => {
              this.cachedColumns = columns.map( (c) => ({ key: c.name, text: c.title } as IDropdownOption) );
              resolve( this.cachedColumns );
            } );
        } else {
          // using cached columns if available to avoid loading spinner every time property pane is refreshed
          return resolve( this.cachedColumns );
        }
      } 
      catch (error) 
      {
        console.log(error);
      }
    }
  });
}
    

    protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.listName!=null&&this.properties.listName!=""&&this.properties.listName!=undefined)
      {
        this.listsDropdownDisabled=false;
        this.columnsDropdownDisabled=false
      }
      else
      {
        this.listsDropdownDisabled = !this.lists;
      this.columnsDropdownDisabled = !this.properties.listName || !this.columns;
      }
      if (this.lists) {
        return;
        
      }
   
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
   
      this.loadLists()
        .then((listOptions: IPropertyPaneDropdownOption[]):  Promise<IPropertyPaneDropdownOption[]> =>  {
          this.lists = listOptions;
          this.listsDropdownDisabled = false;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
         return this.loadColumns(this.properties.listName);
        })
         .then((columnsOptions: IPropertyPaneDropdownOption[]): void => {
          this.columns = columnsOptions;
          this.columnsDropdownDisabled = !this.properties.listName;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.loadViews(this.properties.listName)
          .then((viewsOptions: IPropertyPaneDropdownOption[]): void => {
            this.views = viewsOptions;
            this.context.propertyPane.refresh();
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        });
          this.render();
        });

  }

      
      protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
        if (propertyPath === 'listName' &&
            newValue) {
          // disable item selector until new items are loaded
          this.columnsDropdownDisabled = true;
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
          // communicate loading items
          this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'columns');
          this.cachedColumns="";
          this.columns=null;
          this.views=null;
          this.cachedViews="";
          this.loadViews(this.properties.listName)
            .then((viewOptions: IPropertyPaneDropdownOption[]): void => { 
              this.views=viewOptions;
              this.properties.view="";
              this.context.propertyPane.refresh();
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            });
       
        
          this.loadColumns(this.properties.listName)
            .then((columnOptions: IPropertyPaneDropdownOption[]): void => {            
              this.columns = columnOptions;      
              this.columnsDropdownDisabled = false;
              this.properties.title="";
              this.properties.titleImageUrl="";
              this.properties.description="";
              this.properties.columnName="";
              this.context.statusRenderer.clearLoadingIndicator(this.domElement);
              // re-render the web part as clearing the loading indicator removes the web part body
              this.render();
              this.context.propertyPane.refresh();
            });
        }
        else {
          super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        }
      }
  
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
 
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {   
    this.dropDownOptions=[{key:"Horizontal",text:"Horizontal"},{key:"Vertical",text:"Vertical"}];
    return {
      pages: [
        {         
          groups: [
            {     
              groupFields: [
                
                PropertyPaneButton(' ', {
                text: 'Use the default list',
                onClick: this.createList.bind(this),
                disabled: this.buttonDisabled
                }), 
               
                PropertyPaneDropdown('listName', {
                  label: "Select existing List (*)",
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('view', {
                  label: "Tab List View (*)",
                  options: this.views,
                  disabled: this.columnsDropdownDisabled            
                }),        
                PropertyPaneDropdown('layout', {
                  label: "Layout (*)",
                  options: this.dropDownOptions,
                  disabled:false         
                }),
                PropertyPaneDropdown('title', {
                  label: "Tab Headers (*)",
                  options: this.columns,
                  disabled: this.columnsDropdownDisabled
                }),
                PropertyPaneDropdown('description', {
                  label: "Tab Descriptions (*)",
                  options: this.columns,
                  disabled: this.columnsDropdownDisabled
                }),
                PropertyPaneDropdown('titleImageUrl', {
                  label: "Tab Header Icon",
                  options: this.columns,
                  disabled: this.columnsDropdownDisabled
                })
                ]
            }
          ]
        }
      ]
    };
  }
}