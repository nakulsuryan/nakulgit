import * as React from 'react';
import styles from './TabbedUi.module.scss';
import { ITabbedUiProps } from './ITabbedUiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as $ from 'jquery';
import pnp, { Web } from "sp-pnp-js";
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import * as ReactDOM from 'react-dom';
import Accordion from 'react-responsive-accordion';
import * as strings from 'TabbedUiWebPartStrings';
import './TabbedUI.css';
import { SPHttpClient, ISPHttpClientOptions , SPHttpClientConfiguration,SPHttpClientResponse} from '@microsoft/sp-http';
/*import './CoreCustom.css';*/
//import { IDigestCache, DigestCache } from '@microsoft/sp-http';
export interface IReactSpfxState{   
 
  items:[{}],
} 

export default class TabbedUi extends React.Component<ITabbedUiProps, IReactSpfxState> {
 private desc1="";
 private desc2=<div className={styles.hideData} ></div>;
  public constructor(props: ITabbedUiProps, state: IReactSpfxState){
    super();  
   
    this.state = {  
      items: [{}] 
       
   
 }; 
  }  

  
  public setTitle(event){
    this.props.setTitle(event.target.value);
  }
  

  /*Render the UI */
  public render(): React.ReactElement<ITabbedUiProps> {  
   
    var template=<div></div>;
    var layout=false;
    if(this.props.siteUrl!=undefined && this.props.siteUrl!=null && this.props.siteUrl!="")
    if(this.props.layout!=null && this.props.layout!=undefined && this.props.layout!="" &&
    this.props.listName!=null&&this.props.listName!=undefined && this.props.listName!=""&&
    this.props.view!=null && this.props.view!=undefined && this.props.view!=""&&
    this.props.title!=null&&this.props.title!=undefined && this.props.title!=""&&
    this.props.description!=null&&this.props.description!=undefined && this.props.description!="")
    {
        this.desc1="";
        this.desc2= <div className={styles.hideData} ></div>;
        this.getDatafromSharePointList();   
      
    }
    else
    {
        this.desc1="Configure Avanade Tabbed UI App from property pane";
        this.desc2= <div dangerouslySetInnerHTML={{ __html:this.desc1 }}></div>;
    }
var elWidth=(100/this.state.items.length)-1;
var tabWidth_h="";
if(this.props.layout=="Horizontal")
{  
  tabWidth_h=elWidth.toString()+"%";
}
else
{
  tabWidth_h="";
}
   
    /*Condition to render layout*/         
  if(this.props.layout=="Horizontal" && this.state.items.length<7)
  {
    layout=false;
  }
  else
  {
    layout=true;
  }   
    return (
      <div>  
        <div className="webpartheader">
          { this.props.isEditMode && <textarea onChange={this.setTitle.bind(this)} className="edit" placeholder={strings.TitlePlaceholder} aria-label="Add a title" defaultValue={this.props.titleHeader}></textarea> }
          { !this.props.isEditMode && <span className="view">{this.props.titleHeader}</span> }          
        </div>
      {this.desc2}   
      {window.innerWidth<600
      ?
      <div>
        <Accordion>{this.state.items.map(function(item,key){       
          var title=<div className="contentWrapper tabTableCustom react-tabs__tab"><img className="titleIcon" src={item[this.props.titleImageUrl]}  alt=""></img> <div>{item[this.props.title]}</div></div>;
         var titleActive=<div className={"contentWrapper layoutStyleMobile tabTableCustom react-tabs__tab--selected"}><img className="titleIcon" src={item[this.props.titleImageUrl]}  alt=""></img> <div>{item[this.props.title]}</div></div>;
            return (
              <div className="contentWrapper"  data-trigger={title} data-trigger-when-open={titleActive}>
              <div className="contentWrapper" dangerouslySetInnerHTML={{ __html: item[this.props.description] }}></div></div>
             );},this)} 
            </Accordion>
              </div>  
                 :
      (layout?

      /* Render Vertical layout template*/ 
      <Tabs>
        <table className="tabTableCustom">
        <tr>
          <td className="verticalTabPanel">
            <div className="oops"><TabList>{this.state.items.map(function(item,key){ 
             if(this.props.titleImageUrl!=undefined && this.props.titleImageUrl!=null && this.props.titleImageUrl!="")
               var titleIcon=<img className="titleIcon" src={item[this.props.titleImageUrl]} alt=""></img>;
               else
               titleIcon=null;
                return (<Tab>
                          <div className="contentWrapper verticalTab">{titleIcon}<div>{item[this.props.title]}</div></div> 
                        </Tab>);
                        },this)}
                </TabList></div>
          </td>
          <td className="contentWrapper verticalTabDescription" style={{border:"1px solid #aaa",overflow:"hidden"}}>
            <div>{this.state.items.map(function(item,key){       
                return (  
               <TabPanel>
                 <div className="contentWrapper" dangerouslySetInnerHTML={{ __html: item[this.props.description] }}></div>
              </TabPanel>);},this)}
              </div>
          </td>
        </tr>
        </table>
      </Tabs>
        :
           /* Render Horizontal layout template*/ 
      <div>
      <Tabs>
        <TabList>{this.state.items.map(function(item,key){ 
         if(this.props.titleImageUrl!=undefined && this.props.titleImageUrl!=null && this.props.titleImageUrl!="")
         var titleIcon=<img className="titleIcon" src={item[this.props.titleImageUrl]} alt=""></img>;
         else
         titleIcon=null;
            return (<Tab style={{width:tabWidth_h}}>
            <div className="contentWrapper">{titleIcon}<div>{item[this.props.title]}</div></div>
          </Tab>);},this)}
        </TabList>
        <div className="verticalTabDescription">{this.state.items.map(function(item,key){      
      return (  
          <TabPanel>
          <div className="contentWrapper" dangerouslySetInnerHTML={{ __html: item[this.props.description] }}></div>
        </TabPanel>);},this)}</div>
        </Tabs>
        </div>
       ) }   
      </div>
    );
  }
  
public componentWillMount(){ 
 
}
 

  private getDatafromSharePointList()
  {  
    var reactHandler = this;
    var camlquery="<View><Query>" +this.props.view + "</Query></View>";
    const web = new Web(this.props.siteUrl);
    web.lists.getByTitle(this.props.listName).getItemsByCAMLQuery({'ViewXml':camlquery}).then((res) => {
            reactHandler.setState
            ({
              items:res
              
            });
        }).catch((error)=>
        {
          console.log("Error in getting data from list " + error);
        });       
  }
  
}

