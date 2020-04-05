import * as React from 'react';
import styles from './Roadmap.module.scss';
import { IRoadmapProps } from './IRoadmapProps';
import { escape } from '@microsoft/sp-lodash-subset';

import  './Roadmap.css';
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import Collapsible from 'react-collapsible';
import * as $ from 'jquery'
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
export interface IReactSpfxState{   
  items:[{}],
  categories:[{}]
}

export default class Roadmap extends React.Component<IRoadmapProps, IReactSpfxState> {
  
  public constructor(props: IRoadmapProps, state: IReactSpfxState){
    super();  
    this.state = {  
      items: [{}],
      categories:[{}] 
      
   
 }; 
  } 

  private getDatafromSharePointList(tabParams)
  {
      
    
    // var queryUrl = tabParams.Url +"/_api/Web/Lists/GetByTitle('" + tabParams.ListName + "')/items?" +
    // "$select=Title,Category,Description,ServiceLine,ReleaseDate,Contact,RoadmapURL&$orderby=ReleaseDate desc&$Filter=ServiceLine eq '"+ escape(this.props.ServiceLine) +"'";

    var queryUrl = tabParams.Url +"/_api/Web/Lists/GetByTitle('" + tabParams.ListName + "')/items?" +
    "$select=Title,Stage,Description,Service,ReleaseDate,ReferenceURL,Contact&$orderby=ReleaseDate desc&$Filter=Service eq '"+ escape(this.props.ServiceLine) +"'";

    var me = this;


    var a = $.ajax({  
      url:  queryUrl,
      type: "GET",  
          headers:{'Accept': 'application/json; odata=verbose;'},  
          success: function(resultData) {      
            me.setState({
            items:resultData.d.results
                      });    
           },  
          error : function(jqXHR, textStatus, errorThrown) {  
           
          }  
      }); 
  }
 private getReleaseDate(releaseDate) {
    try {
        var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        var date = new Date(releaseDate);
        releaseDate = monthNames[date.getMonth()] + ', ' + date.getFullYear();
        return releaseDate;
    }
    catch (e) {
        alert("Error occured while fetching data " + e.message);
    }

}

private checkOpen(img,item,obj){

  var srcUp=  String(require("../Images/avaicon_up.jpg"));

  //img=<div className="question"><table><tr><td className="alignImage"><img src={srcUp} /></td><td>{item[obj.props.title]}</td></tr></table></div>; 
  //img=<div className="question" ><table><tr><td className="alignImage"><img src={srcUp} /></td><td>{item["Title"]}</td></tr></table></div>;        
  img = <div className="question"><img src={srcUp}/>{item["Title"]}</div>
  
  return img;

}

  public render(): React.ReactElement<IRoadmapProps> {

    
              

    var tabparameters ={};
      
    var srcDown=  String(require("../Images/avaicon_down.jpg")); 

    //Get list name from property
    if(escape(this.props.ListName)!="")
    {
      tabparameters['ListName'] = escape(this.props.ListName);
    }
    else
    {
      tabparameters['ListName'] = "ITS Roadmap"
    }

    //Get site url of the list from property
    if(escape(this.props.SiteUrl)!="")
    {
      tabparameters['Url'] = escape(this.props.SiteUrl);
    }
    else
    {
      tabparameters['Url'] = "https://avanade.sharepoint.com/teams/ITS/"
    }

   
    if(this.state.items.length<=1)
    {
      this.getDatafromSharePointList(tabparameters);
    }  
    
    
    

    var categories = [{"Title":"Launched"},{"Title":"Rolling Out"},{"Title":"In Development"},{"Title":"Planning"},{"Title":"In the Queue"}];
  
    return (
      
        <div>
<WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateProperty} />

        <Tabs>
  <TabList>
    {categories.map(function(item,key){    
           var elWidth=100/categories.length;
           var tabWidth_h=elWidth.toString()+"%";
              return (
              
              <Tab style={{width:tabWidth_h}}>
                        <div className="contentWrapper"><div>{item['Title']}</div></div> 
                      </Tab>);
                      },this)}
  </TabList>
  <div className="verticalTabDescription">
  {categories.map(function(category,key){
return (
  <TabPanel>

    <div className="contentWrapper">
{this.state.items.map(function(item,key){
if(item["Stage"]==category["Title"])
{
  //var img=<div className="question" ><table><tr><td className="alignImage"></td><td>{item["Title"]}</td></tr></table></div>;        
  //var img=<div className="question" ><table><tr><td className="alignImage"><img src={srcDown} /></td><td>{item["Title"]}</td></tr></table></div>;        
  var img = <div className="question"><img src={srcDown}/>{item["Title"]}</div>

  var ServiceLine = item["Service"]!=""?item["Service"]:"";
  var Description = item["Description"]!=""?item["Description"]:"";
  var ReleaseDate = item["ReleaseDate"]!=null?this.getReleaseDate(item["ReleaseDate"]):"";
  
  var RoadmapURL = item["ReferenceURL"]!=null?item["ReferenceURL"].Url:"";
  var Contact = item["Contact"]!=null?item["Contact"]:"";
  
  
  
  return (  
    <Collapsible 
     
    triggerWhenOpen={this.checkOpen(img,item,this)} trigger={img} >
<div className="answers">
<div className="ava-item-Description" id="avaitemService">Service: {ServiceLine}</div>
<div className="ava-item-Description" id="avaitemDescription">{Description}</div>
{
  ReleaseDate!=""?<div className="ava-item-ReleaseDate" id="avaitemReleaseDate">Release Date: {ReleaseDate}</div>:<div></div>
}
{
  Contact!=""?<div className="ava-item-Description">Contact: <a className="contact-style-anchor" href={Contact.Url} target="_blank">{Contact.Description}</a></div>:<div></div>
}
{
RoadmapURL!=""?<div className="more-info"><a href={RoadmapURL} target="_blank">MORE INFO</a></div>:<div></div>
}

</div>

</Collapsible> 

   );
}

},this)}

</div>
  </TabPanel>    
);
},this)}

</div>
</Tabs>
      </div>   
  );
    
  }
}
