import * as React from 'react';
import styles from './AccordionApp.module.scss';
import { IAccordionAppProps } from './IAccordionAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Accordion from  'react-responsive-accordion';
import Collapsible from 'react-collapsible';
import * as $ from 'jquery';
import pnp, { Web } from "sp-pnp-js";
import * as ReactDOM from 'react-dom';
import { PageContext } from '@microsoft/sp-page-context';
import './Accordion.css';
import * as strings from 'AccordionAppWebPartStrings';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
//Collection created as property
export interface IReactSpfxState{   
  items:[  
    { 
    }]
} 
const IconDown = () => (
  <Icon iconName="ChevronDownSmall" className={styles.iconWhite}  
  />
);
const IconUp = () => (
  <Icon iconName="ChevronUpSmall" className={styles.iconWhite} 
   />
);

export default class AccordionApp extends React.Component<IAccordionAppProps,IReactSpfxState> {
   
  

  public constructor(props: IAccordionAppProps, state: IReactSpfxState){
    super();  
    this.state = {  
      items: [  
        { 
        }  
      ] 
    };     
 
  } 
  public setTitle(event){
    this.props.setTitle(event.target.value);
  }
  //Call the api and load html contents 
  public render(): React.ReactElement<IAccordionAppProps> {  
    var desc1="";  
    var config=false;
    var desc2= <div className={styles.hideData} ></div>;
    if(this.props.siteUrl!=undefined && this.props.siteUrl!=null && this.props.siteUrl!="")
    if(this.props.listName!=null && this.props.listName!=undefined && this.props.listName!="" &&
    this.props.title!=null && this.props.title!=undefined && this.props.title!="" &&
    this.props.description!=null && this.props.description!=undefined && this.props.description!="" &&
    this.props.view!=null && this.props.view!=undefined && this.props.view!="" 
   
  )
    {
      desc1="";
      desc2= <div className={styles.hideData} ></div>;      
     this.getDatafromSharePointList();    
    }
    else
    {
desc1="Configure Avanade Accordion App from property pane";
desc2= <div dangerouslySetInnerHTML={{ __html:desc1 }}></div>;
    }

    //Render the html
    return (
      <div className={styles.accordionApp}>    
       <div className="webpartheader">
          { this.props.isEditMode && <textarea onChange={this.setTitle.bind(this)} className="edit" placeholder={strings.TitlePlaceholder} aria-label="Add a title" defaultValue={this.props.titleHeader}></textarea> }
          { !this.props.isEditMode && <span className="view">{this.props.titleHeader}</span> }          
        </div>
      {desc2}   
        {this.state.items.map(function(item,key){  
          var img=<div className={styles.question} ><table><tr className={styles.row123}><td className={styles.tdClass}><span className={styles.downArrowClass}><IconDown /></span></td><td className={styles.tdClass2}>{item[this.props.title]}</td></tr></table></div>;        
             return (<Collapsible triggerWhenOpen={this.checkOpen(img,item,this)}  trigger={img} >
             <div className={styles.answers} dangerouslySetInnerHTML={{ __html: item[this.props.description] }}></div></Collapsible>);
              },this
              )}         
      </div>
    );
    
  }
  
  //Event handler to change image
  private checkOpen(img,item,obj){
    img=<div className={styles.question2}><table><tr className={styles.row123}><td className={styles.tdClass}><span className={styles.upArrowClass}><IconUp /></span></td><td className={styles.tdClass2}>{item[obj.props.title]}</td></tr></table></div>; 
    return img;

  }
// Rest API to get list 
  private getDatafromSharePointList()
{
  var reactHandler = this;
   var camlquery="<View><Query>" +this.props.view + "</Query></View>";
    const web = new Web(this.props.siteUrl);
    web.lists.getByTitle(this.props.listName).getItemsByCAMLQuery({'ViewXml':camlquery}).then((res) => {
         // console.log(res);
          reactHandler.setState({
            items:res
            
                });
        }).catch((error)=>
        {
          console.log("Error in getting data from list " + error);
        });   
}


}


