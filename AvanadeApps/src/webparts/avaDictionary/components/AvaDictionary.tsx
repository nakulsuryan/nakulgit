import * as React from 'react';
import styles from './AvaDictionary.module.scss';
import { IAvaDictionaryProps } from './IAvaDictionaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from "jquery";
import Collapsible from "react-collapsible";
import { Items } from 'sp-pnp-js/lib/pnp';
export interface IAvaDictionaryState {
  items: [{}],
  RenderArray: any,
  SearchAcronym: any,
  SearchResultAvailable: boolean
}

export default class AvaDictionary extends React.Component<IAvaDictionaryProps, IAvaDictionaryState> {

  public constructor(props: IAvaDictionaryProps, state: IAvaDictionaryState) {
    super(props);
    this.state = {
      items: [{}],
      RenderArray: [],
      SearchAcronym: "",
      SearchResultAvailable: false
    }

    this.BindOnSearch = this.BindOnSearch.bind(this);
    this.clearSearchResult = this.clearSearchResult.bind(this);

    this.getData();
  }

  public render(): React.ReactElement<IAvaDictionaryProps> {

    let srcDown = String(require("../Images/avaicon_down.jpg"));

    return (
      <div>
        <div className={styles.SearchAcronym}>
          <input type="Text" id="SearchInput" placeholder="Search dictionary..." onChange={(e) => this.BindOnSearch(e.target.value)}></input>
          <button className={styles.searchBtn} name="" onClick={this.clearSearchResult}>CLEAR</button>
        </div>
        {
          this.state.SearchAcronym == "" ?
            <div id="contentWrapper" className={styles.contentWrapper}>
              {
                this.state.RenderArray.map(function (category, key) {
                  let img: any = <div className={styles.sform}><img src={srcDown} /><span>{category["Category"]}</span></div>

                  if (category["Values"].length > 0) {
                    return (
                      <Collapsible triggerWhenOpen={this.checkOpen(img, category, this)} trigger={img}>
                        {
                          category["Values"].map(function (item) {
                            var linkurl = item["MoreLinnk"] != null ? item["MoreLinnk"].Url : "";
                            var linkDescription = item["MoreLinnk"] != null ? item["MoreLinnk"].Description : "";
                            
                            var NewLinkurl = item["AdditionalLink"] != null ? item["AdditionalLink"].Url : "";
                            var NewlinkDescription = item["AdditionalLink"] != null ? item["AdditionalLink"].Description : "";
                            
                            return (
                              <div>
                                {(item["Term_x0020_Type"] != null && item["Term_x0020_Type"] != "Acronym") ? (
                                <div className={styles.acronym}>
                                <strong>{item["Acronym_x0020_Short_x0020_Form"]}: </strong>{item["Acronym_x0020_Description"] }
                                <a target="_blank" href={linkurl}>{linkDescription}</a>&nbsp;<a target="_blank" href={NewLinkurl}>{NewlinkDescription}</a><span dangerouslySetInnerHTML={{ __html: item["AcronymEnhancedDescription"] }}></span></div>) :
                                  (<div className={styles.acronym}><strong>
                                    {item["Acronym_x0020_Short_x0020_Form"]}: </strong>
                                    {item["Acronym_x0020_Long_x0020_Form"]}
                                  </div>)}
                              </div>
                            )
                          }, this)}
                      </Collapsible>
                    );
                  }
                  else {
                    return (
                      <Collapsible transitionTime={50} triggerWhenOpen={this.checkOpen(img, category, this)} trigger={img}>
                        {
                          <div>
                            <p className={styles.acronym}>No item available.</p>
                          </div>
                        }
                      </Collapsible>
                    );
                  }
                }, this)}
            </div>
            :
            <div>
              <label id="SearchLabel" className={styles.searchLabel}>Search results for "<span className={styles.searchText}>{this.state.SearchAcronym}</span>"</label>
              <div id="contentWrapper" className={styles.contentWrapper}>
                {
                  this.state.SearchResultAvailable!=true?
                  <div><p className={styles.acronym}>Nothing here matches your search.</p></div>
                  :this.state.RenderArray.map(function (category, key) {
                      return (
                        category["Values"].map(function (item) {

                          var linkurl = item["MoreLinnk"] != null ? item["MoreLinnk"].Url : "";
                          var linkDescription = item["MoreLinnk"] != null ? item["MoreLinnk"].Description : "";
                           
                          var Newlinkurl = item["AdditionalLink"] != null ? item["AdditionalLink"].Url : "";
                          var NewlinkDescription = item["AdditionalLink"] != null ? item["AdditionalLink"].Description : "";

                          var AcronymShortForm = item["Acronym_x0020_Short_x0020_Form"] != null ? item["Acronym_x0020_Short_x0020_Form"] : "";
                          var AcronymLongForm = item["Acronym_x0020_Long_x0020_Form"] != null ? item["Acronym_x0020_Long_x0020_Form"] : "";
                          var AcronymDescription = item["Acronym_x0020_Description"] != null ? item["Acronym_x0020_Description"] : "";
                          var AcronymPoint = item["AcronymEnhancedDescription"] != null ? item["AcronymEnhancedDescription"] : "";

                          var html = '';
                          var regEx = new RegExp("(" + this.state.SearchAcronym + ")", "gi");

                          var index = AcronymShortForm.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                          if (index != -1) {
                            AcronymShortForm = AcronymShortForm.replace(regEx, "<mark>$1</mark>");
                          }

                          if (item["Term_x0020_Type"] != null && item["Term_x0020_Type"] != "Acronym") {

                            var index = AcronymDescription.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                            if (index != -1) {
                             AcronymDescription = AcronymDescription.replace(regEx, "<mark>$1</mark>");
                            }
                            var index = AcronymPoint.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                            if ((index != -1)&&(this.state.SearchAcronym.length>2)) {
                              AcronymPoint = AcronymPoint.replace(regEx, "<span style='background: yellow;'>$1</span>");
                             }

                            var index = linkDescription.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                            if (index != -1) {
                              linkDescription = linkDescription.replace(regEx, "<mark>$1</mark>");
                            }

                            var index = NewlinkDescription.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                            if (index != -1) {
                              NewlinkDescription = NewlinkDescription.replace(regEx, "<mark>$1</mark>");
                            }

                            html = `<strong>` + AcronymShortForm + `: </strong> ` + AcronymDescription + ` <a target="_blank" href=` + linkurl + `>` + linkDescription + ` </a>`+ `&nbsp; <a target="_blank" href=` + Newlinkurl + `>` + NewlinkDescription + ` </a>`+AcronymPoint;
                          }
                          else {

                            var index = AcronymLongForm.toLowerCase().indexOf(this.state.SearchAcronym.toLowerCase());
                            if (index != -1) {
                              AcronymLongForm = AcronymLongForm.replace(regEx, "<mark>$1</mark>");
                            }

                            html = `<strong>` + AcronymShortForm + `: </strong> ` + AcronymLongForm;
                          }

                          return (
                            <div>
                              {
                                <div className={styles.acronym} dangerouslySetInnerHTML={{ __html: html }}></div>
                              }
                            </div>
                          )



                        }, this)



                      )

                    

                  }, this)
                }
              </div>

            </div>
        }

      </div>

    );
  }

  private BindAllData() {

    let categories: any[] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X Y Z"];
    let tempArray: any[] = [];
    var mainArray: any[] = [];

    for (var i = 0; i < categories.length; i++) {
      if (this.state.items.length > 1) {
        var category = categories[i].toString();

        tempArray = this.state.items.filter(function (index) {
          if (category.toLowerCase() == "x y z") {
            if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("x") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("y") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("z")) {
              return index;
            }

          }
          else if (category.toLowerCase() == "a") {

            if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith(category.toLowerCase()) || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("@") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("$") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("#")) {
              return index;
            }
          }
          else {
            if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith(category.toLowerCase())) {
              return index;
            }
          }
        });

        tempArray.sort(function (a, b) {
          if (a.Term_x0020_Type < b.Term_x0020_Type)
            return -1;
          if (a.Term_x0020_Type > b.Term_x0020_Type)
            return 1;
          return 0;
        });

        mainArray.push({ "Category": category, "Values": tempArray });


        tempArray = null;
      }
    }

    this.setState({
      RenderArray: mainArray
    });
  }

  private BindOnSearch(searchtext) {

    if (searchtext != "") {
      this.setState({
        SearchAcronym: searchtext,
        SearchResultAvailable: false
      });
      searchtext = searchtext.toLowerCase();

      let categories: any[] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X Y Z"];
      let tempArray: any[] = [];
      var mainArray: any[] = [];

      for (var i = 0; i < categories.length; i++) {
        if (this.state.items.length > 1) {
          var category = categories[i].toString();

          tempArray = this.state.items.filter(function (index) {

            var AcronymShortForm = index["Acronym_x0020_Short_x0020_Form"] != null ? index["Acronym_x0020_Short_x0020_Form"] : "";
            var AcronymDetails = "";
            var AcronymPoints ="";

            if (index["Term_x0020_Type"] != null && index["Term_x0020_Type"] == "Acronym") {

              AcronymDetails = index["Acronym_x0020_Long_x0020_Form"] != null ? index["Acronym_x0020_Long_x0020_Form"] : "";
            }
            else {
                AcronymDetails = index["Acronym_x0020_Description"]  != null ? index["Acronym_x0020_Description"]  : "";
               AcronymPoints = index["AcronymEnhancedDescription"]  != null ? index["AcronymEnhancedDescription"]  : "";
              }
             

            var AcronymMoreLink = index["MoreLinnk"] != null ? index["MoreLinnk"].Description : "";
            var AcronymNewLink = index["AdditionalLink"] != null ? index["AdditionalLink"].Description : "";



            if (category.toLowerCase() == "x y z") {
              if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("x") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("y") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("z")) {

                if (AcronymShortForm.toLowerCase().indexOf(searchtext) != -1 || AcronymDetails.toLowerCase().indexOf(searchtext) != -1 || AcronymPoints.toLowerCase().indexOf(searchtext) != -1 || AcronymMoreLink.toLowerCase().indexOf(searchtext) != -1 || AcronymNewLink.toLowerCase().indexOf(searchtext) != -1) {
                  return index;
                }

              }

            }
            else if (category.toLowerCase() == "a") {

              if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith(category.toLowerCase()) || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("@") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("$") || index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith("#")) {

                if (AcronymShortForm.toLowerCase().indexOf(searchtext) != -1 || AcronymDetails.toLowerCase().indexOf(searchtext) != -1 || AcronymPoints.toLowerCase().indexOf(searchtext) != -1 || AcronymMoreLink.toLowerCase().indexOf(searchtext) != -1 || AcronymNewLink.toLowerCase().indexOf(searchtext) != -1) {
                  return index;
                }

              }
            }
            else {
              if (index["Acronym_x0020_Short_x0020_Form"].toLowerCase().startsWith(category.toLowerCase())) {

                if (AcronymShortForm.toLowerCase().indexOf(searchtext) != -1 || AcronymDetails.toLowerCase().indexOf(searchtext) != -1 || AcronymPoints.toLowerCase().indexOf(searchtext) != -1 || AcronymMoreLink.toLowerCase().indexOf(searchtext) != -1 || AcronymNewLink.toLowerCase().indexOf(searchtext) != -1) {
                  return index;
                }

              }
            }
          });

          tempArray.sort(function (a, b) {
            if (a.Term_x0020_Type < b.Term_x0020_Type)
              return -1;
            if (a.Term_x0020_Type > b.Term_x0020_Type)
              return 1;
            return 0;
          });

          if (tempArray.length > 0) {
            this.setState({
              SearchResultAvailable: true
            });
          }

          mainArray.push({ "Category": category, "Values": tempArray });


          tempArray = null;
        }
      }


      console.log(this.state.SearchResultAvailable);

      this.setState({
        RenderArray: mainArray
      });
    
    }
    else {
      this.setState({
        SearchAcronym: ""
      });
      this.BindAllData();
    }

  }
  public clearSearchResult() {

    (document.getElementById("SearchInput") as HTMLInputElement).value = "";
    this.BindOnSearch("");

  }
  public checkOpen(img, category, obj) {
    var srcUp = String(require("../Images/avaicon_up.jpg"));
    img = <div className={styles.sform}><img src={srcUp} /><span>{category["Category"]}</span></div>
    return img;
  }

  public checkOpenforX(Ximg, category, obj) {
    var srcUp = String(require("../Images/avaicon_up.jpg"));
    Ximg = <div className={styles.sform}><img src={srcUp} /><span>XYZ</span></div>
    return Ximg;
  }

  public getData() {
    let reactHandler: any = this;
    jquery.ajax({
      url: `${this.props.urlpanel}/_api/web/lists/GetByTitle('${this.props.listName}')/Items?$select=Title,Acronym_x0020_Short_x0020_Form,Acronym_x0020_Long_x0020_Form,Acronym_x0020_Description,MoreLinnk,AcronymEnhancedDescription,AdditionalLink,Term_x0020_Type&$orderby=Acronym_x0020_Short_x0020_Form asc&$top=5000`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        let array: any[] = resultData.d.results;
        console.log(resultData.d.results);
        reactHandler.setState({
          items: resultData.d.results
        });

        reactHandler.BindAllData();
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log("Error while fetching the data." + errorThrown);
      }
    });
  }
}
