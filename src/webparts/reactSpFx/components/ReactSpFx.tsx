import * as React from 'react';
import styles from './ReactSpFx.module.scss';
import { IReactSpFxProps } from './IReactSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, ItemAddResult } from "@pnp/sp";

import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse, MSGraphClient } from '@microsoft/sp-http';

import Slider from "react-slick";
import "../../../../node_modules/slick-carousel/slick/slick.css";
import "../../../../node_modules/slick-carousel/slick/slick-theme.css";
import { any } from 'prop-types';

import { GraphFileBrowser } from '@microsoft/file-browser';
import { graph } from "@pnp/graph";
import { taxonomy, ITermStore, ITermSet, ITerms, ITermData, ITerm } from "@pnp/sp-taxonomy";
import ClassicEditor from 'ckeditor5-classic'
import * as jQuery from 'jquery';

import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IReactItem {
  ID: string,
  Title: string,
  Address: string
  Pic: {
    Description: string,
    Url: string
  }
}

export interface IReactGetItemsState {
  items: IReactItem[],
  selectValue: string
}

declare global {
  interface Window { _graphToken: any; }
}

const logo: any = require('../../assets/panda.jpg');

ClassicEditor.defaultConfig = {
  toolbar: {
    items: [
      'heading',
      '|',
      'bold',
      'italic',
      'fontSize',
      'fontFamily',
      'fontColor',
      'fontBackgroundColor',
      'link',
      'bulletedList',
      'numberedList',
      'imageUpload',
      'insertTable',
      'blockQuote',
      'undo',
      'redo'
    ]
  },
  image: {
    toolbar: [
      'imageStyle:full',
      'imageStyle:side',
      '|',
      'imageTextAlternative'
    ]
  },
  fontFamily: {
    options: [
      'Arial',
      'Helvetica, sans-serif',
      'Courier New, Courier, monospace',
      'Georgia, serif',
      'Lucida Sans Unicode, Lucida Grande, sans-serif',
      'Tahoma, Geneva, sans-serif',
      'Times New Roman, Times, serif',
      'Trebuchet MS, Helvetica, sans-serif',
      'Verdana, Geneva, sans-serif'
    ]
  },
  language: 'en'
}
var myEditor;

export default class ReactSpFx extends React.Component<IReactSpFxProps, IReactGetItemsState> {
  
  public constructor(props: IReactSpFxProps) {
    super(props);
    window._graphToken = props.userToken;
    this.state = {
      items: [],
      selectValue: "Radish"
    };
    sp.setup({
      spfxContext: this.context
    })
    this.next = this.next.bind(this);
    this.previous = this.previous.bind(this);
    graph.setup({
      spfxContext: this.props.context
    });
    
  }  

  public getAuthenticationToken(): Promise<string> {
    return new Promise(resolve => {
      resolve(
        window._graphToken
      );
    });

  }
  //   private async _getListData(): Promise<IReactItem> {
  //     return this.context.spHttpClient.get('https://sharepointTenant/sites/SharepointSite/_api/web/lists/GetByTitle(\'Wlasciwosci_toolbox\')/Items', SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
  //         return response.json();
  //     });
  // }

  private getTermsetWithChildren(termStoreName: string, termsetId: string) {
    return new Promise((resolve, reject) => {
      //const taxonomy = new Session(siteCollectionURL);
      const store: any = taxonomy.termStores.getByName(termStoreName);
      store.getTermSetById(termsetId).terms.select('Name', 'Id', 'Parent').get()
        .then((data: any[]) => {
          let result = [];
          // build termset levels
          do {
            for (let index = 0; index < data.length; index++) {
              let currTerm = data[index];
              if (currTerm.Parent) {
                let parentGuid = currTerm.Parent.Id;
                insertChildInParent(result, parentGuid, currTerm, index);
                index = index - 1;
              } else {
                data.splice(index, 1);
                index = index - 1;
                result.push(currTerm);
              }
            }
          } while (data.length !== 0);
          // recursive insert term in parent and delete it from start data array with index
          function insertChildInParent(searchArray, parentGuid, currTerm, orgIndex) {
            searchArray.forEach(parentItem => {
              if (parentItem.Id == parentGuid) {
                if (parentItem.children) {
                  parentItem.children.push(currTerm);
                } else {
                  parentItem.children = [];
                  parentItem.children.push(currTerm);
                }
                data.splice(orgIndex, 1);
              } else if (parentItem.children) {
                // recursive is recursive is recursive
                insertChildInParent(parentItem.children, parentGuid, currTerm, orgIndex);
              }
            });
          }
          resolve(result);
        }).catch(fail => {
          console.warn(fail);
          reject(fail);
        });
    });
  }

  /* Load CKeditor RTE*/
  public InitializeCKeditor(value): void {
    try {
      /*Replace textarea with classic editor*/
      ClassicEditor
        .create(document.querySelector("#editor1"), {
        })
        .then(editor => {
          myEditor = editor;
          var valueHtml: string;
          valueHtml = value.substr(1).slice(0, -1).replace('\\"/g', '"');
          editor.setData(valueHtml);
          //editor.addCss(".cke_editable{cursor:text; font-size: 14px; font-family: Arial, sans-serif;}");
          //editor.setReadOnly(true);
          editor.isReadOnly = true;
          console.log("CKEditor5 initiated");
        })
        .catch(error => {
          console.log("Error in Classic Editor Create " + error);
        });
    } catch (error) {
      console.log("Error in  InitializeCKeditor " + error);
    }
  }

  public componentDidMount() {
    const obj: string = JSON.stringify(
      {
        "fields": {
          'Title': 'Lin',
          'Company': 'Microsoft'
        }
      }
    );
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("/sites/wendytest123.sharepoint.com,2bf7a991-b669-4537-b0f5-59f7d6452e48,2c7b55b2-e306-407b-b1a2-1ca6fecc99ed/lists/77d9ee4c-9142-40d1-8edb-9bdfd226be2a/items")
          .header('Content-Type', 'application/json')
          .version("v1.0")
          .post(obj, (err, res, success) => {
            if (err) {
              console.log(err);
            }
            if (success) {
              console.log("success");
            }
          })
      });

    // this.props.context.spHttpClient.get(`https://wendytest123.sharepoint.com/sites/itch2/_api/web/lists/getbytitle('List1')/items?select=ID,Title`,
    //   SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
    //     response.json().then((responseJSON: any) => {
    //       reactHandler.setState({
    //         items: responseJSON.value.sort()
    //       });
    //     });
    //   });

    //var dom=this.props.context.pageContext..domElement;
    //   var value="Testvalue";
    //   sp.web.lists.getByTitle("listnotexist").items.add({
    //     Title : value})
    //     .catch((iar: any)  => {
    //      console.log(value,iar.item);

    //      const reader = iar.response.body.getReader();
    //      reader.read().then(({ done, value }) => {
    //        console.log(value);
    //      })
    //  })
    sp.web.lists.getByTitle('MyDoc2').views.getByTitle("All Documents").get().then((view: any) => {
      sp.web.lists.getByTitle('MyDoc2').getItemsByCAMLQuery({
        ViewXml: `<View><Query>` + view.ViewQuery + `</Query></View>`
      }).then((items: any) => {
        console.log(items);
      });
    })
    sp.web.lists.getByTitle('MyDoc2').items.select('Id,FileRef').get().then((items: any) => {
      items.map((item) => {
        console.log(item.FileRef);
      })
    })

    const query = `<Where>
        <Eq><FieldRef Name="File_x0020_Type"/>      
            <Value Type="Text">json</Value> 
        </Eq> 
    </Where>`;
    const xml = '<View Scope="RecursiveAll"><Query>' + query + '</Query></View>';
    sp.web.lists.getByTitle('MyDoc2').getItemsByCAMLQuery({ 'ViewXml': xml }, 'FileRef').then((items: any) => {
      items.map((item) => {
        console.log(item);
        sp.web.getFileByServerRelativeUrl(item.FileRef).getJSON().then((data) => {
          console.log(data);
        })
      })
    })

    sp.web.lists.getByTitle("TestList").items.getById(19).get().then((item: any) => {
      var value = JSON.stringify(item.Description);
      this.InitializeCKeditor(value);
    })
    var store: ITermStore = taxonomy.termStores.getByName("Taxonomy_hAIlyuIrZSNizRU+uUbanA==");
    var set: ITermSet = store.getTermSetById("70719569-ae34-4f24-81b9-0629d68c05aa");
    // load the data into the terms instances
    set.terms.get().then((terms: ITerm[]) => {
      terms.forEach((term: any) => {
        console.log(term['Name']);
        term.LocalCustomProperties._Sys_Nav_TargetUrl
      })
    });

    // this.getTermsetWithChildren(      
    //   'Taxonomy_hAIlyuIrZSNizRU+uUbanA==',
    //   '70719569-ae34-4f24-81b9-0629d68c05aa'
    // ).then(data => {
    //   console.log(data);
    // });

    var reactHandler = this;
    this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TestList')/items?select=ID,Title,Address,Pic`,
      SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          reactHandler.setState({
            items: responseJSON.value.sort()
          });
        });
      });
    // var style = document.createElement('style');
    // style.type = 'text/css';
    // style.innerHTML = "div[class^='pageTitle_']{display: none;}";
    // document.getElementsByTagName('head')[0].appendChild(style); 

    var userJSON = {
      "@odata.id": "https://graph.microsoft.com/v1.0/users/Lee@wendytest123.onmicrosoft.com"
    };

    // graph.groups.getById("0922a22b-82cc-4b51-95f0-4d4494bc31bb").members.add("https://graph.microsoft.com/v1.0/users/Lee@wendytest123.onmicrosoft.com").then(result => {
    //   console.log("user added");
    // });


    // this.props.context.msGraphClientFactory
    //   .getClient()
    //   .then((client: MSGraphClient): void => {
    //     client         
    //      .api("/groups/0922a22b-82cc-4b51-95f0-4d4494bc31bb/members/$ref")
    //      .version("v1.0")
    //       .post(userJSON, (err, res, success) => {
    //         if (err) {  
    //         console.log(err);                 
    //         }                
    //         if (success)
    //         {
    //           console.log("success");
    //         }            
    //       })
    //   });

  }

  protected slider;
  next() {
    this.slider.slickNext();
  }
  previous() {
    this.slider.slickPrev();
  }

  handleChange = (event) => {
    this.setState({ selectValue: event.target.value });
  };

  renderPic(item) {
    if (item.Pic === null) {
      return <img width={150} height={150} /> //use a default image better
    } else {
      return <img width={150} height={150} src={item.Pic.Url} />
    }
  }

  protected AddUserToGroup() {
    var data = myEditor.getData();
    sp.web.lists.getByTitle("TestList").items.getById(14).update({
      Title: "My New Title",
      Description: data
    }).then(i => {
      console.log(i);
    });
    console.log(data);
  }

  protected UploadFile() {
    var files = (document.getElementById('uploadFile') as HTMLInputElement).files;

    var file = files[0];
    if (file != undefined || file != null) {
      let spOpts: ISPHttpClientOptions = {
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: file
      };

      var url = `https://wendytest123.sharepoint.com/sites/modernteamlisa/_api/Web/Lists/getByTitle('MyDoc')/RootFolder/Files/Add(url='${file.name}', overwrite=true)`

      this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
        console.log(`Status code: ${response.status}`);
        console.log(`Status text: ${response.statusText}`);
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);

          let readOpts: ISPHttpClientOptions = {
            headers: {
              "Accept": "application/json"
            }
          }

          this.props.context.spHttpClient.get(responseJSON["@odata.id"] + "/ListItemAllFields", SPHttpClient.configurations.v1, readOpts).then((readResponse: SPHttpClientResponse) => {
            readResponse.json().then((metadataJSON: JSON) => {
              console.log(metadataJSON);
              var metadataType = metadataJSON["@odata.type"].replace('#', '');

              //SP.Data.MyDocItem metadata
              var metaDataBody: string = JSON.stringify({
                '__metadata': {
                  'type': metadataType
                },
                'Title': 'testupdate'
              });
              let updateOpt: ISPHttpClientOptions = {
                headers: {
                  "content-type": "application/json;odata=verbose",
                  "If-Match": "*",
                  "X-HTTP-Method": "MERGE",
                  'odata-version': '3.0'
                },
                body: metaDataBody
              }

              this.props.context.spHttpClient.post(responseJSON["@odata.id"] + "/ListItemAllFields", SPHttpClient.configurations.v1, updateOpt).then((reqResponse: any) => {
                console.log(reqResponse.status)
              })
            })
          })

        });
      });

    }
  }
  private changeStateDepartment = (item: IDropdownOption): void => {
    console.log("dropdown department changed values..." + item.selected + " , " + item.text);
}
  public render(): React.ReactElement<IReactSpFxProps> {    
    const settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 1,
      slidesToScroll: 1
    };

    const sendMail = {
      message: {
        subject: "Test",
        body: {
          contentType: "Text",
          content: "test email in SPFx call."
        },
        toRecipients: [
          {
            emailAddress: {
              address: "lee@wendytest123.onmicrosoft.com"
            }
          }
        ],
        ccRecipients: [
          {
            emailAddress: {
              address: "lee@wendytest123.onmicrosoft.com"
            }
          }
        ]
      },
      saveToSentItems: "false"
    };

    // let res = await client.api('/me/sendMail')
    //   .post(sendMail);
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // get information about the current user from the Microsoft Graph
        client
          // .api('/me')
          // .get((error, response: any, rawResponse?: any) => {
          //   let user = response.displayName;
          // });

          // .api('/me/sendMail')
          // .post(sendMail).then(()=>{
          //   console.log('email send');
          // })
          // .api('/me/drive/recent')
          // .get((error, response: any, rawResponse?: any) => {
          //   console.log(response);
          //   response.value.map((item: any) => {
          //     console.log(item.id);
          //  });
          // });
          .api("sites?search=*")
          .version("v1.0")
          .get((error, response: any, rawResponse?: any) => {
            console.log(response);
          });
      });

    return (
      <div className={styles.reactSpFx}>

        <div className={styles.container}>
          Image Load
        {/* <img src={require('../../assets/panda.jpg')} alt="test" /> */}
          {/* <img src={${require<string>('../../assets/panda.jpg')}} alt="My Company" /> */}
          {/* <div className={styles.img} title="Rencore logo">content</div> */}

          <input type="file" id="uploadFile" />
          <button className="button" onClick={() => this.UploadFile()} >Upload</button>

          <div>
            CKEditor
            <textarea id="editor1"></textarea>
          </div>
          <button className="button" onClick={this.AddUserToGroup}>
            AddUserToGroup
          </button>

          <Dropdown
                  placeHolder="Select Department"
                  label=""
                  multiSelect = {true}                  
                  id="component"                  
                  ariaLabel="Basic dropdown example"
                  options={[
                    { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
                    { key: 'apple', text: 'Apple' },
                    { key: 'banana', text: 'Banana' },
                    { key: 'orange', text: 'Orange', disabled: true },
                    { key: 'grape', text: 'Grape' },
                    { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
                    { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
                    { key: 'broccoli', text: 'Broccoli' },
                    { key: 'carrot', text: 'Carrot' },
                    { key: 'lettuce', text: 'Lettuce' },
                  ]}
                  onChanged={this.changeStateDepartment}
                />

          {(this.state.items || []).map((item, index) => (
            <div key={item.ID} className={(index % 2 == 0) ? styles.rowA : styles.rowB}>{item.Title}
              {index % 2}
              {this.renderPic(item)}
              {/* <div dangerouslySetInnerHTML={{ __html: item.Address.replace(/[\n\r]/g,"<br/>")}}></div>  */}
            </div>
          ))}
        </div>

        <GraphFileBrowser
          getAuthenticationToken={this.getAuthenticationToken}
          endpoint='https://graph.microsoft.com/v1.0/sites/siteid'
          onSuccess={(selectedKeys: any[]) => console.log(selectedKeys)}
          onCancel={(err: Error) => console.log(err.message)}
        />

        <select
          value={this.state.selectValue}
          onChange={this.handleChange}>
          <option value="Orange">Orange</option>
          <option value="Radish">Radish</option>
          <option value="Cherry">Cherry</option>
        </select>
        <div>
          <h2> Single Item</h2>
          <Slider ref={c => (this.slider = c)} {...settings}>
            <div>
              <h3>1</h3>
              <img width="100%" height="300" src="https://i.stack.imgur.com/yG5lu.png"></img>
            </div>
            <div>
              <h3>2</h3>
              <img width="100%" height="300" src="https://media.wired.com/photos/5d09594a62bcb0c9752779d9/master/w_2560%2Cc_limit/Transpo_G70_TA-518126.jpg"></img>
              <div>
                The SharePoint Framework (SPFx) is a page and web part model that provides full support for client-side SharePoint development, easy integration with SharePoint data, and support for open source tooling. With the SharePoint Framework, you can use modern web technologies and tools in your preferred development environment to build productive experiences and apps that are responsive and mobile-ready from day one. The SharePoint Framework works for SharePoint Online and also for on-premises (SharePoint 2016 Feature Pack 2 and SharePoint 2019).
              </div>
            </div>
            <div>
              <h3>3</h3>
              <img width="100%" height="300" src="https://i.stack.imgur.com/yG5lu.png"></img>
            </div>
            <div>
              <h3>4</h3>
            </div>
            <div>
              <h3>5</h3>
            </div>
            <div>
              <h3>6</h3>
            </div>
          </Slider>
          <br />
          <br />
          <div style={{ textAlign: "center" }}>
            <button className="button" onClick={this.previous}>
              Previous
          </button>
            <button className="button" onClick={this.next}>
              Next
          </button>
          </div>
        </div>
      </div>
    );
  }
}
