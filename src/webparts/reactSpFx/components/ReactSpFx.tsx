import * as React from 'react';
import styles from './ReactSpFx.module.scss';
import { IReactSpFxProps } from './IReactSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

import Slider from "react-slick";
import "../../../../node_modules/slick-carousel/slick/slick.css"; 
import "../../../../node_modules/slick-carousel/slick/slick-theme.css";
import { any } from 'prop-types';

import { GraphFileBrowser } from '@microsoft/file-browser';

export interface IReactItem{ 
  ID:string,
  Title:string,
  Address:string
}

export interface IReactGetItemsState{ 
  items:IReactItem[],
  selectValue:string
}

declare global {
  interface Window { _graphToken: any; }
}

export default class ReactSpFx extends React.Component<IReactSpFxProps,IReactGetItemsState> {
  
  public constructor(props: IReactSpFxProps) {
    super(props);
    window._graphToken =props.userToken;
    this.state = {
      items:[],
      selectValue:"Radish"
    };
    this.next = this.next.bind(this);
    this.previous = this.previous.bind(this);    
  }

  public getAuthenticationToken():Promise<string> {    
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
  public componentDidMount() {
    var reactHandler = this;
    this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TestList')/items?select=ID,Title,Address`,
    SPHttpClient.configurations.v1) .then((response: SPHttpClientResponse) => {  
      response.json().then((responseJSON: any) => {            
        reactHandler.setState({
          items: responseJSON.value
        });
      });  
    });   
  }
  
  protected slider;
  next() {
    this.slider.slickNext();
  }
  previous() {
    this.slider.slickPrev();
  }

  handleChange = (event) => {
    this.setState({selectValue:event.target.value});
  };

  public render(): React.ReactElement<IReactSpFxProps> {
    const settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 1,
      slidesToScroll: 1
    };
    
    alert(this.state.selectValue);

    return (
      <div className={styles.reactSpFx}>
      
        <div className={styles.container}>  
        {/* {(this.state.items || []).map(item => (
            <div key={item.ID} className={styles.row}>{item.Title}
            <div dangerouslySetInnerHTML={{ __html: item.Address.replace(/[\n\r]/g,"<br/>")}}></div> 
          </div> 
          ))}                           */}
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
            </div>
            <div>
              <h3>2</h3>
            </div>
            <div>
              <h3>3</h3>
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
          <br/>
          <br/>
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
