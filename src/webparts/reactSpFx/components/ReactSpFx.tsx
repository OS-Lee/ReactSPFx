import * as React from 'react';
import styles from './ReactSpFx.module.scss';
import { IReactSpFxProps } from './IReactSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

import Slider from "react-slick";
import "../../../../node_modules/slick-carousel/slick/slick.css"; 
import "../../../../node_modules/slick-carousel/slick/slick-theme.css";
import { any } from 'prop-types';

export interface IReactItem{ 
  ID:string,
  Title:string,
  Address:string
}

export interface IReactGetItemsState{ 
  items:IReactItem[]
}

export default class ReactSpFx extends React.Component<IReactSpFxProps,IReactGetItemsState> {
  
  public constructor(props: IReactSpFxProps) {
    super(props);
    this.state = {
      items:[]   
    };
    this.next = this.next.bind(this);
    this.previous = this.previous.bind(this);    
  }

  
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

  public render(): React.ReactElement<IReactSpFxProps> {
    const settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 1,
      slidesToScroll: 1
    };
    return (
      <div className={styles.reactSpFx}>
        <div className={styles.container}>  
        {(this.state.items || []).map(item => (
            <div key={item.ID} className={styles.row}>{item.Title}
            <div dangerouslySetInnerHTML={{ __html: item.Address.replace(/[\n\r]/g,"<br/>")}}></div> 
          </div> 
          ))}                          
        </div>
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
