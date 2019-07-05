import * as React from 'react';
import styles from './ReactSpFx.module.scss';
import { IReactSpFxProps } from './IReactSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

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
  public render(): React.ReactElement<IReactSpFxProps> {
    return (
      <div className={styles.reactSpFx}>
        <div className={styles.container}>  
        {(this.state.items || []).map(item => (
            <div key={item.ID} className={styles.row}>{item.Title}
            <div dangerouslySetInnerHTML={{ __html: item.Address}}></div> 
          </div> 
          ))}                          
        </div>
      </div>
    );    
  }
}
