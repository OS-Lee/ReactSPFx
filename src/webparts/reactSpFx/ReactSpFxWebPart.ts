import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'ReactSpFxWebPartStrings';
import ReactSpFx from './components/ReactSpFx';
import { IReactSpFxProps } from './components/IReactSpFxProps';

import { AadHttpClient, HttpClientResponse, AadTokenProvider } from '@microsoft/sp-http';

export interface IReactSpFxWebPartProps {
  description: string;
}

export default class ReactSpFxWebPart extends BaseClientSideWebPart<IReactSpFxWebPartProps> {

 
  public render(): void {
    var tokenId = 'https://graph.microsoft.com'

    this.context.aadTokenProviderFactory.getTokenProvider()
      .then((tokenProvider: AadTokenProvider): Promise<string> => {
        // retrieve access token for the enterprise API secured with Azure AD
        return tokenProvider.getToken(tokenId);
      })
      .then((accessToken: string): void => {
          console.log(accessToken);
          console.log("This solution doesn't show ways to send/receive data to/from service layers");
          const element: React.ReactElement<IReactSpFxProps> = React.createElement(
            ReactSpFx,
          {
            description: this.properties.description,
            context:this.context,
            userToken: accessToken
          });

        ReactDom.render(element, this.domElement);
      });

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
