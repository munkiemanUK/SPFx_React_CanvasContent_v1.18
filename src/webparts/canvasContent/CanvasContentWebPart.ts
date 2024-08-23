import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CanvasContentWebPartStrings';
import CanvasContent from './components/CanvasContent';
import { ICanvasContentProps } from './components/ICanvasContentProps';
import {SPHttpClient} from '@microsoft/sp-http';

export interface ICanvasContentWebPartProps {
  description: string;
  siteUrl: string;
  pageId: number;
  numGroups : number;
  useList : string;
}

export default class CanvasContentWebPart extends BaseClientSideWebPart<ICanvasContentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICanvasContentProps> = React.createElement(
      CanvasContent,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        siteUrl: this.context.pageContext.site.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        context: this.context,
        numGroups : this.properties.numGroups,
        useList : this.properties.useList
      }
    );
    console.log("useList",this.properties.useList);
    console.log("numGroups",this.properties.numGroups);
    if(this.properties.useList){this._renderDataAsync()}

    ReactDom.render(element, this.domElement);
  }

  private async _renderDataAsync() : Promise<void> {
    await this._getData()
      .then((response:any) => {
        console.log('renderData',response);
        this._renderData(response);
        //fetchWPDataFlag = true;
    });
  }
  
  private async _getData() : Promise<any> {
    const endpoint = this.context.pageContext.site.absoluteUrl + `/_api/sitepages/pages(${this.properties.pageId})?$select=CanvasContent1&expand=CanvasContent1`;
    const rawResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    const jsonResponse = await rawResponse.json();
    const jsonCanvasContent = jsonResponse.CanvasContent1;
    const parseCanvasContent = JSON.parse(jsonCanvasContent);
    return parseCanvasContent;        
  }

  private _renderData(items:any): void {    
    items.forEach((item:any)=>{
      if(item.webPartId !== undefined){
        const webPartId : string = item.webPartId;

        if(webPartId === "b513eb44-9a56-4627-a2f3-743ef9090371" || webPartId === "78f2b269-7bab-4490-9c1c-9f407e4bdae0"){
          if(this.properties.numGroups === undefined){            
            //item.webPartData.properties.Slider>this.properties.numGroups || 
            this.properties.numGroups = item.webPartData.properties.Slider;
            console.log("renderData numGroups",this.properties.numGroups,item.webPartData.properties.Slider);
          }
        }
        //console.log("instanceID",this.context.instanceId);
      }
    })
  }

  protected onInit(): Promise<void> {
    this.properties.pageId = 1;

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log("prop pane numgroups",this.properties.numGroups);

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
                }),
                PropertyPaneSlider('numGroups', {
                  label:'How Many Link Groups? (max 10)',
                  min:0,
                  max:10,
                  value:0
                }),
                PropertyPaneToggle('useList', {
                  label: 'Use SharePoint List as link data?',
                  offText: 'No',
                  onText: 'Yes'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
