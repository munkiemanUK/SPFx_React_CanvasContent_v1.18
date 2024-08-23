import * as React from 'react';
import styles from './CanvasContent.module.scss';
import type { ICanvasContentProps } from './ICanvasContentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';

let dataFetchFlag : boolean = false;

export interface IStates {
  useList:string;
  numGroups:number;
  listItems: any[];
}

export default class CanvasContent extends React.Component<ICanvasContentProps, IStates, {}> {  
  
  constructor(props: ICanvasContentProps){
    super(props);

    // set initial state
    this.state = {
      useList: "false",
      numGroups: 0,
      listItems: []
    };
  }

  public render(): React.ReactElement<ICanvasContentProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      useList,
      numGroups
    } = this.props;

    dataFetchFlag = false;

    return (
      <section className={`${styles.canvasContent} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>Use List Props: {escape(useList)}</div>
        <div>Number of Groups: {numGroups}</div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    console.log("component did mount");
    console.log("CDM useList props",this.props.useList);
    console.log("CDM num groups",this.props.numGroups,this.state.listItems.length);
    if(escape(this.props.useList)==='true'){
      this._getWebPartDataAsync();
      if(this.state.listItems.length === 0){ //&& dataFlag===false){
        this._getListData().then((response)=>{ 
          console.log("listItems",response);
        });     
      }  
    }
  }

  public componentDidUpdate(): void {
    console.log("component did update dataflag=",dataFetchFlag);
    //console.log("CDU useList props",this.props.useList);
    //console.log("CDU num groups",this.props.numGroups);
    if(this.state.listItems.length === 0 && !dataFetchFlag){
      this._getListData().then((response)=>{ 
        console.log("listItems",response);
        dataFetchFlag = true;
      });     
    }  
  }

  private _getWebPartDataAsync(): void {
    console.log('render webpart data');
    this.setState({numGroups:this.props.numGroups});
    this._getWebPartData()
      .then((response:any) => {
        this._renderWebPartData(response);
      });
  }

  private async _getWebPartData() {
    const endpoint = `${this.props.siteUrl}/_api/sitepages/pages(1)?$select=CanvasContent1&expand=CanvasContent1`;
    const rawResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    const jsonResponse = await rawResponse.json();
    const jsonCanvasContent = jsonResponse.CanvasContent1;
    const parseCanvasContent = JSON.parse(jsonCanvasContent);
    return parseCanvasContent;
  }

  public _renderWebPartData (items:any): void {
    
    console.log("items",items);
    const listContainer: Element = document.querySelector('#spListContainer')!;
    let html: string="";
    
    for(const item of items){
      if(item.webPartId !== undefined){
        const webPartID : string = item.webPartId;
        if(webPartID === "b513eb44-9a56-4627-a2f3-743ef9090371"){
          console.log("item",item.webPartData.title);
          html += `<div><h1>webpart title: ${item.webPartData.title}</h1></div>`;
        }
      }
    }
    if(listContainer){listContainer.innerHTML = html}
  }

  private async _getListData(): Promise<any> { 
    this.setState({listItems:[]});    

    const endpoint = `${this.props.siteUrl}/_api/web/lists/GetByTitle('Important%20Links')/Items`;
    return await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse)=> {
        return response.json();

      })
    //const jsonResponse = await rawResponse.json();
    //const parseResponse = JSON.parse(jsonResponse);
    //this.setState({listItems: parseResponse});

    //return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Important%20Links')/Items?$orderby=GroupID&$orderby=LinkOrder",SPHttpClient.configurations.v1)
    //.then((response: SPHttpClientResponse) => {
    //  return response.json();
    //});      
  }

}
