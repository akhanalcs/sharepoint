import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  // Adding properties to the property pane. Step 1: Import the respective property pane fields (that are functions) from the framework.
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

// Adding properties to the property pane. Step 2: Update the web part properties to include the new properties.
// This property definition interface is used to define the structure and types for the web part's properties.
// Put simply: these are web part's properties.
// For complete flow, take a look at "Configure the Web part property pane" section in the main README.md file of this repo.
export interface IHelloWorldWebPartProps {
  description: string;
  test: string; // Will store text from PropertyPaneTextField
  test1: boolean; // Will store state from PropertyPaneCheckbox
  test2: string; // Will store selection from PropertyPaneDropdown
  test3: boolean; // Will store state from PropertyPaneToggle
}

export interface ISPList{
  Title: string;
  Id: string;
}

// ISPList interface holds the SharePoint list information that we're connecting to.
export interface ISPLists{
  value: ISPList[];
}

// Main entry point for the web part.
// Any client-side web part should extend the BaseClientSideWebPart class to be defined as a valid web part.
// The web part class is defined to accept a property type IHelloWorldWebPartProps.
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // This model is flexible enough so that web parts can be built in any JavaScript framework and loaded into the DOM element.
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <!-- this.properties comes from "protected get properties(): TProperties;" in BaseWebPart. Found that by Cmd+CLick on this.properties -->
        <!-- this.properties.test comes from IHelloWorldWebPartProps.test -->
        <p>${escape(this.properties.test)}</p>
        <p>${this.properties.test1}</p>
        <p>${escape(this.properties.test2)}</p>
        <p>${this.properties.test3}</p>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
        <ul class="${styles.links}">
          <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
          <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
        </ul>
        <div>Web part description: <strong>${this.properties.description}</strong></div>  
        <div>Web part test: <strong>${escape(this.properties.test)}</strong></div>
        <!-- this.context.pageContext.site means site collection, .web means site. -->
        <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
      </div>
      <div id="spListContainer" />
    </section>`;
    
    this._renderListAsync();
  }

  protected onInit(): Promise<void> {
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Adding properties to the property pane. Step 3: Add the new property pane fields and map them to their respective typed objects.
  // "Mapping" in this context means connecting the UI controls in the property pane (that appear when you click edit on the web part) 
  // to the specific properties defined in your IHelloWorldWebPartProps interface.
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
                }),
                // PropertyPaneTextField is a function that creates a text field control for the property pane.  
                // First parameter (e.g., 'description', 'test') is the property name in IHelloWorldWebPartProps that this control is bound to.
                // Second parameter is an object {/* configuration options */} containing the configuration options for this particular control.
                  
                // The first parameter ('test', 'test1', etc.) is what "maps" the UI control to the property in your interface.
                PropertyPaneTextField('test', {
                  label: 'Multi-line Text Field',
                  multiline: true
                }),
                PropertyPaneCheckbox('test1', {
                  text: 'Checkbox'
                }),
                PropertyPaneDropdown('test2', {
                  label: 'Dropdown',
                  options: [
                    {key: '1', text: 'One'},
                    {key: '2', text: 'Two'},
                    {key: '3', text: 'Three'},
                    {key: '4', text: 'Four'}
                  ]
                }),
                PropertyPaneToggle('test3', {
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
  
  // The method uses the spHttpClient helper class and issues an HTTP GET request.
  // It uses the ISPLists interface and also applies a filter to not retrieve hidden lists.
  private _getListData(): Promise<ISPLists>{
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>{
          return response.json();
        })
        .catch(()=>{});
  }
  
  private _renderList(items: ISPList[]): void{
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}">
          <span class="ms-font-m">${item.Title}</span>    
        </li>
      </ul>`;
    });

    // You'll add "<div id="spListContainer" />" in the render() method later
    if(this.domElement.querySelector('#spListContainer') !== null){
      this.domElement.querySelector('#spListContainer')!.innerHTML = html;
    }
  }
  
  private _renderListAsync():void{
    this._getListData()
        .then((response) => {
          this._renderList(response.value);
        })
        .catch(() => {});
  }
}
