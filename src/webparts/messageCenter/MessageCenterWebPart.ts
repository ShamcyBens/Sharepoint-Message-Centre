import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MessageCenterWebPart.module.scss';
import * as strings from 'MessageCenterWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IMessageCenterWebPartProps {
  description: string;
}

export default class MessageCenterWebPart extends BaseClientSideWebPart<IMessageCenterWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.messageCenter} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
          <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        </div>
        <div id="messagesContainer"></div>
      </section>`;
      
    this.loadMessages();
  }

  protected async onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private async _getEnvironmentMessage(): Promise<string> {
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

  private loadMessages(): void {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Messages')/items?$select=Message Title,Body,Service User,Group`;

    // eslint-disable-next-line no-void
    void this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        this.filterMessages(data.value); 
      });
  }

  private filterMessages(messages: any[]): void {
    const currentUserEmail = this.context.pageContext.user.email;

    // eslint-disable-next-line no-void
    void this.context.msGraphClientFactory.getClient('3')
      .then((client: MSGraphClientV3) => {
        const groupPromises: Promise<any>[] = [];

        messages.forEach(message => {
          if (message.Group) {
            groupPromises.push(client.api(`/groups/${message.Group}/members`).get());
          }
        });

        // eslint-disable-next-line no-void
        void Promise.all(groupPromises).then(groupResults => {
          const groupMembers = groupResults.flatMap((result: any) => result.value.map((user: any) => user.mail || user.userPrincipalName));

          const filteredMessages = messages.filter(message => {
            return message.Recipient === currentUserEmail || groupMembers.includes(currentUserEmail);
          });

          this.renderMessages(filteredMessages);
        });
      });
  }

  private renderMessages(messages: any[]): void {
    const messagesContainer: HTMLElement | null = this.domElement.querySelector('#messagesContainer');
    if (messagesContainer) {
      messagesContainer.innerHTML = '';

      messages.forEach(message => {
        const messageElement = document.createElement('div');
        messageElement.className = styles.message;
        messageElement.innerHTML = `
          <div class="${styles.messageTitle}">${escape(message.Title)}</div>
          <div class="${styles.messageBody}">${escape(message.Body)}</div>
          <div class="${styles.messageRecipient}">To: ${escape(message.Recipient)}${message.Group ? `, Group: ${escape(message.Group)}` : ''}</div>
        `;
        messagesContainer.appendChild(messageElement);
      });
    }
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
