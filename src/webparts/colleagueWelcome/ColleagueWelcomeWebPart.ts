import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

import styles from './ColleagueWelcomeWebPart.module.scss';
import { UserGraphService } from '../../services';

export interface IColleagueWelcomeWebPartProps {
  welcomeprefix: string;
  customcss: string;
}

export default class ColleagueWelcomeWebPart extends BaseClientSideWebPart<IColleagueWelcomeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.colleagueWelcome} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}" style="${escape(this.properties.customcss)}">
        <div id="graphCurrentUserResult"></div>
      </div>      
    </section>`;

    //
    const _userGraphServiceInstance = this.context.serviceScope.consume(UserGraphService.serviceKey);
    _userGraphServiceInstance.getCurrentUserDetails()
      .then((user: MicrosoftGraph.User) => {
        //console.log("User:" + JSON.stringify(user));
        const _container = document.getElementById("graphCurrentUserResult");
        if(_container) {
          if(user.givenName !== null) {          
            _container.innerHTML = escape(this.properties.welcomeprefix)+escape(user.givenName);
          }
        }
      })
      .catch((error: any) => {
        // error
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Presents the colleagues name"
          },
          groups: [
            {
              groupName: "Presentation attributes",
              groupFields: [
                PropertyPaneTextField('welcomeprefix', {
                  label: "Prefix Text"
                }),
                PropertyPaneTextField('customcss', {
                  label: "Custom Css Styles"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
