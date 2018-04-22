import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DevWebPart.module.scss';
import * as strings from 'DevWebPartStrings';

import { $REST, Fabric } from "gd-sprest-js";
import "gd-sprest-js/build/lib/css/fabric.min.css";
import "gd-sprest-js/build/lib/css/fabric.components.min.css";
import "gd-sprest-js/build/lib/css/gd-sprest-js.css";

export interface IDevWebPartProps {
  description: string;
}

export default class DevWebPart extends BaseClientSideWebPart<IDevWebPartProps> {
  private el: HTMLDivElement = null;

  // Render method
  public render(): void {
    // Set the context
    $REST.ContextInfo.setPageContext(this.context.pageContext);

    // Set the webpart template
    this.domElement.innerHTML = `
      <div class="${ styles.dev}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Dev Playground</span>
              <p>"${ this.properties.description}"</p>
              <div id="main"></div>
            </div>
          </div>
        </div>
      </div>`;

    // Set the element
    this.el = this.domElement.querySelector("#main") as HTMLDivElement;

    // Render a panel
    this.renderPanel();
  }

  // Version
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Configuration
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

  /**
   * Private Methods
   */

  // Display a panel
  private renderPanel() {
    // Set the template
    this.el.innerHTML = "<div></div><div></div>";

    // Render a panel
    let panel = Fabric.Panel({
      el: this.el.children[0],
      headerText: "My Panel",
      panelContent: [
        "<h1>Office Fabric-JS Framework</h1>",
        "<p>This panel component was created using the JavaScript framework.</p>"
      ].join('\n')
    });

    // Render a button
    Fabric.Button({
      el: this.el.children[1],
      text: "Show Panel",
      onClick: () => {
        // Display the panel
        panel.show();
      }
    });
  }
}
