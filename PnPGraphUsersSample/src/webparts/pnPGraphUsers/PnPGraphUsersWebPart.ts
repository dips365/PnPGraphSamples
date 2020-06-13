import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { graph } from "@pnp/graph";
import * as strings from 'PnPGraphUsersWebPartStrings';
import PnPGraphUsers from './components/PnPGraphUsers';
import { IPnPGraphUsersProps } from './components/IPnPGraphUsersProps';
import { PnPGraphService } from "../../Services/PnPGraphService";
export interface IPnPGraphUsersWebPartProps {
  description: string;
}

export default class PnPGraphUsersWebPart extends BaseClientSideWebPart <IPnPGraphUsersWebPartProps> {
private PnPGraphServiceInstance:PnPGraphService;

  public async onInit(){
    await super.onInit().then(()=>{
      graph.setup({
        spfxContext:this.context
      });

      this.PnPGraphServiceInstance = new PnPGraphService();
    });
  }

  public render(): void {
    const element: React.ReactElement<IPnPGraphUsersProps> = React.createElement(
      PnPGraphUsers,
      {
        description: this.properties.description,
        PnPGraphServiceInstance:this.PnPGraphServiceInstance
      }
    );

    ReactDom.render(element, this.domElement);
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
