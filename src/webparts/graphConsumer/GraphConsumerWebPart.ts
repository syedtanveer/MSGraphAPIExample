import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { sp } from "@pnp/sp";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "GraphConsumerWebPartStrings";
import GraphConsumer from "./components/GraphConsumer";
import { IGraphConsumerProps } from "./components/IGraphConsumerProps";

export interface IGraphConsumerWebPartProps {
  description: string;
}

export default class GraphConsumerWebPart extends BaseClientSideWebPart<
  IGraphConsumerWebPartProps
> {
  public render(): void {
    const element: React.ReactElement<IGraphConsumerProps> = React.createElement(
      GraphConsumer,
      {
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    return super.onInit();
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
