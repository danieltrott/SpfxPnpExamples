import * as strings from 'PnpProvidersWebPartStrings';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { GraphFI, graphfi, SPFx as graphSPFx } from '@pnp/graph';
import { SPFI, spfi, SPFx as spSPFx } from '@pnp/sp';

import { IPnpProvidersProps } from './components/IPnpProvidersProps';
import PnpProviders from './components/PnpProviders';
import { PnpProvider, PnpProviderType } from './provider/PnpProvider';

export interface IPnpProvidersWebPartProps {
  description: string;
}

export default class PnpProvidersWebPart extends BaseClientSideWebPart<IPnpProvidersWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  private graph: GraphFI;
  private sp: SPFI;

  public render(): void {
    const viewElement: React.ReactElement<IPnpProvidersProps> =
      React.createElement(PnpProviders, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
      });
    const providerElement: React.ReactElement<PnpProviderType> =
      React.createElement(
        PnpProvider,
        {
          graph: this.graph,
          sp: this.sp,
        },
        viewElement
      );

    ReactDom.render(providerElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    await super.onInit();
    this.sp = spfi().using(spSPFx(this.context));
    this.graph = graphfi().using(graphSPFx(this.context));
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
