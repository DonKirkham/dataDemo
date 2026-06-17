// ABOUTME: DataDemo web part entry point with PnP property pane controls for site and list selection.
// ABOUTME: Passes a service factory to the React component, which handles service switching at runtime.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldSitePicker,
  IPropertyFieldSite,
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls';

import { Placeholder } from '@pnp/spfx-controls-react';

import { Logger, LogLevel } from './utilities/logger';

import DataDemo from './components/DataDemo';
import { IDataDemoProps } from './components/IDataDemoProps';
import { ServiceFactory } from './services/ServiceFactory';
import { IListIdentifier } from './models/ISpService';

export interface IStoredList {
  id: string;
  title: string;
  url?: string;
}

export interface IDataDemoWebPartProps {
  sites: IPropertyFieldSite[];
  list: IStoredList | undefined;
  enhancedLogging: boolean;
  apiBaseUrl: string;
  apiResourceUri: string;
}

// Defaults point at the deployed apiDemo Azure Function and its app registration.
// Override per-environment in the property pane.
const DEFAULT_API_BASE_URL = 'https://apidemo-func-4r1iq2.azurewebsites.net';
const DEFAULT_API_RESOURCE_URI = 'api://99b202da-a167-4232-b6d7-7e4bb7a2fdaa';

export default class DataDemoWebPart extends BaseClientSideWebPart<IDataDemoWebPartProps> {

  private _factory: ServiceFactory | undefined;

  private _apiConfig(): { baseUrl: string; resourceUri: string } {
    return {
      baseUrl: this.properties.apiBaseUrl || DEFAULT_API_BASE_URL,
      resourceUri: this.properties.apiResourceUri || DEFAULT_API_RESOURCE_URI
    };
  }

  public render(): void {
    const list = this.properties.list;
    const hasSite = this.properties.sites && this.properties.sites.length > 0;

    Logger.debug(`render: hasSite=${hasSite}, list=${list?.title ?? 'none'}`);

    ReactDom.unmountComponentAtNode(this.domElement);

    // Keep the factory's API config in sync with the property pane between renders.
    if (this._factory) {
      this._factory.api = this._apiConfig();
    }

    if (!hasSite || !list) {
      Logger.debug('render: showing configuration placeholder');
      const element = React.createElement(Placeholder, {
        iconName: 'Edit',
        iconText: 'Configure Data Demo',
        description: 'Select a site and list in the property pane to get started.',
        buttonLabel: 'Configure',
        onConfigure: () => this.context.propertyPane.open()
      });
      ReactDom.render(element, this.domElement);
      return;
    }

    const site = this.properties.sites[0];
    const siteUrl = this._toAbsoluteSiteUrl(site.url);
    const siteId = site.id || this.context.pageContext.site.id.toString();
    const listIdentifier: IListIdentifier = { id: list.id, title: list.title };

    Logger.debug(`render: mounting DataDemo component for site=${siteUrl}, list=${listIdentifier.title}`);

    const element: React.ReactElement<IDataDemoProps> = React.createElement(
      DataDemo,
      {
        factory: this._factory,
        site: { url: siteUrl, id: siteId },
        list: listIdentifier
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    Logger.attachConsoleListener();
    Logger.setLevel(this.properties.enhancedLogging ? LogLevel.Verbose : LogLevel.Warning);

    Logger.info(`onInit: starting (enhancedLogging=${this.properties.enhancedLogging ? 'on' : 'off'})`);

    if (!this.properties.sites || this.properties.sites.length === 0) {
      Logger.debug(`onInit: defaulting site to current web ${this.context.pageContext.web.absoluteUrl}`);
      this.properties.sites = [{
        url: this.context.pageContext.web.absoluteUrl,
        title: this.context.pageContext.web.title,
        id: this.context.pageContext.site.id.toString()
      }];
    }

    this._factory = new ServiceFactory(this.context, this._apiConfig());
    Logger.debug('onInit: web part initialized, service factory ready');
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors, palette } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
    }
  }

  protected onDispose(): void {
    Logger.debug('onDispose: unmounting web part');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _getSiteUrl(): string | undefined {
    const sites = this.properties.sites;
    return sites && sites.length > 0 ? sites[0].url : undefined;
  }

  private _toAbsoluteSiteUrl(siteUrl: string | undefined): string {
    if (!siteUrl || siteUrl === '/') {
      return this.context.pageContext.web.absoluteUrl;
    }
    if (/^https?:\/\//i.test(siteUrl)) {
      return siteUrl;
    }
    const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;
    return `${origin}${siteUrl.startsWith('/') ? '' : '/'}${siteUrl}`;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the data source for the web part.'
          },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyFieldSitePicker('sites', {
                  label: 'Select a site',
                  initialSites: this.properties.sites || [],
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'sitePickerFieldId'
                }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list?.id,
                  disabled: !this._getSiteUrl(),
                  includeHidden: false,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  webAbsoluteUrl: this._getSiteUrl(),
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Advanced settings for the web part.'
          },
          groups: [
            {
              groupName: 'Elevated API',
              groupFields: [
                PropertyPaneTextField('apiBaseUrl', {
                  label: 'API base URL',
                  description: 'Base URL of the apiDemo Azure Function (Simple Auth / Entra App endpoints).',
                  placeholder: DEFAULT_API_BASE_URL
                }),
                PropertyPaneTextField('apiResourceUri', {
                  label: 'API resource URI (Entra App)',
                  description: 'App ID URI the AadHttpClient requests a token for.',
                  placeholder: DEFAULT_API_RESOURCE_URI
                })
              ]
            },
            {
              groupName: 'Logging',
              groupFields: [
                PropertyPaneToggle('enhancedLogging', {
                  label: 'Enhanced logging',
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

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    Logger.debug(`property pane changed: ${propertyPath}`);

    if (propertyPath === 'sites') {
      Logger.debug('site changed, clearing selected list');
      this.properties.list = undefined;
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'enhancedLogging') {
      Logger.setLevel(newValue ? LogLevel.Verbose : LogLevel.Warning);
      Logger.debug(`enhanced logging ${newValue ? 'enabled' : 'disabled'}`);
    }

    // includeListTitleAndUrl returns an IPropertyFieldList object — store it directly
    if (propertyPath === 'list' && newValue && typeof newValue === 'object') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const listObj = newValue as any;
      this.properties.list = {
        id: listObj.id,
        title: listObj.title || '',
        url: listObj.url
      };
    }

    this.render();
  }
}
