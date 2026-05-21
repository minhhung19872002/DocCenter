import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DocCenterWebPartStrings';
import { DocCenter } from './components/DocCenter';
import { IDocCenterProps } from './components/IDocCenterProps';
import { SharePointService } from './services/SharePointService';

export interface IDocCenterWebPartProps {
  documentsLibraryTitle: string;
  hashtagsListTitle: string;
}

const DOC_CENTER_APP_BODY_CLASS = 'doc-center-app-page';

export default class DocCenterWebPart extends BaseClientSideWebPart<IDocCenterWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      document.body.classList.add(DOC_CENTER_APP_BODY_CLASS);
      SharePointService.configure(this.context);
    });
  }

  public render(): void {
    const element: React.ReactElement<IDocCenterProps> = React.createElement(DocCenter, {
      documentsLibraryTitle: this.properties.documentsLibraryTitle || 'Documents',
      hashtagsListTitle: this.properties.hashtagsListTitle || 'Hashtags',
      currentUserLogin: this.context.pageContext.user.loginName,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    document.body.classList.remove(DOC_CENTER_APP_BODY_CLASS);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('documentsLibraryTitle', { label: strings.DocumentsLibraryFieldLabel }),
            PropertyPaneTextField('hashtagsListTitle', { label: strings.HashtagsListFieldLabel })
          ]
        }]
      }]
    };
  }
}
