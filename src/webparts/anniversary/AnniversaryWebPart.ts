import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'AnniversaryWebPartStrings';
import Anniversary from './components/Anniversary';
import { IAnniversaryProps } from './components/IAnniversaryProps';

import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';



import { MSGraphClient } from '@microsoft/sp-http';
export interface IAnniversaryWebPartProps {
  title: string;
  NumberAnniversary: number;
}

export default class AnniversaryWebPart extends BaseClientSideWebPart<IAnniversaryWebPartProps> {
  private graphCLient: MSGraphClient;


  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      // other init code may be present
    });
  }

  public render(): void {
    const element: React.ReactElement<IAnniversaryProps> = React.createElement(
      Anniversary,
      {
        title: this.properties.title,
        NumberAnniversary: this.properties.NumberAnniversary,
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
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
                PropertyPaneTextField('title', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldNumber("NumberAnniversary", {
                  key: "NumberAnniversary",
                  label: strings.NumberUpComingDaysLabel,
                  description: strings.NumberUpComingDaysLabel,
                  value: this.properties.NumberAnniversary,
                  maxValue: 100,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
