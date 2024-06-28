import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnnouncementBannerWebPart.module.scss';
import * as strings from 'AnnouncementBannerWebPartStrings';

export interface IAnnouncementBannerWebPartProps {
  alertTitle: string;
  alertDesc: string;
  colorChoice: string;
  scrollSpeed: number;

  test1: boolean;
  test3: boolean;
}

const colors: Record<string, { lightcolor: string; darkcolor: string; fontcolor: string }> = {
  // red
  darkred: { lightcolor: '#E17B80', darkcolor: 'darkred', fontcolor: 'white'},
  // green
  '#3B7D23': { lightcolor: '#8ED973', darkcolor: '#3B7D23', fontcolor: 'white' },
  // yellow
  '#E2DD00': { lightcolor: '#FFFF81', darkcolor: '#E2DD00',fontcolor: 'black' },
  // blue
  '#0070C0': { lightcolor: '#4BB2FF', darkcolor: '#0070C0', fontcolor: 'white' },
  // orange
  '#FF9933': { lightcolor: '#FFCD9B', darkcolor: '#FF9933', fontcolor: 'white' },
  // purple
  '#7030A0': { lightcolor: '#BF95DF', darkcolor: '#7030A0', fontcolor: 'white' },
  // pink
  '#D86ECC': { lightcolor: '#F1CBEC', darkcolor: '#D86ECC', fontcolor: 'white' },
  // black
  '#000000': { lightcolor: '#828282', darkcolor: '#000000', fontcolor: 'white' },
  // grey
  '#686868': { lightcolor: '#B0B0B0', darkcolor: '#686868', fontcolor: 'white' }
};

export default class AnnouncementBannerWebPart extends BaseClientSideWebPart<IAnnouncementBannerWebPartProps> {


  public render(): void {
    const selectedColor = this.properties.colorChoice;
    const { lightcolor, darkcolor, fontcolor } = colors[selectedColor] || { lightcolor: 'transparent', darkcolor: 'transparent' };
    const scrollSpeed = this.properties.scrollSpeed || 10;

    this.domElement.innerHTML = `
      <section class="${styles.announcementBanner} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.container}">
          <div class="${styles['dark-red']}" style="background-color: ${escape(darkcolor)}; color: ${escape(fontcolor)};">${escape(this.properties.alertTitle)}</div>
          <div class="${styles.marquee}" style="background-color: ${escape(lightcolor)}; color: ${escape(fontcolor)}; --marquee-speed: ${scrollSpeed}s;">
            <span>${escape(this.properties.alertDesc)}</span>
          </div>
        </div>
      </section>`;
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
              PropertyPaneTextField('alertTitle', {
                label: 'Alert Title'
              }),
              PropertyPaneTextField('alertDesc', {
                label: 'Enter a Description for your alert',
                multiline: true
              }),
              PropertyPaneDropdown('colorChoice', {
                label: 'Dropdown',
                options: [
                  { key: 'darkred', text: 'Red' },
                  { key: '#FF9933', text: 'Orange' },
                  { key: '#E2DD00', text: 'Yellow' },
                  { key: '#3B7D23', text: 'Green' },
                  { key: '#0070C0', text: 'Blue' },
                  { key: '#7030A0', text: 'Purple' },
                  { key: '#D86ECC', text: 'Pink' },
                  { key: '#000000', text: 'Black' },
                  { key: '#686868', text: 'Grey' }
                ]}),
              PropertyPaneSlider('scrollSpeed', {
                label: 'Scroll Speed',
                min: 1,
                max: 20,
                value: 10,
                showValue: true
              }),
              PropertyPaneCheckbox('test1', {
                text: 'Checkbox'
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
}
