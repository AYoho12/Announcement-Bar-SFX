import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import styles from './AnnouncementBannerWebPart.module.scss';

export interface IAnnouncementBannerWebPartProps {
  alertTitle: string;
  alertDesc: string;
  colorChoice: string;
  scrollSpeed: number;
  selectedList: string;
  selectedItem: string;
  useListContent: boolean;

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

  public async render(): Promise<void> {
    const selectedColor = this.properties.colorChoice;
    const { lightcolor, darkcolor, fontcolor } = colors[selectedColor] || { lightcolor: 'transparent', darkcolor: 'transparent' };
    const scrollSpeed = this.properties.scrollSpeed || 10;
  
    let title = '';
    let description = '';
  
    if (this.properties.useListContent && this.properties.selectedItem && this.properties.selectedList) {
      const item = await sp.web.lists.getById(this.properties.selectedList).items.getById(parseInt(this.properties.selectedItem)).select("Title", "Description").get();
      title = item.Title;
      description = item.Description;
    } else {
      title = this.properties.alertTitle;
      description = this.properties.alertDesc;
    }
  
    // Check if the alert banner should be displayed
    const shouldDisplayBanner = title || description;
  
    this.domElement.innerHTML = shouldDisplayBanner ? `
      <section class="${styles.announcementBanner} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.container}">
          <div class="${styles.alertMessage}" style="background-color: ${escape(darkcolor)}; color: ${escape(fontcolor)};">${escape(title)}</div>
          <div class="${styles.marquee}" style="background-color: ${escape(lightcolor)}; color: ${escape(fontcolor)}; --marquee-speed: ${scrollSpeed}s;">
            <span>${escape(description)}</span>
          </div>
        </div>
      </section>` : '';
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({ spfxContext: this.context as any });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected itemOptions: { key: string | number, text: string }[] = [];

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'selectedList' && newValue) {
      const items = await sp.web.lists.getById(newValue).items.select("Title", "Id").get();
      this.itemOptions = items.map(item => ({ key: item.Id, text: item.Title }));

      // Reset the selectedItem when the selectedList changes
      this.properties.selectedItem = '';
  
      // Refresh the property pane to update the items dropdown
      this.context.propertyPane.refresh();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    await this.render();
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
  const lists = await sp.web.lists.filter("Hidden eq false").select("Title", "Id").get();
  this.listOptions = lists.map(list => ({ key: list.Id, text: list.Title }));

  if (this.properties.selectedList) {
    const items = await sp.web.lists.getById(this.properties.selectedList).items.select("Title", "Id").get();
    this.itemOptions = items.map(item => ({ key: item.Id, text: item.Title }));
  }

  this.context.propertyPane.refresh();
}

protected listOptions: { key: string | number, text: string }[] = [];

protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  const commonFields = [
    PropertyPaneToggle('useListContent', {
      label: 'Use Content from List',
      onText: 'On',
      offText: 'Off'
    })
  ];

  const manualFields = [
    PropertyPaneTextField('alertTitle', {
      label: 'Alert Title'
    }),
    PropertyPaneTextField('alertDesc', {
      label: 'Enter a Description for your alert',
      multiline: true
    }),
    PropertyPaneDropdown('colorChoice', {
      label: 'Select a Color',
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
      ]
    }),
    PropertyPaneSlider('scrollSpeed', {
      label: 'Scroll Speed',
      min: 1,
      max: 20,
      value: 10,
      showValue: true
    })
  ];

  const listFields = [
    PropertyPaneDropdown('selectedList', {
      label: 'Select a List',
      options: this.listOptions
    }),
    PropertyPaneDropdown('selectedItem', {
      label: 'Select an Item',
      options: this.itemOptions
    }),
    PropertyPaneDropdown('colorChoice', {
      label: 'Select a Color',
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
      ]
    }),
    PropertyPaneSlider('scrollSpeed', {
      label: 'Scroll Speed',
      min: 1,
      max: 20,
      value: 10,
      showValue: true
    })
  ];

  return {
    pages: [
      {
        groups: [
          {
            groupFields: [
              ...commonFields,
              ...(this.properties.useListContent ? listFields : manualFields)
            ]
          }
        ]
      }
    ]
  };
}

}
