import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

// Register the Fluent UI icon font. SPFx sometimes renders web parts before
// the host page has initialized icons, which makes <Icon iconName="..." />
// throw on first render. Calling this once per web part class is idempotent.
import { initializeIcons } from '@fluentui/react/lib/Icons';
initializeIcons();

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'FaqAccordionWebPartStrings';
import FaqAccordion from './components/FaqAccordion';
import { IFaqAccordionProps } from './components/IFaqAccordionProps';
import { ErrorBoundary } from './components/ErrorBoundary';
import { SPFaqService } from './services/SPFaqService';

// =============================================================================
// Style presets — governance-controlled
// =============================================================================
// Each preset defines the full set of CSS overrides for category and question
// headers. End users pick a preset from a dropdown; the individual style
// properties are not exposed. To add or change a preset, edit this list,
// bump the package version, and redeploy.
//
// An empty string for any property means "fall back to the SharePoint theme"
// (the SCSS variables have theme-aware fallbacks).
// =============================================================================
interface IStylePreset {
  key: string;
  label: string;
  categoryBgColor: string;
  categoryTextColor: string;
  categoryFont: string;
  categoryFontWeight: string;
  questionBgColor: string;
  questionTextColor: string;
  questionFont: string;
  questionFontWeight: string;
}

const STYLE_PRESETS: IStylePreset[] = [
  {
    key: 'grey-blue',
    label: 'Grey and Blue',
    categoryBgColor: '#d3d3d3',
    categoryTextColor: '#007a96',
    categoryFont: 'Arial, sans-serif',
    categoryFontWeight: '300',
    questionBgColor: '#ffffff',
    questionTextColor: '#007a96',
    questionFont: 'Arial, sans-serif',
    questionFontWeight: '500'
  },
  {
    key: 'blue-white',
    label: 'White on Blue',
    categoryBgColor: '#007a96',
    categoryTextColor: '#ffffff',
    categoryFont: 'Arial, sans-serif',
    categoryFontWeight: '700',
    questionBgColor: '#ffffff',
    questionTextColor: '#007a96',
    questionFont: 'Arial, sans-serif',
    questionFontWeight: '500'
  }
];

function getPreset(key: string): IStylePreset {
  for (let i = 0; i < STYLE_PRESETS.length; i++) {
    if (STYLE_PRESETS[i].key === key) return STYLE_PRESETS[i];
  }
  return STYLE_PRESETS[0]; // default if missing or unknown
}

export interface IFaqAccordionWebPartProps {
  title: string;
  anchorId: string;  // optional HTML id for jump-to-section links (#anchorId)
  siteUrl: string;   // absolute URL of the site that contains the list; blank = current site
  listId: string;

  // Filtering — by Set (lookup column auto-detected from the FAQ list schema)
  setId: string;      // ID of the chosen Set; '' = no filter

  // Styling — chosen from governance-defined presets (STYLE_PRESETS above).
  stylePreset: string;
}

export default class FaqAccordionWebPart extends BaseClientSideWebPart<IFaqAccordionWebPartProps> {

  private _isDarkTheme: boolean = false;
  // Populated lazily when the property pane opens. The first entry is the
  // "no filter" placeholder so the dropdown always has something to show.
  private _setOptions: IPropertyPaneDropdownOption[] = [
    { key: '', text: 'All sets (no filter)' }
  ];
  // Discovered automatically by reading the "Set" lookup column's schema.
  // Cached here so we don't refetch on every property-pane refresh.
  private _detectedSetsListId: string = '';
  // Internal name of the lookup column on the FAQ list.
  private static readonly SET_COLUMN = 'Set';

  public render(): void {
    // Resolve the selected preset to the eight CSS variable values. If the
    // saved preset key is missing or unknown, fall back to the first preset.
    const preset = getPreset(this.properties.stylePreset);

    const inner: React.ReactElement<IFaqAccordionProps> = React.createElement(
      FaqAccordion,
      {
        title: this.properties.title,
        anchorId: this.properties.anchorId,
        siteUrl: this.properties.siteUrl,
        listId: this.properties.listId,
        setId: this.properties.setId,
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        displayMode: this.displayMode,
        updateTitle: (value: string) => {
          this.properties.title = value;
        },
        styles: {
          categoryBgColor: preset.categoryBgColor,
          categoryTextColor: preset.categoryTextColor,
          categoryFont: preset.categoryFont,
          categoryFontWeight: preset.categoryFontWeight,
          questionBgColor: preset.questionBgColor,
          questionTextColor: preset.questionTextColor,
          questionFont: preset.questionFont,
          questionFontWeight: preset.questionFontWeight
        }
      }
    );

    const wrapped = React.createElement(ErrorBoundary, null, inner);
    ReactDom.render(wrapped, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  // When the property pane opens, load the Set choices from the configured
  // FAQ Sets list so the "Filter by Set" dropdown has values to show.
  protected onPropertyPaneConfigurationStart(): void {
    this._loadSetOptions().then(
      () => { this.context.propertyPane.refresh(); },
      () => { /* errors already logged inside _loadSetOptions */ }
    );
  }

  // Clear dependent properties when their source changes, and reload the
  // Set options whenever the FAQ list (or site URL) is updated. The Sets list
  // is auto-detected from the lookup column's schema.
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'siteUrl' && oldValue !== newValue) {
      this.properties.listId = '';
      this.properties.setId = '';
      this._detectedSetsListId = '';
      this._setOptions = [{ key: '', text: 'All sets (no filter)' }];
      this.context.propertyPane.refresh();
    }
    if (propertyPath === 'listId' && oldValue !== newValue) {
      // FAQ list changed — re-detect the Sets list it points to and reload options.
      this.properties.setId = '';
      this._detectedSetsListId = '';
      this._loadSetOptions().then(
        () => { this.context.propertyPane.refresh(); },
        () => { /* errors already logged inside _loadSetOptions */ }
      );
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private async _loadSetOptions(): Promise<void> {
    this._setOptions = [{ key: '', text: 'All sets (no filter)' }];
    const faqListId = this.properties.listId;
    if (!faqListId) {
      return;
    }
    try {
      const service = new SPFaqService(this.context);
      // 1) Find which list the "Set" column points to (auto-detect from schema).
      if (!this._detectedSetsListId) {
        this._detectedSetsListId = await service.getLookupTargetListId(
          faqListId,
          FaqAccordionWebPart.SET_COLUMN,
          this.properties.siteUrl
        );
      }
      if (!this._detectedSetsListId) {
        // No "Set" column on this list — leave just the "no filter" entry.
        return;
      }
      // 2) Load the rows from the detected Sets list.
      const items = await service.getListChoices(this._detectedSetsListId, this.properties.siteUrl);
      for (let i = 0; i < items.length; i++) {
        this._setOptions.push({
          key: String(items[i].id),
          text: items[i].title || `(no title — ID ${items[i].id})`
        });
      }
    } catch (e) {
      // Leave the placeholder option in place. The user will see the dropdown
      // with only "All sets (no filter)" available, which signals the load failed.
      // eslint-disable-next-line no-console
      console.error('Could not load Set options:', e);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('anchorId', {
                  label: 'Anchor ID (optional)',
                  placeholder: 'e.g. account-help',
                  description: 'Use letters, numbers, and dashes only. Link to this web part from elsewhere on the page with #yourAnchorId.'
                }),
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL (optional)',
                  placeholder: 'https://contoso.sharepoint.com/sites/another-site',
                  description: 'Leave blank to use this site. Paste a full site URL to read a list from another site or site collection.'
                }),
                PropertyFieldListPicker('listId', {
                  label: strings.ListFieldLabel,
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  // Point the picker at the chosen site (defaults to current).
                  // Key includes the URL so the picker remounts when it changes.
                  webAbsoluteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  key: 'listPickerFieldId-' + (this.properties.siteUrl || 'current')
                })
              ]
            },
            {
              groupName: 'Filtering',
              isCollapsed: true,
              groupFields: [
                PropertyPaneDropdown('setId', {
                  label: 'Filter by Set',
                  options: this._setOptions,
                  selectedKey: this.properties.setId || '',
                  // The dropdown stays usable once a FAQ list is picked. If
                  // the list has no "Set" column, only the "no filter" entry
                  // appears, which is the right behaviour.
                  disabled: !this.properties.listId
                })
              ]
            },
            {
              groupName: 'Style',
              isCollapsed: true,
              groupFields: [
                PropertyPaneDropdown('stylePreset', {
                  label: 'Style preset',
                  options: STYLE_PRESETS.map(p => ({ key: p.key, text: p.label })),
                  selectedKey: this.properties.stylePreset || STYLE_PRESETS[0].key
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
