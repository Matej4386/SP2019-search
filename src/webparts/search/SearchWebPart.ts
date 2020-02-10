import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as $ from "jquery";
import { Version, Text } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  IPropertyPaneGroup,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneHorizontalRule,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import { PageOpenBehavior, QueryPathBehavior } from '../../helpers/UrlHelper';
import * as strings from 'SearchResultsWebPartStrings';
import SearchResultsContainer from './components/SearchResultsContainer/SearchResultsContainer';
import { ISearchWebPartProps } from './ISearchWebPartProps';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import ISearchService from '../../services/SearchService/ISearchService';
import ITaxonomyService from '../../services/TaxonomyService/ITaxonomyService';
import ResultsLayoutOption from '../../models/ResultsLayoutOption';
import { TemplateService } from '../../services/TemplateService/TemplateService';
import { isEmpty, find, sortBy, cloneDeep } from '@microsoft/sp-lodash-subset';
import SearchService from '../../services/SearchService/SearchService';
import TaxonomyService from '../../services/TaxonomyService/TaxonomyService';
import ISearchResultsContainerProps from './components/SearchResultsContainer/ISearchResultsContainerProps';
import { SortDirection, Sort } from '@pnp/sp';
import { ISortFieldConfiguration, ISortFieldDirection } from '../../models/ISortFieldConfiguration';
import { ISynonymFieldConfiguration } from '../../models/ISynonymFieldConfiguration';
import { ResultTypeOperator } from '../../models/ISearchResultType';
import IResultService from '../../services/ResultService/IResultService';
import { ResultService, IRenderer } from '../../services/ResultService/ResultService';
import RefinerTemplateOption from '../../models/RefinerTemplateOption';
import RefinersSortOption from '../../models/RefinersSortOptions';
import { SearchComponentType } from '../../models/SearchComponentType';
import ISearchResultSourceData from '../../models/ISearchResultSourceData';
import RefinersLayoutOption from '../../models/RefinersLayoutOptions';
import ISynonymTable from '../../models/ISynonym';
import * as update from 'immutability-helper';
import LocalizationHelper from '../../helpers/LocalizationHelper';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { SearchManagedProperties, ISearchManagedPropertiesProps } from '../../controls/SearchManagedProperties/SearchManagedProperties';
import { PropertyPaneSearchManagedProperties } from '../../controls/PropertyPaneSearchManagedProperties/PropertyPaneSearchManagedProperties';
import { loadTheme } from 'office-ui-fabric-react';

loadTheme({
    palette: {
      themePrimary: '#17375e',
      themeLighterAlt: '#f1f4f9',
      themeLighter: '#c9d6e5',
      themeLight: '#a0b5cf',
      themeTertiary: '#56779f',
      themeSecondary: '#264872',
      themeDarkAlt: '#143155',
      themeDark: '#112a48',
      themeDarker: '#0d1f35',
      neutralLighterAlt: '#f8f8f8',
      neutralLighter: '#f4f4f4',
      neutralLight: '#eaeaea',
      neutralQuaternaryAlt: '#dadada',
      neutralQuaternary: '#d0d0d0',
      neutralTertiaryAlt: '#c8c8c8',
      neutralTertiary: '#a0b5cf',
      neutralSecondary: '#56779f',
      neutralPrimaryAlt: '#264872',
      neutralPrimary: '#17375e',
      neutralDark: '#112a48',
      black: '#0d1f35',
      white: '#ffffff',
    }
  });

export default class SearchWebPart extends BaseClientSideWebPart<ISearchWebPartProps> {
  private _searchService: ISearchService;
  private _taxonomyService: ITaxonomyService;
  private _templateService: BaseTemplateService;
  private _textDialogComponent = null;
  private _propertyFieldCodeEditor = null;
  private _propertyFieldCollectionData = null;
  private _customCollectionFieldType = null;
  private _propertyFieldCodeEditorLanguages = null;
  private _resultService: IResultService;
  private _codeRenderers: IRenderer[];
  private _searchContainer: JSX.Element;
  private _synonymTable: ISynonymTable;
  /**
   * Available property pane options from Web Components
   */
  private _templatePropertyPaneOptions: IPropertyPaneField<any>[];
  private _availableLanguages: IPropertyPaneDropdownOption[];
  /**
   * The template to display at render time
   */
  private _templateContentToDisplay: string;
  /**
   * The list of available managed managed properties (managed globally for all property pane fiels if needed)
   */
  private _availableManagedProperties: IComboBoxOption[];

  public constructor() {
    super();
    this._templateContentToDisplay = '';
    this._availableLanguages = [];
    this._templatePropertyPaneOptions = [];
    this._availableManagedProperties = [];

    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
    this._onUpdateAvailableProperties = this._onUpdateAvailableProperties.bind(this);
  }
  public async render(): Promise<void> {
    // Determine the template content to display
    // In the case of an external template is selected, the render is done asynchronously waiting for the content to be fetched
    await this._initTemplate();

    this.renderCompleted();
  }
  protected get disableReactivePropertyChanges(): boolean {
    // Set this to true if you don't want the reactive behavior.
    return false;
  }

  protected get isRenderAsync(): boolean {
      return true;
  }

  protected renderCompleted(): void {
    super.renderCompleted();
    let renderElement = null;
    let queryTemplate: string = this.properties.queryTemplate;
    let sourceId: string = this.properties.resultSourceId;

    let queryDataSourceValue = null
    if (typeof (queryDataSourceValue) !== 'string') {
        queryDataSourceValue = '';
        this.context.propertyPane.refresh();
    }

    let queryKeywords = (!queryDataSourceValue) ? this.properties.defaultSearchQuery : queryDataSourceValue;

    const currentLocaleId = LocalizationHelper.getLocaleId(this.context.pageContext.cultureInfo.currentCultureName);

    // Configure the provider before the query according to our needs
    this._searchService = update(this._searchService, {
        resultsCount: { $set: this.properties.maxResultsCount },
        queryTemplate: { $set: queryTemplate },
        resultSourceId: { $set: sourceId },
        sortList: { $set: this._convertToSortList(this.properties.sortList) },
        enableQueryRules: { $set: this.properties.enableQueryRules },
        selectedProperties: { $set: this.properties.selectedProperties ? this.properties.selectedProperties.replace(/\s|,+$/g, '').split(',') : [] },
        synonymTable: { $set: this._synonymTable },
        queryCulture: { $set: this.properties.searchQueryLanguage !== -1 ? this.properties.searchQueryLanguage : currentLocaleId },
        refinementFilters: { $set: [] },
        refiners: { $set: this.properties.refinersConfiguration }
    });

    this._searchContainer = React.createElement(
        SearchResultsContainer,
        {
            searchService: this._searchService,
            taxonomyService: this._taxonomyService,
            queryKeywords: queryKeywords,
            sortableFields: this.properties.sortableFields,
            showSearchBox: this.properties.showSearchBox,
            showPaging: this.properties.showPaging,
            showRefinements: this.properties.showRefinements,
            showResultsCount: this.properties.showResultsCount,
            showBlank: this.properties.showBlank,
            displayMode: this.displayMode,
            templateService: this._templateService,
            templateContent: this._templateContentToDisplay,
            templateParameters: this.properties.templateParameters,
            webPartTitle: this.properties.webPartTitle,
            currentUICultureName: this.context.pageContext.cultureInfo.currentUICultureName,
            siteServerRelativeUrl: this.context.pageContext.site.serverRelativeUrl,
            webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
            resultTypes: this.properties.resultTypes,
            useCodeRenderer: this.codeRendererIsSelected(),
            customTemplateFieldValues: this.properties.customTemplateFieldValues,
            rendererId: this.properties.selectedLayout as any,
            enableLocalization: this.properties.enableLocalization,
            refinersConfiguration: this.properties.refinersConfiguration,
            selectedLayoutRefiners: this.properties.selectedLayoutRefiners,
            searchInNewPage: this.properties.searchInNewPage,
            pageUrl: this.properties.pageUrl,
            queryPathBehavior: this.properties.queryPathBehavior,
            queryStringParameter: this.properties.queryStringParameter,
            openBehavior: this.properties.openBehavior,
            enableQuerySuggestions: this.properties.enableQuerySuggestions,
            placeholderText: this.properties.placeholderText,
            domElement: this.domElement,
            onSearchResultsUpdate: async (results, mountingNodeId, searchService) => {
                this._resultService.updateResultData(results, this.properties.selectedLayout as any, mountingNodeId, this.properties.customTemplateFieldValues);
            }
        } as ISearchResultsContainerProps
    );

    renderElement = this._searchContainer;

    ReactDom.render(renderElement, this.domElement);
  }
  protected async onInit(): Promise<void> {
    
    $("#workbenchPageContent").prop("style", "max-width: none"); 
    $(".SPCanvas-canvas").prop("style", "max-width: none");
    $(".CanvasZone").prop("style", "max-width: none");

    this.initializeRequiredProperties();

    
    this._taxonomyService = new TaxonomyService(this.context.pageContext.site.absoluteUrl);

    let timeZoneBias = {
        WebBias: this.context.pageContext.legacyPageContext.webTimeZoneData.Bias,
        WebDST: this.context.pageContext.legacyPageContext.webTimeZoneData.DaylightBias,
        UserBias: null,
        UserDST: null
    };
    if (this.context.pageContext.legacyPageContext.userTimeZoneData) {
        timeZoneBias.UserBias = this.context.pageContext.legacyPageContext.userTimeZoneData.Bias;
        timeZoneBias.UserDST = this.context.pageContext.legacyPageContext.userTimeZoneData.DaylightBias;
    }

    this._searchService = new SearchService(this.context.pageContext, this.context.spHttpClient);
    this._templateService = new TemplateService(this.context.spHttpClient, this.context.pageContext.cultureInfo.currentUICultureName, this._searchService, timeZoneBias, this.context);
    
    this._resultService = new ResultService();
    this._codeRenderers = this._resultService.getRegisteredRenderers();
    // Set the default search results layout
    this.properties.selectedLayout = this.properties.selectedLayout ? this.properties.selectedLayout : ResultsLayoutOption.DetailsList;
    this._synonymTable = this._convertToSynonymTable(this.properties.synonymList);

    return super.onInit();
  }

  private _convertToSortConfig(sortList: string): ISortFieldConfiguration[] {
    let pairs = sortList.split(',');
    return pairs.map(sort => {
        let direction;
        let kvp = sort.split(':');
        if (kvp[1].toLocaleLowerCase().trim() === "ascending") {
            direction = ISortFieldDirection.Ascending;
        } else {
            direction = ISortFieldDirection.Descending;
        }

        return {
            sortField: kvp[0].trim(),
            sortDirection: direction
        } as ISortFieldConfiguration;
    });
  }

  private _convertToSynonymTable(synonymList: ISynonymFieldConfiguration[]): ISynonymTable {
    let synonymsTable: ISynonymTable = {};

    if (synonymList) {
        synonymList.forEach(item => {
            const currentTerm = item.Term.toLowerCase();
            const currentSynonyms = this._splitSynonyms(item.Synonyms);

            //add to array
            synonymsTable[currentTerm] = currentSynonyms;

            if (item.TwoWays) {
                // Loop over the list of synonyms
                let tempSynonyms: string[] = currentSynonyms;
                tempSynonyms.push(currentTerm.trim());

                currentSynonyms.forEach(s => {
                    synonymsTable[s.toLowerCase().trim()] = tempSynonyms.filter(f => { return f !== s; });
                });
            }
        });
    }
    return synonymsTable;
  }

  private _splitSynonyms(value: string) {
    return value.split(",").map(v => { return v.toLowerCase().trim().replace(/\"/g, ""); });
  }

  private _convertToSortList(sortList: ISortFieldConfiguration[]): Sort[] {
    return sortList.map(e => {

        let direction;

        switch (e.sortDirection) {
            case ISortFieldDirection.Ascending:
                direction = SortDirection.Ascending;
                break;

            case ISortFieldDirection.Descending:
                direction = SortDirection.Descending;
                break;

            default:
                direction = SortDirection.Ascending;
                break;
        }

        return {
            Property: e.sortField,
            Direction: direction
        } as Sort;
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Initializes the Web Part required properties if there are not present in the manifest (i.e. during an update scenario)
   */
  private initializeRequiredProperties() {

    this.properties.queryTemplate = this.properties.queryTemplate ? this.properties.queryTemplate : "{searchTerms} Path:{Site}";

    if (!Array.isArray(this.properties.sortList) && !isEmpty(this.properties.sortList)) {
        this.properties.sortList = this._convertToSortConfig(this.properties.sortList);
    }

    this.properties.sortList = Array.isArray(this.properties.sortList) ? this.properties.sortList : [
        {
            sortField: "Created",
            sortDirection: ISortFieldDirection.Ascending
        },
        {
            sortField: "Size",
            sortDirection: ISortFieldDirection.Descending
        }
    ];

    this.properties.sortableFields = Array.isArray(this.properties.sortableFields) ? this.properties.sortableFields : [];
    this.properties.selectedProperties = this.properties.selectedProperties ? this.properties.selectedProperties : "Title,Path,Created,Filename,SiteLogo,PreviewUrl,PictureThumbnailURL,ServerRedirectedPreviewURL,ServerRedirectedURL,HitHighlightedSummary,FileType,contentclass,ServerRedirectedEmbedURL,DefaultEncodingURL,owstaxidmetadataalltagsinfo";
    this.properties.maxResultsCount = this.properties.maxResultsCount ? this.properties.maxResultsCount : 10;
    this.properties.resultTypes = Array.isArray(this.properties.resultTypes) ? this.properties.resultTypes : [];
    this.properties.synonymList = Array.isArray(this.properties.synonymList) ? this.properties.synonymList : [];
    this.properties.searchQueryLanguage = this.properties.searchQueryLanguage ? this.properties.searchQueryLanguage : -1;
    this.properties.templateParameters = this.properties.templateParameters ? this.properties.templateParameters : {}; 
    /**
    * Refiners
    */
   if (<any>this.properties.refinersConfiguration === "") {
    this.properties.refinersConfiguration = [];
    }

    if (Array.isArray(this.properties.refinersConfiguration)) {

        this.properties.refinersConfiguration = this.properties.refinersConfiguration.map(config => {
        if (!config.template) {
            config.template = RefinerTemplateOption.CheckBox;
        }
        if (!config.refinerSortType) {
            config.refinerSortType = RefinersSortOption.ByNumberOfResults;
        }

        return config;
        });

    } else {
        // Default setup
        this.properties.refinersConfiguration = [
        {
            refinerName: "Created",
            displayValue: "Created Date",
            template: RefinerTemplateOption.CheckBox,
            refinerSortType: RefinersSortOption.ByNumberOfResults,
            showExpanded: false
        },
        {
            refinerName: "Size",
            displayValue: "Size of the file",
            template: RefinerTemplateOption.CheckBox,
            refinerSortType: RefinersSortOption.ByNumberOfResults,
            showExpanded: false
        },
        {
            refinerName: "owstaxidmetadataalltagsinfo",
            displayValue: "Tags",
            template: RefinerTemplateOption.CheckBox,
            refinerSortType: RefinersSortOption.ByNumberOfResults,
            showExpanded: false
        }
        ];
    }

    this.properties.selectedLayoutRefiners = this.properties.selectedLayoutRefiners ? this.properties.selectedLayoutRefiners : RefinersLayoutOption.Vertical;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    const templateParametersGroup = this._getTemplateFieldsGroup();

    let stylingPageGroups: IPropertyPaneGroup[] = [
        {
            groupName: strings.StylingSettingsGroupName,
            groupFields: this._getStylingFields(),
            isCollapsed: false
        },                        
    ];

    if (templateParametersGroup) {
        stylingPageGroups.push(templateParametersGroup);
    }

    return {
        pages: [
            {                   
                groups: [
                    {
                        groupFields: this._getSearchQueryFields(),
                        isCollapsed: false,
                        groupName: strings.SearchQuerySettingsGroupName
                    },
                    {
                        groupFields: this._getSearchSettingsFields(),
                        isCollapsed: false,
                        groupName: strings.SearchSettingsGroupName
                    }
                ],
                displayGroupsAsAccordion: true
            },
            {
                groups: stylingPageGroups,
                displayGroupsAsAccordion: true
            },
            {
                groups: [
                    {
                        groupName: strings.SearchBoxNewPage,
                        groupFields: this._getSearchBehaviorOptionsFields()
                    },
                ],
                displayGroupsAsAccordion: true
            }, 
            {
                groups: [
                    {
                        groupName: strings.RefinersConfigurationGroupName,
                        groupFields: this._getRefinerSettings()
                    },
                    {
                        groupName: strings.StylingSettingsGroupName,
                        groupFields: this._getStylingFieldsRefiners()
                    }
                ],
                displayGroupsAsAccordion: true
            }
        ]
    };
  }
  /**
   * Determines the group fields for the search options inside the property pane
   */
  private _getSearchBehaviorOptionsFields(): IPropertyPaneField<any>[] {

    let searchBehaviorOptionsFields: IPropertyPaneField<any>[]  = [
        PropertyPaneToggle("showSearchBox", {
            checked: true,
            label: strings.SearchBoxEnable
          })
    ];
    if (this.properties.showSearchBox) {
        searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
            PropertyPaneToggle("enableQuerySuggestions", {
                checked: false,
                label: strings.SearchBoxEnableQuerySuggestions
            }),
            PropertyPaneHorizontalRule(),
            PropertyPaneTextField('placeholderText', {
                label: strings.SearchBoxPlaceholderTextLabel
            }),
            PropertyPaneHorizontalRule(),
            PropertyPaneToggle("searchInNewPage", {
                checked: false,
                label: strings.SearchBoxSearchInNewPageLabel
            })
        ]);
    }
    if ((this.properties.showSearchBox)&&(this.properties.searchInNewPage)) {
      searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
        PropertyPaneTextField('pageUrl', {
          disabled: !this.properties.searchInNewPage,
          label: strings.SearchBoxPageUrlLabel,
          onGetErrorMessage: this._validatePageUrl.bind(this)
        }),
        PropertyPaneDropdown('openBehavior', {
          label:  strings.SearchBoxPageOpenBehaviorLabel,
          options: [
            { key: PageOpenBehavior.Self, text: strings.SearchBoxSameTabOpenBehavior },
            { key: PageOpenBehavior.NewTab, text: strings.SearchBoxNewTabOpenBehavior }
          ],
          disabled: !this.properties.searchInNewPage,
          selectedKey: this.properties.openBehavior
        }),
        PropertyPaneDropdown('queryPathBehavior', {
          label:  strings.SearchBoxQueryPathBehaviorLabel,
          options: [
            { key: QueryPathBehavior.URLFragment, text: strings.SearchBoxUrlFragmentQueryPathBehavior },
            { key: QueryPathBehavior.QueryParameter, text: strings.SearchBoxQueryStringQueryPathBehavior }
          ],
          disabled: !this.properties.searchInNewPage,
          selectedKey: this.properties.queryPathBehavior
        })
      ]);
    }

    if (this.properties.searchInNewPage && this.properties.queryPathBehavior === QueryPathBehavior.QueryParameter) {
      searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
        PropertyPaneTextField('queryStringParameter', {
          disabled: !this.properties.searchInNewPage || this.properties.searchInNewPage && this.properties.queryPathBehavior !== QueryPathBehavior.QueryParameter,
          label: strings.SearchBoxQueryStringParameterName,
          onGetErrorMessage: (value) => {
            if (this.properties.queryPathBehavior === QueryPathBehavior.QueryParameter) {
              if (value === null ||
                value.trim().length === 0) {
                return strings.SearchBoxQueryParameterNotEmpty;
              }              
            }
            return '';
          }
        })
      ]);
    }

    return searchBehaviorOptionsFields;
  }
  /**
   * Determines the group fields for styling options inside the property pane
   */
  private _getStylingFieldsRefiners(): IPropertyPaneField<any>[] {

    // Options for the search results layout 
    const layoutOptions = [
      {
        iconProps: {
          officeFabricIconFontName: 'BulletedList2'
        },
        text: 'Vertical',
        key: RefinersLayoutOption.Vertical,
      },
      {
        iconProps: {
          officeFabricIconFontName: 'ClosePane'
        },
        text: 'Panel',
        key: RefinersLayoutOption.LinkAndPanel
      }
    ] as IPropertyPaneChoiceGroupOption[];

    // Sets up styling fields
    let stylingFields: IPropertyPaneField<any>[] = [
      PropertyPaneChoiceGroup('selectedLayoutRefiners', {
        label: strings.RefinerLayoutLabel,
        options: layoutOptions
      })
    ];

    return stylingFields;
  }
  /**
   * Determines the group fields for refiner settings
   */
  private _getRefinerSettings(): IPropertyPaneField<any>[] {

    const refinerSettings = [
        PropertyPaneToggle('showRefinements', {
            label: strings.UseRefinementWebPartLabel,
            checked: this.properties.showRefinements,
        }),
        this._propertyFieldCollectionData('refinersConfiguration', {
        manageBtnLabel: strings.Refiners.EditRefinersLabel,
        key: 'refiners',
        enableSorting: true,
        panelHeader: strings.Refiners.EditRefinersLabel,
        panelDescription: strings.Refiners.RefinersFieldDescription,
        label: strings.Refiners.RefinersFieldLabel,
        value: this.properties.refinersConfiguration,
        fields: [
          {
            id: 'refinerName',
            title: strings.Refiners.RefinerManagedPropertyField,
            type: this._customCollectionFieldType.custom,
            onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
              // Need to specify a React key to avoid item duplication when adding a new row
              return React.createElement("div", {key : `${field.id}-${itemId}`},
                  React.createElement(SearchManagedProperties, {
                  defaultSelectedKey: item[field.id] ? item[field.id] : '',
                  onUpdate: (newValue: any, isSortable: boolean) => { 
                    onUpdate(field.id, newValue);
                    onCustomFieldValidation(field.id, '');
                  },
                  searchService: this._searchService,
                  validateSortable: false,
                  availableProperties: this._availableManagedProperties,
                  onUpdateAvailableProperties: this._onUpdateAvailableProperties
              } as ISearchManagedPropertiesProps));
            } 
          },
          {
            id: 'displayValue',
            title: strings.Refiners.RefinerDisplayValueField,
            type: this._customCollectionFieldType.string
          },
          {
            id: 'template',
            title: "Refiner template",
            type: this._customCollectionFieldType.dropdown,
            options: [
              {
                key: RefinerTemplateOption.CheckBox,
                text: strings.Refiners.Templates.RefinementItemTemplateLabel
              },
              {
                key: RefinerTemplateOption.CheckBoxMulti,
                text: strings.Refiners.Templates.MutliValueRefinementItemTemplateLabel
              },
              {
                key: RefinerTemplateOption.DateRange,
                text: strings.Refiners.Templates.DateRangeRefinementItemLabel,
              }
            ]
          },
          {
            id: 'refinerSortType',
            title: strings.Refiners.Templates.RefinerSortTypeLabel,
            type: this._customCollectionFieldType.dropdown,
            options: [
              {
                key: RefinersSortOption.ByNumberOfResults,
                text: strings.Refiners.Templates.RefinerSortTypeByNumberOfResults
              },
              {
                key: RefinersSortOption.Alphabetical,
                text: strings.Refiners.Templates.RefinerSortTypeAlphabetical
              }
            ]
          },
          {
            id: 'showExpanded',
            title: strings.Refiners.ShowExpanded,
            type: this._customCollectionFieldType.boolean
          }
        ]
      }),
    ];

    return refinerSettings;
  }
  protected async loadPropertyPaneResources(): Promise<void> {

    // Code editor component for result types
    this._textDialogComponent = await import(
        /* webpackChunkName: 'search-property-pane' */
        '../../controls/TextDialog'
    );

    // tslint:disable-next-line:no-shadowed-variable
    const { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } = await import(
        /* webpackChunkName: 'search-property-pane' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );
    this._propertyFieldCodeEditor = PropertyFieldCodeEditor;
    this._propertyFieldCodeEditorLanguages = PropertyFieldCodeEditorLanguages;

    const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
        /* webpackChunkName: 'search-property-pane' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
    );
    this._propertyFieldCollectionData = PropertyFieldCollectionData;
    this._customCollectionFieldType = CustomCollectionFieldType;

    if (this._availableLanguages.length == 0) {
        const languages = await this._searchService.getAvailableQueryLanguages();

        this._availableLanguages = languages.map(language => {
            return {
                key: language.Lcid,
                text: `${language.DisplayName} (${language.Lcid})`
            };
        });
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string) {
    if (propertyPath.localeCompare('queryKeywords') === 0) {
        // Update data source information
        this._saveDataSourceInfo();
    }
    if (!this.properties.useDefaultSearchQuery) {
        this.properties.defaultSearchQuery = '';
    }
    if (this.properties.enableLocalization) {

        let udpatedProperties: string[] = this.properties.selectedProperties.split(',');
        if (udpatedProperties.indexOf('UniqueID') === -1) {
            udpatedProperties.push('UniqueID');
        }
        // Add automatically the UniqueID managed property for subsequent queries
        this.properties.selectedProperties = udpatedProperties.join(',');
    }

    if (propertyPath.localeCompare('selectedLayout') === 0) {
        // Refresh setting the right template for the property pane
        if (!this.codeRendererIsSelected()) {
            await this._initTemplate();
        }
        if (this.codeRendererIsSelected) {
            this.properties.customTemplateFieldValues = undefined;
        }

        this.context.propertyPane.refresh();
    }

    // Detect if the layout has been changed to custom...
    if (propertyPath.localeCompare('inlineTemplateText') === 0) {

        // Automatically switch the option to 'Custom' if a default template has been edited
        // (meaning the user started from a the list or tiles template)
        if (this.properties.inlineTemplateText && this.properties.selectedLayout !== ResultsLayoutOption.Custom) {
            this.properties.selectedLayout = ResultsLayoutOption.Custom;

            // Reset also the template URL
            this.properties.externalTemplateUrl = '';
        }
    }

    this._synonymTable = this._convertToSynonymTable(this.properties.synonymList);
  }

  protected async onPropertyPaneConfigurationStart() {
    await this.loadPropertyPaneResources();
  }

  /**
  * Save the useful information for the connected data source. 
  * They will be used to get the value of the dynamic property if this one fails.
  */
  private _saveDataSourceInfo() {
    this.properties.sourceId = null;
    this.properties.propertyId = null;
    this.properties.propertyPath = null;
  }

  /**
   * Checks if a field if empty or not
   * @param value the value to check
   */
  private _validateEmptyField(value: string): string {

    if (!value) {
        return strings.EmptyFieldErrorMessage;
    }

    return '';
  }

  /**
   * Ensures the result source id value is a valid GUID
   * @param value the result source id
   */
  private validateSourceId(value: string): string {
    if (value.length > 0) {
        if (!/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/.test(value)) {
            return strings.InvalidResultSourceIdMessage;
        }
    }

    return '';
  }

  /**
   * Init the template according to the property pane current configuration
   * @returns the template content as a string
   */
  private async _initTemplate(): Promise<void> {

    if (this.properties.selectedLayout === ResultsLayoutOption.Custom) {
        
        // Reset options
        this._templatePropertyPaneOptions = [];

        if (this.properties.externalTemplateUrl) {
            this._templateContentToDisplay = await this._templateService.getFileContent(this.properties.externalTemplateUrl);
        } else {
            this._templateContentToDisplay = this.properties.inlineTemplateText ? this.properties.inlineTemplateText : TemplateService.getTemplateContent(ResultsLayoutOption.Custom);
        }
    } else {

        // Builtin templates with options
        this._templateContentToDisplay = TemplateService.getTemplateContent(this.properties.selectedLayout);
        this._templatePropertyPaneOptions = this._templateService.getTemplateParameters(this.properties.selectedLayout, this.properties, this._onUpdateAvailableProperties, this._availableManagedProperties);
    }

    // Register result types inside the template      
    this._templateService.registerResultTypes(this.properties.resultTypes);
  }

  /**
   * Custom handler when the external template file URL
   * @param value the template file URL value
   */
  private async _onTemplateUrlChange(value: string): Promise<String> {

    try {
        // Doesn't raise any error if file is empty (otherwise error message will show on initial load...)
        if (isEmpty(value)) {
            return '';
        }
        // Resolves an error if the file isn't a valid .htm or .html file
        else if (!TemplateService.isValidTemplateFile(value)) {
            return strings.ErrorTemplateExtension;
        }
        // Resolves an error if the file doesn't answer a simple head request
        else {
            await this._templateService.ensureFileResolves(value);
            return '';
        }
    } catch (error) {
        return Text.format(strings.ErrorTemplateResolve, error);
    }
  }

  /**
   * Determines the group fields for the search settings options inside the property pane
   */
  private _getSearchSettingsFields(): IPropertyPaneField<any>[] {

    // Sets up search settings fields
    const searchSettingsFields: IPropertyPaneField<any>[] = [
        PropertyPaneTextField('queryTemplate', {
            label: strings.QueryTemplateFieldLabel,
            value: this.properties.queryTemplate,
            disabled: false,
            multiline: true,
            resizable: true,
            placeholder: strings.SearchQueryPlaceHolderText,
            deferredValidationTime: 300
        }),
        PropertyPaneTextField('resultSourceId', {
            label: strings.ResultSourceIdLabel,
            multiline: false,
            onGetErrorMessage: this.validateSourceId.bind(this),
            deferredValidationTime: 300
        }),
        this._propertyFieldCollectionData('sortList', {
            manageBtnLabel: strings.Sort.EditSortLabel,
            key: 'sortList',
            enableSorting: true,
            panelHeader: strings.Sort.EditSortLabel,
            panelDescription: strings.Sort.SortListDescription,
            label: strings.Sort.SortPropertyPaneFieldLabel,
            value: this.properties.sortList,
            fields: [
                {
                    id: 'sortField',
                    title: "Field name",
                    type: this._customCollectionFieldType.custom,
                    required: true,
                    onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {

                        // Need to specify a React key to avoid item duplication when adding a new row
                        return React.createElement("div", {key : `${field.id}-${itemId}`},
                            React.createElement(SearchManagedProperties, {
                            defaultSelectedKey: item[field.id] ? item[field.id] : '',
                            onUpdate: (newValue: any, isSortable: boolean) => { 

                                if (!isSortable) {
                                    onCustomFieldValidation(field.id, strings.Sort.SortInvalidSortableFieldMessage);
                                } else {
                                    onUpdate(field.id, newValue);
                                    onCustomFieldValidation(field.id, '');
                                }
                            },
                            searchService: this._searchService,
                            validateSortable: true,
                            availableProperties: this._availableManagedProperties,
                            onUpdateAvailableProperties: this._onUpdateAvailableProperties
                        } as ISearchManagedPropertiesProps));
                    }
                },
                {
                    id: 'sortDirection',
                    title: "Direction",
                    type: this._customCollectionFieldType.dropdown,
                    required: true,
                    options: [
                        {
                            key: ISortFieldDirection.Ascending,
                            text: strings.Sort.SortDirectionAscendingLabel
                        },
                        {
                            key: ISortFieldDirection.Descending,
                            text: strings.Sort.SortDirectionDescendingLabel
                        }
                    ]
                }
            ]
        }),
        this._propertyFieldCollectionData('sortableFields', {
            manageBtnLabel: strings.Sort.EditSortableFieldsLabel,
            key: 'sortableFields',
            enableSorting: true,
            panelHeader: strings.Sort.EditSortableFieldsLabel,
            panelDescription: strings.Sort.SortableFieldsDescription,
            label: strings.Sort.SortableFieldsPropertyPaneField,
            value: this.properties.sortableFields,
            fields: [
                {
                    id: 'sortField',
                    title: strings.Sort.SortableFieldManagedPropertyField,
                    type: this._customCollectionFieldType.custom,
                    required: true,
                    onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                        // Need to specify a React key to avoid item duplication when adding a new row
                        return (
                        React.createElement("div", {key : `${field.id}-${itemId}`}, React.createElement(SearchManagedProperties, {
                            defaultSelectedKey: item[field.id] ? item[field.id] : '',
                            onUpdate: (newValue: any, isSortable: boolean) => { 
                                if (!isSortable) {
                                    onCustomFieldValidation(field.id, strings.Sort.SortInvalidSortableFieldMessage);
                                } else {
                                    onUpdate(field.id, newValue);
                                    onCustomFieldValidation(field.id, '');
                                }
                            },
                            searchService: this._searchService,
                            validateSortable: true,
                            availableProperties: this._availableManagedProperties,
                            onUpdateAvailableProperties: this._onUpdateAvailableProperties
                        } as ISearchManagedPropertiesProps)));
                    }
                },          
                {
                    id: 'displayValue',
                    title: strings.Sort.SortableFieldDisplayValueField,
                    type: this._customCollectionFieldType.string
                }
            ]
        }),
        PropertyPaneToggle('enableQueryRules', {
            label: strings.EnableQueryRulesLabel,
            checked: this.properties.enableQueryRules,
        }),
        new PropertyPaneSearchManagedProperties('selectedProperties', {
            label: strings.SelectedPropertiesFieldLabel,
            description: strings.SelectedPropertiesFieldDescription,
            allowMultiSelect: true,
            availableProperties: this._availableManagedProperties,
            defaultSelectedKeys: this.properties.selectedProperties.split(","),
            onPropertyChange: (propertyPath: string, newValue: any) => { 
                this.properties[propertyPath] = newValue.join(','); 
                this.onPropertyPaneFieldChanged(propertyPath);

                // Refresh the WP with new selected properties
                this.render();
            },
            onUpdateAvailableProperties: this._onUpdateAvailableProperties,
            searchService: this._searchService,
        }),
        PropertyPaneSlider('maxResultsCount', {
            label: strings.MaxResultsCount,
            max: 50,
            min: 1,
            showValue: true,
            step: 1,
            value: 50,
        }),
        PropertyPaneToggle('enableLocalization', {
            checked: this.properties.enableLocalization,
            label: strings.EnableLocalizationLabel,
            onText: strings.EnableLocalizationOnLabel,
            offText: strings.EnableLocalizationOffLabel
        }),
        PropertyPaneDropdown('searchQueryLanguage', {
            label: strings.QueryCultureLabel,
            options: [{
                key: -1,
                text: strings.QueryCultureUseUiLanguageLabel
            } as IDropdownOption].concat(sortBy(this._availableLanguages, ['text'])),
            selectedKey: this.properties.searchQueryLanguage ? this.properties.searchQueryLanguage : 0
        }),
        this._propertyFieldCollectionData('synonymList', {
            manageBtnLabel: strings.Synonyms.EditSynonymLabel,
            key: 'synonymList',
            enableSorting: false,
            panelHeader: strings.Synonyms.EditSynonymLabel,
            panelDescription: strings.Synonyms.SynonymListDescription,
            label: strings.Synonyms.SynonymPropertyPanelFieldLabel,
            value: this.properties.synonymList,
            fields: [
                {
                    id: 'Term',
                    title: strings.Synonyms.SynonymListTerm,
                    type: this._customCollectionFieldType.string,
                    required: true,
                    placeholder: strings.Synonyms.SynonymListTermExemple
                },
                {
                    id: 'Synonyms',
                    title: strings.Synonyms.SynonymListSynonyms,
                    type: this._customCollectionFieldType.string,
                    required: true,
                    placeholder: strings.Synonyms.SynonymListSynonymsExemple
                },
                {
                    id: 'TwoWays',
                    title: strings.Synonyms.SynonymIsTwoWays,
                    type: this._customCollectionFieldType.boolean,
                    required: false
                }
            ]
        })
    ];

    return searchSettingsFields;
  }
  /**
   * Determines the group fields for the search query options inside the property pane
   */
  private _getSearchQueryFields(): any {

    let defaultSearchQueryFields: IPropertyPaneField<any>[] = [];

    defaultSearchQueryFields.push(
      PropertyPaneCheckbox('useDefaultSearchQuery', {
          text: strings.UseDefaultSearchQueryKeywordsFieldLabel
      })
    );
    if (this.properties.useDefaultSearchQuery) {
      defaultSearchQueryFields.push(
          PropertyPaneTextField('defaultSearchQuery', {
              label: strings.DefaultSearchQueryKeywordsFieldLabel,
              description: strings.DefaultSearchQueryKeywordsFieldDescription,
              multiline: true,
              resizable: true,
              placeholder: strings.SearchQueryPlaceHolderText,
              onGetErrorMessage: this._validateEmptyField.bind(this),
              deferredValidationTime: 500
          })
      );
    }
    defaultSearchQueryFields.push(
      PropertyPaneTextField('queryKeywords', {
        label: strings.SearchQueryKeywordsFieldLabel,
        description: strings.SearchQueryKeywordsFieldDescription,
        multiline: true,
        resizable: true,
        placeholder: strings.SearchQueryPlaceHolderText,
        deferredValidationTime: 500
      })
    );

    return defaultSearchQueryFields
  }

  /**
   * Determines the group fields for styling options inside the property pane
   */
  private _getStylingFields(): IPropertyPaneField<any>[] {

    // Options for the search results layout 
    const layoutOptions = [
        {
            iconProps: {
                officeFabricIconFontName: 'List'
            },
            text: strings.SimpleListLayoutOption,
            key: ResultsLayoutOption.SimpleList,
        },
        {
            iconProps: {
                officeFabricIconFontName: 'Table'
            },
            text: strings.DetailsListLayoutOption,
            key: ResultsLayoutOption.DetailsList,
        },
        {
            iconProps: {
                officeFabricIconFontName: 'Tiles'
            },
            text: strings.TilesLayoutOption,
            key: ResultsLayoutOption.Tiles
        },
        {
            iconProps: {
                officeFabricIconFontName: 'People'
            },
            text: strings.PeopleLayoutOption,
            key: ResultsLayoutOption.People
        },
        {
            iconProps: {
                officeFabricIconFontName: 'Code'
            },
            text: strings.DebugLayoutOption,
            key: ResultsLayoutOption.Debug
        }
    ] as IPropertyPaneChoiceGroupOption[];

    layoutOptions.push(...this.getCodeRenderers());
    layoutOptions.push({
        iconProps: {
            officeFabricIconFontName: 'CodeEdit'
        },
        text: strings.CustomLayoutOption,
        key: ResultsLayoutOption.Custom,
    });

    const canEditTemplate = this.properties.externalTemplateUrl && this.properties.selectedLayout === ResultsLayoutOption.Custom ? false : true;

    let dialogTextFieldValue;
    if (!this.codeRendererIsSelected()) {
        switch (this.properties.selectedLayout) {
            case ResultsLayoutOption.DetailsList:
                dialogTextFieldValue = BaseTemplateService.getDefaultResultTypeListItem();
                break;

            case ResultsLayoutOption.Tiles:
                dialogTextFieldValue = BaseTemplateService.getDefaultResultTypeTileItem();
                break;

            default:
                dialogTextFieldValue = BaseTemplateService.getDefaultResultTypeCustomItem();
                break;
        }
    }

    // Sets up styling fields
    let stylingFields: IPropertyPaneField<any>[] = [
        PropertyPaneTextField('webPartTitle', {
            label: strings.WebPartTitle
        }),
        PropertyPaneToggle('showBlank', {
            label: strings.ShowBlankLabel,
            checked: this.properties.showBlank,
        }),
        PropertyPaneToggle('showResultsCount', {
            label: strings.ShowResultsCountLabel,
            checked: this.properties.showResultsCount,
        }),
        PropertyPaneToggle('showPaging', {
            label: strings.UsePaginationWebPartLabel,
            checked: this.properties.showPaging,
        }),
        PropertyPaneHorizontalRule(),
        PropertyPaneChoiceGroup('selectedLayout', {
            label: strings.ResultsLayoutLabel,
            options: layoutOptions
        }),
    ];

    if (!this.codeRendererIsSelected()) {
        stylingFields.push(
            this._propertyFieldCodeEditor('inlineTemplateText', {
                label: strings.DialogButtonLabel,
                panelTitle: strings.DialogTitle,
                initialValue: this._templateContentToDisplay,
                deferredValidationTime: 500,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: !canEditTemplate,
                key: 'inlineTemplateTextCodeEditor',
                language: this._propertyFieldCodeEditorLanguages.Handlebars
            }),
            this._propertyFieldCollectionData('resultTypes', {
                manageBtnLabel: strings.ResultTypes.EditResultTypesLabel,
                key: 'resultTypes',
                panelHeader: strings.ResultTypes.EditResultTypesLabel,
                panelDescription: strings.ResultTypes.ResultTypesDescription,
                enableSorting: true,
                label: strings.ResultTypes.ResultTypeslabel,
                value: this.properties.resultTypes,
                fields: [
                    {
                        id: 'property',
                        title: strings.ResultTypes.ConditionPropertyLabel,
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            // Need to specify a React key to avoid item duplication when adding a new row
                            return React.createElement("div", {key : itemId},
                            React.createElement(SearchManagedProperties, {
                            defaultSelectedKey: item[field.id] ? item[field.id] : '',
                            onUpdate: (newValue: any, isSortable: boolean) => { 
                                onUpdate(field.id, newValue);
                            },
                            searchService: this._searchService,
                            validateSortable: false,
                            availableProperties: this._availableManagedProperties,
                            onUpdateAvailableProperties: this._onUpdateAvailableProperties
                            } as ISearchManagedPropertiesProps));
                        }
                    },
                    {
                        id: 'operator',
                        title: strings.ResultTypes.CondtionOperatorValue,
                        type: this._customCollectionFieldType.dropdown,
                        defaultValue: ResultTypeOperator.Equal,
                        required: true,
                        options: [
                            {
                                key: ResultTypeOperator.Equal,
                                text: strings.ResultTypes.EqualOperator
                            },
                            {
                                key: ResultTypeOperator.Contains,
                                text: strings.ResultTypes.ContainsOperator
                            },
                            {
                                key: ResultTypeOperator.StartsWith,
                                text: strings.ResultTypes.StartsWithOperator
                            },
                            {
                                key: ResultTypeOperator.NotNull,
                                text: strings.ResultTypes.NotNullOperator
                            },
                            {
                                key: ResultTypeOperator.GreaterOrEqual,
                                text: strings.ResultTypes.GreaterOrEqualOperator
                            },
                            {
                                key: ResultTypeOperator.GreaterThan,
                                text: strings.ResultTypes.GreaterThanOperator
                            },
                            {
                                key: ResultTypeOperator.LessOrEqual,
                                text: strings.ResultTypes.LessOrEqualOperator
                            },
                            {
                                key: ResultTypeOperator.LessThan,
                                text: strings.ResultTypes.LessThanOperator
                            }
                        ]
                    },
                    {
                        id: 'value',
                        title: strings.ResultTypes.ConditionValueLabel,
                        type: this._customCollectionFieldType.string,
                        required: false,
                    },
                    {
                        id: "inlineTemplateContent",
                        title: strings.ResultTypes.InlineTemplateContentLabel,
                        type: this._customCollectionFieldType.custom,
                        onCustomRender: (field, value, onUpdate) => {
                            return (
                                React.createElement("div", null,
                                    React.createElement(this._textDialogComponent.TextDialog, {
                                        language: this._propertyFieldCodeEditorLanguages.Handlebars,
                                        dialogTextFieldValue: value ? value : dialogTextFieldValue,
                                        onChanged: (fieldValue) => onUpdate(field.id, fieldValue),
                                        strings: {
                                            cancelButtonText: strings.CancelButtonText,
                                            dialogButtonText: strings.DialogButtonText,
                                            dialogTitle: strings.DialogTitle,
                                            saveButtonText: strings.SaveButtonText
                                        }
                                    })
                                )
                            );
                        }
                    },
                    {
                        id: 'externalTemplateUrl',
                        title: strings.ResultTypes.ExternalUrlLabel,
                        type: this._customCollectionFieldType.url,
                        onGetErrorMessage: this._onTemplateUrlChange.bind(this),
                        placeholder: 'https://mysite/Documents/external.html'
                    },
                ]
            })
        );
    }
    // Only show the template external URL for 'Custom' option
    if (this.properties.selectedLayout === ResultsLayoutOption.Custom) {
        stylingFields.splice(6, 0, PropertyPaneTextField('externalTemplateUrl', {
            label: strings.TemplateUrlFieldLabel,
            placeholder: strings.TemplateUrlPlaceholder,
            deferredValidationTime: 500,
            onGetErrorMessage: this._onTemplateUrlChange.bind(this)
        }));
    }

    if (this.codeRendererIsSelected()) {
        const currentCodeRenderer = find(this._codeRenderers, (renderer) => renderer.id === (this.properties.selectedLayout as any));
        if (!this.properties.customTemplateFieldValues) {
            this.properties.customTemplateFieldValues = currentCodeRenderer.customFields.map(field => {
                return {
                    fieldName: field,
                    searchProperty: ''
                };
            });
        }
        if (currentCodeRenderer && currentCodeRenderer.customFields && currentCodeRenderer.customFields.length > 0) {
            const searchPropertyOptions = this.properties.selectedProperties.split(',').map(prop => {
                return ({
                    key: prop,
                    text: prop
                });
            });
            stylingFields.push(this._propertyFieldCollectionData('customTemplateFieldValues', {
                key: 'customTemplateFieldValues',
                label: strings.customTemplateFieldsLabel,
                panelHeader: strings.customTemplateFieldsPanelHeader,
                manageBtnLabel: strings.customTemplateFieldsConfigureButtonLabel,
                value: this.properties.customTemplateFieldValues,
                fields: [
                    {
                        id: 'fieldName',
                        title: strings.customTemplateFieldTitleLabel,
                        type: this._customCollectionFieldType.string,
                    },
                    {
                        id: 'searchProperty',
                        title: strings.customTemplateFieldPropertyLabel,
                        type: this._customCollectionFieldType.dropdown,
                        options: searchPropertyOptions
                    }
                ]
            }));
        }
    }

    return stylingFields;
  }

  /**
   * Gets template parameters fields
   */
  private _getTemplateFieldsGroup(): IPropertyPaneGroup {

      let templateFieldsGroup: IPropertyPaneGroup = null;

      if (this._templatePropertyPaneOptions.length > 0) {

          templateFieldsGroup = {
              groupFields: this._templatePropertyPaneOptions,
              isCollapsed: false,
              groupName: strings.TemplateParameters.TemplateParametersGroupName
          };
      } 

      return templateFieldsGroup;
  }

  protected getCodeRenderers(): IPropertyPaneChoiceGroupOption[] {
      const registeredRenderers = this._codeRenderers;
      if (registeredRenderers && registeredRenderers.length > 0) {
          return registeredRenderers.map(ca => {
              return {
                  key: ca.id,
                  text: ca.name,
                  iconProps: {
                      officeFabricIconFontName: ca.icon
                  },
              };
          });
      } else {
          return [];
      }
  }

  protected codeRendererIsSelected(): boolean {
      const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
      return guidRegex.test(this.properties.selectedLayout as any);
  }

  public getPropertyDefinitions(): ReadonlyArray<any> {

      // Use the Web Part title as property title since we don't expose sub properties
      return [
          {
              id: SearchComponentType.SearchResultsWebPart,
              title: this.properties.webPartTitle ? this.properties.webPartTitle : this.title
          }
      ];
  }

  public getPropertyValue(propertyId: string): ISearchResultSourceData {

    const searchResultSourceData: ISearchResultSourceData = {
        queryKeywords: '',
        refinementResults: (this._resultService && this._resultService.results) ? this._resultService.results.RefinementResults : [],
        paginationInformation: (this._resultService && this._resultService.results) ? this._resultService.results.PaginationInformation : {
            CurrentPage: 1,
            MaxResultsPerPage: this.properties.maxResultsCount,
            TotalRows: 0
        },
        searchServiceConfiguration: this._searchService.getConfiguration(),
        verticalsInformation: []
    };

    switch (propertyId) {
        case SearchComponentType.SearchResultsWebPart:
            return searchResultSourceData;
    }

    throw new Error('Bad property id');
  }

  /**
   * Handler when the list of available managed properties is fetched by a property pane control¸or a field in a collection data control
   * @param properties the fetched properties
   */
  private _onUpdateAvailableProperties(properties: IComboBoxOption[]) {
      // Save the value in the root Web Part class to avoid fetching it again if the property list is requested again by any other property pane control
      this._availableManagedProperties = cloneDeep(properties);

      // Refresh all fields so other property controls can use the new list 
      this.context.propertyPane.refresh();
      this.render();
  }
  /**
   * Verifies if the string is a correct URL
   * @param value the URL to verify
   */
  private _validatePageUrl(value: string) {    
    
    if ((!/^(https?):\/\/[^\s/$.?#].[^\s]*/.test(value) || !value) && this.properties.searchInNewPage) {
      return strings.SearchBoxUrlErrorMessage;
    }
    
    return '';
  }
}
