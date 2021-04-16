import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Text, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneHorizontalRule,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  IPropertyPaneGroup,
} from '@microsoft/sp-webpart-base';
import * as strings from 'NbsSearchWebPartStrings';
import { INbsSearchWebPartProps } from './INbsSearchWebPartProps';
import SearchResultsContainer from './components/SearchResultsContainer/SearchResultsContainer';
import ISearchResultsContainerProps from './components/SearchResultsContainer/ISearchResultsContainerProps';
import BaseTemplateService from '../../services/TemplateService/BaseTemplateService';
import RefinerTemplateOption from '../../models/RefinerTemplateOption';
import ISearchService from '../../services/SearchService/ISearchService';
import RefinersSortOption from '../../models/RefinersSortOptions';
import RefinersSortDirection from '../../models/RefinersSortDirection';
import ITaxonomyService from '../../services/TaxonomyService/ITaxonomyService';
import RefinersLayoutOption from '../../models/RefinersLayoutOptions';
import ResultsLayoutOption from '../../models/ResultsLayoutOption';
import { TemplateService } from '../../services/TemplateService/TemplateService';
import { isEmpty, find, sortBy, cloneDeep, findIndex } from '@microsoft/sp-lodash-subset';
import SearchService from '../../services/SearchService/SearchService';
import { BaseSuggestionProvider } from '../../providers/BaseSuggestionProvider';
import TaxonomyService from '../../services/TaxonomyService/TaxonomyService';
import { SortDirection, Sort } from '@pnp/sp';
import { ISortFieldConfiguration, ISortFieldDirection } from '../../models/ISortFieldConfiguration';
import { ISynonymFieldConfiguration } from '../../models/ISynonymFieldConfiguration';
import { ResultTypeOperator } from '../../models/ISearchResultType';
import IResultService from '../../services/ResultService/IResultService';
import { ResultService, IRenderer } from '../../services/ResultService/ResultService';
import { IRefinementFilter} from '../../models/ISearchResult';
import { SearchComponentType } from '../../models/SearchComponentType';
import ISynonymTable from '../../models/ISynonym';
import * as update from 'immutability-helper';
import LocalizationHelper from '../../helpers/LocalizationHelper';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { SearchManagedProperties, ISearchManagedPropertiesProps } from '../../controls/SearchManagedProperties/SearchManagedProperties';
import { PropertyPaneSearchManagedProperties } from '../../controls/PropertyPaneSearchManagedProperties/PropertyPaneSearchManagedProperties';
import { ExtensibilityService } from '../../services/ExtensibilityService/ExtensibilityService';
import IExtensibilityService from '../../services/ExtensibilityService/IExtensibilityService';
import { IComponentDefinition } from '../../services/ExtensibilityService/IComponentDefinition';
import { AvailableComponents } from '../../components/AvailableComponents';
import { IQueryModifierDefinition } from '../../services/ExtensibilityService/IQueryModifierDefinition';
import { ISuggestionProviderDefinition } from '../../services/ExtensibilityService/ISuggestionProviderDefinition';
import { IQueryModifierInstance } from '../../services/ExtensibilityService/IQueryModifierInstance';
import { ObjectCreator } from '../../services/ExtensibilityService/ObjectCreator';
import { BaseQueryModifier } from '../../services/ExtensibilityService/BaseQueryModifier';
import { Toggle } from 'office-ui-fabric-react';
import IQueryModifierConfiguration from '../../models/IQueryModifierConfiguration';
import { ISuggestionProviderInstance } from '../../services/ExtensibilityService/ISuggestionProviderInstance';
import { SharePointDefaultSuggestionProvider } from '../../providers/SharePointDefaultSuggestionProvider';
import { SearchHelper } from '../../helpers/SearchHelper';
import IUserService from '../../services/UserService/IUserService';
import { UserService } from '../../services/UserService/UserService';
import PnPTelemetry from "@pnp/telemetry-js";
import { UrlHelper } from '../../../lib/helpers/UrlHelper';

export default class NbsSearchWebPart extends BaseClientSideWebPart<INbsSearchWebPartProps> {
  private _userService: IUserService;
  /**
   * Refinement
   */
   private _selectedRefinementFilters: IRefinementFilter[];
   private _isDirty: boolean;

  /**
   * Search box
   */
  private _foundCustomSuggestionProviders: boolean = false;
  private _suggestionProviderInstances: ISuggestionProviderInstance<any>[];
  /**
   * searc results
   */
  private _searchService: ISearchService;
  private _taxonomyService: ITaxonomyService;
  private _templateService: BaseTemplateService;
  private _extensibilityService: IExtensibilityService;
  private _textDialogComponent = null;
  private _propertyFieldCodeEditor = null;
  private _placeholder = null;
  private _propertyFieldCollectionData = null;
  private _customCollectionFieldType = null;
  private _queryModifierInstance: IQueryModifierInstance<any> = null;
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
  private _initComplete = false;
  /**
   * Information about time zone bias (current user or web)
   */
  private _timeZoneBias: any;

  /**
   * The available web component definitions (not registered yet)
   */
  private availableWebComponentDefinitions: IComponentDefinition<any>[] = AvailableComponents.BuiltinComponents;

  /**
   * The available query modifier definitions (not instancied yet)
   */
  private availableQueryModifierDefinitions: IQueryModifierDefinition<any>[] = [];
  private queryModifierSelected: boolean = false;

  /**
   * The default selected filters
   */
  private defaultSelectedFilters: IRefinementFilter[] = [];

  /**
   * The current page number
   */
  private currentPageNumber: number = 1;

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

    if (!this._initComplete) {
        // Don't render until all init is complete
        return;
    }

    // Determine the template content to display
    // In the case of an external template is selected, the render is done asynchronously waiting for the content to be fetched
    await this._initTemplate();

    if (this.displayMode === DisplayMode.Edit) {
        const { Placeholder } = await import(
            /* webpackChunkName: 'search-property-pane' */
            '@pnp/spfx-controls-react/lib/Placeholder'
        );
        this._placeholder = Placeholder;
    }

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
    let selectedFilters: IRefinementFilter[] = [];
    let queryTemplate: string = this.properties.queryTemplate;
    let sourceId: string = this.properties.resultSourceId;

    // Get default selected refiners from the URL
    this.defaultSelectedFilters = SearchHelper.getRefinementFiltersFromUrl();
    selectedFilters = this.defaultSelectedFilters;

    let queryDataSourceValue = this.properties.queryKeywords;

    let queryKeywords = queryDataSourceValue ? queryDataSourceValue : this.properties.defaultSearchQuery;
    
    const searchK: string = UrlHelper.getQueryStringParam('k', null);
    if (searchK) {
        queryKeywords = decodeURIComponent(searchK);
    }
    
    // Get data from connected sources
    if (this.properties.useRefiners) {
        if (this._isDirty) {
            selectedFilters = this._selectedRefinementFilters;
            // Reset the default filters provided in URL when user starts to select/unselected values manually
            this.defaultSelectedFilters = [];
        }
    }    

    const currentLocaleId = LocalizationHelper.getLocaleId(this.context.pageContext.cultureInfo.currentCultureName);
    const queryModifier = this._queryModifierInstance && this._queryModifierInstance.isInitialized ? this._queryModifierInstance.instance : null;
    const enableSuggestions = this.properties.enableQuerySuggestions && this.properties.suggestionProviders.some(sp => sp.providerEnabled);
    // Configure the provider before the query according to our needs
    this._searchService = update(this._searchService, {
        timeZoneId: { $set: this._timeZoneBias && this._timeZoneBias.Id ? this._timeZoneBias.Id : null },
        resultsCount: { $set: this.properties.paging.itemsCountPerPage },
        queryTemplate: { $set: queryTemplate },
        resultSourceId: { $set: sourceId },
        sortList: { $set: this._searchService.sortList || this._convertToSortList(this.properties.sortList) },
        enableQueryRules: { $set: this.properties.enableQueryRules },
        selectedProperties: { $set: this.properties.selectedProperties ? this.properties.selectedProperties.replace(/\s|,+$/g, '').split(',') : [] },
        synonymTable: { $set: this._synonymTable },
        queryCulture: { $set: this.properties.searchQueryLanguage !== -1 ? this.properties.searchQueryLanguage : currentLocaleId },
        refinementFilters: { $set: selectedFilters.length > 0 ? selectedFilters : [] },
        refiners: { $set: this.properties.refinersConfiguration },
        queryModifier: { $set: queryModifier },
    });
    
    const isValueConnected = !!this.properties.queryKeywords;
    this._searchContainer = React.createElement(
        SearchResultsContainer,
        {
            userService: this._userService,
            useRefiners: this.properties.useRefiners,
            useSearchBox: this.properties.useSearchBox,
            defaultSelectedRefinementFilters: this.defaultSelectedFilters,
            language: this.context.pageContext.cultureInfo.currentUICultureName,
            refinersSelectedLayout: this.properties.refinersSelectedLayout,
            refinersConfiguration: this.properties.refinersConfiguration,
            enableQuerySuggestions: enableSuggestions,
            placeholderText: this.properties.placeholderText,
            suggestionProviders: this._suggestionProviderInstances,
            domElement: this.domElement,
            searchService: this._searchService,
            taxonomyService: this._taxonomyService,
            queryKeywords: queryKeywords,
            sortList: this.properties.sortList,
            sortableFields: this.properties.sortableFields,
            showResultsCount: this.properties.showResultsCount,
            showBlank: this.properties.showBlank,
            displayMode: this.displayMode,
            templateService: this._templateService,
            templateContent: this._templateContentToDisplay,
            templateParameters: this.properties.templateParameters,
            currentUICultureName: this.context.pageContext.cultureInfo.currentUICultureName,
            siteServerRelativeUrl: this.context.pageContext.site.serverRelativeUrl,
            webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
            resultTypes: this.properties.resultTypes,
            useCodeRenderer: this.codeRendererIsSelected(),
            customTemplateFieldValues: this.properties.customTemplateFieldValues,
            rendererId: this.properties.selectedLayout as any,
            enableLocalization: this.properties.enableLocalization,
            selectedPage: this.currentPageNumber,
            selectedLayout: this.properties.selectedLayout,
            onSearchResultsUpdate: async (results, mountingNodeId, searchService) => {
                /*
                if (this.properties.selectedLayout in ResultsLayoutOption) {
                    let node = document.getElementById(mountingNodeId);
                    if (node) {
                        //ReactDom.render(null,node);
                    }
                }*/
                this._resultService.updateResultData(results, this.properties.selectedLayout as any, mountingNodeId, this.properties.customTemplateFieldValues);
            },
            onRefinersUpdate: async (isDirty, selectedRefinements) => {
                this._isDirty = isDirty;
                if (isDirty) {
                    this._selectedRefinementFilters = selectedRefinements;
                    this.currentPageNumber = 1;
                    this.render();
                }
                
            },
            pagingSettings: this.properties.paging,
            instanceId: this.instanceId
        } as ISearchResultsContainerProps
    );

    if (isValueConnected && !this.properties.useDefaultSearchQuery ||
        isValueConnected && this.properties.useDefaultSearchQuery && this.properties.defaultSearchQuery ||
        !isValueConnected && !isEmpty(queryKeywords)) {
        renderElement = this._searchContainer;
    } else {
        if (this.displayMode === DisplayMode.Edit) {
            const placeholder: React.ReactElement<any> = React.createElement(
                this._placeholder,
                {
                    iconName: strings.PlaceHolderEditLabel,
                    iconText: strings.PlaceHolderIconText,
                    description: strings.PlaceHolderDescription,
                    buttonLabel: strings.PlaceHolderConfigureBtnLabel,
                    onConfigure: this._setupWebPart.bind(this)
                }
            );
            renderElement = placeholder;
        } else {
            renderElement = React.createElement('div', null);
        }
    }

    ReactDom.render(renderElement, this.domElement);
  }
  protected async onInit(): Promise<void> {
    // Disable PnP Telemetry
    const telemetry = PnPTelemetry.getInstance();
    if (telemetry.optOut) telemetry.optOut();

    this.initializeRequiredProperties();
    
    this._taxonomyService = new TaxonomyService(this.context.pageContext.site.absoluteUrl);

    this._timeZoneBias = {
        WebBias: this.context.pageContext.legacyPageContext.webTimeZoneData.Bias,
        WebDST: this.context.pageContext.legacyPageContext.webTimeZoneData.DaylightBias,
        UserBias: null,
        UserDST: null,
        Id: this.context.pageContext.legacyPageContext.webTimeZoneData.Id
    };
    if (this.context.pageContext.legacyPageContext.userTimeZoneData) {
        this._timeZoneBias.UserBias = this.context.pageContext.legacyPageContext.userTimeZoneData.Bias;
        this._timeZoneBias.UserDST = this.context.pageContext.legacyPageContext.userTimeZoneData.DaylightBias;
        this._timeZoneBias.Id = this.context.pageContext.legacyPageContext.webTimeZoneData.Id;
    }

    this._searchService = new SearchService(this.context.pageContext, this.context.spHttpClient);
    this._templateService = new TemplateService(this.context.spHttpClient, this.context.pageContext.cultureInfo.currentUICultureName, this._searchService, this._timeZoneBias, this.context);

    this._userService = new UserService(this.context.pageContext);
    this._resultService = new ResultService();
    this._extensibilityService = new ExtensibilityService();
    this._codeRenderers = this._resultService.getRegisteredRenderers();

    await this.initSuggestionProviders();

    // Load extensibility library if present
    const extensibilityLibrary = await this._extensibilityService.loadExtensibilityLibrary();

    // Load extensibility additions
    if (extensibilityLibrary) {

        // Add custom web components if any
        this.availableWebComponentDefinitions = this.availableWebComponentDefinitions.concat(extensibilityLibrary.getCustomWebComponents());

        // Get custom query modifiers if present
        this.availableQueryModifierDefinitions = extensibilityLibrary.getCustomQueryModifiers ? extensibilityLibrary.getCustomQueryModifiers() : [];

        // Initializes query modifiers property for selection
        this.properties.queryModifiers = this.availableQueryModifierDefinitions.map(definition => {
            return {
                queryModifierDisplayName: definition.displayName,
                queryModifierDescription: definition.description,
                queryModifierEnabled: this.properties.selectedQueryModifierDisplayName && this.properties.selectedQueryModifierDisplayName === definition.displayName ? true : false
            } as IQueryModifierConfiguration;
        });

        // If we have a query modifier selected from config, we ensure it exists and is actually loaded fron the extensibility library
        const queryModifierDefinition = this.availableQueryModifierDefinitions.filter(definition => definition.displayName === this.properties.selectedQueryModifierDisplayName);
        if (this.properties.selectedQueryModifierDisplayName && queryModifierDefinition.length === 1) {
            this.queryModifierSelected = true;
            this._queryModifierInstance = await this._initQueryModifierInstance(queryModifierDefinition[0]);
        } else {
            this.properties.selectedQueryModifierDisplayName = null;
        }
    }

    // Set the default search results layout
    this.properties.selectedLayout = (this.properties.selectedLayout !== undefined && this.properties.selectedLayout !== null) ? this.properties.selectedLayout : ResultsLayoutOption.DetailsList;

    // Registers web components
    this._templateService.registerWebComponents(this.availableWebComponentDefinitions);

    this._synonymTable = this._convertToSynonymTable(this.properties.synonymList);

    this._initComplete = true;

    // Bind web component events
    this.bindPagingEvents();

    return super.onInit();
  }

  private async _initQueryModifierInstance(queryModifierDefinition: IQueryModifierDefinition<any>): Promise<IQueryModifierInstance<any>> {

    if (!queryModifierDefinition) {
        return null;
    }

    let isInitialized = false;
    let instance: BaseQueryModifier = null;

    try {
        instance = ObjectCreator.createEntity(queryModifierDefinition.class, this.context);
        await instance.onInit();
        isInitialized = true;
    }
    catch (error) {
        console.log(`Unable to initialize query modifier '${queryModifierDefinition.displayName}'. ${error}`);
    }
    finally {
        return {
            ...queryModifierDefinition,
            instance,
            isInitialized
        };
    }
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

  /**
   * Initializes the Web Part required properties if there are not present in the manifest (i.e. during an update scenario)
   */
  private initializeRequiredProperties() {

    this.properties.queryTemplate = this.properties.queryTemplate ? this.properties.queryTemplate : "{searchTerms}";

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

    // Ensure the minmal managed properties are here        
    const defaultManagedProperties =    [
                                            "Title",
                                            "Path",
                                            "OriginalPath",
                                            "SiteLogo",
                                            "contentclass",
                                            "FileExtension",
                                            "Filename",
                                            "ServerRedirectedURL",
                                            "DefaultEncodingURL",
                                            "IsDocument",
                                            "IsContainer",
                                            "IsListItem",
                                            "FileType",
                                            "HtmlFileType",
                                            "NormSiteID",
                                            "NormListID",
                                            "NormUniqueID",
                                            "Created",
                                            "PreviewUrl",
                                            "PictureThumbnailURL",
                                            "ServerRedirectedPreviewURL",
                                            "HitHighlightedSummary",
                                            "ServerRedirectedEmbedURL",
                                            "ParentLink",
                                            "owstaxidmetadataalltagsinfo",
                                            "Author",
                                            "AuthorOWSUSER",
                                            "SPSiteUrl",
                                            "SiteTitle",
                                            "SiteId",
                                            "WebId",
                                            "UniqueID"
                                        ];

    if (this.properties.selectedProperties) {

        let properties = this.properties.selectedProperties.split(',');

        defaultManagedProperties.map(property => {

            const idx = findIndex(properties, (item:string) => property.toLowerCase() === item.toLowerCase());                
            if (idx === -1) {
                properties.push(property);
            }
        });

        this.properties.selectedProperties = properties.join(',');
    } else {
        this.properties.selectedProperties = defaultManagedProperties.join(',');
    }
    
    this.properties.resultTypes = Array.isArray(this.properties.resultTypes) ? this.properties.resultTypes : [];
    this.properties.synonymList = Array.isArray(this.properties.synonymList) ? this.properties.synonymList : [];
    this.properties.searchQueryLanguage = this.properties.searchQueryLanguage ? this.properties.searchQueryLanguage : -1;
    this.properties.templateParameters = this.properties.templateParameters ? this.properties.templateParameters : {};
    this.properties.queryModifiers = !isEmpty(this.properties.queryModifiers) ? this.properties.queryModifiers : [];

    if (!this.properties.paging) {

        this.properties.paging = {
            itemsCountPerPage: 10,
            pagingRange: 5,
            showPaging: true,
            hideDisabled: true,
            hideFirstLastPages: false,
            hideNavigation: false
        };
    }
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
            }
          ],
          displayGroupsAsAccordion: false
        },
        {
          groups: [
            {
                groupFields: this._getSearchSettingsFields(),
                isCollapsed: false,
                groupName: strings.SearchSettingsGroupName
            },
            {
                groupName: strings.Paging.PagingOptionsGroupName,
                groupFields: this.getPagingGroupFields()
            }
          ],
          displayGroupsAsAccordion: true
        },
        {
          groups: stylingPageGroups,
          displayGroupsAsAccordion: false
        },
        {
            groups: [
                {
                  groupName: strings.SearchBoxQuerySettings,
                  groupFields: this._getSearchBoxFields()
                }
              ],
              displayGroupsAsAccordion: true
        },
        {
            groups: [
                {
                  groupFields: this._getRefiners(),
                  isCollapsed: false,
                  groupName: strings.RefinersSettingsGroupName
                }
              ],
              displayGroupsAsAccordion: false
        }
      ]
    };
  }
  protected async loadPropertyPaneResources(): Promise<void> {

    // tslint:disable-next-line:no-shadowed-variable
    const { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } = await import(
        /* webpackChunkName: 'search-property-pane' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );
    this._propertyFieldCodeEditor = PropertyFieldCodeEditor;
    this._propertyFieldCodeEditorLanguages = PropertyFieldCodeEditorLanguages;

    // Code editor component for property pane controls
    this._textDialogComponent = await import(
        /* webpackChunkName: 'search-property-pane' */
        '../../controls/TextDialog'
    );

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

    if (!this.properties.useDefaultSearchQuery) {
        this.properties.defaultSearchQuery = '';
    }

    // clean out duplicate ones
    let allProps = this.properties.selectedProperties.split(',');
    allProps = allProps.filter((item, index) => {
        return allProps.indexOf(item) === index;
    });
    this.properties.selectedProperties = allProps.join(',');


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

    if (propertyPath.localeCompare('queryModifiers') === 0) {

        // Load only the selected query modifier (can only have one at once, blocked by the UI)
        const configuredQueryModifiers = this.properties.queryModifiers.filter(m => m.queryModifierEnabled);

        if (configuredQueryModifiers.length === 1) {

            // Get the corresponding query modifier definition
            const queryModifierDefinition = this.availableQueryModifierDefinitions.filter(definition => definition.displayName === configuredQueryModifiers[0].queryModifierDisplayName);
            if (queryModifierDefinition.length === 1) {

                this.properties.selectedQueryModifierDisplayName = queryModifierDefinition[0].displayName;
                this._queryModifierInstance = await this._initQueryModifierInstance(queryModifierDefinition[0]);
            }

        } else {
            this.properties.selectedQueryModifierDisplayName = null;
            this._queryModifierInstance = null;
        }

        //this.context.dynamicDataSourceManager.notifyPropertyChanged(SearchComponentType.SearchResultsWebPart);
    }
    if (propertyPath.localeCompare('suggestionProviders') === 0) {
        await this.initSuggestionProviders();
    }
  }
  protected async onPropertyPaneConfigurationStart() {
    await this.loadPropertyPaneResources();
  }
  /**
   * Opens the Web Part property pane
   */
  private _setupWebPart() {
    this.context.propertyPane.open();
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

  private _validateNumber(value: string): string {
    let number = parseInt(value);
    if (isNaN(number)) {
        return strings.InvalidNumberIntervalMessage;
    }
    if (number < 1 || number > 500) {
        return strings.InvalidNumberIntervalMessage;
    }
    return '';
  }

  /**
   * Ensures the result source id value is a valid GUID
   * @param value the result source id
   */
  private validateSourceId(value: string): string {
    if (value.length > 0) {
        if (!(/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/).test(value)) {
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

    await this._templateService.optimizeLoadingForTemplate(this._templateContentToDisplay);
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
  private _getSearchBoxFields(): IPropertyPaneField<any>[] {
    let searchBehaviorOptionsFields: IPropertyPaneField<any>[] = [];
    let useSearchBox = this.properties.useSearchBox;
    if (this.properties.useSearchBox) {
        useSearchBox = false;
    }

    searchBehaviorOptionsFields.push(
        PropertyPaneToggle('useSearchBox', {
            label: strings.UseSearchBoxWebPartLabel,
            checked: useSearchBox
        }),
    );
    if (this.properties.useSearchBox) {
        searchBehaviorOptionsFields.push(
            PropertyPaneToggle("enableQuerySuggestions", {
                checked: false,
                label: strings.SearchBoxEnableQuerySuggestions
            })
        );
        if (this._foundCustomSuggestionProviders) {
            searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
            this._propertyFieldCollectionData('suggestionProviders', {
            manageBtnLabel: strings.SuggestionProviders.EditSuggestionProvidersLabel,
            key: 'suggestionProviders',
            panelHeader: strings.SuggestionProviders.EditSuggestionProvidersLabel,
            panelDescription: strings.SuggestionProviders.SuggestionProvidersDescription,
            disableItemCreation: true,
            disableItemDeletion: true,
            disabled: !this.properties.enableQuerySuggestions,
            label: strings.SuggestionProviders.SuggestionProvidersLabel,
            value: this.properties.suggestionProviders,
            fields: [
                {
                    id: 'providerEnabled',
                    title: strings.SuggestionProviders.EnabledPropertyLabel,
                    type: this._customCollectionFieldType.custom,
                    onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return (
                        React.createElement("div", null,
                            React.createElement(Toggle, { key: itemId, checked: value, onChanged: (checked) => {
                            onUpdate(field.id, checked);
                            }})
                        )
                        );
                    }
                },
                {
                    id: 'providerDisplayName',
                    title: strings.SuggestionProviders.ProviderNamePropertyLabel,
                    type: this._customCollectionFieldType.custom,
                    onCustomRender: (field, value) => {
                        return (
                        React.createElement("div", { style: { 'fontWeight': 600 } }, value)
                        );
                    }
                },
                {
                    id: 'providerDescription',
                    title: strings.SuggestionProviders.ProviderDescriptionPropertyLabel,
                    type: this._customCollectionFieldType.custom,
                    onCustomRender: (field, value) => {
                        return (
                        React.createElement("div", null, value)
                        );
                    }
                }
            ]
            })
        ]);
        }
        searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
            PropertyPaneHorizontalRule(),
            PropertyPaneTextField('placeholderText', {
                label: strings.SearchBoxPlaceholderTextLabel
                }),
        ]);
    }
    return searchBehaviorOptionsFields;
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
            deferredValidationTime: 1000
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
                        return React.createElement("div", { key: `${field.id}-${itemId}` },
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
                        return React.createElement("div", { key: `${field.id}-${itemId}` },
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
                    id: 'displayValue',
                    title: strings.Sort.SortableFieldDisplayValueField,
                    type: this._customCollectionFieldType.string
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
        /*
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
        }),*/
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
  private _getRefiners(): any {
    let defaultRefinersFields: IPropertyPaneField<any>[] = [];
    let useRefiners = this.properties.useRefiners;
    if (this.properties.useRefiners) {
        useRefiners = false;
    }

    defaultRefinersFields.push(
        PropertyPaneToggle('useRefiners', {
            label: strings.UseRefinersWebPartLabel,
            checked: useRefiners
        }),
    );
    if (this.properties.useRefiners) {
        defaultRefinersFields.push(
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
                      return React.createElement("div", { key: `${field.id}-${itemId}` },
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
                    title: strings.Refiners.RefinerTemplateField,
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
                      },
                      {
                        key: RefinerTemplateOption.FixedDateRange,
                        text: strings.Refiners.Templates.FixedDateRangeRefinementItemLabel,
                      },
                      {
                        key: RefinerTemplateOption.Persona,
                        text: strings.Refiners.Templates.PersonaRefinementItemLabel,
                      },
                      {
                        key: RefinerTemplateOption.FileType,
                        text: strings.Refiners.Templates.FileTypeRefinementItemTemplateLabel
                      },
                      {
                        key: RefinerTemplateOption.FileTypeMulti,
                        text: strings.Refiners.Templates.FileTypeMutliValueRefinementItemTemplateLabel
                      },
                      {
                        key: RefinerTemplateOption.ContainerTree,
                        text: strings.Refiners.Templates.ContainerTreeRefinementItemTemplateLabel
                      }
                    ]
                  },
                  {
                    id: 'refinerSortType',
                    title: strings.Refiners.Templates.RefinerSortTypeLabel,
                    type: this._customCollectionFieldType.dropdown,
                    options: [
                      {
                        key: RefinersSortOption.Default,
                        text: "--"
                      },
                      {
                        key: RefinersSortOption.ByNumberOfResults,
                        text: strings.Refiners.Templates.RefinerSortTypeByNumberOfResults,
                        ariaLabel: strings.Refiners.Templates.RefinerSortTypeByNumberOfResults
                      },
                      {
                        key: RefinersSortOption.Alphabetical,
                        text: strings.Refiners.Templates.RefinerSortTypeAlphabetical,
                        ariaLabel: strings.Refiners.Templates.RefinerSortTypeAlphabetical
                      }
                    ]
                  },
                  {
                    id: 'refinerSortDirection',
                    title: strings.Refiners.Templates.RefinerSortTypeSortOrderLabel,
                    type: this._customCollectionFieldType.dropdown,
                    options: [
                      {
                        key: RefinersSortDirection.Ascending,
                        text: strings.Refiners.Templates.RefinerSortTypeSortDirectionAscending,
                        ariaLabel: strings.Refiners.Templates.RefinerSortTypeSortDirectionAscending
                      },
                      {
                        key: RefinersSortDirection.Descending,
                        text: strings.Refiners.Templates.RefinerSortTypeSortDirectionDescending,
                        ariaLabel: strings.Refiners.Templates.RefinerSortTypeSortDirectionDescending
                      }
                    ]
                  },
                  {
                    id: 'showExpanded',
                    title: strings.Refiners.ShowExpanded,
                    type: this._customCollectionFieldType.boolean
                  },
                  {
                    id: 'showValueFilter',
                    title: strings.Refiners.showValueFilter,
                    type: this._customCollectionFieldType.boolean
                  }
                ]
              })
        );
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
        defaultRefinersFields.push(
            PropertyPaneChoiceGroup('refinersSelectedLayout', {
                label: strings.RefinerLayoutLabel,
                options: layoutOptions
            })
        );
    }
    

    return defaultRefinersFields;
  }
  
  /**
   * Determines the group fields for the search query options inside the property pane
   */
  private _getSearchQueryFields(): any {
    let defaultSearchQueryFields: IPropertyPaneField<any>[] = [];
    let queryModifiersFields: IPropertyPaneField<any>[] = [];

    // Query modifier fields
    if (this.properties.queryModifiers.length > 0) {
        queryModifiersFields = [
            PropertyPaneHorizontalRule(),
            ...this._getQueryModfiersFields()
        ];
    }

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
                deferredValidationTime: 1000
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
        deferredValidationTime: 1000
      }),
      ...queryModifiersFields
    );

    return defaultSearchQueryFields;
}

private _getQueryModfiersFields(): IPropertyPaneField<any>[] {

    let queryModificationFields: IPropertyPaneField<any>[] = [
      this._propertyFieldCollectionData('queryModifiers', {
        manageBtnLabel: strings.QueryModifier.ConfigureBtn,
        key: 'queryModifiers',
        panelHeader: strings.QueryModifier.PanelHeader,
        panelDescription: strings.QueryModifier.PanelDescription,
        enableSorting: false,
        label: strings.QueryModifier.FieldLbl,
        disableItemCreation: true,
        disableItemDeletion: true,
        disabled: this.availableQueryModifierDefinitions.length === 0,
        value: this.properties.queryModifiers,
        fields: [
          {
              id: 'queryModifierEnabled',
              title: strings.QueryModifier.EnableColumnLbl,
              type: this._customCollectionFieldType.custom,
              required: true,
              onCustomRender: (field, value, onUpdate, item, itemId) => {
                  return (
                      React.createElement("div", null,
                          React.createElement(Toggle, {
                              key: itemId,
                              checked: value,
                              disabled: this.queryModifierSelected && this.queryModifierSelected !== item[field.id] ? true : false,
                              onChange: ((evt, checked) => {
                                  // Reset every time the selected modifier. This will be determined when the field will be saved
                                  this.properties.selectedQueryModifierDisplayName = null;
                                  this.queryModifierSelected = !value;
                                  onUpdate(field.id, checked);
                              }).bind(this)
                          })
                      )
                  );
              }
          },
          {
              id: 'queryModifierDisplayName',
              title: strings.QueryModifier.DisplayNameColumnLbl,
              type: this._customCollectionFieldType.custom,
              onCustomRender: (field, value, onUpdate, item, itemId) => {
                  return (
                      React.createElement("div", { style: { 'fontWeight': 600 } }, value)
                  );
              }
          },
          {
              id: 'queryModifierDescription',
              title: strings.QueryModifier.DescriptionColumnLbl,
              type: this._customCollectionFieldType.custom,
              onCustomRender: (field, value, onUpdate, item, itemId) => {
                  return (
                      React.createElement("div", null, value)
                  );
              }
          }
        ]
      })
    ];

    if (this.properties.selectedQueryModifierDisplayName) {
      queryModificationFields.push(
        PropertyPaneLabel('', {
            text: Text.format(strings.QueryModifier.SelectedQueryModifierLbl, this.properties.selectedQueryModifierDisplayName)
        })
      );
    }

    return queryModificationFields;
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
                officeFabricIconFontName: 'Slider'
            },
            text: strings.SliderLayoutOption,
            key: ResultsLayoutOption.Slider
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
        PropertyPaneToggle('showBlank', {
            label: strings.ShowBlankLabel,
            checked: this.properties.showBlank,
        }),
        PropertyPaneToggle('showResultsCount', {
            label: strings.ShowResultsCountLabel,
            checked: this.properties.showResultsCount,
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
                            return React.createElement("div", { key: itemId },
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
                                key: ResultTypeOperator.NotEqual,
                                text: strings.ResultTypes.NotEqualOperator
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
            title: 'My Search'
        }
    ];
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
     * Binds event fired from pagination web components
     */
    private bindPagingEvents() {

      this.domElement.addEventListener('pageNumberUpdated', ((ev: CustomEvent) => {
          // We ensure the event if not propagated outside the component (i.e. other Web Part instances)
          ev.stopImmediatePropagation();

          // These information comes from the PaginationWebComponent class
          this.currentPageNumber = ev.detail.pageNumber;

          this.render();

      }).bind(this));
  }

  /**
   * Returns property pane 'Paging' group fields
   */
  private getPagingGroupFields(): IPropertyPaneField<any>[] {

      let groupFields: IPropertyPaneField<any>[] = [];


      groupFields.push(
          PropertyPaneToggle('paging.showPaging', {
              label: strings.Paging.ShowPagingFieldName,
          }),
          PropertyPaneTextField('paging.itemsCountPerPage', {
              label: strings.Paging.ItemsCountPerPageFieldName,
              value: this.properties.paging.itemsCountPerPage.toString(),
              maxLength: 3,
              deferredValidationTime: 300,
              onGetErrorMessage: this._validateNumber.bind(this),
          }),
          PropertyPaneSlider('paging.pagingRange', {
              label: strings.Paging.PagingRangeFieldName,
              max: 50,
              min: 0, // 0 = no page numbers displayed
              step: 1,
              showValue: true,
              value: this.properties.paging.pagingRange,
              disabled: !this.properties.paging.showPaging
          }),
          PropertyPaneHorizontalRule(),
          PropertyPaneToggle('paging.hideNavigation', {
              label: strings.Paging.HideNavigationFieldName,
              disabled: !this.properties.paging.showPaging
          }),
          PropertyPaneToggle('paging.hideFirstLastPages', {
              label: strings.Paging.HideFirstLastPagesFieldName,
              disabled: !this.properties.paging.showPaging
          }),
          PropertyPaneToggle('paging.hideDisabled', {
              label: strings.Paging.HideDisabledFieldName,
              disabled: !this.properties.paging.showPaging
          })
      );
      return groupFields;
  }
  

private async initSuggestionProviders(): Promise<void> {

  this.properties.suggestionProviders = await this.getAllSuggestionProviders();

  this._suggestionProviderInstances = await this.initSuggestionProviderInstances(this.properties.suggestionProviders);

}

private async getAllSuggestionProviders(): Promise<ISuggestionProviderDefinition<any>[]> {
  const [ defaultProviders, customProviders ] = await Promise.all([
      this.getDefaultSuggestionProviders(),
      this.getCustomSuggestionProviders()
  ]);

  //Track if we have any custom suggestion providers
  if (customProviders && customProviders.length > 0) {
    this._foundCustomSuggestionProviders = true;
  }

  //Merge all providers together and set defaults
  const savedProviders = this.properties.suggestionProviders && this.properties.suggestionProviders.length > 0 ? this.properties.suggestionProviders : [];
  const providerDefinitions = [ ...defaultProviders, ...customProviders ].map(provider => {
      const existingSavedProvider = find(savedProviders, sp => sp.providerName === provider.providerName);

      provider.providerEnabled = existingSavedProvider && undefined !== existingSavedProvider.providerEnabled
                                  ? existingSavedProvider.providerEnabled
                                  : undefined !== provider.providerEnabled
                                    ? provider.providerEnabled
                                    : true;

      return provider;
  });
  return providerDefinitions;
}

private async getDefaultSuggestionProviders(): Promise<ISuggestionProviderDefinition<any>[]> {
  return [{
      providerName: SharePointDefaultSuggestionProvider.ProviderName,
      providerDisplayName: SharePointDefaultSuggestionProvider.ProviderDisplayName,
      providerDescription: SharePointDefaultSuggestionProvider.ProviderDescription,
      providerClass: SharePointDefaultSuggestionProvider
  }];
}

private async getCustomSuggestionProviders(): Promise<ISuggestionProviderDefinition<any>[]> {
  let customSuggestionProviders: ISuggestionProviderDefinition<any>[] = [];

  // Load extensibility library if present
  const extensibilityLibrary = await this._extensibilityService.loadExtensibilityLibrary();

  // Load extensibility additions
  if (extensibilityLibrary && extensibilityLibrary.getCustomSuggestionProviders) {

      // Add custom suggestion providers if any
      customSuggestionProviders = extensibilityLibrary.getCustomSuggestionProviders();
  }

  return customSuggestionProviders;
}

private async initSuggestionProviderInstances(providerDefinitions: ISuggestionProviderDefinition<any>[]): Promise<ISuggestionProviderInstance<any>[]> {

  const webpartContext = this.context;

  let providerInstances = await Promise.all(providerDefinitions.map<Promise<ISuggestionProviderInstance<any>>>(async (provider) => {
    let isInitialized = false;
    let instance: BaseSuggestionProvider = null;

    try {
      instance = ObjectCreator.createEntity(provider.providerClass, webpartContext);
      await instance.onInit();
      isInitialized = true;
    }
    catch (error) {
      console.log(`Unable to initialize '${provider.providerName}'. ${error}`);
    }
    finally {
      return {
        ...provider,
        instance,
        isInitialized
      };
    }
  }));

  // Keep only the onces that initialized successfully
  providerInstances = providerInstances.filter(pi => pi.isInitialized);

  return providerInstances;
    }
}
