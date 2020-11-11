import ResultsLayoutOption from '../../models/ResultsLayoutOption';
import { ISortFieldConfiguration } from '../../models/ISortFieldConfiguration';
import ISortableFieldConfiguration from '../../models/ISortableFieldConfiguration';
import { ISearchResultType } from '../../models/ISearchResultType';
import { ICustomTemplateFieldValue } from '../../services/ResultService/ResultService';
import { ISynonymFieldConfiguration} from '../../models/ISynonymFieldConfiguration';
import IQueryModifierConfiguration from '../../models/IQueryModifierConfiguration';
import { IPagingSettings } from '../../models/IPagingSettings';
import { ISuggestionProviderDefinition } from '../../services/ExtensibilityService/ISuggestionProviderDefinition';
import IRefinerConfiguration from "../../models/IRefinerConfiguration";
import RefinersLayoutOption from "../../models/RefinersLayoutOptions";

export interface INbsSearchWebPartProps {
    /**
     * Search results properties
     */
    queryKeywords: string;
    defaultSearchQuery: string;
    useDefaultSearchQuery: boolean;
    queryTemplate: string;
    resultSourceId: string;
    sortList: ISortFieldConfiguration[];
    enableQueryRules: boolean;
    selectedProperties: string;
    sortableFields: ISortableFieldConfiguration[];
    showResultsCount: boolean;
    showBlank: boolean;
    selectedLayout: ResultsLayoutOption;
    externalTemplateUrl: string;
    inlineTemplateText: string;
    resultTypes: ISearchResultType[];
    rendererId: string;
    customTemplateFieldValues: ICustomTemplateFieldValue[];
    enableLocalization: boolean;
    useRefiners: boolean;
    useSearchBox: boolean;
    paginationDataSourceReference: string;
    synonymList: ISynonymFieldConfiguration[];
    searchQueryLanguage: number;
    templateParameters: { [key:string]: any };
    queryModifiers: IQueryModifierConfiguration[];
    selectedQueryModifierDisplayName: string;
    paging: IPagingSettings;
    /**
     * Search box properties
     */
    enableQuerySuggestions: boolean;
    placeholderText: string;
    suggestionProviders: ISuggestionProviderDefinition<any>[];
    /**
     * Search refiners
     */
    refinersConfiguration: IRefinerConfiguration[];
    refinersSelectedLayout: RefinersLayoutOption;
}
