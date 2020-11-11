import { PageOpenBehavior, QueryPathBehavior } from '../../../../../../../helpers/UrlHelper';
import ISearchService from       '../../../../../../../services/SearchService/ISearchService';
import { ISuggestionProviderInstance } from '../../../../../../../services/ExtensibilityService/ISuggestionProviderInstance';

export interface ISearchBoxContainerProps {
    onSearch: (searchQuery: string) => void;
    enableQuerySuggestions: boolean;
    suggestionProviders: ISuggestionProviderInstance<any>[];
    searchService: ISearchService;
    inputValue: string;
    placeholderText: string;
    domElement: HTMLElement;
}
