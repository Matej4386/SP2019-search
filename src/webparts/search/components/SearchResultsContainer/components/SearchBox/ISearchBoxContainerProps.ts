import { PageOpenBehavior, QueryPathBehavior } from '../../../../../../helpers/UrlHelper';
import ISearchService from       '../../../../../../services/SearchService/ISearchService';

export interface ISearchBoxContainerProps {
    onSearch: (searchQuery: string) => void;
    searchInNewPage: boolean;
    enableQuerySuggestions: boolean;
    searchService: ISearchService;
    pageUrl: string;
    openBehavior: PageOpenBehavior;
    queryPathBehavior: QueryPathBehavior;
    queryStringParameter: string;
    inputValue: string;
    placeholderText: string;
    domElement: HTMLElement;
}
