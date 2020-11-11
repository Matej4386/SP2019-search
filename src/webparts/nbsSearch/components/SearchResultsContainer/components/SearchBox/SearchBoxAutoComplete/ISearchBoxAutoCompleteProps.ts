import { ISuggestionProviderInstance } from '../../../../../../../services/ExtensibilityService/ISuggestionProviderInstance';

export interface ISearchBoxAutoCompleteProps {
  placeholderText: string;
  suggestionProviders: ISuggestionProviderInstance<any>[];
  inputValue: string;
  onSearch: (queryText: string, isReset?: boolean) => void;
  domElement: HTMLElement;
}
