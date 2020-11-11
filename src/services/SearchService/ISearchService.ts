import { ISearchResults, } from '../../models/ISearchResult';
import { ISearchServiceConfiguration } from '../../models/ISearchServiceConfiguration';
import IManagedPropertyInfo from '../../models/IManagedPropertyInfo';

interface ISearchService extends ISearchServiceConfiguration {
    /**
     * Perfoms a search query.
     * @returns ISearchResults object. Use the 'RelevantResults' property to acces results proeprties (returned as key/value pair object => item.[<Managed property name>])
     */
    search(kqlQuery: string, pageNumber?: number, useOldSPIcons?: boolean): Promise<ISearchResults>;

    /**
     * Retrieves search query suggestions
     * @param query the term to suggest from
     */
    suggest(query: string): Promise<string[]>;

    /**
     * Retrieve the configuration of the search service
     */
    getConfiguration(): ISearchServiceConfiguration;

    /**
     * Gets available search managed properties in the search schema
     */
    getAvailableManagedProperties(): Promise<IManagedPropertyInfo[]>;

    /**
     * Checks if the provided manage property is sortable or not
     * @param property the managed property to verify
     */
    validateSortableProperty(property: string): Promise<boolean>;

    /**
     * Gets all available languages for the search query
     */
    getAvailableQueryLanguages(): Promise<any[]>;
}

 export default ISearchService;