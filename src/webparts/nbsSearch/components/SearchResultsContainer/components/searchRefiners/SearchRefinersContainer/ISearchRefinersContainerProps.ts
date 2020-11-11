import { IRefinementResult, IRefinementFilter, IRefinementValue } from "../../../../../../../models/ISearchResult";
import IRefinerConfiguration from "../../../../../../../models/IRefinerConfiguration";
import { DisplayMode } from "@microsoft/sp-core-library";
import RefinersLayoutOption from "../../../../../../../models/RefinersLayoutOptions";
import IUserService from '../../../../../../../services/UserService/IUserService';

export interface ISearchRefinersContainerProps {
  
  /**
   * Default selected refinement filters
   */
  defaultSelectedRefinementFilters: IRefinementFilter[];

  /**
   * List of available refiners from the connected search results Web Part
   */
  availableRefiners: IRefinementResult[];

  /**
   * The Web Part refiners configuration
   */
  refinersConfiguration: IRefinerConfiguration[];

  /**
   * The selected layout
   */
  refinersSelectedLayout: RefinersLayoutOption;

  /**
   * Handler method when a filter value is updated in children components
   */
  onUpdateFilters: (filters: IRefinementFilter[]) => void;

  /**
   * Indicates if we should show blank if no refinement result
   */
  showBlank: boolean;

  /**
   * The current UI language
   */
  language: string;

  /**
   * The current search query
   */
  query: string;

  /**
   * UserService
   */
  userService: IUserService;
}
