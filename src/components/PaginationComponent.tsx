import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Pagination from        'react-js-pagination';
import { Icon } from 'office-ui-fabric-react';
import styles from './PaginationComponent.module.scss';
import { BaseWebComponent } from './BaseWebComponent';

export type PageUpdateCallback = (pageNumber: number) => void;

export interface IPaginationComponentProps {

    /**
     * The total number of items retrieved from the data source if known.
     */
    totalItems: string;
    
    /**
     * The number of items per page
     */
    itemsCountPerPage: string;

    /**
     * The current page number
     */
    currentPageNumber: string;

    /**
     * The page range to display 
     */
    range?: string;

    /**
     * Handler method called when a page is updated
     */
    onPageUpdate: PageUpdateCallback;

    /**
     * Hide first/last navigation buttons for the paging control
     */
    hideFirstLastPages: boolean;

    /**
     * Hide navigation buttons (prev page, next page)
     */
    hideNavigation: boolean;

    /**
     * Hide navigation buttons (prev, next, first, last) if they are disabled.
     */
    hideDisabled: boolean;

    /**
     * The pre-calculated page links.
     */
    pageLinks: string;

}

export interface IPaginationComponentState {

    /**
     * The current page number
     */
    currentPageNumber: number;
}

export class PaginationComponent extends React.Component<IPaginationComponentProps, IPaginationComponentState> {

    constructor(props: IPaginationComponentProps) {
        super(props);

        this.state = {
            currentPageNumber: parseInt(props.currentPageNumber)
        };
    }

    public render() {
        // Calculate the total items count
        let totalCount: number = parseInt(this.props.totalItems);
        let itemsCountPerPage: number = parseInt(this.props.itemsCountPerPage);
        

        // Show nothing if there is only one page
        if (totalCount > itemsCountPerPage) {

            return(
                <div className={styles.paginationContainer}>
                    <div className={styles.paginationContainer__paginationContainer}>
                        <div className={`${styles.paginationContainer__paginationContainer__pagination}`}>
                            <div className={`${styles.standard}`}>
                                <Pagination
                                    activePage={this.state.currentPageNumber}
                                    firstPageText={<Icon iconName='DoubleChevronLeft'/>}
                                    lastPageText={<Icon iconName='DoubleChevronRight'/>}
                                    prevPageText={<Icon iconName='ChevronLeft' />}
                                    nextPageText={<Icon iconName='ChevronRight'/>}
                                    activeLinkClass={ `${styles.active} ${''}`}  
                                    itemsCountPerPage={ this.props.itemsCountPerPage }
                                    totalItemsCount={ totalCount }
                                    hideDisabled={ this.props.hideDisabled ? this.props.hideDisabled : false  }
                                    hideNavigation={ this.props.hideNavigation ? this.props.hideNavigation : false }
                                    hideFirstLastPages={ this.props.hideFirstLastPages ? this.props.hideFirstLastPages : false }
                                    pageRangeDisplayed={ this.props.range ? this.props.range : 5 }
                                    onChange={((pageNumber: number) => {

                                        this.setState({
                                            currentPageNumber: pageNumber
                                        });

                                        this.props.onPageUpdate(pageNumber);
                                        
                                    }).bind(this)}
                                />    
                            </div>
                        </div>
                    </div>
                </div>
            );
        
        } else {
            return null;
        }
    }
}

export class PaginationWebComponent extends BaseWebComponent {
   
    public constructor() {
        super(); 
    }
 
    public async connectedCallback() {
    
        // Handle events       
        let props: any = this.resolveAttributes();
        const pagination = <PaginationComponent {...props} onPageUpdate={(pageNumber: number, pageLink: string) => {

            // Bubble event through the DOM
            this.dispatchEvent(new CustomEvent('pageNumberUpdated', { 
                detail: {
                    pageNumber: pageNumber,
                    pageLink: pageLink
                }, 
                bubbles: true,
                cancelable: true
            }));
        }}/>;
        ReactDOM.render(pagination, this);
    }    
}