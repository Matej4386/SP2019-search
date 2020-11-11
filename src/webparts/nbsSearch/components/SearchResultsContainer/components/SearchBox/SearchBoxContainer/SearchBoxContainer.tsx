import * as React from                               'react';
import { ISearchBoxContainerProps } from             './ISearchBoxContainerProps';
import * as strings from                             'NbsSearchWebPartStrings';
import ISearchBoxContainerState from                 './ISearchBoxContainerState';
import { MessageBar, MessageBarType } from           'office-ui-fabric-react/lib/MessageBar';
import styles from '../SearchBoxWebPart.module.scss';
import SearchBoxAutoComplete from '../SearchBoxAutoComplete/SearchBoxAutoComplete';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { IconButton } from 'office-ui-fabric-react/lib/Button';

export default class SearchBoxContainer extends React.Component<ISearchBoxContainerProps, ISearchBoxContainerState> {

  public constructor(props: ISearchBoxContainerProps) {

    super(props);

    this.state = {
      searchInputValue: (props.inputValue) ? decodeURIComponent(props.inputValue) : '',
      errorMessage: null,
      showClearButton: !!props.inputValue,
    };

    this._onSearch = this._onSearch.bind(this);
  }

  private renderSearchBoxWithAutoComplete(): JSX.Element {
    return (
      <SearchBoxAutoComplete
        inputValue={this.props.inputValue}
        onSearch={this._onSearch}
        placeholderText={this.props.placeholderText}
        suggestionProviders={this.props.suggestionProviders}
        domElement={this.props.domElement}
      />
    );
  }

  private renderBasicSearchBox(): JSX.Element {
    return (
      <div className={styles.searchBoxWrapper}>
        <SearchBox
          placeholder={this.props.placeholderText ? this.props.placeholderText : strings.SearchInputPlaceholder}
          className={ styles.searchTextField }
          value={ this.state.searchInputValue }
          autoComplete= "off"
          onChange={(value) => this.setState({ searchInputValue: value })}
          onSearch={() => this._onSearch(this.state.searchInputValue)}
          onClear={() => this._onSearch('', true)}
        />
        <div className={styles.searchButton}>
          {this.state.searchInputValue &&
            <IconButton
              onClick={() => this._onSearch(this.state.searchInputValue)}
              iconProps={{iconName: 'Search' }}
            />
          }
        </div>
      </div>
    );
  }

  /**
   * Handler when a user enters new keywords
   * @param queryText The query text entered by the user
   */
  public async _onSearch(queryText: string, isReset: boolean = false) {

    // Don't send empty value
    if (queryText || isReset) {

      this.setState({
        searchInputValue: queryText,
        showClearButton: !isReset
      });
      this.props.onSearch(queryText);
    }
  }


  public componentWillReceiveProps(nextProps: ISearchBoxContainerProps) {
    this.setState({
      searchInputValue: decodeURIComponent(nextProps.inputValue),
    });
  }

  public render(): React.ReactElement<ISearchBoxContainerProps> {
    let renderErrorMessage: JSX.Element = null;

    if (this.state.errorMessage) {
      renderErrorMessage = <MessageBar messageBarType={ MessageBarType.error }
                                        dismissButtonAriaLabel='Close'
                                        isMultiline={ false }
                                        onDismiss={ () => {
                                          this.setState({
                                            errorMessage: null,
                                          });
                                        }}
                                        className={styles.errorMessage}>
                                        { this.state.errorMessage }</MessageBar>;
    }

    const renderSearchBox = this.props.enableQuerySuggestions ?
                          this.renderSearchBoxWithAutoComplete() :
                          this.renderBasicSearchBox();
    return (
      <div className={styles.searchBox}>
        { renderErrorMessage }
        { renderSearchBox }
      </div>
    );
  }
}
