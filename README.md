# SP2019-search
SP2019 Search webpart (searchBox, Refiners, searchResults)

SP2019 on-prem version for https://microsoft-search.github.io/pnp-modern-search/

This solution consists of 1 webpart - it renders Search results webpart which was modified to handle search box web part, filters web part and pagination from original solution.
Main webpart renders SearchResultsContainer with components SearchBox and SearchRefiners. Due to this you cant turn off Search results and have for example just search box on page.
Due to the orginal version which is using Dynamic connectors - i needed to make some variables (like query string from search box, filters, ...) part of state of search results. 

New version:
1. support search box, filters, pagination, search results with templates
2. Search box, filters, pagination can be turned off
3. Search localization is turned of (settings are commented in code - you can uncomment it and test it on your enviroment- on my enviroment it caused error (GUID))
4. Theme support was deleted - it is not supported on SP2019. I am using separate extension to set my custom theme.
5. Search results templates: Simple list, Details list, Tiles, Carousel, People, Debug, Custom
6. In case you want your own template please see documation here: https://microsoft-search.github.io/pnp-modern-search/

Be aware - i am still testing/developing/updating this web part (in case of time), some changes are not fancy and not nice but i did my best. There is still bunch of code that can be deleted/modified/... Localization is for english and I think that there can be some strings that are not used :-) Localizaton is in one file.

This webpart is mainly for at least a little bit experienced user with ability to build, package solution and maybe made some changes in this solution.
Please leave a comment with any suggestion. If you need more information about this solution please let me know - i will try to responde as soon as possible.





