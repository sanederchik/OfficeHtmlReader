# OfficeHtmlReader

This Office app is designed to insert HTML content into Excel sheets / PowerPoint slides in order to make these apps more interactive.

### Installing

1. Open an `Insert tab` in your Excel/PowerPoint file
2. Click a `Get Add-ins` button and choose a `Store` tab
3. In a `Search` field find a `WebPage Loader` app and click `Add`

### Using

If you want to insert a **local HTML** file:

1. Create a `WebPage Loader` container by clicking `Insert` -> `My Add-ins` -> `WebPage Loader`
2. Click a `Local file` button
3. Click a `Choose file` button and choose your file from local computer storage

If you want to insert a **web page url**:

1. Create a `WebPage Loader` container by clicking `Insert` -> `My Add-ins` -> `WebPage Loader`
2. Click a `Website` button and insert your URL. Notice that a `AJAX` request with `'X-Requested-With': 'XMLHttpRequest'` is used to retrieve data
3. Click a `Insert URL` button

## License

This project is licensed under the Apache 2.0 License - see the [LICENSE](https://github.com/sanederchik/OfficeHtmlReader/blob/master/LICENSE) file for details
