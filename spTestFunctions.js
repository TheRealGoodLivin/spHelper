function spCreateDocumentLibFolder(listTitle, itemName, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listGetURL = siteURL + "/_api/lists/GetByTitle('" + listTitle + "')/RootFolder";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        var listPostURL = siteURL + "/_api/Web/Folders/add('" + data.d.ServerRelativeUrl + "/" + itemName + "')";
        var postOptions = {
            method: 'POST',
            headers: new Headers({
                'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose'
            }),
            credentials: 'include'
        };

        return fetch(listPostURL, postOptions);
    }).then(function(response) {
        if (!response.ok) {
            throw new Error('Error!');
        }
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}
