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

function spUploadDocumentLibItem (listTitle, itemLocation = '', itemFiles, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest (_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listGetURL = siteURL + "/_api/lists/GetByTitle('" + listTitle + "')/RootFolder";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    return fetch (listGetURL, getOptions).then(response => response.json()).then(data => {
        return Promise.all([].map.call(itemFiles, function (file) {
            return new Promise(function (resolve, reject) {
                var readFile = new FileReader();
                readFile.onloadend = function () {
                    resolve({ result: readFile.result, file: file });
                };
                reader.readAsArrayBuffer(file);
            });
        })).then (function (results) {
            return results.forEach (function (result) {
                var listPostURL = siteURL + "/_api/web/GetFolderByServerRelativeUrl('" + data.d.ServerRelativeUrl + "/" + itemLocation + "')/Files/add(url='" + result.file.name + "',overwrite=true)";
                var postOptions = {
                    method: 'POST',
                    data: result,
                    async: false,
                    processData: false,
                    headers: new Headers({
                        'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                        'Accept': 'application/json; odata=verbose',
                        'Content-Type': 'application/json; odata=verbose',
                        'Content-Length': result.byteLength
                    }),
                    credentials: 'include'
                };
                return fetch (listPostURL, postOptions);
            });
        });
    }).then(function (response) {
        if (!response.ok) {
            throw new Error('Error!');
        }
    }).catch(error => {
        console.error (error);
        return Promise.reject (error);
    });
}
