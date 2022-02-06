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

function spUploadDocumentLibItem(listTitle, itemLocation = '', itemFile, siteURL = _spPageContextInfo.webAbsoluteUrl) {
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
        return new Promise((resolve, reject) => {
            let fr = new FileReader();
            fr.onload = x=> resolve(fr.result);
            fr.readAsArrayBuffer(itemFile); // fr.readAsDataURL(file);
	    }).then(function (results) {
            //var listPostURL = siteURL + "/_api/web/GetFolderByServerRelativeUrl('" + data.d.ServerRelativeUrl + "/" + itemLocation + "')/Files/add(url='" + results.file.name + "',overwrite=true)";
            var listPostURL = siteURL + "/_api/Web/GetFolderByServerRelativeUrl(@target)/Files/add(overwrite=true, url='" + results.file.name + "')?@target='" + data.d.ServerRelativeUrl + "/" + itemLocation + "'&$expand=ListItemAllFields"; 
            var postOptions = {
                method: 'POST',
                data: results,
                async: false,
                processData: false,
                headers: new Headers({
                    'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                    'Accept': 'application/json; odata=verbose',
                    'Content-Type': 'application/json; odata=verbose',
                    'Content-Length': results.byteLength
                }),
                credentials: 'include'
            };
            return fetch (listPostURL, postOptions);
        });
    }).then(function (response) {
        if (response.ok) return response.json();
        else throw new Error('Error!');
    }).catch(error => {
        console.error (error);
        return Promise.reject (error);
    });
}
