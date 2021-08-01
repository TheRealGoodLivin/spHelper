/*
    AUTHOR: Austin Livengood
    DATE: 22 Jul 2021

    NOTE: if the script you are using is on the same site as the list, remove siteURL variable from function. It will grab the same site URL.
*/

/*
    DESCRIPTION: Deletes an entire list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteList('test').then(res => { console.log('List was deleted!') });
        });
*/
function spDeleteList(listTitle, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listPostURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/";
    var postOptions = {
        method: 'POST',
        headers: new Headers({
            'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose',
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*'
        }),
        credentials: 'include'
    };

    return fetch(listPostURL, postOptions).then(function(response) {
        if (!response.ok) {
            throw new Error('Error');
        }
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}

/*
    DESCRIPTION: Deletes an item from a list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteListItem('test', '1').then(res => { console.log('Item was deleted!') });
        });
*/
function spDeleteListItem(listTitle, itemId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listPostURL = siteURL +"/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";
    var postOptions = {
        method: 'POST',
        headers: new Headers({
            'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose',
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*'
        }),
        credentials: 'include'
    };

    return fetch(listPostURL, postOptions).then(function(response) {
        if (!response.ok) {
            throw new Error('Error!');
        }
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}

/*
    DESCRIPTION: Grabs current user ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetCurrentUserId().then(res => { console.log(res) });
        });
*/
function spGetCurrentUserId() {
    var listGetURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        return data.d.Id;
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}

/*
    DESCRIPTION: Grabs user by ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetUserById(10).then(res => { console.log(res) });
        });
*/
function spGetUserById(userId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/Web/SiteUserInfoList/Items?$filter=Id eq " + userId;
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        return data;
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}

/*
    DESCRIPTION: Grabs all items from a List

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetListColumns('List1').then(res => { console.log(res) });
        });
*/
function spGetListItems(listTitle, listParameters = '', siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listGetURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items" + listParameters;
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    var results = [];
    return spGetItems();

    function spGetItems() {
        return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                return results;
            }
        }).catch(error => {
            console.error(error);
            return Promise.reject(error);
        });
    }
}

/*
    DESCRIPTION: Create list item

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spCreateListItem('test', itemProperties).then(res => { console.log('Item was created!') });
        });
*/
function spCreateListItem(listTitle, itemProperties, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listGetURL = siteURL + "/_api/lists/GetByTitle('" + listTitle + "')?$select=ListItemEntityTypeFullName";
    var listPostURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items";
    var getOptions = { 
        method: 'GET',
        headers: new Headers({
            'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
            'Accept': 'application/json; odata=verbose'
        }),
        credentials: 'include'
    }
    
    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        var itemPayload = {
            __metadata: {
                type: data.d.ListItemEntityTypeFullName
            }
        }
        for (var prop in itemProperties) {
            itemPayload[prop] = itemProperties[prop];
        }

        var postOptions = {
            method: 'POST',
            headers: new Headers({
                'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose'
            }),
            credentials: 'include',
            body: JSON.stringify(itemPayload)
        }
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

/*
    DESCRIPTION: Update list item by id

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spUpdateListItemById('test', 2, itemProperties).then(res => { console.log('Item was updated!') });
        });
*/
function spUpdateListItemById(listTitle, itemId, itemProperties, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var listGetURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";
    var listPostURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";
    var getOptions = { 
        method: 'GET',
        headers: new Headers({
            'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
            'Accept': 'application/json; odata=verbose'
        }),
        credentials: 'include'
    }

    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        var itemPayload = {
            __metadata: {
                type: data.d.__metadata.type
            }
        }
        for (var prop in itemProperties) {
            itemPayload[prop] = itemProperties[prop];
        }
        
        var postOptions = {
            method: 'POST',
            headers: new Headers({
                'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                'Accept': 'application/json; odata=verbose',
                'Content-Type': 'application/json; odata=verbose',
                'X-HTTP-Method': 'MERGE',
                'If-Match': data.d.__metadata.etag
            }),
            credentials: 'include',
            body: JSON.stringify(itemPayload)
        }
        
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

/*
    DESCRIPTION: Get List Item by Id

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetListItemById('test', 2).then(res => { console.log(res) });
        });
*/
function spGetListItemById(listTitle, itemId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
        return data;
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
}

/*
    DESCRIPTION: Get List Column Names and Column Types

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetListColumns('test').then(res => { console.log(res) });
        });
*/
function spGetListColumns(listTitle, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/web/lists/getbytitle('" + listTitle + "')/fields?$select=Title,TypeAsString,TypeDisplayName,Required&$filter=InternalName eq 'Title' and Hidden eq false and ReadOnlyField eq false";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    var results = [];
    return spGetItems();

    function spGetItems() {
        return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                return results;
            }
        }).catch(function(error) {
            console.error(error);
            return Promise.reject(error);
        });
    }
}

/*
    DESCRIPTION: Get All Site Lists

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetAllLists().then(res => { console.log(res) });
        });
*/
function spGetAllLists(siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/web/lists";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    var results = [];
    return spGetItems();

    function spGetItems() {
        return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                return results;
            }
        }).catch(error => {
            console.error(error);
            return Promise.reject(error);
        });
    }
}

/*
    DESCRIPTION: SharePoint Modal with URL or HTML

    NOTE: 
        -modalRefresh is if you want the main page to refresh when user clicks on save of the List Form
            --true=Page Refreshes
            --false=Page does not Refresh

    USE:
        HTML:
            document.addEventListener("DOMContentLoaded", function(event){
                var htmlElement = document.createElement("p");
                var messageNode = document.createTextNode("Content of dialog");
                htmlElement.appendChild(messageNode);

                spModalOpen('Test', 500, 500, htmlElement, 'html');
            });

        URL:
            document.addEventListener("DOMContentLoaded", function(event){
                spModalOpen('Test', 500, 500, 'https://google.com', 'url');
            });
*/
function spModalOpen(modalTitle, modalWidth, modalHeight, modalInformation, modalOption = 'url', modalAutoResize = false, modalRefresh = false) {
    var options = {
        allowMaximize: false,
        showClose: true,
        autoSize: modalAutoResize,
        width: modalWidth,
        height: modalHeight,
        title: modalTitle,
        dialogReturnValueCallback: function dialogReturnValueCallback(dialogResult) {
            if (dialogResult != SP.UI.DialogResult.cancel) {
                if(modalRefresh) SP.UI.ModalDialog.RefreshPage(dialogResult);
            }
        }
    };
    options[modalOption] = modalInformation;

    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
}
