/*
    AUTHOR: Austin Livengood
    DATE: 22 Jul 2021

    NOTE: if the script you are using is on the same site as the list, remove siteURL variable from function. It will use the same site URL.
*/

/*
    DESCRIPTION: Deletes an entire list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteList('test');
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

    fetch(listPostURL, postOptions).then(function(response) {
        if (response.ok) {
            console.log('Response was good!');
        } else {
            throw new Error('Error');
        }
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Deletes an item from a list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteListItem('test', '1');
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

    fetch(listPostURL, postOptions).then(function(response) {
        if (response.ok) {
            console.log('Response was good!');
        } else {
            throw new Error('Error!');
        }
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Grabs current user ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var currentUserId = spGetCurrentUserId();
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

    fetch(listGetURL, getOptions).then(function(response) {
        if (response.ok) {
            console.log('Response was good!');
            return response.json();
        } else {
            throw new Error('Error!');
        }
    }).then(function(data) {
        console.log('Success: Current User ID ' + data.d.Id + '.');
        return data.d.Id;
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Grabs user by ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var userData = spGetUserById(10);
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

    fetch(listGetURL, getOptions).then(function(response) {
        if (response.ok) {
            console.log('Response was good!');
            return response.json();
        } else {
            throw new Error('Error!');
        }
    }).then(function(data) {
        return data;
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Grabs all items from a List

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listItems = spGetListItems('test');
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
    spGetItems();

    function spGetItems() {
        fetch(listGetURL, getOptions).then(function(response) {
            if (response.ok) {
                console.log('Response was good!');
                return response.json();
            } else {
                throw new Error('Error!');
            }
        }).then(function(data) {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                return results;
            }
        }).catch(function(error) {
            console.error(error);
        });
    }
}

/*
    DESCRIPTION: Create list item

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spCreateListItem('test', itemProperties);
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
    
    fetch(listGetURL, getOptions).then(function(response) {
        if (response.ok) {
            return response.json();
        } else {
            throw new Error('Error!');
        }
    }).then(function(data) {
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
        if (response.ok) {
            console.log('Response was good!');
        } else {
            throw new Error('Error!');
        }
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Update list item by id

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spUpdateListItemByID('test', 2, itemProperties);
        });
*/
function spUpdateListItemByID(listTitle, itemId, itemProperties, siteURL = _spPageContextInfo.webAbsoluteUrl) {
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

    fetch(listGetURL, getOptions).then(function(response) {
        if (response.ok) {
            return response.json();
        } else {
            throw new Error('Error!');
        }
    }).then(function(data) {
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
        if (response.ok) {
            console.log('Response was good!');
        }
        else {
            throw new Error('Error!');
        }
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Get List Item by Id

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listItem = spGetListItemByID('test', 2);
        });
*/
function spGetListItemByID(listTitle, itemId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";
    var getOptions = {
        method: 'GET',
        headers: new Headers({
            'Accept': 'application/json; odata=verbose',
            'content-type': 'application/json; odata=verbose'
        })
    };

    fetch(listGetURL, getOptions).then(function(response) {
        if (response.ok) {
            console.log('Response was good!');
            return response.json();
        } else {
            throw new Error('Error!');
        }
    }).then(function(data) {
        return data;
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Get List Column Names and Column Types

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listColumns = spGetListColumns('test');
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
    spGetItems();

    function spGetItems() {
        fetch(listGetURL, getOptions).then(function(response) {
            if (response.ok) {
                console.log('Response was good!');
                return response.json();
            } else {
                throw new Error('Error!');
            }
        }).then(function(data) {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                return results;
            }
        }).catch(function(error) {
            console.error(error);
        });
    }
}

/*
    DESCRIPTION: Get All Site Lists

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listColumns = spGetAllLists('test');
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
    spGetItems();

    function spGetItems() {
        fetch(listGetURL, getOptions).then(function(response) {
            if (response.ok) {
                console.log('Response was good!');
                return response.json();
            } else {
                throw new Error('Error!');
            }
        }).then(function(data) {
            results = results.concat(data.d.results);
            if (data.d.__next) {
                listGetURL = data.d.__next;
                spGetItems();
            } else {
                console.log(results);
                return results;
            }
        }).catch(function(error) {
            console.error(error);
        });
    }
}

/*
    DESCRIPTION: SharePoint Modal with URL or HTML

    NOTE: 
        -refreshOnSave is if you want the main page to refresh when user clicks on save of the List Form
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
function spModalOpen(modalTitle, modalWidth, modalHeight, modalInformation, modalOption = 'url', modalRefresh = 'false') {
    var options = {
        allowMaximize: false,
        showClose: true,
        width: modalWidth,
        height: modalHeight,
        title: modalTitle,
        dialogReturnValueCallback: function dialogReturnValueCallback(dialogResult) {
            if (dialogResult != SP.UI.DialogResult.cancel) {
                if(refreshOnSave) SP.UI.ModalDialog.RefreshPage(dialogResult);
            }
        }
    };
    options[modalOption] = modalInformation;

    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
}