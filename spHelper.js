/*
    DATE: 22 Jul 2021
    UPDATED: 23 Feb 2022
    
    MIT License

    Copyright (c) 2021 Austin Livengood

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    NOTE: if the script you are using is on the same site as the list, remove siteURL variable from function. It will grab the same site URL.
*/

/*
    DESCRIPTION: Deletes an entire list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
    DESCRIPTION: Grabs current user data using fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetCurrentUser().then(res => { console.log(res) });
        });
*/
function spGetCurrentUser() {
    var listGetURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser/?$expand=groups";
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
    DESCRIPTION: Grabs current user ID.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
            console.log(spGetCurrentUserId());
        });
*/
function spGetCurrentUserId() {
    var userId = _spPageContextInfo.userId;
    return userId;
}

/*
    DESCRIPTION: Grabs user by ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetUserById(10).then(res => { console.log(res) });
        });
*/
function spGetUserById(userId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var listGetURL = siteURL + "/_api/Web/SiteUserInfoList/Items?$filter=Id eq " + userId + "&$expand=groups";
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
        document.addEventListener("DOMContentLoaded", function(event) {
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
    DESCRIPTION: Create Document Library Folder

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spCreateDocumentLibFolder('Document Library', 'Test Folder').then(res => { console.log(res) });
        });
*/
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

/*
    DESCRIPTION: Create Document Library Folder

    USE: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spUploadDocumentLibItem('Document Library', 'Test Folder', document.querySelector('#fileSelectorHTML').files[0]).then(res => { console.log(res) });
        });
*/
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

    return new Promise((resolve, reject) => {
        var reader = new FileReader();
        reader.onload = function(e) { resolve(e.target.result); }
        reader.onerror = function(e) { reject(e.target.error); }
        reader.readAsArrayBuffer(itemFile);
    }).then(function (file) {
        return fetch(listGetURL, getOptions).then(response => response.json()).then(data => {
            var listPostURL = siteURL + "/_api/Web/GetFolderByServerRelativeUrl(@target)/Files/add(overwrite=true, url='" + itemFile.name + "')?@target='" + data.d.ServerRelativeUrl + "/" + itemLocation + "'&$expand=ListItemAllFields"; 
            var postOptions = {
                method: 'POST',
                body: file,
                async: false,
                binaryStringRequestBody: true,
                headers: new Headers({
                    'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
                    'Accept': 'application/json; odata=verbose'
                }),
                credentials: 'include'
            };
            return fetch(listPostURL, postOptions);
        }).then(function(response) {
            if(response.ok) return response.json();
            else throw new Error('Error!');
        }).catch(error => {
            console.error(error);
            return Promise.reject(error);
        });
    }).catch(error => {
        console.error(error);
        return Promise.reject(error);
    });
} 

/*
    DESCRIPTION: SharePoint Modal with URL or HTML

    NOTE: 
        -modalRefresh is if you want the main page to refresh when user clicks on save of the List Form
            --true=Page Refreshes
            --false=Page does not Refresh

    USE:
        HTML:
            document.addEventListener("DOMContentLoaded", function(event) {
                var htmlElement = document.createElement("p");
                var messageNode = document.createTextNode("Content of dialog");
                htmlElement.appendChild(messageNode);

                spModalOpen('Test', 500, 500, htmlElement, 'html');
            });

        URL:
            document.addEventListener("DOMContentLoaded", function(event) {
                spModalOpen('Test', 500, 500, 'https://google.com', 'url');
            });
*/
function spModalOpen(modalTitle, modalWidth, modalHeight, modalInformation, modalOption = 'url', modalAutoResize = false, modalRefresh = false, modalRedirect = false, modalRedirectOption = 'alert', modalRedirectInformation = '') {
    var options = {
        allowMaximize: false,
        showClose: true,
        autoSize: modalAutoResize,
        title: modalTitle,
        dialogReturnValueCallback: function dialogReturnValueCallback(dialogResult) {
            if (dialogResult != SP.UI.DialogResult.cancel) {
                if(modalRefresh) SP.UI.ModalDialog.RefreshPage(dialogResult);

                if(modalRedirect) {
                  if (modalRedirectOption === 'url') {
                    window.location.href = modalRedirectInformation;
                  } else if (modalRedirectOption === 'alert') {
                    alert(modalRedirectInformation);
                  }
                }
            }
        }
    };

    if (modalWidth === 0 && modalAutoResize === true) {
    } else if (modalWidth !== 0 && modalAutoResize === false) {
      options['width'] = modalWidth;
    } else if (modalWidth === 0 && modalAutoResize === false) {
      options['width'] = 100;
    }

    if (modalHeight === 0 && modalAutoResize === true) {
    } else if (modalHeight !== 0 && modalAutoResize === false) {
      options['height'] = modalHeight;
    } if (modalHeight === 0 && modalAutoResize === false) {
      options['height'] = 100;
    } 

    options[modalOption] = modalInformation;

    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
}
