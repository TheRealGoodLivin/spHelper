/*
    AUTHOR: Austin Livengood
    DATE: 22 Jul 2021

    NOTE: if the script you are using is on the same site as the list, leave siteURL variable blank.
*/

/*
    DESCRIPTION: Deletes an entire list using Fetch. Use with caution.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteList('test');
        });
*/
function spDeleteList(listTitle, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    siteURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/";

    var postHeaders = new Headers({
        'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
        'Accept': 'application/json; odata=verbose',
        'content-type': 'application/json; odata=verbose',
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*'
    });

    var postOptions = {
        method: 'POST',
        headers: postHeaders,
        credentials: 'include'
    };

    fetch(siteURL, postOptions).then(function(response) {
        if (response.ok) {
            console.log('Success: ' + listTitle + ' has been removed.')
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
    siteURL = siteURL + "/_api/web/lists/GetByTitle('" + listTitle + "')/items(" + itemId.toString() + ")";

    var postHeaders = new Headers({
        'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value,
        'Accept': 'application/json; odata=verbose',
        'content-type': 'application/json; odata=verbose',
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*'
    });

    var postOptions = {
        method: 'POST',
        headers: postHeaders,
        credentials: 'include'
    };

    fetch(siteURL, postOptions).then(function(response) {
        if (response.ok) {
            console.log('Success: Item ' + itemId + ' has been removed from '  + listTitle + '.')
        } else {
            throw new Error('Error');
        }
    }).catch(function(error) {
        console.error(error);
    });
}

/*
    DESCRIPTION: Grabs current user ID using Fetch.

    USE: 
        document.addEventListener("DOMContentLoaded", function(event){
            spGetCurrentUserId();
        });
*/
function spGetCurrentUserId() {
    var siteURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser";

    var getHeaders = new Headers({
        'Accept': 'application/json; odata=verbose',
        'content-type': 'application/json; odata=verbose'
    });

    var getOptions = {
        method: 'GET',
        headers: getHeaders
    };

    fetch(siteURL, getOptions).then(response => response.json()).then(function(data) {
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
            spGetUserById(10);
        });
*/
function spGetUserById(userId, siteURL = _spPageContextInfo.webAbsoluteUrl) {
    var siteURL = siteURL + "/_api/Web/SiteUserInfoList/Items?$filter=Id eq " + userId;

    var getHeaders = new Headers({
        'Accept': 'application/json; odata=verbose',
        'content-type': 'application/json; odata=verbose'
    });

    var getOptions = {
        method: 'GET',
        headers: getHeaders
    };

    fetch(siteURL, getOptions).then(response => response.json()).then(function(data) {
        console.log('Success: Got User by ID ' + userId + '.');
        return data;
    }).catch(function(error) {
        console.error(error);
    });
}