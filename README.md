# LICENSE
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


# spHelper
A Javascript library to assist with common/useful SharePoint REST List functions using pure javascript.

> Written with Visual Studio Code (https://code.visualstudio.com/) 
> - Author: Austin Livengood
> - Original Reference Author (JQuery Version): William Bechard (https://github.com/ury2ok2000/SPHelper)

# Notes
 > - Purely written in Javascript without the requirements of additional libraries.
 > - For **siteURL**, if the List is on the same site as your script, remove the variable and it will using the current site.

# Future Changes
 > - Custom Modal Dialog for custom list fields.
 > - Updating Groups for User

## How To Use
In SharePoint, the best place to store a .js file is within your Site Assets (Document Library) folder within your Site. You can use 'spHelper.js' or the 'spHelper.min.js'. Once you have uploaded the file, you just reference it like any other JavaScript library reference:

#### spHelper Main Use:
    <script type="text/javascript"  src="LOCATION/spHelper.js"></script>
    
#### spHelper Min Main Use:
    <script type="text/javascript"  src="LOCATION/spHelper.min.js"></script>

# Functions
## SharePoint -- Delete List:
    Without Site URL: spDeleteList('ListName').then(res => { console.log('List was deleted!') });
    With Site URL: spDeleteList('ListName', SiteURL).then(res => { console.log('List was deleted!') });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spDeleteList('ListName').then(res => { console.log('List was deleted!') });
        });
        
## SharePoint -- Delete List Item:
    Without Site URL: spDeleteListItem('ListName', ItemID).then(res => { console.log('Item was deleted!') });
    With Site URL: spDeleteListItem('ListName', ItemID, SiteURL).then(res => { console.log('Item was deleted!') });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spDeleteListItem('ListName', ItemID).then(res => { console.log('Item was deleted!') });
        });

## SharePoint -- Get Current User Data:
    spGetCurrentUser().then(res => { console.log(res) });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetCurrentUserId().then(res => { console.log(res) });
        });
        
## SharePoint -- Get Current User ID:
    spGetCurrentUserId();
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            console.log(spGetCurrentUserId());
        });

## SharePoint -- Get User Data By ID:
    spGetUserById(UserID).then(res => { console.log(res) });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetUserById(10).then(res => { console.log(res) });
        });

## SharePoint -- Grabs All Items From a List:
    Without Site URL: spGetListItems(ListName).then(res => { console.log(res) });
    With Site URL: spGetListItems(ListName, SiteURL).then(res => { console.log(res) });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetListColumns('List1').then(res => { console.log(res) });
        });

## SharePoint -- Create List Item:
    Without Site URL: spCreateListItem('ListName', itemProperties).then(res => { console.log('Item was created!') });
    With Site URL: spCreateListItem('ListName', ItemProperties, SiteURL).then(res => { console.log('Item was created!') });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spCreateListItem('test', itemProperties).then(res => { console.log('Item was created!') });
        });

## SharePoint -- Update List Item By ID:
    Without Site URL: spUpdateListItemById('ListName', ItemID, ItemProperties)).then(res => { console.log('Item was updated!') });
    With Site URL: spUpdateListItemById('ListName', ItemID, ItemProperties, SiteURL)).then(res => { console.log('Item was updated!') });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spUpdateListItemById('test', 2, itemProperties).then(res => { console.log('Item was updated!') });
        });

## SharePoint -- Get List Item by ID:
    Without Site URL: spGetListItemById('ListName', ItemID).then(res => { console.log(res) });
    With Site URL: spGetListItemById('ListName', ItemID, ItemProperties, SiteURL).then(res => { console.log(res) });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetListItemById('test', 2).then(res => { console.log(res) });
        });

## SharePoint -- Get List Columns (Title, Type, DisplayName, Required):
    Without Site URL: spGetListColumns('ListName').then(res => { console.log(res) });
    With Site URL: spGetListColumns('ListName', SiteURL).then(res => { console.log(res) });
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetListColumns('test').then(res => { console.log(res) });
        });

## SharePoint -- Get All List:
    Without Site URL: spGetListColumns();
    With Site URL: spGetListColumns(SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event) {
            spGetAllLists().then(res => { console.log(res) });
        });

## SharePoint -- Modal Dialog (URL, HTML):
    Without Site URL: spGetListColumns();
    With Site URL: spGetListColumns(SiteURL);
    
#### Use: 
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
