# spHelper
A Javascript library to assist with common/useful SharePoint REST List functions using pure javascript.

> Written with Visual Studio Code (https://code.visualstudio.com/) 
> - Author: Austin Livengood
> - Original Reference with JQuery: William Bechard (https://github.com/ury2ok2000/SPHelper)

## How To Use
In SharePoint, the best place to store a .js file is within your Site Assets (Document Library) folder within your Site. Once you have uploaded the file, you just reference it like any other JavaScript library reference:

    <script type="text/javascript"  src="LOCATION/spHelper.js"></script>
   
# Notes
 > - Purely written in Javascript without the requirements of additional libraries.
 > - For **siteURL**, if the List is on the same site as your script, remove the variable and it will using the current site.

# Functions
## SharePoint -- Delete List:
    Without Site URL: spDeleteList(ListName);
    With Site URL: spDeleteList(ListName, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteList('Documents');
        });
        
## SharePoint -- Delete List Item:
    Without Site URL: spDeleteListItem(ListName, ItemID);
    With Site URL: spDeleteListItem(ListName, ItemID, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            spDeleteListItem('test', '1');
        });

## SharePoint -- Get Current User ID:
    spGetCurrentUserId();
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var currentUserID = spGetCurrentUserId();
        });

## SharePoint -- Get User Data By ID:
    spGetUserById(UserID);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var userData = spGetUserById(10);
        });

## SharePoint -- Grabs All Items From a List:
    Without Site URL: spGetListItems(ListName, ItemID);
    With Site URL: spGetListItems(ListName, ItemID, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listItems = spGetListItems('test');
        });

## SharePoint -- Create List Item:
    Without Site URL: spCreateListItem(ListName, ItemProperties);
    With Site URL: spCreateListItem(ListName, ItemProperties, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spCreateListItem('test', itemProperties);
        });

## SharePoint -- Update List Item By ID:
    Without Site URL: spUpdateListItemById(ListName, ItemID, ItemProperties);
    With Site URL: spUpdateListItemById(ListName, ItemID, ItemProperties, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var itemProperties = {};
            itemProperties['Title'] = 'test';

            spUpdateListItemById('test', 2, itemProperties);
        });

## SharePoint -- Get List Item by ID:
    Without Site URL: spGetListItemById(ListName, ItemID);
    With Site URL: spGetListItemById(ListName, ItemID, ItemProperties, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listItem = spGetListItemById('test', 2);
        });

## SharePoint -- Get List Columns (Title, Type, DisplayName, Required):
    Without Site URL: spGetListColumns(ListName);
    With Site URL: spGetListColumns(ListName, SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listItem = spGetListItemById('test', 2);
        });

## SharePoint -- Get All List:
    Without Site URL: spGetListColumns();
    With Site URL: spGetListColumns(SiteURL);
    
#### Use: 
        document.addEventListener("DOMContentLoaded", function(event){
            var listColumns = spGetAllLists('test');
        });

## SharePoint -- Modal Dialog (URL, HTML):
    Without Site URL: spGetListColumns();
    With Site URL: spGetListColumns(SiteURL);
    
#### Use: 
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





