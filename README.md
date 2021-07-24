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
