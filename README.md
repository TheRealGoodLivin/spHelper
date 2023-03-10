# MIT License

Copyright (c) 2021 Austin L

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
All code examples have been placed above the functions in the JS main file. Examples have been removed in min file.
