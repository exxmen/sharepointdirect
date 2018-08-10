# SharePoint Direct

### SharePointDirect is a command line interface that can be used by Robotic Process Automation solutions that do not have the capability to connect to SharePoint Online directly and needs to execute some Sharepoint actions (see functionality below).

### The CLI was created with C# and uses the CSOM API for .NET.

### git clone https://github.com/exxmen/sharepointdirect.git if you would like to contribute

### Functions

* GetNumberOfItems - Gets the number of items in the provided list.
* GetItemId - Gets the ID of the item by title
* AddItem - adds an item to the list
* DeleteItemById - deletes an item from the list using a provided ID
* UploadFileWithMeta - uploads a file to a specified library and adds the given properties
* UploadFileWithMeta - uploads a file to a specified library with no properties
* GetOneItem - Gets an item from the list based on a given title and writes the result to a csv file

* All functions writes the result to the console and a textfile/csv file.

### Usage:

*SharePointDirect args[]*

### Examples:

#### GetNumberOfItems:
*SharePointDirect GetNumberOfItems URL Listname*

#### GetItemId:
*SharePointDirect GetItemId URL Listname ItemTitle*

#### AddItem:
*SharePointDirect AddItem URL Listname Field Value*

#### DeleteItemById:
*SharePointDirect DeleteItemById URL Listname ID*

#### UploadFileWithMeta:
*SharePointDirect UploadFileWithMeta URL Listname Filepath PropertyKey, PropertyValue...*

#### UploadFileNoMeta:
*SharePointDirect UploadFileNoMeta URL Listname Filepath*

#### GetOneItem:
*SharePointDirect GetOneItem URL Listname Title*