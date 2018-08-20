# SharePointDirect CLI

SharePointDirect is a command line interface that can be used by Robotic Process Automation solutions that do not have the capability to connect to SharePoint Online directly and needs to execute some Sharepoint actions (see functionalities/methods available below).

## Built with
* [C#](https://maven.apache.org/) - Language used
* [CSOM](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-client-library-code) - API used for accessing Sharepoint Online
* [SharePointPnPCoreOnline](https://www.nuget.org/packages/SharePointPnPCoreOnline/) - Extension methods for CSOM

## Methods

### All methods writes the result to the console or a textfile or to a csv file.

* **GetNumberOfItems** - Gets the number of items in the provided list. Writes the result to **_C:\Apps\GetNumberOfItems.txt_**
* **GetItemId** - Gets the ID of the item by title. Writes the result to **_C:\Apps\GetItemId.txt_**
* **AddItem** - adds an item to the list Writes the result to **_C:\Apps\AddItem.txt_**
* **DeleteItemById** - deletes an item from the list using a provided ID.  Writes the result to **_C:\Apps\DeleteItemById.txt_**
* **UploadFileWithMeta** - uploads a file to a specified library and adds the given properties.  Writes the result to **_C:\Apps\UploadFileWithMeta.txt_**. Please note that the current version does not work when uploading to sub folders.
* **UploadFileNoMeta** - uploads a file to a specified library with no properties. Writes the result to **_C:\Apps\UploadFileNoMeta.txt_**. Please note that the current version does not work when uploading to sub folders.
* **GetOneIte**m - Gets an item from the list based on a given title and writes the requested field name and field value to a txt file. Writes the result to **_C:\Apps\GetOneItem.txt_**.
* **GetOldestItem** - Gets the oldest item from the list. When successful, writes the result to **_C:\Apps\GetOldestItem.csv_**.

## Usage:

```
copy the debug/release folder to any destination then rename it

cd [renamed folder]

SharePointDirect Method args[]
```

In your Robotic Process Automation solution, execute the program via the command prompt with the required arguments

## Examples:

#### GetNumberOfItems:
```
SharePointDirect GetNumberOfItems URL Listname
```

#### GetItemId:
```
SharePointDirect GetItemId URL Listname ItemTitle
```

#### AddItem:
```
SharePointDirect AddItem URL Listname Field Value
```

#### DeleteItemById:
```
SharePointDirect DeleteItemById URL Listname ID
```
#### UploadFileWithMeta:
```
SharePointDirect UploadFileWithMeta URL Listname Filepath PropertyKey, PropertyValue...
```

#### UploadFileNoMeta:
```
SharePointDirect UploadFileNoMeta URL Listname Filepath
```

#### GetOneItem:
```
SharePointDirect GetOneItem URL Listname FieldName1 FieldName2...
```

#### GetOldestItem:
```
SharePointDirect GetOldestItem URL Listname
```

## Authors

* **Exx Navarro** - *Initial work* - [exxmen](https://github.com/exxmen)

See also the list of [contributors](https://github.com/exxmen/sharepointdirect/graphs/contributors) who participated in this project.

## Contributing
Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct, and the process for submitting pull requests to us.

Please note that you need to have Visual Studio and .NET 4.1 installed on your machine to contribute

```
git clone https://github.com/exxmen/sharepointdirect.git

cd SharePointDirect

Build Solution in Visual Studio
```

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone whose code was used. If you want to be mentioned, please let me know.
* Inspired by Marcin's awesome Excel SharePoint REST API based tool
