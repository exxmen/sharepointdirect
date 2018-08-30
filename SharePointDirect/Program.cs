using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace SpOnlineDirectConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

                string version;

                version = fileVersionInfo.ProductVersion;

                if (args.Length != 0)
                {
                    string method = args[0].ToLower();

                    if (!Directory.Exists(@"C:\Apps"))
                    {
                        Directory.CreateDirectory(@"C:\Apps");
                    }

                    switch (method)
                    {
                        case "getnumberofitems":
                            GetNumberOfItems(args[1], args[2]);
                            break;
                        case "getitemid":
                            GetItemId(args[1], args[2], args[3]);
                            break;
                        case "additem":
                            var addItemMap = new Dictionary<string, string>();
                            for (int i = 3; i <= args.Length - 1; i += 2)
                            {
                                addItemMap.Add(args[i], args[i + 1]);
                            }
                            AddItem(args[1], args[2], addItemMap);
                            break;
                        case "deleteitembyid":
                            int itemId = 0;
                            Int32.TryParse(args[3], out itemId);
                            DeleteItemById(args[1], args[2], itemId);
                            break;
                        case "uploadfilewithmeta":
                            var propertiesMap = new Dictionary<string, string>();
                            for (int i = 4; i <= args.Length - 1; i += 2)
                            {
                                propertiesMap.Add(args[i], args[i + 1]);
                            }
                            UploadFileWithMeta(args[1], args[2], args[3], propertiesMap);
                            break;
                        case "getoneitem":
                            var fieldsToReturn = new List<string>();
                            for (int i = 4; i < args.Length; i++)
                            {
                                fieldsToReturn.Add(args[i]);
                            }
                            GetOneItem(args[1], args[2], args[3], fieldsToReturn);
                            break;
                        case "getoldestitem":
                            GetOldestItem(args[1], args[2]);
                            break;
                        case "uploadfilenometa":
                            UploadFileNoMeta(args[1], args[2], args[3]);
                            break;
                        case "-v":
                        case "--version":
                            Console.WriteLine(version);
                            Console.WriteLine("Press any key to exit. ");
                            Console.ReadKey();
                            break;
                        case "-h":
                        case "--help":
                            Console.WriteLine(" ");
                            Console.WriteLine("Welcome to the SharePointDirect CLI. ");
                            Console.WriteLine("This tool is brought to you by Exx Navarro (and contributors). ");
                            Console.WriteLine(" ");
                            Console.WriteLine("use -h or --help to show this screen");
                            Console.WriteLine("use -v or --version to check CLI version");
                            Console.WriteLine(" ");
                            Console.WriteLine("Usage:");
                            Console.WriteLine("SharePointDirect [method] <options>");
                            Console.WriteLine(" ");
                            Console.WriteLine("Methods:");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"GetNumberOfItems\" to get the number of items in a list");
                            Console.WriteLine("Example: SharePointDirect GetNumberOfItems <URL> <Listname>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"GetItemId\" to get the ID for a certain item by using the title as the criteria. ");
                            Console.WriteLine("Example: SharePointDirect GetItemId <URL> <Listname> <Title>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"AddItem\" to add a new item to the list. ");
                            Console.WriteLine("Example: SharePointDirect GetItemId <URL> <Listname> <Field1> <Value1> <...>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"DeleteItemById\" to delete and item from the list using the item ID. ");
                            Console.WriteLine("Example: SharePointDirect DeleteItemById <URL> <Listname> <ID>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"UploadFileWithMeta\" to upload a file and include metadata. ");
                            Console.WriteLine("Example: SharePointDirect UploadFileWithMeta <URL> <Listname> <Property1> <Property2> <...>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"UploadFileNoMeta\" to upload a file with no defined metadata. ");
                            Console.WriteLine("Example: SharePointDirect UploadFileNoMeta <URL> <Listname>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"GetOneItem\" to get an item from the list based on the title. ");
                            Console.WriteLine("Example: SharePointDirect GetOneItem <URL> <Listname> <Field1> <Value1> <...>");
                            Console.WriteLine(" ");
                            Console.WriteLine("use \"GetOldestItem\" to get the oldest item from the list based on the title. ");
                            Console.WriteLine("Example: SharePointDirect GetOldestItem <URL> <Listname>");
                            Console.WriteLine(" ");
                            Console.WriteLine("More information on this link: https://github.com/exxmen/sharepointdirect/blob/master/README.md");
                            Console.WriteLine(" ");
                            Console.WriteLine("Press any key to exit. ");
                            Console.ReadKey();
                            break;
                    }
                }
                else
                {
                    using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\SharePointDirectError.txt"))
                    {
                        Console.WriteLine("Error: You need to pass at least one argument. \"-h\" to get information on the accepted arguments.");
                        Console.WriteLine("Press ENTER key to exit. ");
                        var input = Console.ReadLine();
                        sw.WriteLine("Error: You need to pass at least one argument. \"-h\" to get information on the accepted arguments.");
                    }
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\SharePointDirectError.txt"))
                {
                    Console.WriteLine("Error: " + e.Message);
                    Console.WriteLine("Press ENTER key to exit. ");
                    var input = Console.ReadLine();
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// gets the item count of the provided list in the specified sharepoint site
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        public static void GetNumberOfItems(string URL, string ListName)
        {
            try
            {
                AuthenticationManager authenticationManager = new AuthenticationManager();

                int itemCount = 0;

                var context = authenticationManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                context.Load(list);
                context.ExecuteQuery();

                itemCount = list.ItemCount;

                var message = itemCount.ToString();
                Console.WriteLine(message);

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetNumberOfItems.txt"))
                {
                    sw.WriteLine(itemCount);
                }
            }
            catch (Exception e)
            {

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetNumberOfItems.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }

        }

        /// <summary>
        /// gets the Id of the given title from the specified sharepoint list
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        /// <param name="ItemTitle"></param>
        public static void GetItemId(string URL, string ListName, string ItemTitle)
        {
            string itemId;

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                CamlQuery query = new CamlQuery();
                var viewXML = "<View><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE'/></OrderBy><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>" +
                    ItemTitle
                    + "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
                query.ViewXml = viewXML;
                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                ListItemCollection listItems = list.GetItems(query);
                context.Load(listItems);
                context.ExecuteQuery();

                itemId = listItems[0].Id.ToString();

                var message = itemId.ToString();
                Console.WriteLine(message);

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetItemId.txt"))
                {
                    sw.WriteLine(message);
                }
            }
            catch (Exception e)
            {

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetItemId.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// adds one item to the specified sharepoint list
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        /// <param name="Field"></param>
        /// <param name="ValuePairs"></param>
        /// <returns></returns>
        public static void AddItem(string URL, string ListName, Dictionary<string, string> ValuePairs)
        {
            string itemId;

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(newItemInfo);

                foreach (var pair in ValuePairs)
                {
                    newItem["" + pair.Key + ""] = pair.Value;
                }

                newItem.Update();
                context.Load(newItem);
                context.ExecuteQuery();

                itemId = newItem.Id.ToString();

                var message = itemId.ToString();
                Console.WriteLine(message);

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\AddItem.txt"))
                {
                    sw.WriteLine(message);
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\AddItem.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// deletes an item from the specified sharepoint list based on the given ID
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        /// <param name="Id"></param>
        public static void DeleteItemById(string URL, string ListName, int Id)
        {

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                ListItem itemToBeDeleted = list.GetItemById(Id);
                itemToBeDeleted.DeleteObject();
                context.ExecuteQuery();

                var message = "Item deleted.";
                Console.WriteLine(message);

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\DeleteItemById.txt"))
                {
                    sw.WriteLine(message);
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\DeleteItemById.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }

        }

        /// <summary>
        /// uploads the file and adds metadata
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="FolderName"></param>
        /// <param name="Filepath"></param>
        /// <param name="Pairs"></param>
        /// <remarks>
        /// the pairs argument accepts key value pairs of strings only
        /// </remarks>
        public static void UploadFileWithMeta(string URL, string FolderName, string Filepath, Dictionary<string,string> Pairs)
        {
            string Filename;
            try
            {
                Filename = Path.GetFileName(Filepath);

                AuthenticationManager authManager = new AuthenticationManager();

                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List library = web.Lists.GetByTitle(FolderName);
                Folder folder = library.RootFolder;
                context.Load(folder);
                context.ExecuteQuery();

                using (FileStream fs = new FileStream(Filepath, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = library.RootFolder.ServerRelativeUrl + "/" + Filename;
                    fileInfo.Overwrite = true;
                    Microsoft.SharePoint.Client.File file = folder.Files.Add(fileInfo);

                    foreach (var pair in Pairs)
                    {
                        file.ListItemAllFields["" + pair.Key + ""] = "" + pair.Value + "";
                    }

                    file.ListItemAllFields.Update();
                    context.Load(file);
                    context.ExecuteQuery();

                    if (file.CheckOutType != CheckOutType.None)
                    {
                        file.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    }

                    var message = "File uploaded with metadata";
                    Console.WriteLine(message);

                    using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\UploadFileWithMeta.txt"))
                    {
                        sw.WriteLine(message);
                    }
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\UploadFileWithMeta.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// uploads the file with no metadata
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="FolderName"></param>
        /// <param name="Filepath"></param>
        public static void UploadFileNoMeta(string URL, string FolderName, string Filepath)
        {
            string Filename;

            try
            {
                Filename = Path.GetFileName(Filepath);
                                
                AuthenticationManager authManager = new AuthenticationManager();

                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List library = web.Lists.GetByTitle(FolderName);
                Folder folder = library.RootFolder;
                context.Load(folder);
                context.ExecuteQuery();

                using (FileStream fs = new FileStream(Filepath, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = library.RootFolder.ServerRelativeUrl + "/" + Filename;
                    fileInfo.Overwrite = true;
                    Microsoft.SharePoint.Client.File file = folder.Files.Add(fileInfo);
                    context.Load(file);
                    context.ExecuteQuery();

                    if (file.CheckOutType != CheckOutType.None)
                    {
                        file.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    }

                    var message = "File uploaded";
                    Console.WriteLine(message);

                    using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\UploadFileNoMeta.txt"))
                    {
                        sw.WriteLine(message);
                    }
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\UploadFileNoMeta.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// gets one item from the list and writes the data to a csv file
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        /// <param name="SearchTitle"></param>
        /// <param name="FieldsToReturn"></param>
        public static void GetOneItem(string URL, string ListName, string SearchTitle, List<string> FieldsToReturn)
        {
            int itemId;

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                CamlQuery query = new CamlQuery();
                var viewXML = "<View><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE'/></OrderBy><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                    SearchTitle
                    + "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
                query.ViewXml = viewXML;
                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                ListItemCollection listItems = list.GetItems(query);
                context.Load(listItems);
                context.ExecuteQuery();

                itemId = listItems[0].Id;

                ListItem item = list.GetItemById(itemId);
                context.Load(item);
                context.ExecuteQuery();


                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetOneItem.csv"))
                {
                    foreach (string field in FieldsToReturn)
                    {
                        sw.Write(field + "," + item.FieldValues[field] + ",");
                    }
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetOneItem.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }

        /// <summary>
        /// gets the oldest item from the list provided
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        public static void GetOldestItem(string URL, string ListName)
        {
            int itemId;

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                CamlQuery query = new CamlQuery();
                var viewXML = "<View><Query><OrderBy><FieldRef Name='Created' Ascending='TRUE'/></OrderBy></Query><RowLimit>1</RowLimit></View>";
                query.ViewXml = viewXML;
                var context = authManager.GetWebLoginClientContext(URL);
                Web web = context.Web;
                List list = web.Lists.GetByTitle(ListName);
                ListItemCollection listItems = list.GetItems(query);
                context.Load(listItems);
                context.ExecuteQuery();

                itemId = listItems[0].Id;

                ListItem item = list.GetItemById(itemId);
                context.Load(item);
                context.ExecuteQuery();

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetOldestItem.txt"))
                {
                    //sw.Write(itemId + ";" + item.FieldValues["Title"]);
                    sw.Write(item.FieldValues["Title"]);
                }
            }
            catch (Exception e)
            {
                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetOldestItem.txt"))
                {
                    sw.WriteLine("Error: " + e.Message);
                }
            }
        }
    }
}
