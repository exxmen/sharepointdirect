using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.IO;

namespace SpOnlineDirectConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 0)
            {
                switch (args[0])
                {
                    case "GetNumberOfItems":
                        GetNumberOfItems(args[1], args[2]);
                        break;
                    case "GetItemId":
                        GetItemId(args[1], args[2], args[3]);
                        break;
                    case "AddItem":
                        AddItem(args[1], args[2], args[3], args[4]);
                        break;
                    case "DeleteItemById":
                        int itemId = 0;
                        Int32.TryParse(args[3], out itemId);
                        DeleteItemById(args[1], args[2], itemId);
                        break;
                    case "UploadFileWithMeta":
                        var map = new Dictionary<string, string>();
                        for (int i = 4; i <= args.Length; i+=2)
                        {
                            map.Add(args[i], args[i + 1]);
                        }
                        UploadFileWithMeta(args[1], args[2], args[3], map);
                        break;
                    case "-v":
                        Console.WriteLine("V 0.1");
                        break;
                    case "-h":
                        Console.WriteLine("Welcome to the SharePointDirect CLI. ");
                        Console.WriteLine("Brought to you by Exx Navarro. ");
                        Console.WriteLine("use -v to check tool version");
                        Console.WriteLine("to use the CLI, use one of the choices below as the first argument, the second argument should be the url, the third is the list or library name. The arguments that are needed for the specific comes after these 3 required arguments.");
                        Console.WriteLine("use \"GetNumberOfItems\" to get the number of items in a list");
                        Console.WriteLine("use \"GetItemId\" to get the ID for a certain item by using the title as the criteria. ");
                        Console.WriteLine("use \"AddItem\" to add a new item to the list. ");
                        Console.WriteLine("use \"DeleteItemById\" to delete and item from the list using the item ID. ");
                        Console.WriteLine("Press any key to exit. ");
                        var input = Console.ReadLine();
                        break;

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
            AuthenticationManager authenticationManager = new AuthenticationManager();

            int itemCount = 0;

            var context = authenticationManager.GetWebLoginClientContext(URL);
            Web web = context.Web;
            List list = web.Lists.GetByTitle(ListName);
            context.Load(list);
            context.ExecuteQuery();

            itemCount = list.ItemCount;

            Console.WriteLine(itemCount);

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
            {
                sw.WriteLine(itemCount);
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

            AuthenticationManager authManager = new AuthenticationManager();

            CamlQuery query = new CamlQuery();
            var viewXML = "<View><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE'/></OrderBy><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
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

            Console.WriteLine(itemId);

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
            {
                sw.WriteLine(itemId);
            }
        }

        /// <summary>
        /// adds one item to the specified sharepoint list
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="ListName"></param>
        /// <param name="Field"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static void AddItem(string URL, string ListName, string Field, string Value)
        {
            string itemId;

            AuthenticationManager authManager = new AuthenticationManager();

            var context = authManager.GetWebLoginClientContext(URL);
            Web web = context.Web;
            List list = web.Lists.GetByTitle(ListName);
            ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
            ListItem newItem = list.AddItem(newItemInfo);
            newItem["" + Field + ""] = Value;
            newItem.Update();
            context.Load(newItem);
            context.ExecuteQuery();

            itemId = newItem.Id.ToString();

            Console.WriteLine(itemId);

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
            {
                sw.WriteLine(itemId);
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

            AuthenticationManager authManager = new AuthenticationManager();

            var context = authManager.GetWebLoginClientContext(URL);
            Web web = context.Web;
            List list = web.Lists.GetByTitle(ListName);
            ListItem itemToBeDeleted = list.GetItemById(Id);
            itemToBeDeleted.DeleteObject();
            context.ExecuteQuery();

            Console.WriteLine("Item deleted.");

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
            {
                sw.WriteLine("Item Deleted.");
            }

        }

        /// <summary>
        /// uploads the file and adds metadata
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="FolderName"></param>
        /// <param name="Filepath"></param>
        /// <param name="pairs"></param>
        /// <remarks>
        /// the pairs accepts key value pairs of strings
        /// </remarks>
        public static void UploadFileWithMeta(string URL, string FolderName, string Filepath, Dictionary<string,string> pairs)
        {
            string Filename;
            Filename = Filepath.Substring(Filepath.IndexOf("\\"));

            AuthenticationManager authManager = new AuthenticationManager();
 
            var context = authManager.GetWebLoginClientContext(URL);
            Web web = context.Web;
            List library = web.Lists.GetByTitle(FolderName);
            context.Load(library);
            Folder folder = library.RootFolder;

            using (FileStream fs = new FileStream(Filepath, FileMode.Open))
            {
                FileCreationInformation fileInfo = new FileCreationInformation();
                fileInfo.ContentStream = fs;
                fileInfo.Url = Filename;
                fileInfo.Overwrite = true;
                Microsoft.SharePoint.Client.File file = folder.Files.Add(fileInfo);
                
                foreach (var pair in pairs)
                {
                    file.ListItemAllFields["" + pair.Key + ""] = "" + pair.Value +"";
                }

                folder.Context.Load(file);
                folder.Context.ExecuteQueryRetry();

                Console.WriteLine("File uploaded.");

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
                {
                    sw.WriteLine("File uploaded.");
                }
            }
        }
    }
}
