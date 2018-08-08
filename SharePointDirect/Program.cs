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
                    case "DeleteItem":
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
                var file = folder.Files.Add(fileInfo);
                folder.Context.Load(file);
                folder.Context.ExecuteQueryRetry();

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\spresult.txt"))
                {
                    sw.WriteLine("File uploaded.");
                }
            }
        }
    }
}
