using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.IO;

namespace SharePointDirectLibrary
{
    public class Library
    {
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

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\GetItemIdResult.txt"))
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

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\AddItemResult.txt"))
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

            using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\DeleteItemByIdResult.txt"))
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
        /// the pairs accepts key value pairs of strings only
        /// </remarks>
        public static void UploadFileWithMeta(string URL, string FolderName, string Filepath, Dictionary<string, string> pairs)
        {
            string Filename;
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

                foreach (var pair in pairs)
                {
                    file.ListItemAllFields["" + pair.Key + ""] = "" + pair.Value + "";
                }

                file.ListItemAllFields.Update();
                context.Load(file);
                context.ExecuteQuery();

                Console.WriteLine("File uploaded.");

                using (StreamWriter sw = System.IO.File.CreateText("C:\\Apps\\UploadFileWithMetaResult.txt"))
                {
                    sw.WriteLine("File uploaded.");
                }
            }
        }
    }
}
