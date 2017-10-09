using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Threading;
using System.Text.RegularExpressions;
// Uses workbook made from SCExporter to upload data back to sharpcloud
namespace SCTags
{
    class Program
    {

        static void Main(string[] args)
        {
            // Get info from config file
            var team = ConfigurationManager.AppSettings["team"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var sheet = ConfigurationManager.AppSettings["sheet"];
            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd);
            var teamBook = sc.StoriesTeam(team);
            string fileName = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString() + "\\" + sheet + ".xlsx";
            // Load data from excel
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                // Generate data from worksheets
                var itemSheet = xlPackage.Workbook.Worksheets.First();
                var itemRows = itemSheet.Dimension.End.Row;
                var itemColumns = itemSheet.Dimension.End.Column;
                // Creating the objects to be used in the threads
                // add attribute to story

                foreach(var book in teamBook)
                {
                    var story = sc.LoadStory(book.Id);
                    Console.WriteLine("Adding tags to " + story.Name);
                    for (var k = 2; k < itemRows; k++)
                    {
                        if (story.ItemTag_FindByName(itemSheet.Cells[k, 1].Value.ToString()) == null)
                        {
                            story.ItemTag_AddNew(itemSheet.Cells[k, 1].Value.ToString(), " ", itemSheet.Cells[k, 3].Value.ToString());
                        }
                    }
                    story.Save();
                }
            }
        }
    }
}