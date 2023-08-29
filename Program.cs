using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;

namespace OneDriveWithMSGraph
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Working with Graph and One Drive is fun!");

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

            // Initialize the auth provider with values from appsettings.json
            var authProvider = new DeviceCodeAuthProvider(appId, scopes);

            // Request a token to sign in the user
            var accessToken = authProvider.GetAccessToken().Result;
            GraphHelper.Initialize(authProvider);

            int choice = -1;
            while (choice != 0)
            {
                Console.WriteLine("Chọn 1 lựa chọn :");
                Console.WriteLine("0. Thoát");
                Console.WriteLine("1. Hiện thị access token");
                Console.WriteLine("2. Get your OneDrive root folder");
                //Console.WriteLine("3. Lấy các file trong onedrive");
                Console.WriteLine("4. Lấy các file excel trong onedrive");
                Console.WriteLine("5. Đọc dữ liệu từ 1 file ");
                Console.WriteLine("6. Đọc dữ liệu từ theo range theo worksheet trong file");
                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }
                var driveContents = await GraphHelper.GetDriveContents();
                var excelContents = await GraphHelper.GetExcelFiles();
                switch (choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;

                    case 1:
                        Console.WriteLine(string.Empty);
                        Console.ForegroundColor = ConsoleColor.Green;
                        // Display access token
                        Console.WriteLine($"Access token is : {accessToken}");
                        Console.ForegroundColor = ConsoleColor.White;

                        break;

                    case 2:
                        // Get OneDrive Info
                        Console.WriteLine(string.Empty);
                        Console.ForegroundColor = ConsoleColor.Green;
                        var driveInfo = await GraphHelper.GetOneDrive();
                        Console.WriteLine(FormatDriveInfo(driveInfo));
                        Console.ForegroundColor = ConsoleColor.White;
                        break;

                    //case 3:
                    //    // Get OneDrive contents
                    //    Console.WriteLine(string.Empty);
                    //    Console.ForegroundColor = ConsoleColor.Green;
                    //    Console.WriteLine(ListOneDriveContents(driveContents.ToList()));
                    //    Console.ForegroundColor = ConsoleColor.White;
                    //    break;

                    case 4:
                        Console.WriteLine(string.Empty);
                        Console.ForegroundColor = ConsoleColor.Green;
                        var contentList = await ListOneDriveContents(excelContents.ToList());
                        Console.WriteLine(contentList);

                        Console.ForegroundColor = ConsoleColor.White;
                        break;

                    case 5:
                        Console.WriteLine("Nhập vào id file muốn xem  : ");
                        string fileId = Console.ReadLine();
                        Console.WriteLine(string.Empty);
                        Console.ForegroundColor = ConsoleColor.Green;
                        var data = await GraphHelper.ReadAllWorksheets(fileId);
                        await ShowContentExcelsAsync(data, null);
                        Console.ForegroundColor = ConsoleColor.White;
                        break;

                    case 6:
                        Console.WriteLine("Nhập vào id file muốn xem : ");
                        string idFile = Console.ReadLine();
                        Console.WriteLine("Nhập vào tên worksheet (ví dụ Sheet1) : ");
                        string worksheet = Console.ReadLine();
                        Console.WriteLine("Nhập vào khoảng muốn xem ví dụ (A1:C3) : ");
                        string address = Console.ReadLine();
                        Console.WriteLine(string.Empty);
                        Console.ForegroundColor = ConsoleColor.Green;
                        string[,] table = await GraphHelper.GetRangeValues(idFile, worksheet, address);
                        await ShowContentExcelsAsync(null, table);
                        Console.ForegroundColor = ConsoleColor.White;
                        break;

                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }

            return appConfig;
        }

        private static string FormatDriveInfo(Drive drive)
        {
            var str = new StringBuilder();
            str.AppendLine($" OneDrive Name is: {drive.Name}");
            str.AppendLine($" OneDrive Ownder is: {drive.Owner.User.DisplayName}");
            str.AppendLine($" OneDrive id is: {drive.Id}");
            //str.AppendLine($"The OneDrive was modified last by: {drive?.LastModifiedBy?.User?.DisplayName}");

            return str.ToString();
        }

        //private static string ListOneDriveContents(List<DriveItem> contents)
        //{
        //    if (contents == null || contents.Count == 0)
        //    {
        //        return "No content found";
        //    }

        //    var str = new StringBuilder();
        //    foreach (var item in contents)
        //    {
        //        if (item.Folder != null)
        //        {
        //            str.AppendLine($"'{item.Name}' là folder, ID : {item.Id}");
        //        }
        //        else if (item.File != null)
        //        {
        //            str.AppendLine($"'{item.Name}' là 1 file, ID : {item.Id}");
        //        }
        //        else if (item.Audio != null)
        //        {
        //            str.AppendLine($"'{item.Audio.Title}' là 1 audio, ID : {item.Id}");
        //        }
        //        else
        //        {
        //            str.AppendLine($"Generic drive item found with name {item.Name}, ID : {item.Id}");
        //        }
        //    }

        //    return str.ToString();
        //}

        private static async Task<string> ListOneDriveContents(List<DriveItem> contents)
        {
            if (contents == null || contents.Count == 0)
            {
                return "No content found";
            }

            var str = new StringBuilder();
            foreach (var item in contents)
            {
                if (item.File != null)
                {
                    str.AppendLine($"'{item.Name}' is a file, ID: {item.Id}");
                    var worksheets = await GraphHelper.GetWorkSheet(item.Id).ConfigureAwait(false);

                    if (worksheets != null && worksheets.Any())
                    {
                        foreach (var worksheet in worksheets)
                        {
                            str.AppendLine($"\t- Name: '{worksheet.Name}', ID: {worksheet.Id}");
                        }
                    }
                    else
                    {
                        str.AppendLine("No worksheets found.");
                    }
                }
                else
                {
                    str.AppendLine($"Generic drive item found with name {item.Name}, ID: {item.Id}");
                }
            }

            return str.ToString();
        }

        // show nội dung trong worksheet
        private static async Task ShowContentExcelsAsync(List<string[,]> worksheetTables, string[,] worksheetContents)
        {
            if (worksheetTables != null)
            {
                for (int tableIndex = 0; tableIndex < worksheetTables.Count; tableIndex++)
                {
                    string[,] table = worksheetTables[tableIndex];
                    int numRows = table.GetLength(0);
                    int numCols = table.GetLength(1);

                    Console.WriteLine($"Table {tableIndex + 1}:");
                    for (int i = 0; i < numRows; i++)
                    {
                        for (int j = 0; j < numCols; j++)
                        {
                            Console.Write($"{table[i, j],-15}\t");
                        }
                        Console.WriteLine();
                    }
                    Console.WriteLine();
                }
            }
           
            if (worksheetContents != null)
            {
                int numRows = worksheetContents.GetLength(0);
                int numCols = worksheetContents.GetLength(1);

                Console.WriteLine("Single Worksheet Contents:");
                for (int i = 0; i < numRows; i++)
                {
                    for (int j = 0; j < numCols; j++)
                    {
                        Console.Write($"{worksheetContents[i, j],-15}\t");
                    }
                    Console.WriteLine();
                }
            }
            
        }
    }
}