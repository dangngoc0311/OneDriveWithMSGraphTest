using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OneDriveWithMSGraph
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
            graphClient.BaseUrl = "https://graph.microsoft.com/beta";
        }

        // Lấy thông tin cá nhân
        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        public static async Task<Drive> GetOneDrive()
        {
            try
            {
                // GET /me
                return await graphClient.Me.Drive.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting OneDrive data: {ex.Message}");
                return null;
            }
        }

        // Lấy các file trong onedrive
        public static async Task<IEnumerable<DriveItem>> GetDriveContents()
        {
            try
            {
                return await graphClient.Me.Drive.Root.Children.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting One Drive contents: {ex.Message}");
                return null;
            }
        }

        // Lấy tất cả các file excel
        public static async Task<IEnumerable<DriveItem>> GetExcelFiles()
        {
            try
            {
                var driveContents = await graphClient.Me.Drive.Root.Children.Request().GetAsync();

                // Filter the items to include only Excel files
                var excelFiles = driveContents.Where(item =>
                   item.File != null &&
                   item.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)); // Check for the ".xlsx" extension

                return excelFiles;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting OneDrive contents: {ex.Message}");
                return null;
            }
        }

        // Lấy danh sách các worksheet trong file
        public static async Task<IEnumerable<WorkbookWorksheet>> GetWorkSheet(string fileId)
        {
            try
            {
                var workbook = await graphClient.Me.Drive.Items[fileId].Workbook.Worksheets.Request().GetAsync();
                if (workbook != null)
                {
                    return workbook.CurrentPage;
                }
                return null;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error reading worksheets: {ex.Message}");
                return null;
            }
        }

        // Đọc nội dung tất cả các workshift trong file
        public static async Task<List<string[,]>> ReadAllWorksheets(string fileId)
        {
            List<string[,]> tables = new List<string[,]>();

            try
            {
                var workbook = await graphClient.Me.Drive.Items[fileId].Workbook.Worksheets.Request().GetAsync();

                if (workbook != null)
                {
                    foreach (var worksheet in workbook.CurrentPage)
                    {
                        var worksheetId = worksheet.Id;
                        var usedRange = await graphClient.Me.Drive.Items[fileId].Workbook.Worksheets[worksheetId].UsedRange().Request().GetAsync();

                        if (usedRange != null && usedRange.Values != null)
                        {
                            var valuesArray = usedRange.Text as JArray;
                            if (valuesArray != null)
                            {
                                int numRows = valuesArray.Count;
                                int numCols = valuesArray[0].Count();

                                string[,] table = new string[numRows, numCols];

                                for (int i = 0; i < numRows; i++)
                                {
                                    for (int j = 0; j < numCols; j++)
                                    {
                                        string value = valuesArray[i][j]?.ToString()?.Replace("\n", "") ?? "";

                                        // Check if the value is in the "m/dd/yyyy" format
                                        if (DateTime.TryParseExact(value, "M/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                        {
                                            // Convert and store the date in the "dd/MM/yyyy" format
                                            table[i, j] = parsedDate.ToString("dd/MM/yyyy");
                                        }
                                        else
                                        {
                                            table[i, j] = value;
                                        }
                                    }
                                }

                                tables.Add(table);
                            }
                        }
                    }

                    return tables;
                }
                else
                {
                    Console.WriteLine("No workbook found with the specified fileId.");
                    return null;
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error reading worksheets: {ex.Message}");
                return null;
            }
        }


        // lấy nội dung theo cell
        public static async Task<string> GetCellValue(string fileId, string worksheetId, string cellAddress)
        {
            try
            {
                var cell = await graphClient.Me.Drive.Items[fileId].Workbook.Worksheets[worksheetId].Range(cellAddress).Request().GetAsync();

                if (cell != null)
                {
                    return cell.Values?[0][0]?.ToString() ?? "";
                }

                return null;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting cell value: {ex.Message}");
                return null;
            }
        }

        // Lấy nội dung theo range

        public static async Task<string[,]> GetRangeValues(string fileId, string worksheetId, string rangeAddress)
        {
            try
            {
                var range = await graphClient.Me.Drive.Items[fileId].Workbook.Worksheets[worksheetId].Range(rangeAddress).Request().GetAsync();

                if (range != null && range.Values != null)
                {
                    var valuesArray = range.Values as JArray;
                    if (valuesArray != null)
                    {
                        int numRows = valuesArray.Count;
                        int numCols = valuesArray[0].Count();

                        string[,] table = new string[numRows, numCols];

                        for (int i = 0; i < numRows; i++)
                        {
                            for (int j = 0; j < numCols; j++)
                            {
                                table[i, j] = valuesArray[i][j]?.ToString() ?? "";
                            }
                        }

                        return table;
                    }
                }

                return null;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting range values: {ex.Message}");
                return null;
            }
        }
    }
}