using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelManipulation
{
    class DateAndSize
    {
        public string date { get; set; } = string.Empty;
        public string size { get; set; } = string.Empty;

        public override bool Equals(object obj)
        {
            if (obj is DateAndSize other)
            {
                return size == other.size && date == other.date;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return size.GetHashCode() ^ date.GetHashCode();
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            // Set the license context before using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial

            // Change the directory path as needed
            string directoryPath = @"C:\CP_Volume";
            string outPutPath = @"C:\CP_Volume\Volume_output\output.xlsx";

            // Get all Excel files in the specified directory
            string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the directory.");
                return;
            }


            // Set the console encoding to UTF-8
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            if (File.Exists(outPutPath))
            {
                File.Delete(outPutPath);
            }
            foreach (string excelFile in excelFiles)
            {
                Console.WriteLine($"Processing Excel file: {excelFile}");

                //Size declare
                string sizeJb = "Jb";
                string size4XL = "4XL";
                string size3XL = "3XL";
                string size2XLPlus = "2XL+";
                string size2XL = "2XL";
                string sizeXL = "XL";
                string sizeLL = "LL";
                string sizeL = "L";
                string sizeM = "M";
                string sizeS = "S";
                string[] sizes = { "Jb", "4XL", "3XL", "2XL+", "2XL", "XL", "LL", "L", "M", "S" };

                int sizeJbValue = 0;
                int size4XLValue = 0;
                int size3XLValue = 0;
                int size2XLPlusValue = 0;
                int size2XLValue = 0;
                int sizeXLValue = 0;
                int sizeLLValue = 0;
                int sizeLValue = 0;
                int sizeMValue = 0;
                int sizeSValue = 0;
                Dictionary<DateAndSize, int> sizeCounts = new Dictionary<DateAndSize, int>();
                int state = 1;
                
                // Load the Excel file using EPPlus
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFile)))
                {
                    // Access the first worksheet
                    if (excelPackage.Workbook.Worksheets.Count > 0)
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0]; // Change the index if needed

                        // Identify the columns corresponding to specific dates (excluding "Total" columns)
                        var dateColumns = worksheet.Cells[1, 2, 1, worksheet.Dimension.Columns]
                            .Where(cell => cell.Text.StartsWith("Total", StringComparison.OrdinalIgnoreCase) == false)
                            .Select(cell => cell.Text).ToList();

                        string date;
                        int columncount = 1;

                        

                        for (int row = 3; row <= worksheet.Dimension.Rows-3; row++)
                        {
                            var nameProduct = worksheet.Cells[row, 1].Text;
                            columncount = 1;
                            Console.WriteLine($"Row {row}:");

                            for (int col = 3; col <= worksheet.Dimension.Columns-3; col += 3)
                            {
                                date = dateColumns[columncount];
                                var cellValue = worksheet.Cells[row, col].Text;
                                Console.WriteLine($"{nameProduct} -> {date}: {cellValue}");

                                if (nameProduct.Contains("STD") &&
                                        nameProduct.Contains("กุ้งขาว") &&
                                        cellValue != "" &&
                                        nameProduct.Contains("KG") &&
                                        !nameProduct.Contains("ตัว/กก"))
                                {
                                    string size = DetermineSize(nameProduct);
                                    DateAndSize key = new DateAndSize { date = date, size = size };
                                    if (sizeCounts.TryGetValue(key, out int weight) && cellValue != "")
                                    {
                                        weight += Int32.Parse(cellValue.Replace(",", ""));
                                        sizeCounts[key] = weight;
                                    }
                                    else if (size == "Jb")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeJb},  Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "4XL")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = size4XL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "3XL")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = size3XL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if ((size == "2XL+"))
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = size2XLPlus }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "2XL")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = size2XL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "XL")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeXL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "LL")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeLL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "L")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeL }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "M")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeM }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (size == "S")
                                    {
                                        sizeCounts.Add(new DateAndSize { date = date, size = sizeS }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                }
                                columncount += 3;
                            }

                            Console.WriteLine(); // Move to the next line after completing a row

                        }
                    }
                    else
                    {
                        Console.WriteLine("The workbook does not contain any worksheets.");
                    }
                }

                var sortedSizeCounts = sizeCounts.OrderBy(entry => entry.Key.date)
                                         .ThenBy(entry => Array.IndexOf(sizes, entry.Key.size))
                                         .ToDictionary(entry => entry.Key, entry => entry.Value);
                using (ExcelPackage excelPackage = new ExcelPackage(outPutPath))
                {
                    // Add a new worksheet to the Excel package
                    string sheetName = Path.GetFileNameWithoutExtension(excelFile);
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);

                    // Headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "Size";
                    worksheet.Cells[1, 3].Value = "Weight";

                    int row = 2;
                    string previousDate = null;
                    // Populate the worksheet with data from the dictionary
                    foreach (var kvp in sortedSizeCounts)
                    {
                        if (previousDate != null && kvp.Key.date != previousDate)
                        {
                            row++;
                        }
                        worksheet.Cells[row, 1].Value = kvp.Key.date;
                        worksheet.Cells[row, 2].Value = kvp.Key.size;
                        worksheet.Cells[row, 3].Value = kvp.Value;

                        row++;
                        previousDate = kvp.Key.date;
                    }

                    // Save the Excel package to a file
                    if(File.Exists(outPutPath))
                    {
                        excelPackage.Save();
                    }
                    else
                    {
                        excelPackage.SaveAs(new FileInfo(outPutPath));
                    }
                }
                foreach (var sizeCount in sizeCounts.OrderBy(kvp => kvp.Key.date).ThenBy(kvp => kvp.Key.size))
                {
                    Console.WriteLine(1);
                    Console.WriteLine($"{sizeCount.Key.date}: size = {sizeCount.Key.size} Weight = {sizeCount.Value}");
                }
                string region = Path.GetFileNameWithoutExtension(excelFile);
                Console.WriteLine($"{region}");
                Console.WriteLine($"{nameof(sizeJb)} {sizeJbValue}");
                Console.WriteLine($"{nameof(size4XL)} {size4XLValue}");
                Console.WriteLine($"{nameof(size3XL)} {size3XLValue}");
                Console.WriteLine($"{nameof(size2XLPlus)} {size2XLPlusValue}");
                Console.WriteLine($"{nameof(size2XL)} {size2XLValue}");
                Console.WriteLine($"{nameof(sizeXL)} {sizeXLValue}");
                Console.WriteLine($"{nameof(sizeLL)} {sizeLLValue}");
                Console.WriteLine($"{nameof(sizeL)} {sizeLValue}");
                Console.WriteLine($"{nameof(sizeM)} {sizeMValue}");
                Console.WriteLine($"{nameof(sizeS)} {sizeSValue}");

                Console.WriteLine(); // Add a separator line between files
            }

        }

        private static string DetermineSize(string nameProduct)
        {
            string size = "";
            string sizeJb = "Jb";
            string size4XL = "4XL";
            string size3XL = "3XL";
            string size2XLPlus = "2XL+";
            string size2XL = "2XL";
            string sizeXL = "XL";
            string sizeLL = "LL";
            string sizeL = "L";
            string sizeM = "M";
            string sizeS = "S";

            if (nameProduct.Contains("STD Jumbo"))
                size = sizeJb;
            else if (nameProduct.Contains("STD 4XL") || nameProduct.Contains("STD XXXXL"))
                size = size4XL;
            else if (nameProduct.Contains("STD 3XL") || nameProduct.Contains("STD XXXL"))
                size = size3XL;
            else if (nameProduct.Contains("STD 2XL Plus") || nameProduct.Contains("STD XXL Plus"))
                size = size2XLPlus;
            else if (nameProduct.Contains("STD 2XL") || nameProduct.Contains("STD XXL"))
                size = size2XL;
            else if (nameProduct.Contains("STD XL"))
                size = sizeXL;
            else if (nameProduct.Contains("STD LL") || nameProduct.Contains("STD L+"))
                size = sizeLL;
            else if (nameProduct.Contains("STD L"))
                size = sizeL;
            else if (nameProduct.Contains("STD M"))
                size = sizeM;
            else if (nameProduct.Contains("STD S"))
                size = sizeS;

            return size;
        }
    }
}
