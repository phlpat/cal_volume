using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelManipulation
{
    class DateAndSku
    {
        public string date { get; set; } = string.Empty;
        public string sku { get; set; } = string.Empty;

        public override bool Equals(object obj)
        {
            if (obj is DateAndSku other)
            {
                return sku == other.sku && date == other.date;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return sku.GetHashCode() ^ date.GetHashCode();
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            TTVolume();
            MTMakroVolume();
            MTLotusVolume();
        }

        private static string DetermineSize(string nameProduct)
        {
            string premium_3_kg = "premium_3_kg";
            string premium_5_kg = "premium_5_kg";
            string premium_10_kg = "premium_10_kg";
            string premium_30_kg = "premium_30_kg";
            string premium_nb_10_kg = "premium_nb_10_kg";

            string std_Jb_5kg = "std_Jb_5kg";
            string std_4XL_5kg = "std_4XL_5kg";
            string std_3XL_5kg = "std_3XL_5kg";
            string std_2XLPlus_5kg = "std_2XLPlus_5kg";
            string std_2XL_5kg = "std_2XL_5kg";
            string std_XL_5kg = "std_XL_5kg";
            string std_LL_5kg = "std_LL_5kg";
            string std_L_5kg = "std_L_5kg";
            string std_M_5kg = "std_M_5kg";
            string std_S_5kg = "std_S_5kg";

            string std_Jb_10kg = "std_Jb_10kg";
            string std_4XL_10kg = "std_4XL_10kg";
            string std_3XL_10kg = "std_3XL_10kg";
            string std_2XLPlus_10kg = "std_2XLPlus_10kg";
            string std_2XL_10kg = "std_2XL_10kg";
            string std_XL_10kg = "std_XL_10kg";
            string std_LL_10kg = "std_LL_10kg";
            string std_L_10kg = "std_L_10kg";
            string std_M_10kg = "std_M_10kg";
            string std_S_10kg = "std_S_10kg";

            string std_Jb_18kg = "std_Jb_18kg";
            string std_4XL_18kg = "std_4XL_18kg";
            string std_3XL_18kg = "std_3XL_18kg";
            string std_2XLPlus_18kg = "std_2XLPlus_18kg";
            string std_2XL_18kg = "std_2XL_18kg";
            string std_XL_18kg = "std_XL_18kg";
            string std_LL_18kg = "std_LL_18kg";
            string std_L_18kg = "std_L_18kg";
            string std_M_18kg = "std_M_18kg";
            string std_S_18kg = "std_S_18kg";

            string mixedSize = "mixedSize";

            string sure = "sure";

            string premiumSoft = "premiumSoft";

            string remainPondGrade1 = "remainPondGrade1";
            string remainPondGrade234 = "remainPondGrade234";
            string remainPMGrade1 = "remainPMGrade1";
            string remainPMGrade234 = "remainPMGrade234";
            string remainSTDGrade1 = "remainSTDGrade1";
            string remainSTDGrade234 = "remainSTDGrade234";
            string foamGrade2 = "foamGrade2";

            string sku = "";

            if (nameProduct.Contains("FM"))
            {
                if (nameProduct.Contains("3 KG"))
                {
                    sku = premium_3_kg;
                }
                else if (nameProduct.Contains("5 KG"))
                {
                    sku = premium_5_kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = premium_10_kg;
                }
                else if (nameProduct.Contains("30 KG"))
                {
                    sku = premium_30_kg;
                }
            }

            else if (nameProduct.Contains("NB") && nameProduct.Contains("10 KG"))
            {
                sku = premium_nb_10_kg;
            }
               
            else if ((nameProduct.Contains("STD Jumbo") || nameProduct.Contains("STD Jb")))
            {
                if (nameProduct.Contains("5 KG") || nameProduct.Contains("5KG"))
                {
                    sku = std_Jb_5kg;
                }
                else if (nameProduct.Contains("10 KG") || nameProduct.Contains("10KG"))
                {
                    sku = std_Jb_10kg;
                }
                else if (nameProduct.Contains("18 KG") || nameProduct.Contains("18KG"))
                {
                    sku = std_Jb_18kg;
                }
            }
                
            else if ((nameProduct.Contains("STD 4XL") || nameProduct.Contains("STD XXXXL")))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_4XL_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                { 
                    sku = std_4XL_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_4XL_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD 3XL") || nameProduct.Contains("STD XXXL"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_3XL_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_3XL_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_3XL_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD 2XL Plus") || nameProduct.Contains("STD XXL Plus"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_2XLPlus_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_2XLPlus_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_2XLPlus_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD 2XL") || nameProduct.Contains("STD XXL"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_2XL_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_2XL_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_2XL_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD XL"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_XL_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_XL_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_XL_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD LL") || nameProduct.Contains("STD L+"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_LL_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_LL_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_LL_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD L"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_L_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_L_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_L_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD M"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_M_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_M_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_M_18kg;
                }
            }
                
            else if (nameProduct.Contains("STD S"))
            {
                if (nameProduct.Contains("5 KG"))
                {
                    sku = std_S_5kg;
                }
                else if (nameProduct.Contains("10 KG"))
                {
                    sku = std_S_10kg;
                }
                else if (nameProduct.Contains("18 KG"))
                {
                    sku = std_S_18kg;
                }
            }

            else if (nameProduct.Contains("ไซซ์รวม"))
            {
                sku = mixedSize;
            }
               

            else if (nameProduct.Contains("SURE"))
            {
                sku = sure;
            }

            else if (nameProduct.Contains("Premium") && nameProduct.Contains("นิ่ม"))
            {
                sku = premiumSoft;
            }

            else if (nameProduct.Contains("ปากบ่อ") && nameProduct.Contains("กก"))
            {
                if (nameProduct.Contains("นิ่ม") || nameProduct.Contains("น่วม") || nameProduct.Contains("เสีย"))
                {
                    sku = remainPondGrade234;
                }
                else
                {
                    sku = remainPondGrade1;
                }
            }
                
            else if (nameProduct.Contains("WET") && nameProduct.Contains("กก"))
            {
                if (nameProduct.Contains("นิ่ม") || nameProduct.Contains("น่วม") || nameProduct.Contains("เสีย"))
                {
                    sku = remainPMGrade234;
                }
                else
                {
                    sku = remainPMGrade1;
                }
            }

            else if (nameProduct.Contains("STD") && nameProduct.Contains("กก"))
            {
                if (nameProduct.Contains("นิ่ม") || nameProduct.Contains("น่วม") || nameProduct.Contains("เสีย"))
                {
                    sku = remainSTDGrade234;
                }
                else
                {
                    sku = remainSTDGrade1;
                }
            }

            else if (nameProduct.Contains("เกรด2 แพ็คโฟม"))
            {
                sku = foamGrade2;  
            }

            return sku;
        }

        private static string DetermineSizeMT(string nameProduct)
        {

            string super_premiumMT = "super_premiumMT";
            string foamMT = "foamMT";
            string foamFM_10kg = "foamFM_10kg";
            string foamFM_5kg = "foamFM_5kg";
            string foamFM_3kg = "foamFM_3kg";
            string foamPM = "foamPM";
            string stdMT_10kg = "stdMT_10kg";
            string lotusB2B = "lotusB2B";
            string lotusSTD = "lotusSTD";

            string sku = "";

            if (nameProduct.Contains("ตาข่าย"))
            {
                sku = super_premiumMT;
            }

            else if (nameProduct.Contains(" MT ") && nameProduct.Contains("กล่องโฟม"))
            {
                sku = foamMT;
            }

            else if (nameProduct.Contains(" FM ") && nameProduct.Contains("กล่องโฟม"))
            {
                if (nameProduct.Contains("10 KG"))
                {
                    sku = foamFM_10kg;
                }
                else if (nameProduct.Contains("5 KG"))
                {
                    sku = foamFM_5kg;
                }
                else if (nameProduct.Contains("3 KG"))
                {
                    sku = foamFM_3kg;
                }
            }

            else if (nameProduct.Contains(" PM ") && nameProduct.Contains("กล่องโฟม"))
            {
                sku = foamPM;
            }

            else if ((nameProduct.Contains(" MT ") && nameProduct.Contains(" STD ")) || (nameProduct.Contains("10KG") && nameProduct.Contains("STD ")))
            {
                sku = stdMT_10kg;
            }

            else if (nameProduct.Contains("Lotus") && nameProduct.Contains("B2B"))
            {
                sku = lotusB2B;
            }

            else if (nameProduct.Contains("Lotus") && nameProduct.Contains("STD"))
            {
                sku = lotusSTD;
            }

            return sku;
        }

        private static void TTVolume()
        {
            // Set the license context before using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial

            // Change the directory path as needed
            string directoryPath = @"C:\Volume_Sales\TT";
            string outPutPath = @"C:\Volume_Sales\Volume_output\outputTT.xlsx";

            // Get all Excel files in the specified directory
            string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the directory.");
                return;
            }


            // Set the console encoding to UTF-8
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            foreach (string excelFile in excelFiles)
            {
                Console.WriteLine($"Processing Excel file: {excelFile}");

                //SKU declare
                string premium_3_kg = "premium_3_kg";
                string premium_5_kg = "premium_5_kg";
                string premium_10_kg = "premium_10_kg";
                string premium_30_kg = "premium_30_kg";
                string premium_nb_10_kg = "premium_nb_10_kg";

                string std_Jb_5kg = "std_Jb_5kg";
                string std_4XL_5kg = "std_4XL_5kg";
                string std_3XL_5kg = "std_3XL_5kg";
                string std_2XLPlus_5kg = "std_2XLPlus_5kg";
                string std_2XL_5kg = "std_2XL_5kg";
                string std_XL_5kg = "std_XL_5kg";
                string std_LL_5kg = "std_LL_5kg";
                string std_L_5kg = "std_L_5kg";
                string std_M_5kg = "std_M_5kg";
                string std_S_5kg = "std_S_5kg";

                string std_Jb_10kg = "std_Jb_10kg";
                string std_4XL_10kg = "std_4XL_10kg";
                string std_3XL_10kg = "std_3XL_10kg";
                string std_2XLPlus_10kg = "std_2XLPlus_10kg";
                string std_2XL_10kg = "std_2XL_10kg";
                string std_XL_10kg = "std_XL_10kg";
                string std_LL_10kg = "std_LL_10kg";
                string std_L_10kg = "std_L_10kg";
                string std_M_10kg = "std_M_10kg";
                string std_S_10kg = "std_S_10kg";

                string std_Jb_18kg = "std_Jb_18kg";
                string std_4XL_18kg = "std_4XL_18kg";
                string std_3XL_18kg = "std_3XL_18kg";
                string std_2XLPlus_18kg = "std_2XLPlus_18kg";
                string std_2XL_18kg = "std_2XL_18kg";
                string std_XL_18kg = "std_XL_18kg";
                string std_LL_18kg = "std_LL_18kg";
                string std_L_18kg = "std_L_18kg";
                string std_M_18kg = "std_M_18kg";
                string std_S_18kg = "std_S_18kg";

                string mixedSize = "mixedSize";

                string sure = "sure";

                string premiumSoft = "premiumSoft";

                string remainPondGrade1 = "remainPondGrade1";
                string remainPondGrade234 = "remainPondGrade234";
                string remainPMGrade1 = "remainPMGrade1";
                string remainPMGrade234 = "remainPMGrade234";
                string remainSTDGrade1 = "remainSTDGrade1";
                string remainSTDGrade234 = "remainSTDGrade234";
                string foamGrade2 = "foamGrade2";

                var skuOrderTT = new List<string>
                {
                    "premium_3_kg",
                    "premium_5_kg",
                    "premium_10_kg",
                    "premium_30_kg",
                    "premium_nb_10_kg",
                    "std_Jb_5kg",
                    "std_4XL_5kg",
                    "std_3XL_5kg",
                    "std_2XLPlus_5kg",
                    "std_2XL_5kg",
                    "std_XL_5kg",
                    "std_LL_5kg",
                    "std_L_5kg",
                    "std_M_5kg",
                    "std_S_5kg",
                    "std_Jb_10kg",
                    "std_4XL_10kg",
                    "std_3XL_10kg",
                    "std_2XLPlus_10kg",
                    "std_2XL_10kg",
                    "std_XL_10kg",
                    "std_LL_10kg",
                    "std_L_10kg",
                    "std_M_10kg",
                    "std_S_10kg",
                    "std_Jb_18kg",
                    "std_4XL_18kg",
                    "std_3XL_18kg",
                    "std_2XLPlus_18kg",
                    "std_2XL_18kg",
                    "std_XL_18kg",
                    "std_LL_18kg",
                    "std_L_18kg",
                    "std_M_18kg",
                    "std_S_18kg",
                    "mixedSize",
                    "sure",
                    "premiumSoft",
                    "remainPondGrade1",
                    "remainPondGrade234",
                    "remainPMGrade1",
                    "remainPMGrade234",
                    "remainSTDGrade1",
                    "remainSTDGrade234",
                    "foamGrade2"
                };

                Dictionary<DateAndSku, int> sizeCounts = new Dictionary<DateAndSku, int>();

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



                        for (int row = 3; row <= worksheet.Dimension.Rows - 3; row++)
                        {
                            var nameProduct = worksheet.Cells[row, 1].Text;
                            columncount = 1;
                            Console.WriteLine($"Row {row}:");

                            for (int col = 3; col <= worksheet.Dimension.Columns - 3; col += 3)
                            {
                                date = dateColumns[columncount];
                                var cellValue = worksheet.Cells[row, col].Text;
                                Console.WriteLine($"{nameProduct} -> {date}: {cellValue}");

                                if (nameProduct.Contains("กุ้ง") && cellValue != "")
                                {
                                    string sku = DetermineSize(nameProduct);
                                    DateAndSku key = new DateAndSku { date = date, sku = sku };
                                    if (sizeCounts.TryGetValue(key, out int weight) && cellValue != "")
                                    {
                                        weight += Int32.Parse(cellValue.Replace(",", ""));
                                        sizeCounts[key] = weight;
                                    }
                                    else if (sku == "premium_3_kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premium_3_kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "premium_5_kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premium_5_kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "premium_10_kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premium_10_kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if ((sku == "premium_30_kg"))
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premium_30_kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "premium_nb_10_kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premium_nb_10_kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_Jb_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_Jb_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_4XL_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_4XL_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_3XL_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_3XL_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XLPlus_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XLPlus_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XL_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XL_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_XL_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_XL_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_LL_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_LL_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_L_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_L_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_M_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_M_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_S_5kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_S_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_Jb_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_Jb_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_4XL_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_4XL_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_3XL_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_3XL_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XLPlus_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XLPlus_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XL_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XL_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_XL_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_XL_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_LL_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_LL_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_L_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_L_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_M_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_M_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_S_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_S_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_Jb_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_Jb_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_4XL_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_4XL_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_3XL_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_3XL_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XLPlus_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XLPlus_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_2XL_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_2XL_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_XL_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_XL_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_LL_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_LL_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_L_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_L_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_M_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_M_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "std_S_18kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = std_S_18kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "mixedSize")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = mixedSize }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "sure")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = sure }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "premiumSoft")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = premiumSoft }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainPondGrade1")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainPondGrade1 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainPondGrade234")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainPondGrade234 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainPMGrade1")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainPMGrade1 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainPMGrade234")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainPMGrade234 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainSTDGrade1")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainSTDGrade1 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "remainSTDGrade234")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = remainSTDGrade234 }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamGrade2")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamGrade2 }, Int32.Parse(cellValue.Replace(",", "")));
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

                var sortedSizeCounts = sizeCounts
                                        .OrderBy(entry => entry.Key.date)
                                        .ThenBy(entry => skuOrderTT.IndexOf(entry.Key.sku))
                                        .ToDictionary(entry => entry.Key, entry => entry.Value);

                if (File.Exists(outPutPath))
                {
                    File.Delete(outPutPath);
                }
                using (ExcelPackage excelPackage = new ExcelPackage(outPutPath))
                {
                    // Add a new worksheet to the Excel package
                    string sheetName = Path.GetFileNameWithoutExtension(excelFile);
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                    worksheet.Column(4).Style.Numberformat.Format = "0.000";

                    // Headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "Size";
                    worksheet.Cells[1, 3].Value = "Weight";
                    worksheet.Cells[1, 4].Value = "Weight(Ton)";

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
                        worksheet.Cells[row, 2].Value = kvp.Key.sku;
                        worksheet.Cells[row, 3].Value = kvp.Value;
                        worksheet.Cells[row, 4].Value = kvp.Value / 1000f;

                        row++;
                        previousDate = kvp.Key.date;
                    }

                    // Save the Excel package to a file
                    if (File.Exists(outPutPath))
                    {
                        excelPackage.Save();
                    }
                    else
                    {
                        excelPackage.SaveAs(new FileInfo(outPutPath));
                    }
                }
                foreach (var sizeCount in sizeCounts.OrderBy(kvp => kvp.Key.date).ThenBy(kvp => kvp.Key.sku))
                {
                    Console.WriteLine(1);
                    Console.WriteLine($"{sizeCount.Key.date}: size = {sizeCount.Key.sku} Weight = {sizeCount.Value}");
                }

                Console.WriteLine(); // Add a separator line between files
            }

        }

        private static void MTMakroVolume()
        {
            // Set the license context before using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial

            // Change the directory path as needed
            string directoryPath = @"C:\Volume_Sales\MT(Makro)";
            string outPutPath = @"C:\Volume_Sales\Volume_output\outputMT(Makro).xlsx";

            // Get all Excel files in the specified directory
            string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the directory.");
                return;
            }


            // Set the console encoding to UTF-8
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            foreach (string excelFile in excelFiles)
            {
                Console.WriteLine($"Processing Excel file: {excelFile}");

                //SKU declare
                string super_premiumMT = "super_premiumMT";
                string foamMT = "foamMT";
                string foamFM_10kg = "foamFM_10kg";
                string foamFM_5kg = "foamFM_5kg";
                string foamFM_3kg = "foamFM_3kg";
                string foamPM = "foamPM";
                string stdMT_10kg = "stdMT_10kg";
                string lotusB2B = "lotusB2B";
                string lotusSTD = "lotusSTD";

                var skuOrderMT = new List<string>
                {
                    "super_premiumMT",
                    "foamMT",
                    "foamFM_10kg",
                    "foamFM_5kg",
                    "foamFM_3kg",
                    "foamPM",
                    "stdMT_10kg",
                    "lotusB2B",
                    "lotusSTD"
                };

                Dictionary<DateAndSku, int> sizeCounts = new Dictionary<DateAndSku, int>();

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



                        for (int row = 3; row <= worksheet.Dimension.Rows - 3; row++)
                        {
                            var nameProduct = worksheet.Cells[row, 1].Text;
                            columncount = 1;
                            Console.WriteLine($"Row {row}:");

                            for (int col = 3; col <= worksheet.Dimension.Columns - 3; col += 3)
                            {
                                date = dateColumns[columncount];
                                var cellValue = worksheet.Cells[row, col].Text;
                                Console.WriteLine($"{nameProduct} -> {date}: {cellValue}");

                                if (nameProduct.Contains("กุ้ง") && !nameProduct.Contains("แช่แข็ง") && !nameProduct.Contains("หมึก")  && !nameProduct.Contains("PDTO") && !nameProduct.Contains("BS") && cellValue != "")
                                {
                                    string sku = DetermineSizeMT(nameProduct);
                                    DateAndSku key = new DateAndSku { date = date, sku = sku };
                                    if (sizeCounts.TryGetValue(key, out int weight) && cellValue != "")
                                    {
                                        weight += Int32.Parse(cellValue.Replace(",", ""));
                                        sizeCounts[key] = weight;
                                    }
                                    else if (sku == "super_premiumMT")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = super_premiumMT }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamMT")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamMT }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamFM_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if ((sku == "foamFM_5kg"))
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamFM_3kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_3kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamPM")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamPM }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "stdMT_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = stdMT_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "lotusB2B")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = lotusB2B }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "lotusSTD")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = lotusSTD }, Int32.Parse(cellValue.Replace(",", "")));
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

                var sortedSizeCounts = sizeCounts
                                        .OrderBy(entry => entry.Key.date)
                                        .ThenBy(entry => skuOrderMT.IndexOf(entry.Key.sku))
                                        .ToDictionary(entry => entry.Key, entry => entry.Value);

                if (File.Exists(outPutPath))
                {
                    File.Delete(outPutPath);
                }

                using (ExcelPackage excelPackage = new ExcelPackage(outPutPath))
                {
                    // Add a new worksheet to the Excel package
                    string sheetName = Path.GetFileNameWithoutExtension(excelFile);
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                    worksheet.Column(4).Style.Numberformat.Format = "0.000";

                    // Headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "Size";
                    worksheet.Cells[1, 3].Value = "Weight";
                    worksheet.Cells[1, 4].Value = "Weight(Ton)";

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
                        worksheet.Cells[row, 2].Value = kvp.Key.sku;
                        worksheet.Cells[row, 3].Value = kvp.Value;
                        worksheet.Cells[row, 4].Value = kvp.Value/1000f;

                        row++;
                        previousDate = kvp.Key.date;
                    }

                    // Save the Excel package to a file
                    if (File.Exists(outPutPath))
                    {
                        excelPackage.Save();
                    }
                    else
                    {
                        excelPackage.SaveAs(new FileInfo(outPutPath));
                    }
                }
                foreach (var sizeCount in sizeCounts.OrderBy(kvp => kvp.Key.date).ThenBy(kvp => kvp.Key.sku))
                {
                    Console.WriteLine(1);
                    Console.WriteLine($"{sizeCount.Key.date}: size = {sizeCount.Key.sku} Weight = {sizeCount.Value}");
                }

                Console.WriteLine(); // Add a separator line between files
            }

        }

        private static void MTLotusVolume()
        {
            // Set the license context before using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial

            // Change the directory path as needed
            string directoryPath = @"C:\Volume_Sales\MT(Lotus)";
            string outPutPath = @"C:\Volume_Sales\Volume_output\outputMT(Lotus).xlsx";

            // Get all Excel files in the specified directory
            string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the directory.");
                return;
            }


            // Set the console encoding to UTF-8
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            foreach (string excelFile in excelFiles)
            {
                Console.WriteLine($"Processing Excel file: {excelFile}");

                //SKU declare
                string super_premiumMT = "super_premiumMT";
                string foamMT = "foamMT";
                string foamFM_10kg = "foamFM_10kg";
                string foamFM_5kg = "foamFM_5kg";
                string foamFM_3kg = "foamFM_3kg";
                string foamPM = "foamPM";
                string stdMT_10kg = "stdMT_10kg";
                string lotusB2B = "lotusB2B";
                string lotusSTD = "lotusSTD";

                var skuOrderMT = new List<string>
                {
                    "super_premiumMT",
                    "foamMT",
                    "foamFM_10kg",
                    "foamFM_5kg",
                    "foamFM_3kg",
                    "foamPM",
                    "stdMT_10kg",
                    "lotusB2B",
                    "lotusSTD"
                };

                Dictionary<DateAndSku, int> sizeCounts = new Dictionary<DateAndSku, int>();

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



                        for (int row = 3; row <= worksheet.Dimension.Rows - 3; row++)
                        {
                            var nameProduct = worksheet.Cells[row, 1].Text;
                            columncount = 1;
                            Console.WriteLine($"Row {row}:");

                            for (int col = 3; col <= worksheet.Dimension.Columns - 3; col += 3)
                            {
                                date = dateColumns[columncount];
                                var cellValue = worksheet.Cells[row, col].Text;
                                Console.WriteLine($"{nameProduct} -> {date}: {cellValue}");

                                if (nameProduct.Contains("กุ้ง") && !nameProduct.Contains("แช่แข็ง") && !nameProduct.Contains("หมึก") && !nameProduct.Contains("PDTO") && !nameProduct.Contains("BS") && cellValue != "")
                                {
                                    string sku = DetermineSizeMT(nameProduct);
                                    DateAndSku key = new DateAndSku { date = date, sku = sku };
                                    if (sizeCounts.TryGetValue(key, out int weight) && cellValue != "")
                                    {
                                        weight += Int32.Parse(cellValue.Replace(",", ""));
                                        sizeCounts[key] = weight;
                                    }
                                    else if (sku == "super_premiumMT")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = super_premiumMT }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamMT")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamMT }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamFM_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if ((sku == "foamFM_5kg"))
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_5kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamFM_3kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamFM_3kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "foamPM")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = foamPM }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "stdMT_10kg")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = stdMT_10kg }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "lotusB2B")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = lotusB2B }, Int32.Parse(cellValue.Replace(",", "")));
                                    }
                                    else if (sku == "lotusSTD")
                                    {
                                        sizeCounts.Add(new DateAndSku { date = date, sku = lotusSTD }, Int32.Parse(cellValue.Replace(",", "")));
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

                var sortedSizeCounts = sizeCounts
                                        .OrderBy(entry => entry.Key.date)
                                        .ThenBy(entry => skuOrderMT.IndexOf(entry.Key.sku))
                                        .ToDictionary(entry => entry.Key, entry => entry.Value);

                if (File.Exists(outPutPath))
                {
                    File.Delete(outPutPath);
                }

                using (ExcelPackage excelPackage = new ExcelPackage(outPutPath))
                {
                    // Add a new worksheet to the Excel package
                    string sheetName = Path.GetFileNameWithoutExtension(excelFile);
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                    worksheet.Column(4).Style.Numberformat.Format = "0.000";

                    // Headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "Size";
                    worksheet.Cells[1, 3].Value = "Weight";
                    worksheet.Cells[1, 4].Value = "Weight(Ton)";

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
                        worksheet.Cells[row, 2].Value = kvp.Key.sku;
                        worksheet.Cells[row, 3].Value = kvp.Value;
                        worksheet.Cells[row, 4].Value = kvp.Value / 1000f;

                        row++;
                        previousDate = kvp.Key.date;
                    }

                    // Save the Excel package to a file
                    if (File.Exists(outPutPath))
                    {
                        excelPackage.Save();
                    }
                    else
                    {
                        excelPackage.SaveAs(new FileInfo(outPutPath));
                    }
                }
                foreach (var sizeCount in sizeCounts.OrderBy(kvp => kvp.Key.date).ThenBy(kvp => kvp.Key.sku))
                {
                    Console.WriteLine(1);
                    Console.WriteLine($"{sizeCount.Key.date}: size = {sizeCount.Key.sku} Weight = {sizeCount.Value}");
                }

                Console.WriteLine(); // Add a separator line between files
            }

        }
    }
}
