using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToJson
{
    public class UpdateExcel
    {
        /// <summary>
        /// 污染等级
        /// </summary>
        public static List<string> PollutionDegree
        {
            get
            {
                var list = new List<string>() { "优", "良", "轻度污染", "中度污染", "重度污染" };
                return list;
            }
        }

        public static bool UpdateExcelData(string excelFilePath, string excelName)
        {
            XSSFWorkbook excelBook = new XSSFWorkbook(File.Open(excelFilePath, FileMode.Open));
            var sheet = excelBook.GetSheetAt(0);
            var headerRow = sheet.GetRow(0);
            //total columns
            int cellCount = headerRow.LastCellNum;
            //total rows
            int rowCount = sheet.LastRowNum;

            switch (excelName)
            {
                case Constant.Water:
                    UpdateWaterData(sheet, cellCount, rowCount);
                    break;
                case Constant.Air:
                    UpdateAirData(sheet, cellCount, rowCount);
                    break;
            }

            try
            {
                FileStream writefile = new FileStream("d:\\" + excelName + ".xlsx", FileMode.Create, FileAccess.Write);
                excelBook.Write(writefile);
                writefile.Close();
                return true;
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// 更新空气Excel数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellCount"></param>
        /// <param name="rowCount"></param>
        private static void UpdateAirData(ISheet sheet, int cellCount, int rowCount)
        {
            for (var i = (sheet.FirstRowNum + 1); i < rowCount; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                for (var j = row.FirstCellNum; j < cellCount; j++)
                {
                    System.Threading.Thread.Sleep(10);
                    if (j == 1)//污染等级
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(PollutionDegree[new Random().Next(0, 4)]);
                    }
                    else if (j == 3)//pm2.5
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else if (j == 4)//悬浮物
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else if (j == 5)//so2
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else
                    {
                        if (row.GetCell(j) != null)
                        {
                            row.GetCell(j).SetCellValue(row.GetCell(j).ToString());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 更新水体质量Excel数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellCount"></param>
        /// <param name="rowCount"></param>
        private static void UpdateWaterData(ISheet sheet, int cellCount, int rowCount)
        {
            for (var i = (sheet.FirstRowNum + 1); i < rowCount; i++)
            {
                var row = sheet.GetRow(i);

                if (row == null)
                {
                    continue;
                }

                System.Threading.Thread.Sleep(10);
                for (var j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (j == 1)//污染等级
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(PollutionDegree[new Random().Next(0, 4)]);
                    }
                    else if (j == 3)//ph值
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().Next(0, 9));
                    }
                    else if (j == 4)//悬浮物
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else if (j == 5)//cod
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().Next(0, 250));
                    }
                    else if (j == 6)//铅含量
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else if (j == 7)//汞含量
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                        }
                        row.GetCell(j).SetCellValue(new Random().NextDouble().ToString("0.000"));
                    }
                    else
                    {
                        if (row.GetCell(j) != null)
                        {
                            row.GetCell(j).SetCellValue(row.GetCell(j).ToString());
                        }
                    }
                }
            }
        }

    }
}
