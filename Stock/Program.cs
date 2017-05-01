using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Stock
{
    class Program
    {
        static List<YearData> YearDataList = new List<YearData>();

        static void Main(string[] args)
        {            
            if (args[0] == "Download")
            {
                //取股市資料
                string param = "qdate=" + args[1] + "/" + args[2] + "/" + args[3] + "&select2=" + args[4] + "&Sort_kind=STKNO&download=html";
                byte[] paramBytes = Encoding.ASCII.GetBytes(param);
                string webResponse = StockWeb(paramBytes, "http://www.twse.com.tw/ch/trading/exchange/BWIBBU/BWIBBU_d.php", Encoding.UTF8);

                //將資料寫入Excel(建立來源檔案)
                HtmlDocument stockDataHtmlDoc = new HtmlDocument();
                stockDataHtmlDoc.LoadHtml(webResponse);
                if (webResponse != "null" && !webResponse.Contains("查無資料"))
                {
                    string path = args[5] + "StockDataExport_" + args[1] + "_" + args[2] + "_" + args[3] + "_" + args[4] + ".xls";
                    RenderDataToExcel(stockDataHtmlDoc, path);
                    Console.WriteLine(args[1] + "/" + args[2] + "/" + args[3] + " 下載完成!!");
                }
                else
                {
                    Console.WriteLine(args[1] + "/" + args[2] + "/" + args[3] + " 查無資料!!");
                }                
            }
            else if (args[0] == "Analyze")
            {
                Console.WriteLine("分析開始!!");
                StockAnalysis(args[1], args[2]);
                Console.WriteLine("分析完成!!");
            }
        }

        public static void RenderDataToExcel(HtmlDocument StockDataHtmlDoc, string Path)
        {
            HtmlDocument StockTableDataHtmlDoc = new HtmlDocument();
            MemoryStream ms = new MemoryStream();
            FileStream DownloadFileStream = new FileStream(Path, FileMode.Create, FileAccess.Write);
            
            //建立Excel資料
            int Row = 0;
            int Count = 0;
            HSSFWorkbook ExcelWorkBook = new HSSFWorkbook();
            ISheet ExcelSheet = ExcelWorkBook.CreateSheet("股市資料");
            ExcelSheet.CreateRow(0).CreateCell(0).SetCellValue("證券代號");
            ExcelSheet.GetRow(0).CreateCell(1).SetCellValue("證券名稱");
            ExcelSheet.GetRow(0).CreateCell(2).SetCellValue("本益比");
            ExcelSheet.GetRow(0).CreateCell(3).SetCellValue("殖利率(%)");
            ExcelSheet.GetRow(0).CreateCell(4).SetCellValue("股價淨值比");
            foreach (HtmlNode tdnode in StockDataHtmlDoc.DocumentNode.SelectNodes("//tbody//td"))
            {
                if ((Count % 5) == 0)
                {
                    Row++;
                    ExcelSheet.CreateRow(Row);
                }
                ExcelSheet.GetRow(Row).CreateCell((Count % 5)).SetCellValue(tdnode.InnerText);
                Count++;
            }

            //建立Excel檔案
            ExcelWorkBook.Write(ms);           
            DownloadFileStream.Write(ms.ToArray(), 0, ms.ToArray().Length);
            DownloadFileStream.Flush();
            DownloadFileStream.Close();
        }

        private static void StockAnalysis(string Source, string Target)
        {
            HSSFWorkbook ReadWriteTargetFile = new HSSFWorkbook(new FileStream(Target + "AnalysisData.xls", FileMode.Open, FileAccess.ReadWrite));

            //取得來源資料夾內所有檔案
            foreach (string FileName in Directory.GetFiles(Source, "*.xls", SearchOption.TopDirectoryOnly))
            {
                string[] FileNameArray = FileName.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
                HSSFWorkbook ReadSourceFile = new HSSFWorkbook(new FileStream(FileName, FileMode.Open, FileAccess.Read));
                if (ReadSourceFile.NumberOfSheets == 0)
                {
                    Console.Write("File:" + FileName + "沒有資料!!");
                    continue;
                }
                else
                {
                    //建立要處理資料的物件(讀取目標檔案)
                    InitData(FileNameArray[1], FileNameArray[2], ReadWriteTargetFile);
                    //將來源檔案資料存到物件中
                    ProcessData(FileNameArray[1], FileNameArray[2], ReadSourceFile);                    
                }
            }
            //將物件存到目標檔案中
            SaveData(ReadWriteTargetFile, Target);
        }

        private static void SaveData(HSSFWorkbook ReadWriteTargetFile, string Target)
        {
            foreach (YearData YearDataObj in YearDataList)
            {
                ISheet TargetFileSheet = ReadWriteTargetFile.GetSheetAt(Convert.ToInt16(YearDataObj.Id));
                foreach (MonthData MonthDataObj in YearDataObj.MonthData)
                {
                    foreach (StockData StockDataObj in MonthDataObj.StockData)
                    {
                        int BasicColumn = GetColumnByMonth(MonthDataObj.Month);
                        IRow TargetFileRow = TargetFileSheet.GetRow(StockDataObj.Row);
                        if(TargetFileRow == null)
                        {
                            TargetFileRow = TargetFileSheet.CreateRow(StockDataObj.Row);
                        }
                        ICell TargetFilePerCell = TargetFileRow.GetCell(BasicColumn);
                        if(TargetFilePerCell == null)
                        {
                            TargetFilePerCell = TargetFileRow.CreateCell(BasicColumn);
                        }
                        TargetFilePerCell.SetCellValue(Math.Round((StockDataObj.Per / StockDataObj.PerCount), 2));
                        ICell TargetFileYieldCell = TargetFileRow.GetCell(BasicColumn + 1);
                        if (TargetFileYieldCell == null)
                        {
                            TargetFileYieldCell = TargetFileRow.CreateCell(BasicColumn + 1);
                        }
                        TargetFileYieldCell.SetCellValue(Math.Round((StockDataObj.Yield / StockDataObj.YieldCount), 2));
                        ICell TargetFilePbrCell = TargetFileRow.GetCell(BasicColumn + 2);
                        if (TargetFilePbrCell == null)
                        {
                            TargetFilePbrCell = TargetFileRow.CreateCell(BasicColumn + 2);
                        }
                        TargetFilePbrCell.SetCellValue(Math.Round((StockDataObj.Pbr / StockDataObj.PbrCount), 2));
                        if (StockDataObj.IsNew)  //新增的資料
                        {
                            ICell TargetFileIdCell = TargetFileRow.CreateCell(0);
                            TargetFileIdCell.SetCellValue(Convert.ToInt32(StockDataObj.Id));
                            ICell TargetFileNameCell = TargetFileRow.CreateCell(1);
                            TargetFileNameCell.SetCellValue(StockDataObj.Name);
                        }
                    }
                }
            }
            ReadWriteTargetFile.Write(new FileStream(Target + "AnalysisData.xls", FileMode.Open, FileAccess.Write));
        }

        private static int GetColumnByMonth(string Month)
        {
            switch (Month)
            {
                case "01": return 14;
                case "02": return 17;
                case "03": return 20;
                case "04": return 23;
                case "05": return 26;
                case "06": return 29;
                case "07": return 32;
                case "08": return 35;
                case "09": return 2;
                case "10": return 5;
                case "11": return 8;
                case "12": return 11; 
                default: return 38;
            }
        }

        private static void ProcessData(string Year, string Month, HSSFWorkbook ReadSourceFile)
        {
            string YearId = GetYearDataID(Year, Month);
            int YearDataIndex = YearDataList.FindIndex(row => row.Id == YearId);
            int MonthDataIndex = YearDataList[YearDataIndex].MonthData.FindIndex(row => row.Month == Month);
            ISheet SourceFileSheet = ReadSourceFile.GetSheetAt(0);
            for (int i = 1; i <= SourceFileSheet.LastRowNum; i++)
            {
                IRow SourceFileRow = SourceFileSheet.GetRow(i);
                int StockDataIndex = YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData.FindIndex(row => row.Id == SourceFileRow.GetCell(0).StringCellValue);
                //證劵代號不存在就新增
                if (StockDataIndex == -1)  
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData.Add(new StockData { Id = SourceFileRow.GetCell(0).StringCellValue, Name = SourceFileRow.GetCell(1).StringCellValue, PerCount = 0, YieldCount = 0, PbrCount = 0, Pbr = 0, Per = 0, Yield = 0, IsNew = true });
                    StockDataIndex = YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData.Count() - 1;
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].Row = YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex - 1].Row + 1;
                }
                //本益比
                if (SourceFileRow.GetCell(2).StringCellValue != "-" && SourceFileRow.GetCell(2).StringCellValue != "0.00")
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].PerCount++;
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].Per += Convert.ToDouble(SourceFileRow.GetCell(2).StringCellValue);
                }
                //殖利率
                if (SourceFileRow.GetCell(3).StringCellValue != "-" && SourceFileRow.GetCell(3).StringCellValue != "0.00")
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].YieldCount++;
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].Yield += Convert.ToDouble(SourceFileRow.GetCell(3).StringCellValue);
                }
                //股價淨值比
                if (SourceFileRow.GetCell(4).StringCellValue != "-" && SourceFileRow.GetCell(4).StringCellValue != "0.00")
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].PbrCount++;
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData[StockDataIndex].Pbr += Convert.ToDouble(SourceFileRow.GetCell(4).StringCellValue);
                }
            }
        }

        private static void InitData(string Year, string Month, HSSFWorkbook ReadWriteTargetFile)
        {
            //第幾年
            string YearId = GetYearDataID(Year, Month);
            if (YearDataList.Where(row => row.Id == YearId).Count() == 0)
            {
                YearDataList.Add(new YearData { Id = YearId });
            }

            //第幾個月
            int YearDataIndex = YearDataList.FindIndex(row => row.Id == YearId);
            if (YearDataList[YearDataIndex].MonthData == null)
            {
                YearDataList[YearDataIndex].MonthData = new List<MonthData>();
            }           
            if (YearDataList[YearDataIndex].MonthData.Where(row => row.Month == Month).Count() == 0)
            {
                YearDataList[YearDataIndex].MonthData.Add(new MonthData { Month = Month });
                int MonthDataIndex = YearDataList[YearDataIndex].MonthData.FindIndex(row => row.Month == Month);
                ISheet TargetFileSheet = ReadWriteTargetFile.GetSheetAt(Convert.ToInt16(YearId));
                if (YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData == null)
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData = new List<StockData>();
                }
                for (int i = 3; i <= TargetFileSheet.LastRowNum; i++)
                {
                    YearDataList[YearDataIndex].MonthData[MonthDataIndex].StockData.Add(new StockData { Id = TargetFileSheet.GetRow(i).GetCell(0).NumericCellValue.ToString(), Name = TargetFileSheet.GetRow(i).GetCell(1).StringCellValue, PerCount = 0, YieldCount = 0, PbrCount = 0, Pbr = 0, Per = 0, Yield = 0, Row = i, IsNew = false });
                }
            }
        }

        private static string GetYearDataID(string Year, string Month)
        {
            int MonthInteger = 0;
            int YearInteger = Convert.ToInt16(Year);
            int YearId = 0;
            if (Month.Substring(0, 1) == "0")
            {
                MonthInteger = Convert.ToInt16(Month.Substring(1, 1));
            }
            else
            {
                MonthInteger = Convert.ToInt16(Month);
            }

            //以94/9/1為基準
            if ((YearInteger - 94) == 0)
            {
                YearId = 1;
            }
            else
            {
                YearId = YearInteger - 94;
                if (MonthInteger >= 9)
                {
                    YearId++;
                }
            }

            return YearId.ToString();
        }

        public static string StockWeb(byte[] bs, string url, Encoding code)
        {
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            req.ContentLength = bs.Length;
            req.Headers.Add("Access-Control-Allow-Origin", "*");
            req.GetRequestStream().Write(bs, 0, bs.Length);
            return new StreamReader(req.GetResponse().GetResponseStream(), code).ReadToEnd();
        }

    }

    class StockData
    {
        public string Id { set; get; }
        public string Name { set; get; }
        public double Per { set; get; }  //本益比
        public double Yield { set; get; }  //殖利率
        public double Pbr { set; get; }  //股價淨值比
        public int PerCount { set; get; }
        public int YieldCount { set; get; }
        public int PbrCount { set; get; }
        public int Row { set; get; }
        public bool IsNew { set; get; }
    }

    class MonthData
    {
        public string Month { set; get; }
        public List<StockData> StockData { set; get; }
    }

    class YearData
    {
        public string Id { set; get; }
        public List<MonthData> MonthData { set; get; }
    }
}
