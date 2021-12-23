using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Extensions.Options;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using VersionControlHelper.Config;
using VersionControlHelper.Domain;

namespace VersionControlHelper.Service
{
    public abstract class BaseCommitHistoryService : IDisposable, ICommitHistoryService
    {
        protected bool disposedValue;

        protected readonly AppConfig _config;
        protected readonly VssCredentials _clientCredentials;
        protected readonly VssConnection _connection;

        public abstract int? Top { get; set; }
        public abstract string Project { get; set; }
        public abstract Enums.VersionControlType Type { get; }

        public BaseCommitHistoryService(IOptionsSnapshot<AppConfig> configSnapshot)
        {
            this._config = configSnapshot?.Value ?? throw new ConfigFileException(nameof(AppConfig));

            var serverUrl = new Uri(_config.ServerUrl);

            ICredentials credentials = this._config.Credential.UseLoginUser
                ? CredentialCache.DefaultCredentials
                : new NetworkCredential(this._config.Credential.UserName, this._config.Credential.Password, this._config.Credential.Domain);

            this._clientCredentials = new VssCredentials(new WindowsCredential(credentials));

            this._connection = new VssConnection(serverUrl, this._clientCredentials);
        }

        public abstract IAsyncEnumerable<ICommitResult> GetCommitsAsync();

        public virtual async Task ExportExcelAsync(string folderPath = null, string fileName = null)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                folderPath = this._config.ExportPath;

            if (string.IsNullOrWhiteSpace(fileName))
                fileName = this._config.Project;

            using ExcelPackage package = new ExcelPackage();

            ExcelWorksheet myWorkSheet = package.Workbook.Worksheets.Add(fileName);

            int rowIndex = 2;

            // 使用 TimeZoneInfo 先取得台北時區
            var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time");




            await foreach (ICommitResult commitResult in this.GetCommitsAsync())
            {
                try
                {
                    int commitBeginRowIndex = rowIndex;

                    // 再使用 TimeZoneInfo 來變更時間
                    var convertedDateTime = TimeZoneInfo.ConvertTime(commitResult.CreatedDate, timeZone);

                    myWorkSheet.Cells[$"A{rowIndex}"].Value = convertedDateTime.ToString("yyyy-MM-dd\nHH:mm");
                    myWorkSheet.Cells[$"B{rowIndex}"].Value = commitResult.Committer;
                    myWorkSheet.Cells[$"C{rowIndex}"].Value = commitResult.Comment;

                    foreach (var change in commitResult.Changes)
                    {
                        myWorkSheet.Cells[$"D{rowIndex}"].Value = change.Path;
                        myWorkSheet.Cells[$"E{rowIndex}"].Value = change.ChangeType.ToString();
                        //myWorkSheet.Cells[$"F{rowIndex}"].Value = change.IsFolder ? "是" : "否";

                        rowIndex++;
                    }

                    //合併儲存格
                    //myWorkSheet.Cells[commitBeginRowIndex < rowIndex ? $"A{commitBeginRowIndex}:A{rowIndex - 1}" : $"A{rowIndex - 1}:A{commitBeginRowIndex}"].Merge = true;
                    //myWorkSheet.Cells[commitBeginRowIndex < rowIndex ? $"B{commitBeginRowIndex}:B{rowIndex - 1}" : $"B{rowIndex - 1}:B{commitBeginRowIndex}"].Merge = true;
                    //myWorkSheet.Cells[commitBeginRowIndex < rowIndex ? $"C{commitBeginRowIndex}:C{rowIndex - 1}" : $"C{rowIndex - 1}:C{commitBeginRowIndex}"].Merge = true;

                    if (commitBeginRowIndex == rowIndex)
                        rowIndex++;
                }
                catch (Exception ex)
                {
                    var aa = ex;
                }
            }

            int dataCount = rowIndex - 1;

            string[] titles = new[]
            {
                "變更日期",
                "作者",
                "註解",
                "變動路徑",
                "變動類型",
                //"是否為資料夾",
            };

            SetAsTemplate(ref myWorkSheet, fileName, dataCount, eOrientation.Landscape, titles);

            myWorkSheet.Column(4).Width = 100;
            myWorkSheet.Cells[$"C{3}:E{dataCount}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            byte[] result = package.GetAsByteArray();

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            await File.WriteAllBytesAsync(Path.Combine(folderPath, $"{fileName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx"), result);
        }

        #region 樣板
        /// <summary>
        /// 設定邊界、紙張大小、縮放至單一頁面
        /// </summary>
        /// <param name="workSheet">工作表</param>
        /// <param name="Margin">邊界值(單位inch)</param>
        /// <param name="Orient">方向預設直向</param>
        /// <param name="PaperSize">紙張大小，預設A4/param>
        /// <param name="FitToWidth">將所有欄放入單一頁面</param>
        /// <param name="FitToHeight">將所有列放入單一頁面</param>
        protected void SetPrinterSettings(ref ExcelWorksheet workSheet, eOrientation Orient = eOrientation.Portrait, decimal Margin = 0.2m, ePaperSize PaperSize = ePaperSize.A4, int FitToWidth = 1, int FitToHeight = 0)
        {
            //列印方向
            workSheet.PrinterSettings.Orientation = Orient;

            //上下左右邊界
            workSheet.PrinterSettings.TopMargin = Margin;
            workSheet.PrinterSettings.BottomMargin = Margin;
            workSheet.PrinterSettings.LeftMargin = Margin;
            workSheet.PrinterSettings.RightMargin = Margin;

            if (Orient == eOrientation.Portrait)
            {
                workSheet.PrinterSettings.BottomMargin = Margin * 2;
            }

            //頁首頁尾邊界
            workSheet.PrinterSettings.HeaderMargin = Margin;
            workSheet.PrinterSettings.FooterMargin = 0.04m;
            //紙張大小
            workSheet.PrinterSettings.PaperSize = PaperSize;
            //調整至一頁大小
            workSheet.PrinterSettings.FitToPage = true;
            workSheet.PrinterSettings.FitToWidth = FitToWidth;
            workSheet.PrinterSettings.FitToHeight = FitToHeight;
            //設定列印標題
            workSheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$2");
            //頁尾顯示頁碼
            workSheet.HeaderFooter.differentOddEven = false;
            workSheet.HeaderFooter.OddFooter.CenteredText = $"第{ExcelHeaderFooter.PageNumber}頁，共{ExcelHeaderFooter.NumberOfPages}頁";
            workSheet.HeaderFooter.EvenFooter.CenteredText = $"第{ExcelHeaderFooter.PageNumber}頁，共{ExcelHeaderFooter.NumberOfPages}頁";
        }

        /// <summary>
        /// 設定框線、自動換行
        /// </summary>
        /// <param name="range">範圍</param>
        /// <param name="IsSetBorder">是否設定框線</param>
        /// <param name="IsSetWrapText">是否設定自動換行</param>
        protected void SetBorderOrWrap(ref ExcelRange range, bool IsSetBorder = true, bool IsSetWrapText = true, ExcelBorderStyle BorderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center; //垂直置中

            if (IsSetBorder)
            {
                range.Style.Border.Bottom.Style = BorderStyle;  //下框線
                range.Style.Border.Top.Style = BorderStyle; //上框線
                range.Style.Border.Right.Style = BorderStyle; //右框線
                range.Style.Border.Left.Style = BorderStyle; //左框線
            }

            range.Style.WrapText = IsSetWrapText; //自動換行
        }


        /// <summary>
        /// 設定活頁簿為樣板
        /// </summary>
        /// <param name="myWorkSheet">ref 活頁簿</param>
        /// <param name="bigTitleName">大標題</param>
        /// <param name="dataCount">資料的數量</param>
        /// <param name="orientation">列印方向</param>
        /// <param name="titles">params 標題名稱</param>
        protected void SetAsTemplate(ref ExcelWorksheet myWorkSheet, string bigTitleName, int dataCount, eOrientation orientation = eOrientation.Landscape, params string[] titles)
        {
            int endRow = dataCount;
            int endCol = titles.Count();

            #region 全頁
            //字型
            myWorkSheet.Cells.Style.Font.Name = "標楷體";
            //文字大小
            myWorkSheet.Cells.Style.Font.Size = 13;
            //設定文字水平對齊，此範例設定水平置中...還有其他屬性可以設定
            myWorkSheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            #endregion

            #region 大標
            //大標題
            myWorkSheet.Cells["A1"].Value = bigTitleName;
            //文字大小
            myWorkSheet.Cells["A1"].Style.Font.Size = 16;
            //合併儲存格
            myWorkSheet.Cells[1, 1, 1, endCol].Merge = true;
            #endregion

            #region 小標
            for (int col = 1; col <= endCol; col++)
            {
                myWorkSheet.Cells[2, col].Value = titles[col - 1];
                myWorkSheet.Column(col).Width = 30;
            }

            //一定要加這行..不然會報錯
            myWorkSheet.Cells[2, 1, 2, endCol].Style.Fill.PatternType = ExcelFillStyle.Solid;  //從第幾Row，col開始～第幾Row，col結束
            //上色
            myWorkSheet.Cells[2, 1, 2, endCol].Style.Fill.BackgroundColor.SetColor(Color.LightGray);   //從第幾Row，col開始～第幾Row，col結束
            #endregion

            #region 設定框線、列印設定
            ExcelRange printRange = myWorkSheet.Cells[1, 1, endRow, endCol];
            SetBorderOrWrap(ref printRange);
            SetPrinterSettings(ref myWorkSheet, orientation);
            #endregion
        }
        #endregion

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: 處置受控狀態 (受控物件)
                    this._connection.Dispose();
                }

                // TODO: 釋出非受控資源 (非受控物件) 並覆寫完成項
                // TODO: 將大型欄位設為 Null
                disposedValue = true;
            }
        }

        // // TODO: 僅有當 'Dispose(bool disposing)' 具有會釋出非受控資源的程式碼時，才覆寫完成項
        // ~TFSHistoryService()
        // {
        //     // 請勿變更此程式碼。請將清除程式碼放入 'Dispose(bool disposing)' 方法
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // 請勿變更此程式碼。請將清除程式碼放入 'Dispose(bool disposing)' 方法
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
