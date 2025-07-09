using Binance.Net.Clients;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WebSocketRESTAPI.Helpers;
using WebSocketRESTAPI.UI;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;



namespace WebSocketRESTAPI
{
    public partial class Sheet1
    {
        private WebSocketClient _socketClient;
        private Microsoft.Office.Tools.Excel.Controls.Button btnPlotChart;
        private async void Sheet1_Startup(object sender, EventArgs e)
        {
            this.Change += Sheet1_Change;
            await InitializeDropdown();



            
            //CreateCandlestickButton();
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

        private async Task InitializeDropdown()
        {
            
            var symbols = await RestClient.GetSymbolsAsync();
            UI.DropdownManager.CreateSymbolDropdown(this.Range["H5"], symbols);
            this.Range["H5"].Style = "Accent3";
            this.Range["G3"].Value2 = "Threshold";

            this.Range["A7"].Value2 = "Interval";
            var intervals = new string[] { "1m", "5m", "15m", "1h", "4h", "1d" };
            string intervalList = string.Join(",", intervals);

            // Add dropdown to cell B7
            Excel.Range intervalCell = this.Range["B7"];
            intervalCell.Validation.Delete();
            intervalCell.Validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Excel.XlFormatConditionOperator.xlBetween,
                intervalList
            );
            intervalCell.Value2 = "1h";
        }

        private async void Sheet1_Change(Microsoft.Office.Interop.Excel.Range target)
        {
            if (target.Address[false, false] == "H5")
            {
                string symbol = target.Value2?.ToString();
                if (!string.IsNullOrEmpty(symbol))
                {
                    StartStreaming(symbol);
                    _ = LoadAndPlotCandles(symbol);
                }
            }

            //try
            //{
            //    string symbol = this.Range["B1"].Value2?.ToString();
            //    string interval = this.Range["B7"].Value2?.ToString();
            //    DateTime start = DateTime.Parse(this.Range["B8"].Value2?.ToString());
            //    DateTime end = DateTime.Parse(this.Range["B9"].Value2?.ToString());

            //    var klines = await RestClient.GetKlinesAsync(symbol, interval, start, end);
            //    PlotCandlestickChart(klines);
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            //}
        }
        private async Task LoadAndPlotCandles(string symbol)
        {
            try
            {
                // You can set these as fixed or read from Excel
                string interval = this.Range["B7"].Value2?.ToString() ?? "1h"; // or this.Range["B7"].Value2?.ToString();
                DateTime end = DateTime.UtcNow;
                DateTime start;

                switch (interval)
                {
                    case "1m": start = end.AddHours(-2); break;      // 120 candles
                    case "5m": start = end.AddHours(-10); break;     // 120 candles
                    case "15m": start = end.AddHours(-30); break;
                    case "1h": start = end.AddDays(-2); break;       // 48 candles
                    case "4h": start = end.AddDays(-10); break;
                    case "1d": start = end.AddDays(-60); break;      // 60 candles
                    default: start = end.AddDays(-1); break;
                }

                var klines = await RestClient.GetKlinesAsync(symbol, interval, start, end);
                PlotCandlestickChart(klines);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load candles: " + ex.Message);
            }
        }

        private async void StartStreaming(string symbol)
        {
            if (_socketClient != null)
            {
                await _socketClient.DisconnectAsync();
            }

            _socketClient = new WebSocketClient(symbol);
            _socketClient.OnBookTickerUpdate += UpdateBidAsk;

            await _socketClient.ConnectAsync();
        }

        private void UpdateBidAsk(decimal bid, decimal ask)
        {
            this.Range["G4"].Value2 = "Bid";
            this.Range["I4"].Value2 = "Ask";
            this.Range["G5"].Value2 = bid;
            this.Range["I5"].Value2 = ask;
            this.Range["G5"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            this.Range["I5"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);


            UI.FlashManager.FlashIfSpreadTooHigh(this.InnerObject, bid, ask);
        }



        //private async void BtnPlotChart_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        string symbol = this.Range["B1"].Value2?.ToString();
        //        string interval = this.Range["B7"].Value2?.ToString();
        //        DateTime start = DateTime.Parse(this.Range["B8"].Value2?.ToString());
        //        DateTime end = DateTime.Parse(this.Range["B9"].Value2?.ToString());

        //        var klines = await RestClient.GetKlinesAsync(symbol, interval, start, end);
        //        PlotCandlestickChart(klines);
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
        //    }
        //}

        private void PlotCandlestickChart(List<KlineEntry> klines)
        {
            var sheet = this;
            int startRow = 12;
            int row = startRow;

            // Headers
            sheet.Cells[row, 1].Value2 = "Time";
            sheet.Cells[row, 2].Value2 = "Open";
            sheet.Cells[row, 3].Value2 = "High";
            sheet.Cells[row, 4].Value2 = "Low";
            sheet.Cells[row, 5].Value2 = "Close";
            row++;

            // Data
            foreach (var k in klines)
            {
                sheet.Cells[row, 1].Value2 = k.OpenTime;
                sheet.Cells[row, 2].Value2 = k.Open;
                sheet.Cells[row, 3].Value2 = k.High;
                sheet.Cells[row, 4].Value2 = k.Low;
                sheet.Cells[row, 5].Value2 = k.Close;
                row++;
            }

            // Add chart
            var range = sheet.Range["B" + startRow, "E" + (row - 1)];
            var chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)sheet.ChartObjects();
            var chartObject = chartObjects.Add(500, 20, 600, 400);
            var chart = chartObject.Chart;

            chart.SetSourceData(range);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlStockOHLC;

            chart.HasTitle = true;
            chart.ChartTitle.Text = "Candlestick Chart: " + this.Range["B1"].Value2;
        }
    }
}
