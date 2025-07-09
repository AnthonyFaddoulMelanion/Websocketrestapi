using Binance.Net.Clients;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WebSocketRESTAPI.Helpers
{
    public class ExchangeInfo
    {
        public List<Symbol> symbols { get; set; }
    }
    public class KlineEntry
    {
        public DateTime OpenTime { get; set; }
        public decimal Open { get; set; }
        public decimal High { get; set; }
        public decimal Low { get; set; }
        public decimal Close { get; set; }
    }
    public class Symbol
    {
        public string symbol { get; set; }
        public string status { get; set; }
        public string baseAsset { get; set; }
        public string quoteAsset { get; set; }
    }
    public static class RestClient
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public static async Task<List<string>> GetSymbolsAsync()
        {
            string url = "https://api.binance.com/api/v3/exchangeInfo";

            try
            {
                HttpResponseMessage response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string json = await response.Content.ReadAsStringAsync();

                ExchangeInfo exchangeInfo = JsonSerializer.Deserialize<ExchangeInfo>(json);

                var activeSymbols = new List<string>();
                foreach (var symbol in exchangeInfo.symbols)
                {
                    if (symbol.status == "TRADING")
                        activeSymbols.Add(symbol.symbol);
                }

                return activeSymbols;
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to fetch symbols from Binance REST API: " + ex.Message, ex);
            }
        }

        public static async Task<List<KlineEntry>> GetKlinesAsync(string symbol, string interval, DateTime start, DateTime end)
        {
            var client = new HttpClient();
            string url = $"https://api.binance.com/api/v3/klines?symbol={symbol.ToUpper()}&interval={interval}&startTime={ToUnixMillis(start)}&endTime={ToUnixMillis(end)}";

            var response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var content = await response.Content.ReadAsStringAsync();
            var json = JsonDocument.Parse(content).RootElement;

            var result = new List<KlineEntry>();

            foreach (var entry in json.EnumerateArray())
            {
                result.Add(new KlineEntry
                {
                    OpenTime = DateTimeOffset.FromUnixTimeMilliseconds(entry[0].GetInt64()).DateTime,
                    Open = decimal.Parse(entry[1].GetString()),
                    High = decimal.Parse(entry[2].GetString()),
                    Low = decimal.Parse(entry[3].GetString()),
                    Close = decimal.Parse(entry[4].GetString()),
                });
            }

            return result;
        }

        private static long ToUnixMillis(DateTime dt)
        {
            return new DateTimeOffset(dt).ToUnixTimeMilliseconds();
        }
    }
}
