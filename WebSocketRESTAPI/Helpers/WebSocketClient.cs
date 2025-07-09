using Binance.Net.Clients;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Reactive.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Websocket.Client;

namespace WebSocketRESTAPI.Helpers
{
    public class WebSocketClient
    {
        private readonly string _symbol;
        private WebsocketClient _client;

        public event Action<decimal, decimal> OnBookTickerUpdate;

        public WebSocketClient(string symbol)
        {
            _symbol = symbol.ToLower(); 
        }

        public async Task ConnectAsync()
        {
            var url = new Uri($"wss://stream.binance.com:9443/ws/{_symbol}@bookTicker");

            _client = new WebsocketClient(url);
            _client.ReconnectTimeout = TimeSpan.FromSeconds(30);
            _client.ReconnectionHappened.Subscribe(info =>
                Console.WriteLine($"Reconnection happened, type: {info.Type}"));

            _client.MessageReceived
                .Where(msg => !string.IsNullOrEmpty(msg.Text))
                .Subscribe(msg =>
                {
                    Console.WriteLine("Received WebSocket msg: " + msg.Text);  // 👈 DEBUG HERE
                    ParseMessage(msg.Text);
                });

            await _client.Start();
        }

        public async Task DisconnectAsync()
        {
            if (_client != null)
            {
                await _client.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "User closed");
                _client.Dispose();
                _client = null;
            }
        }

        private void ParseMessage(string json)
        {
            try
            {
                var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.TryGetProperty("b", out var bidProp) &&
                    root.TryGetProperty("a", out var askProp))
                {
                    decimal bid = decimal.Parse(bidProp.GetString());
                    decimal ask = decimal.Parse(askProp.GetString());

                    OnBookTickerUpdate?.Invoke(bid, ask);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Parse error: " + ex.Message);
            }
        }

}
}
