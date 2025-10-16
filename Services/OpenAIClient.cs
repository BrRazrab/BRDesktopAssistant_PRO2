using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using BRDesktopAssistant.Models;

namespace BRDesktopAssistant.Services
{
    public class OpenAIClient
    {
        private readonly HttpClient _http;
        private readonly string _apiKey;
        private readonly string _baseUrl;
        private readonly string _model;

        public OpenAIClient()
        {
            _http = new HttpClient();
            _apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? "";
            if (string.IsNullOrWhiteSpace(_apiKey))
                throw new InvalidOperationException("Не найден OPENAI_API_KEY в переменных окружения.");

            _baseUrl = "https://api.openai.com/v1/chat/completions";
            _model = "gpt-4o-mini";
        }

        public async Task<string> GetChatCompletionAsync(ChatMessage[] messages)
        {
            var payload = new ChatCompletionRequest
            {
                model = _model,
                messages = messages
            };

            using var req = new HttpRequestMessage(HttpMethod.Post, _baseUrl);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            using var res = await _http.SendAsync(req);
            var json = await res.Content.ReadAsStringAsync();
            if (!res.IsSuccessStatusCode)
                throw new Exception($"OpenAI error {res.StatusCode}: {json}");

            var data = JsonSerializer.Deserialize<ChatCompletionResponse>(json);
            return data?.choices?[0]?.message?.content ?? "(пустой ответ)";
        }

        private class ChatCompletionRequest
        {
            public string model { get; set; } = string.Empty;
            public ChatMessage[] messages { get; set; } = Array.Empty<ChatMessage>();
            [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public double? temperature { get; set; }
        }

        private class ChatCompletionResponse
        {
            public Choice[]? choices { get; set; }
        }

        private class Choice
        {
            public ChatMessage? message { get; set; }
        }
    }
}
