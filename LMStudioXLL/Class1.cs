using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

public static class LMStudioFunctions
{
    private static readonly string defaultDomain = "localhost:1234";
    private static readonly HttpClient client = new HttpClient();

    [ExcelFunction(Description = "Queries LM Studio API and returns a response.")]
    public static object QueryLMStudio(string prompt, string domainPort = null, string model = "default-model")
    {
        try
        {
            string server = string.IsNullOrEmpty(domainPort) ? defaultDomain : domainPort;
            return QueryLMStudioAsync(prompt, server, model).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    // Async method to send API request to LM Studio
    private static async Task<string> QueryLMStudioAsync(string prompt, string domainPort, string model)
    {
        string serverURL = $"http://{domainPort}/v1/completions";

        try
        {
            // Build the JSON request payload
            var requestBody = new
            {
                model = model,
                prompt = prompt,
                max_tokens = 100
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            HttpContent content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            // Send the HTTP request
            HttpResponseMessage response = await client.PostAsync(serverURL, content);

            if (response.IsSuccessStatusCode)
            {
                string responseString = await response.Content.ReadAsStringAsync();
                dynamic jsonResponse = JsonConvert.DeserializeObject(responseString);
                return jsonResponse.choices[0].text.ToString();
            }
            else
            {
                return $"Error: {response.ReasonPhrase}";
            }
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }
}