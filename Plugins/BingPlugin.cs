using Microsoft.SemanticKernel;
using System.ComponentModel;
using System.Diagnostics;
using System.Net.Http;

namespace BMCGen
{
    public class BingPlugin
    {        
        [KernelFunction("Search")]
        [Description("Searches Bing for the specified query and returns the results")]
        public async Task<BingResults> SearchAsync(string query, string site)
        {
            string? subscriptionKey = Environment.GetEnvironmentVariable("bingKey");            
            string endpoint = "https://api.bing.microsoft.com/v7.0/search";
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", subscriptionKey);

            Trace.WriteLine("Searching the Web for: " + query);
            Logging.WriteLog("Searching the Web for: " + query);

            // Construct the URI of the search request
            var uriQuery = endpoint + "?q=" + Uri.EscapeDataString(query) + 
                            "site:" + site +
                            "&responseFilter=webpages";

            // Run the query
            HttpResponseMessage httpResponseMessage = await client.GetAsync(uriQuery).ConfigureAwait(false);

            // Deserialize the response content
            var responseContentString = await httpResponseMessage.Content.ReadAsStringAsync().ConfigureAwait(false);
            Newtonsoft.Json.Linq.JObject responseObjects = Newtonsoft.Json.Linq.JObject.Parse(responseContentString);

            // Convert the response to a BingResults object
            var results = new BingResults
            {
                Results = new List<BingResult>()
            };

            if (responseObjects!= null && responseObjects["webPages"] != null && responseObjects["webPages"]["value"] != null)
            {
                foreach (var result in responseObjects["webPages"]["value"])
                {
                    results.Results.Add(new BingResult
                    {
                        Title = result["name"].ToString(),
                        Description = result["snippet"].ToString(),
                        Url = result["url"].ToString()
                    });
                }
            }

            return results;
        }
    }
}

public class BingResults
{
    public List<BingResult>? Results { get; set; }
}

public class BingResult
{
    public string? Title { get; set; }
    public string? Description { get; set; }
    public string? Url { get; set; }
}
