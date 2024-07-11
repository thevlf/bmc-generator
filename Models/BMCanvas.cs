using Newtonsoft.Json;

namespace BMCGen
{
    public class BMC
    {
        [JsonProperty("Customer Segments")]
        public List<string> CustomerSegments { get; set; } = new List<string>();

        [JsonProperty("Value Propositions")]
        public List<string> ValuePropositions { get; set; } = new List<string>();

        public List<string> Channels { get; set; } = new List<string>();

        [JsonProperty("Customer Relationships")]
        public List<string> CustomerRelationships { get; set; } = new List<string>();

        [JsonProperty("Key Resources")]
        public List<string> KeyResources { get; set; } = new List<string>();

        [JsonProperty("Key Activities")]
        public List<string> KeyActivities { get; set; } = new List<string>();

        [JsonProperty("Key Partners")]
        public List<string> KeyPartners { get; set; } = new List<string>();

        [JsonProperty("Cost Structure")]
        public List<string> CostStructure { get; set; } = new List<string>();

        [JsonProperty("Revenue Streams")]
        public List<string> RevenueStreams { get; set; } = new List<string>();
    }
}