namespace BMCGen
{
    public class BingResponse
    {
        public string _type { get; set; }
        public QueryContext queryContext { get; set; }
        public WebPages webPages { get; set; }
        public Images images { get; set; }
        public News news { get; set; }
        public RelatedSearches relatedSearches { get; set; }
        public Videos videos { get; set; }
        public RankingResponse rankingResponse { get; set; }
    }

    public class About
    {
        public string _type { get; set; }
        public AggregateRating aggregateRating { get; set; }
        public string readLink { get; set; }
        public string name { get; set; }
    }

    public class AggregateRating
    {
        public int ratingValue { get; set; }
        public int reviewCount { get; set; }
    }

    public class ContractualRule
    {
        public string _type { get; set; }
        public string targetPropertyName { get; set; }
        public int targetPropertyIndex { get; set; }
        public bool mustBeCloseToContent { get; set; }
        public License license { get; set; }
        public string licenseNotice { get; set; }
        public string text { get; set; }
    }

    public class Creator
    {
        public string name { get; set; }
    }

    public class Hint
    {
        public string text { get; set; }
    }

    public class Image
    {
        public string contentUrl { get; set; }
        public Thumbnail thumbnail { get; set; }
    }

    public class Images
    {
        public string id { get; set; }
        public string readLink { get; set; }
        public string webSearchUrl { get; set; }
        public bool isFamilyFriendly { get; set; }
        public List<Value> value { get; set; }
    }

    public class InsightsMetadata
    {
        public int shoppingSourcesCount { get; set; }
        public int recipeSourcesCount { get; set; }
    }

    public class Item
    {
        public string _type { get; set; }
        public string text { get; set; }
        public string url { get; set; }
        public string answerType { get; set; }
        public Value value { get; set; }
        public int? resultIndex { get; set; }
    }

    public class Label
    {
        public string text { get; set; }
    }

    public class License
    {
        public string name { get; set; }
        public string url { get; set; }
    }

    public class Mainline
    {
        public List<Item> items { get; set; }
    }

    public class News
    {
        public string id { get; set; }
        public string readLink { get; set; }
        public List<Value> value { get; set; }
    }

    public class PrimaryImageOfPage
    {
        public string thumbnailUrl { get; set; }
        public int width { get; set; }
        public int height { get; set; }
        public string imageId { get; set; }
    }

    public class Provider
    {
        public string _type { get; set; }
        public string name { get; set; }
        public Image image { get; set; }
    }

    public class Publisher
    {
        public string name { get; set; }
    }

    public class QueryContext
    {
        public string originalQuery { get; set; }
    }

    public class RankingResponse
    {
        public Mainline mainline { get; set; }
        public Sidebar sidebar { get; set; }
    }

    public class RelatedSearches
    {
        public string id { get; set; }
        public List<Value> value { get; set; }
    }

    public class RichFact
    {
        public Label label { get; set; }
        public List<Item> items { get; set; }
        public Hint hint { get; set; }
    }    

    public class Sidebar
    {
        public List<Item> items { get; set; }
    }

    public class Thumbnail
    {
        public int width { get; set; }
        public int height { get; set; }
        public string contentUrl { get; set; }
    }

    public class Value
    {
        public string id { get; set; }
        public List<ContractualRule> contractualRules { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public string thumbnailUrl { get; set; }
        public bool isFamilyFriendly { get; set; }
        public string displayUrl { get; set; }
        public string snippet { get; set; }
        public object dateLastCrawled { get; set; }
        public PrimaryImageOfPage primaryImageOfPage { get; set; }
        public string cachedPageUrl { get; set; }
        public string language { get; set; }
        public bool isNavigational { get; set; }
        public List<RichFact> richFacts { get; set; }
        public bool noCache { get; set; }
        public List<About> about { get; set; }
        public DateTime? datePublished { get; set; }
        public string datePublishedDisplayText { get; set; }
        public string webSearchUrl { get; set; }
        public string contentUrl { get; set; }
        public string hostPageUrl { get; set; }
        public string contentSize { get; set; }
        public string encodingFormat { get; set; }
        public string hostPageDisplayUrl { get; set; }
        public int width { get; set; }
        public int height { get; set; }
        public Thumbnail thumbnail { get; set; }
        public InsightsMetadata insightsMetadata { get; set; }
        public Image image { get; set; }
        public string description { get; set; }
        public List<Provider> provider { get; set; }
        public string category { get; set; }
        public string text { get; set; }
        public string displayText { get; set; }
        public List<Publisher> publisher { get; set; }
        public int viewCount { get; set; }
        public bool isSuperfresh { get; set; }
        public Creator creator { get; set; }
        public bool? isAccessibleForFree { get; set; }
        public string duration { get; set; }
        public string motionThumbnailUrl { get; set; }
        public string embedHtml { get; set; }
        public bool? allowHttpsEmbed { get; set; }
        public bool? allowMobileEmbed { get; set; }
    }

    public class Videos
    {
        public string id { get; set; }
        public string readLink { get; set; }
        public string webSearchUrl { get; set; }
        public bool isFamilyFriendly { get; set; }
        public List<Value> value { get; set; }
        public string scenario { get; set; }
    }

    public class WebPages
    {
        public string webSearchUrl { get; set; }
        public int totalEstimatedMatches { get; set; }
        public List<Value> value { get; set; }
    }
}
