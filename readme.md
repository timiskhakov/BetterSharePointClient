# BetterSharePointClient

A small library to retrieve data from SharePoint 2016 lists.

## Usage

```csharp

class Subscription
{
    public int Id { get; set; }
    public DateTime Created { get; set; }
    public string User { get; set; }
    public string SubscriptionType { get; set; }
}

// ...

var url = "http://sharepoint2016";
var credentials = CredentialCache.DefaultNetworkCredentials;
var listName = "Subscriptions";
var queryString = @"<View></View>";

var fields = new[] { "ID", "Created", "CompanyEmployee", "ApplicationSubscription" };
Func<Dictionary<string, object>, Subscription> mapper = fieldValues => new Subscription
{
    Id = (int)fieldValues["ID"],
    Created = (DateTime)fieldValues["Created"],
    User = (fieldValues["CompanyEmployee"] as FieldUserValue)?.LookupValue,
    SubscriptionType = fieldValues["ApplicationSubscription"] as string
};

using (var client = new Client(url, credentials))
{
    List<Subscription> subscriptions;
    try
    {
        subscriptions = client.GetEntities<Subscription>(listName, queryString, fields, mapper);
    }
    catch (WebException ex)
    {
        Console.WriteLine(ex);
    }
    // ...
}
```