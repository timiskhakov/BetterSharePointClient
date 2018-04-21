# BetterSharePointClient

A small library to retrieve data from SharePoint 2016 lists.

## Usage

```csharp
var url = "http://sharepoint2016";
var credentials = CredentialCache.DefaultNetworkCredentials;
var listName = "Subscriptions";

var fields = new[] { "Created", "CompanyEmployee", "ApplicationSubscription" };
Expression<Func<ListItem, bool>> filter = li => (int) li["ID"] <= 100;

using (var client = new Client(url, credentials))
{
    List<Dictionary<string, object>> subscriptions;
    try
    {
        subscriptions = client.GetEntities(listName, fields, filter);
    }
    catch (WebException ex)
    {
        Console.WriteLine(ex);
    }
    // ...
}
```