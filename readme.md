# BetterSharePointClient

A small library to manipulate SharePoint 2016 data.

## Usage

Setup:
```csharp
var url = "http://sharepoint2016";
var credentials = CredentialCache.DefaultNetworkCredentials;
var listName = "Subscriptions";
```

### GetEntities

```csharp
var fields = new[] { "Created", "CompanyEmployee", "ApplicationSubscription" };
Expression<Func<ListItem, bool>> filter = li => (int) li["ID"] <= 100;
using (var client = new Client(url, credentials))
{
    List<Dictionary<string, object>> subscriptions = client.GetEntities(listName, fields, filter);
}
```

### MoveItemToFolder

```csharp
using (var client = new Client(url, credentials))
{
    client.MoveItemToFolder(listName, 5, $"Lists/{listName}/SomeFolder");
}
```
