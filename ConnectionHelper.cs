using System.Security.Cryptography.X509Certificates;

namespace BetterSharePointClient
{
    internal static class ConnectionHelper
    {
        internal static X509Certificate2 GetCertificate(string serialNumber)
        {
            var store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            var collection = store.Certificates.Find(X509FindType.FindBySerialNumber, serialNumber, false);
            store.Close();
            return collection.Count > 0
                ? collection[0]
                : null;
        }
    }
}
