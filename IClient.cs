using System;
using System.Collections.Generic;

namespace BetterSharePointClient
{
    public interface IClient
    {
        List<T> GetEntities<T>(
            string listName,
            IEnumerable<string> fields,
            Func<Dictionary<string, object>, T> mapper,
            int threshold = 5000) where T : class;
    }
}
