﻿using CamlexNET;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net;

namespace BetterSharePointClient
{
    public class Client : IDisposable
    {
        private const int SharePointThreshold = 5000;
        private readonly ClientContext _clientContext;

        /// <summary>
        /// Creates an instance of Client
        /// </summary>
        /// <param name="baseUrl">SharePoint web full url</param>
        /// <param name="credentials">SharePoint credentials</param>
        /// <param name="certificateSerialNumber">Certificate serial number</param>
        public Client(string baseUrl, NetworkCredential credentials, string certificateSerialNumber = null)
        {
            _clientContext = new ClientContext(baseUrl)
            {
                Credentials = credentials
            };
            if (!string.IsNullOrEmpty(certificateSerialNumber))
            {
                _clientContext.ExecutingWebRequest += (s, e) =>
                {
                    var request = e.WebRequestExecutor.WebRequest;
                    var certificate = ConnectionHelper.GetCertificate(certificateSerialNumber);
                    request.ClientCertificates.Add(certificate);
                };
            };
        }

        #region Public API

        /// <summary>
        /// Returns a list of items represented by a dictionary: field name, field value
        /// </summary>
        /// <param name="listName">List name</param>
        /// <param name="fields">Fields to select</param>
        /// <param name="filter">Field filter</param>
        /// <param name="threshold">SharePoint list threshold</param>
        /// <returns>List of items</returns>
        /// <exception cref="WebException">Occurs when something is wrong with a request to SharePoint</exception>
        public List<Dictionary<string, object>> GetEntities(
            string listName,
            IEnumerable<string> fields,
            Expression<Func<ListItem, bool>> filter = null,
            int threshold = SharePointThreshold)
        {
            var result = new List<Dictionary<string, object>>();

            List list = _clientContext.Web.Lists.GetByTitle(listName);
            var maxId = GetMaxId(list);
            if (maxId == 0)
            {
                return result;
            }

            var includes = fields.Select(f =>
            {
                Expression<Func<ListItem, object>> lambda = li => li[f];
                return lambda;
            }).ToArray();

            var min = 0;
            var max = threshold;
            while (min < maxId)
            {
                var filters = GetFilters(min, max, filter)
                    .ToArray();
                var query = Camlex.Query()
                    .WhereAll(filters)
                    .ToCamlQuery(); ;
                var items = list.GetItems(query);
                _clientContext.Load(items, item => item.Include(includes));
                ExecuteQueryWithCustomErrorMessage($"Error while retrieving items from the list {listName}");

                IEnumerable<Dictionary<string, object>> range = items
                    .AsEnumerable()
                    .Select(li => li.FieldValues);
                result.AddRange(range);

                min += threshold;
                max += threshold;
            }

            return result;
        }

        /// <summary>
        /// Moves an item to a specified folder inside the same list
        /// </summary>
        /// <param name="listName">List title</param>
        /// <param name="itemId">Item ID</param>
        /// <param name="relativeFolderUrl">Relative folder URL</param>
        public void MoveItemToFolder(string listTitle, int itemId, string relativeFolderUrl)
        {
            List list = _clientContext.Web.Lists.GetByTitle(listTitle);
            ListItem item = list.GetItemById(itemId);
            _clientContext.Load(item,
                i => i["ID"],
                i => i["FileLeafRef"],
                i => i["FileDirRef"],
                i => i["FileRef"]);
            ExecuteQueryWithCustomErrorMessage("Error while retrieving item information");

            var file = _clientContext.Web.GetFileByServerRelativeUrl(item["FileRef"].ToString());
            var fileUrl = item["FileRef"].ToString().Replace(item["FileDirRef"].ToString(), relativeFolderUrl);
            file.MoveTo(fileUrl, MoveOperations.Overwrite);
            ExecuteQueryWithCustomErrorMessage($"Error while updating the item {item["ID"]}");
        }

        #endregion

        public void Dispose()
        {
            _clientContext?.Dispose();
        }

        #region Private Methods

        private int GetMaxId(List list)
        {
            var caml = Camlex.Query()
                .OrderBy(i => i["ID"] as Camlex.Desc)
                .Take(1)
                .ToCamlQuery();
            ListItemCollection items = list.GetItems(caml);
            _clientContext.Load(items, li => li.Include(i => i["ID"]));
            ExecuteQueryWithCustomErrorMessage($"Error while retrieving max id from the list");

            if (items.Count == 0)
            {
                return 0;
            }
            var maxId = (int)items.First().FieldValues["ID"];
            return maxId;
        }

        private static IEnumerable<Expression<Func<ListItem, bool>>> GetFilters(int min, int max, Expression<Func<ListItem, bool>> filter)
        {
            yield return li => (int)li["ID"] > min && (int)li["ID"] <= max;
            if (filter != null)
            {
                yield return filter;
            }
        }

        private void ExecuteQueryWithCustomErrorMessage(string errorMessage)
        {
            try
            {
                _clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw new WebException($"{errorMessage}. Exception message: {ex.Message}");
            }
        }

        #endregion

    }
}
