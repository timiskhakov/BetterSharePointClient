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
        private readonly ClientContext _clientContext;

        /// <summary>
        /// Creates an instance of Client
        /// </summary>
        /// <param name="baseUrl">SharePoint web full url</param>
        /// <param name="credentials">SharePoint credentials</param>
        /// <param name="certificateSerialNumber">Certificate serial number</param>
        public Client(string baseUrl, NetworkCredential credentials)
        {
            _clientContext = new ClientContext(baseUrl)
            {
                Credentials = credentials
            };
        }

        #region Public API

        /// <summary>
        /// Returns a list of items represented by a dictionary: field name, field value
        /// </summary>
        /// <param name="listName">List name</param>
        /// <param name="fields">Fields to select</param>
        /// <param name="threshold">SharePoint list threshold</param>
        /// <returns>List of items</returns>
        /// <exception cref="WebException">Occurs when something is wrong with a request to SharePoint</exception>
        public List<Dictionary<string, object>> GetEntities(
            string listName,
            IEnumerable<string> fields,
            int threshold = 5000)
        {
            var result = new List<Dictionary<string, object>>();

            List list = _clientContext.Web.Lists.GetByTitle(listName);
            _clientContext.Load(list, l => l.ItemCount);
            ExecuteQueryWithCustomErrorMessage($"Error while getting information about the list {listName}");

            var includes = fields.Select(f =>
            {
                Expression<Func<ListItem, object>> lambda = li => li[f];
                return lambda;
            }).ToArray();

            var min = 0;
            var max = threshold;
            while (min < list.ItemCount)
            {
                var query = Camlex.Query()
                    .Where(li =>
                        (int)li["ID"] >= min &&
                        (int)li["ID"] < max)
                    .ToCamlQuery(); ;
                var items = list.GetItems(query);
                _clientContext.Load(items, item => item.Include(includes));
                ExecuteQueryWithCustomErrorMessage($"Error while getting items from the list {listName}");

                min += threshold;
                max += threshold;
                IEnumerable<Dictionary<string, object>> range = items
                    .AsEnumerable()
                    .Select(li => li.FieldValues);
                result.AddRange(range);
            }

            return result;
        }

        #endregion

        public void Dispose()
        {
            _clientContext?.Dispose();
        }

        #region Private Methods

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
