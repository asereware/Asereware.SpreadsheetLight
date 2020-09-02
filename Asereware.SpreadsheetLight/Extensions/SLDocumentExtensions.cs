// ***********************************************************************
// Assembly         : Metis.Applications.Import
// Author           : asereware
// Created          : 10-10-2018
//
// Last Modified By : asereware
// Last Modified On : 10-10-2018
// ***********************************************************************
// <copyright file="SLDocumentExtensions.cs" company="Asereware">
//     Copyright © Asereware 2014
// </copyright>
// <summary></summary>
// ***********************************************************************
using SpreadsheetLight;
using System;
using System.Data;

namespace SpreadsheetLight.Extensions
{
    /// <summary>
    /// Class SpreadSheetLightExtensions.
    /// </summary>
    public static class SLDocumentExtensions
    {
        /// <summary>
        /// Creates the data table.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="hasHeaders">if set to <c>true</c> [has headers].</param>
        /// <returns>DataTable.</returns>
        public static DataTable CreateDataTable(this SLDocument sheet, bool hasHeaders = true)
        {
            DataTable dt = new DataTable(sheet.GetCurrentWorksheetName());
            var stats = sheet.GetWorksheetStatistics();
            var columnsCount = stats.NumberOfColumns;
            var rowsCount = stats.NumberOfRows;
            string colName;
            for (int i = 0; i < columnsCount; i++)
            {
                colName = null;
                var index = i + 1;
                if (hasHeaders)
                {
                    colName = sheet.GetCellValueAsString(1, i + 1);
                    if (!String.IsNullOrWhiteSpace(colName) && dt.Columns.Contains(colName))
                    {
                        //Set unknown alias due to the column will be ignored.
                        //This solution avoids alternative that is to count the similar names to increment a counter.
                        colName = String.Concat(colName, "-", Guid.NewGuid().ToString("N"));
                    }
                }

                if(String.IsNullOrWhiteSpace(colName))                
                {
                    colName = $"Col{index.ToString().PadLeft(4, '0')}";
                }

                dt.Columns.Add(new DataColumn(colName));
            }

            for (int i = 1; i <= rowsCount; i++)
            {
                if (hasHeaders && i == 1)
                    continue;

                DataRow dr = dt.NewRow();
                for (int j = 1; j <= columnsCount; j++)
                {
                    dr[j - 1] = sheet.GetCellValueAsString(i, j);
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        /// <summary>
        /// Creates the data view.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="hasHeaders">if set to <c>true</c> [has headers].</param>
        /// <returns>DataView.</returns>
        public static DataView CreateDataView(this SLDocument sheet, bool hasHeaders = true)
        {
            var dt = CreateDataTable(sheet, hasHeaders);
            DataView dv = new DataView(dt);
            return dv;
        }
    }
}
