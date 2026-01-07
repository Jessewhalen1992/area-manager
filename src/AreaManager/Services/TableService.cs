using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AreaManager.Models;

namespace AreaManager.Services
{
    public static class TableService
    {
        public static Table CreateTempAreaTable(Point3d insertionPoint, IEnumerable<TempAreaRow> rows)
        {
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            var rowList = new List<TempAreaRow>(rows);
            table.SetSize(rowList.Count + 2, 8);
            table.SetRowHeight(2.5);
            table.SetColumnWidth(12.0);

            table.Cells[0, 0].TextString = "TEMPORARY AREA(S) INFORMATION";
            table.MergeCells(CellRange.Create(table, 0, 0, 0, 7));

            var headers = new[]
            {
                "DESCRIPTION",
                "ID",
                "WIDTH",
                "LENGTH",
                "AREA (ha)",
                "WITHIN EXISTING DISPOSITIONS",
                "EXISTING CUT DISTURBANCE (ha)",
                "NEW CUT DISTURBANCE (ha)"
            };

            for (var col = 0; col < headers.Length; col++)
            {
                table.Cells[1, col].TextString = headers[col];
            }

            var rowIndex = 2;
            foreach (var row in rowList)
            {
                table.Cells[rowIndex, 0].TextString = row.Description ?? string.Empty;
                table.Cells[rowIndex, 1].TextString = row.Identifier ?? string.Empty;
                table.Cells[rowIndex, 2].TextString = row.Width ?? string.Empty;
                table.Cells[rowIndex, 3].TextString = row.Length ?? string.Empty;
                table.Cells[rowIndex, 4].TextString = row.AreaHa ?? string.Empty;
                table.Cells[rowIndex, 5].TextString = row.WithinExistingDisposition ?? string.Empty;
                table.Cells[rowIndex, 6].TextString = row.ExistingCutDisturbance ?? string.Empty;
                table.Cells[rowIndex, 7].TextString = row.NewCutDisturbance ?? string.Empty;
                rowIndex++;
            }

            return table;
        }

        public static Table CreateWorkspaceTotalsTable(Point3d insertionPoint, IEnumerable<WorkspaceAreaRow> rows)
        {
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            var rowList = new List<WorkspaceAreaRow>(rows);
            table.SetSize(rowList.Count, 9);

            var rowIndex = 0;
            foreach (var row in rowList)
            {
                var withinHa = row.ExistingDispositionHa;
                var withinAc = ConvertHaToAc(withinHa);
                var outsideHa = row.TotalHa - withinHa;
                var outsideAc = ConvertHaToAc(outsideHa);
                var totalAc = ConvertHaToAc(row.TotalHa);

                table.Cells[rowIndex, 0].TextString = row.WorkspaceId ?? string.Empty;
                table.Cells[rowIndex, 1].TextString = withinHa.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 2].TextString = withinAc.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 3].TextString = outsideHa.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 4].TextString = outsideAc.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 5].TextString = row.TotalHa.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 6].TextString = totalAc.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 7].TextString = row.ExistingCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);
                table.Cells[rowIndex, 8].TextString = row.NewCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);
                rowIndex++;
            }

            return table;
        }

        private static double ConvertHaToAc(double hectares)
        {
            return hectares / 0.4047;
        }
    }
}
