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
            // Create a table without built‑in headings. The consumer of this service can
            // overlay their own title and column headers later. We set up only the data
            // rows here so that the user has full control over the final appearance.
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            // Same reasoning as CreateTempAreaTable: keep AutoCAD from treating row 0/1 as
            // title/header rows when styles are applied later.
            table.IsTitleSuppressed = true;
            table.IsHeaderSuppressed = true;

            // IMPORTANT (2015/2025): Suppress the built-in title/header rows BEFORE calling
            // SetSize() / touching Cells[]. If we size/fill first, AutoCAD can treat row 0/1
            // as Title/Header (depending on style), which can lead to unexpected merges and
            // even fatal errors when styles are applied later.
            table.IsTitleSuppressed = true;
            table.IsHeaderSuppressed = true;

            // Convert the enumerable to a list for efficient indexing and counting
            var rowList = new List<TempAreaRow>(rows ?? Array.Empty<TempAreaRow>());

            // There are always eight columns: description, id, width, length, area, within dispositions,
            // existing cut disturbance, and new cut disturbance.
            table.SetSize(rowList.Count, 8);

            // Configure the row height and individual column widths.
            table.SetRowHeight(25.0);

            double[] columnWidths = { 162.0, 40.0, 50.0, 50.0, 50.0, 120.0, 120.0, 120.0 };
            for (int col = 0; col < columnWidths.Length; col++)
            {
                table.Columns[col].Width = columnWidths[col];
            }

            int rowIndex = 0;
            foreach (var row in rowList)
            {
                // Populate the description and identifier
                table.Cells[rowIndex, 0].TextString = row.Description ?? string.Empty;
                table.Cells[rowIndex, 1].TextString = row.Identifier ?? string.Empty;

                // Determine irregular
                bool isIrregular =
                    string.Equals(row.Width, "IRREGULAR", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(row.Length, "IRREGULAR", StringComparison.OrdinalIgnoreCase);

                // Fill width/length cells first (formatting happens before merge to avoid touching covered cells)
                if (isIrregular)
                {
                    table.Cells[rowIndex, 2].TextString = row.Width ?? string.Empty;
                    table.Cells[rowIndex, 3].TextString = string.Empty; // covered after merge
                }
                else
                {
                    table.Cells[rowIndex, 2].TextString = row.Width ?? string.Empty;
                    table.Cells[rowIndex, 3].TextString = row.Length ?? string.Empty;
                }

                // Fill remaining columns
                table.Cells[rowIndex, 4].TextString = row.AreaHa ?? string.Empty;
                table.Cells[rowIndex, 5].TextString = row.WithinExistingDisposition ?? string.Empty;
                table.Cells[rowIndex, 6].TextString = row.ExistingCutDisturbance ?? string.Empty;
                table.Cells[rowIndex, 7].TextString = row.NewCutDisturbance ?? string.Empty;

                // Apply formatting BEFORE merging (safer across versions)
                for (int col = 0; col < 8; col++)
                {
                    var cell = table.Cells[rowIndex, col];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = 10.0;
                }

                // Merge after formatting
                if (isIrregular)
                {
                    table.MergeCells(CellRange.Create(table, rowIndex, 2, rowIndex, 3));
                }

                rowIndex++;
            }

            // IMPORTANT:
            // Do NOT call GenerateLayout() here. The caller will apply the final TableStyle
            // (and title/header suppression flags) and then call GenerateLayout() once
            // right before insertion into the database.
            return table;
        }

        public static Table CreateWorkspaceTotalsTable(Point3d insertionPoint, IEnumerable<WorkspaceAreaRow> rows)
        {
            // Create a table without header rows. The caller applies style and generates layout.
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            // Same rationale as above: suppress title/header before sizing/filling.
            table.IsTitleSuppressed = true;
            table.IsHeaderSuppressed = true;

            var rowList = new List<WorkspaceAreaRow>(rows ?? Array.Empty<WorkspaceAreaRow>());

            // There are nine columns.
            table.SetSize(rowList.Count, 9);

            // Set a uniform row height and assign widths to each column
            table.SetRowHeight(25.0);
            double[] columnWidths2 = { 132.0, 60.0, 60.0, 60.0, 60.0, 50.0, 50.0, 120.0, 120.0 };
            for (int col = 0; col < columnWidths2.Length; col++)
            {
                table.Columns[col].Width = columnWidths2[col];
            }

            var summaryColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 254);

            int i = 0;
            foreach (var row in rowList)
            {
                bool isTotalsRow = string.IsNullOrEmpty(row.WorkspaceId);

                if (isTotalsRow)
                {
                    table.Cells[i, 0].TextString = string.Empty;
                    for (int col = 1; col <= 4; col++)
                    {
                        table.Cells[i, col].TextString = string.Empty;
                    }

                    table.Cells[i, 5].TextString = row.TotalHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 6].TextString = ConvertHaToAc(row.TotalHa).ToString("0.000", CultureInfo.InvariantCulture);

                    table.Cells[i, 5].BackgroundColor = summaryColor;
                    table.Cells[i, 6].BackgroundColor = summaryColor;

                    table.Cells[i, 7].TextString = string.Empty;
                    table.Cells[i, 8].TextString = string.Empty;

                    // Format BEFORE merge
                    for (int col = 0; col < 9; col++)
                    {
                        var cell = table.Cells[i, col];
                        cell.Alignment = CellAlignment.MiddleCenter;
                        cell.TextHeight = 10.0;
                    }

                    // Merge cells A–E and H–I for the summary row
                    table.MergeCells(CellRange.Create(table, i, 0, i, 4));
                    table.MergeCells(CellRange.Create(table, i, 7, i, 8));
                }
                else
                {
                    double withinHa = row.ExistingDispositionHa;
                    double withinAc = ConvertHaToAc(withinHa);
                    double outsideHa = row.TotalHa - withinHa;
                    double outsideAc = ConvertHaToAc(outsideHa);
                    double totalAc = ConvertHaToAc(row.TotalHa);

                    table.Cells[i, 0].TextString = row.WorkspaceId ?? string.Empty;
                    table.Cells[i, 1].TextString = withinHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 2].TextString = withinAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 3].TextString = outsideHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 4].TextString = outsideAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 5].TextString = row.TotalHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 6].TextString = totalAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 7].TextString = row.ExistingCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 8].TextString = row.NewCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);

                    for (int col = 0; col < 9; col++)
                    {
                        var cell = table.Cells[i, col];
                        cell.Alignment = CellAlignment.MiddleCenter;
                        cell.TextHeight = 10.0;
                    }
                }

                i++;
            }

            // IMPORTANT:
            // Do NOT call GenerateLayout() here; the caller will do that after applying the final style.
            return table;
        }

        private static double ConvertHaToAc(double hectares)
        {
            return hectares / 0.4047;
        }
    }
}
