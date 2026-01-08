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
            // Create a table without built‑in headings.  The consumer of this service can
            // overlay their own title and column headers later.  We set up only the data
            // rows here so that the user has full control over the final appearance.
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            // Convert the enumerable to a list for efficient indexing and counting
            var rowList = new List<TempAreaRow>(rows);

            // There are always eight columns: description, id, width, length, area, within dispositions,
            // existing cut disturbance, and new cut disturbance.  The row count matches the number
            // of data entries; no extra rows are created for headings.
            table.SetSize(rowList.Count, 8);

            // Configure the row height and individual column widths to mirror the provided sample
            // (values in drawing units).  This yields a table with wide description and equal
            // spacing for the other numeric columns.
            table.SetRowHeight(25.0);
            // Column widths are based on the provided sample.  The description column
            // should be narrower (162 units) and the identifier column ("#") much narrower
            // (40 units) than our previous implementation.  The area column is only 50
            // units wide, while the last three disturbance columns remain wide (120 units).
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

                // Determine if this row represents an irregular area.  If so, merge the width
                // and length columns into a single cell and place the text once.  Otherwise
                // populate both cells separately.
                bool isIrregular = string.Equals(row.Width, "IRREGULAR", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(row.Length, "IRREGULAR", StringComparison.OrdinalIgnoreCase);
                if (isIrregular)
                {
                    // Merge the cells spanning columns 2 and 3 for this row
                    table.MergeCells(CellRange.Create(table, rowIndex, 2, rowIndex, 3));
                    table.Cells[rowIndex, 2].TextString = row.Width ?? string.Empty;
                }
                else
                {
                    table.Cells[rowIndex, 2].TextString = row.Width ?? string.Empty;
                    table.Cells[rowIndex, 3].TextString = row.Length ?? string.Empty;
                }

                // Fill in the remaining columns
                table.Cells[rowIndex, 4].TextString = row.AreaHa ?? string.Empty;
                table.Cells[rowIndex, 5].TextString = row.WithinExistingDisposition ?? string.Empty;
                table.Cells[rowIndex, 6].TextString = row.ExistingCutDisturbance ?? string.Empty;
                table.Cells[rowIndex, 7].TextString = row.NewCutDisturbance ?? string.Empty;

                // Center the text horizontally and vertically for every cell in the row and set
                // the text height to 10.0.  Using MiddleCenter ensures a consistent look.
                for (int col = 0; col < 8; col++)
                {
                    var cell = table.Cells[rowIndex, col];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = 10.0;
                }

                rowIndex++;
            }

            // Generate the layout to apply column widths, merges and alignment settings before
            // returning the table to the caller.
            table.GenerateLayout();

            return table;
        }

        public static Table CreateWorkspaceTotalsTable(Point3d insertionPoint, IEnumerable<WorkspaceAreaRow> rows)
        {
            // Create a table for crown area usage without any header rows.  This method mirrors
            // the example layout provided by setting explicit sizes for rows and columns and
            // applying consistent alignment and text height across all cells.
            var table = new Table
            {
                TableStyle = ObjectId.Null,
                Position = insertionPoint
            };

            var rowList = new List<WorkspaceAreaRow>(rows);
            // There are nine columns: ID/description, within ha, within ac, outside ha,
            // outside ac, total ha, total ac, existing cut disturbance ha, new cut disturbance ha.
            table.SetSize(rowList.Count, 9);

            // Set a uniform row height and assign widths to each column based on the sample
            table.SetRowHeight(25.0);
            double[] columnWidths2 = { 132.0, 60.0, 60.0, 60.0, 60.0, 50.0, 50.0, 120.0, 120.0 };
            for (int col = 0; col < columnWidths2.Length; col++)
            {
                table.Columns[col].Width = columnWidths2[col];
            }

            // Define colours for shading: use a very light colour (e.g. index 254) for
            // the grand total row and no shading for per‑activity totals.  Colour 254
            // corresponds to a light grey in most AutoCAD colour tables.  By using
            // FromColorIndex, we avoid specifying RGB values that might look different
            // under various colour schemes.
            var summaryColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 254);

            int i = 0;
            foreach (var row in rowList)
            {
                bool isTotalsRow = string.IsNullOrEmpty(row.WorkspaceId);

                if (isTotalsRow)
                {
                    // For the grand total row, leave the description and intermediate
                    // columns blank.  Populate only the total ha and total ac cells with
                    // the aggregated values.  Colour these cells using the summary colour.
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
                }
                else
                {
                    // Compute the derived area usage values.  Conversions to acres (Ac.) are
                    // performed here so they can be inserted directly into the table.  These
                    // represent existing disposition, outside existing dispositions and totals.
                    double withinHa = row.ExistingDispositionHa;
                    double withinAc = ConvertHaToAc(withinHa);
                    double outsideHa = row.TotalHa - withinHa;
                    double outsideAc = ConvertHaToAc(outsideHa);
                    double totalAc = ConvertHaToAc(row.TotalHa);

                    // Populate each column of the row
                    table.Cells[i, 0].TextString = row.WorkspaceId ?? string.Empty;
                    table.Cells[i, 1].TextString = withinHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 2].TextString = withinAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 3].TextString = outsideHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 4].TextString = outsideAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 5].TextString = row.TotalHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 6].TextString = totalAc.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 7].TextString = row.ExistingCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);
                    table.Cells[i, 8].TextString = row.NewCutDisturbanceHa.ToString("0.000", CultureInfo.InvariantCulture);
                    // Do not set a background colour for per‑activity totals; these cells
                    // remain white (default) so the user can identify them as calculated
                    // values without shading.  Note that formulas are not inserted via
                    // API: the numeric values are calculated directly.
                }

                // Apply centre alignment and set the text height for each cell in the row
                for (int col = 0; col < 9; col++)
                {
                    var cell = table.Cells[i, col];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = 10.0;
                }

                i++;
            }

            // Finalize the table's layout before returning
            table.GenerateLayout();
            return table;
        }

        private static double ConvertHaToAc(double hectares)
        {
            return hectares / 0.4047;
        }
    }
}