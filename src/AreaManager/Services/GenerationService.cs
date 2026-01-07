using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AreaManager.Models;

namespace AreaManager.Services
{
    public static class GenerationService
    {
        public static void GenerateTemporaryAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var pairs = BlockAttributeService.GetUniqueAttributePairs(editor, database);
            if (pairs.Count == 0)
            {
                editor.WriteMessage("\nNo matching blocks found.");
                return;
            }

            var rows = BuildTempAreaRows(pairs);
            InsertTable(database, editor, rows);
        }

        public static void GenerateWorkspaceAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var tableSelection = PromptForWorkspaceTable(editor);
            if (tableSelection == ObjectId.Null)
            {
                return;
            }

            var workspaceRows = WorkspaceAreaService.CalculateWorkspaceAreasFromTable(editor, database, tableSelection);
            if (workspaceRows.Count == 0)
            {
                editor.WriteMessage("\nNo workspace rows found in the selected table.");
                return;
            }

            InsertWorkspaceTable(database, editor, workspaceRows);
        }

        private static List<TempAreaRow> BuildTempAreaRows(IEnumerable<Tuple<string, string>> pairs)
        {
            var rows = new List<TempAreaRow>();

            foreach (var pair in pairs)
            {
                var identifier = pair.Item1?.Trim() ?? string.Empty;
                var enterText = pair.Item2?.Trim() ?? string.Empty;

                var description = MapIdentifierDescription(identifier);
                var row = BuildTempAreaRow(description, identifier, enterText);
                rows.Add(row);
            }

            return rows
                .OrderBy(row => ParseIdentifierPrefix(row.Identifier))
                .ThenBy(row => ParseIdentifierNumber(row.Identifier))
                .ThenBy(row => row.Identifier, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static TempAreaRow BuildTempAreaRow(string description, string identifier, string enterText)
        {
            var row = new TempAreaRow
            {
                Description = description,
                Identifier = identifier,
                WithinExistingDisposition = "No"
            };

            if (enterText.IndexOf("/P=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                enterText.IndexOf("\\P=", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                row.Width = "IRREGULAR";
                row.Length = "IRREGULAR";
                row.AreaHa = ExtractAreaFromIrregular(enterText);
                row.NewCutDisturbance = "0.000";
                row.ExistingCutDisturbance = row.AreaHa;
                return row;
            }

            if (enterText.Contains("x"))
            {
                var parts = enterText.Split('x');
                if (parts.Length >= 2)
                {
                    var part1 = TextParsingService.ExtractLastNumber(parts[0]);
                    var part2 = TextParsingService.ExtractLastNumber(parts[1]);

                    if (IsNumeric(part1) && IsNumeric(part2))
                    {
                        var width = TextParsingService.ParseDoubleOrDefault(part1);
                        var length = TextParsingService.ParseDoubleOrDefault(part2);
                        var area = (width * length) / 10000.0;

                        row.Width = width.ToString("0.0");
                        row.Length = length.ToString("0.0");
                        row.AreaHa = area.ToString("0.000");
                        row.ExistingCutDisturbance = row.AreaHa;
                        row.NewCutDisturbance = "0.000";
                        return row;
                    }
                }
            }

            row.Width = "IRREGULAR";
            row.Length = "IRREGULAR";
            row.AreaHa = TextParsingService.ExtractAllNumbers(enterText);
            row.ExistingCutDisturbance = row.AreaHa;
            row.NewCutDisturbance = "0.000";
            return row;
        }

        private static string ExtractAreaFromIrregular(string enterText)
        {
            var index = enterText.IndexOf("/P=", StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                index = enterText.IndexOf("\\P=", StringComparison.OrdinalIgnoreCase);
            }

            var areaText = index >= 0 ? enterText.Substring(index + 3) : enterText;
            var number = TextParsingService.ExtractLastNumber(areaText);
            if (IsNumeric(number))
            {
                var area = TextParsingService.ParseDoubleOrDefault(number);
                return area.ToString("0.000");
            }

            return "N/A";
        }

        private static bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        private static string MapIdentifierDescription(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return string.Empty;
            }

            var upper = identifier.ToUpperInvariant();
            if (upper.StartsWith("W") && upper.Length > 1 && char.IsDigit(upper[1]))
            {
                return "WORKSPACE";
            }

            if (upper.StartsWith("LD"))
            {
                return "LOG DECK";
            }

            if (upper.StartsWith("AR"))
            {
                return "ACCESS ROAD";
            }

            if (upper.StartsWith("BP"))
            {
                return "BORROW PIT";
            }

            if (upper.StartsWith("BS"))
            {
                return "BANK STABILIZATION";
            }

            return string.Empty;
        }

        private static void InsertTable(Database database, Editor editor, IEnumerable<TempAreaRow> rows)
        {
            var promptPoint = editor.GetPoint("\nPick insertion point for Temporary Areas table: ");
            if (promptPoint.Status != PromptStatus.OK)
            {
                return;
            }

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = TableService.CreateTempAreaTable(promptPoint.Value, rows);
                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                modelSpace.AppendEntity(table);
                transaction.AddNewlyCreatedDBObject(table, true);
                transaction.Commit();
            }
        }

        private static void InsertWorkspaceTable(Database database, Editor editor, IEnumerable<WorkspaceAreaRow> rows)
        {
            var promptPoint = editor.GetPoint("\nPick insertion point for Crown Area Usage table: ");
            if (promptPoint.Status != PromptStatus.OK)
            {
                return;
            }

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = TableService.CreateWorkspaceTotalsTable(promptPoint.Value, rows);
                ApplyTableStyle(transaction, database, editor, table, "Induction Bend");
                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                modelSpace.AppendEntity(table);
                transaction.AddNewlyCreatedDBObject(table, true);
                transaction.Commit();
            }
        }

        private static ObjectId PromptForWorkspaceTable(Editor editor)
        {
            var options = new PromptEntityOptions("\nSelect workspace source table: ")
            {
                AllowNone = false
            };
            options.SetRejectMessage("\nOnly tables are allowed.");
            options.AddAllowedClass(typeof(Table), true);

            var result = editor.GetEntity(options);
            return result.Status == PromptStatus.OK ? result.ObjectId : ObjectId.Null;
        }

        private static void ApplyTableStyle(Transaction transaction, Database database, Editor editor, Table table, string styleName)
        {
            var dictionary = (DBDictionary)transaction.GetObject(database.TableStyleDictionaryId, OpenMode.ForRead);
            if (!dictionary.Contains(styleName))
            {
                editor.WriteMessage($"\nWarning: Table style '{styleName}' was not found.");
                return;
            }

            table.TableStyle = dictionary.GetAt(styleName);
        }

        private static string ParseIdentifierPrefix(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return string.Empty;
            }

            var trimmed = identifier.Trim();
            var index = 0;
            while (index < trimmed.Length && char.IsLetter(trimmed[index]))
            {
                index++;
            }

            return trimmed.Substring(0, index).ToUpperInvariant();
        }

        private static int ParseIdentifierNumber(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return int.MaxValue;
            }

            var trimmed = identifier.Trim();
            var index = 0;
            while (index < trimmed.Length && char.IsLetter(trimmed[index]))
            {
                index++;
            }

            var numberStart = index;
            while (index < trimmed.Length && char.IsDigit(trimmed[index]))
            {
                index++;
            }

            if (numberStart == index)
            {
                return int.MaxValue;
            }

            if (int.TryParse(trimmed.Substring(numberStart, index - numberStart), out var number))
            {
                return number;
            }

            return int.MaxValue;
        }
    }
}
