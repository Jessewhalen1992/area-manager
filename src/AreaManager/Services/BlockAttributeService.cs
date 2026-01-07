using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace AreaManager.Services
{
    public static class BlockAttributeService
    {
        private static readonly HashSet<string> SupportedBlockNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "Dyn_temp_area",
            "TempArea-Blue",
            "TempArea_County",
            "Work_Area_Stretchy",
            "Temp_Area_Pink"
        };

        public static List<Tuple<string, string>> GetUniqueAttributePairs(Editor editor, Database database)
        {
            var result = new List<Tuple<string, string>>();
            var uniquePairs = new HashSet<string>(StringComparer.Ordinal);

            var selection = editor.GetSelection(new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT")
            }));

            if (selection.Status != PromptStatus.OK)
            {
                return result;
            }

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selected in selection.Value)
                {
                    if (selected == null)
                    {
                        continue;
                    }

                    var blockReference = transaction.GetObject(selected.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blockReference == null)
                    {
                        continue;
                    }

                    var blockName = blockReference.EffectiveName;
                    if (!SupportedBlockNames.Contains(blockName))
                    {
                        continue;
                    }

                    var tempArea = string.Empty;
                    var enterText = string.Empty;

                    foreach (ObjectId attributeId in blockReference.AttributeCollection)
                    {
                        var attribute = transaction.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;
                        if (attribute == null)
                        {
                            continue;
                        }

                        if (attribute.Tag.Equals("TEMP_AREA_W1", StringComparison.Ordinal))
                        {
                            tempArea = attribute.TextString;
                        }
                        else if (attribute.Tag.Equals("ENTER_TEXT", StringComparison.Ordinal))
                        {
                            enterText = attribute.TextString;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(tempArea) || string.IsNullOrWhiteSpace(enterText))
                    {
                        continue;
                    }

                    var key = $"{tempArea}|{enterText}";
                    if (uniquePairs.Add(key))
                    {
                        result.Add(Tuple.Create(tempArea, enterText));
                    }
                }

                transaction.Commit();
            }

            return result;
        }
    }
}
