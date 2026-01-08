using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AreaManager.Services;
using AreaManager.UI;

namespace AreaManager
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("\nArea Manager loaded. Type AMUI to open the tool.");
        }

        public void Terminate()
        {
        }

        [CommandMethod("AMUI")]
        public void ShowAreaManagerUi()
        {
            var form = new MainForm();
            Application.ShowModelessDialog(form);
        }

        [CommandMethod("AMTEMP")]
        public void GenerateTemporaryAreasTable()
        {
            GenerationService.GenerateTemporaryAreasTable();
        }

        [CommandMethod("AMWORK")]
        public void GenerateWorkspaceAreasTable()
        {
            GenerationService.GenerateWorkspaceAreasTable();
        }

        [CommandMethod("AMODSHAPES")]
        public void AddObjectDataToShapes()
        {
            WorkspaceObjectDataService.AddObjectDataToShapes();
        }
    }
}
