using System;
using System.Windows.Forms;
using AreaManager.Services;

namespace AreaManager.UI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void tempAreasButton_Click(object sender, EventArgs e)
        {
            GenerationService.GenerateTemporaryAreasTable();
        }

        private void workspaceAreasButton_Click(object sender, EventArgs e)
        {
            GenerationService.GenerateWorkspaceAreasTable();
        }
    }
}
