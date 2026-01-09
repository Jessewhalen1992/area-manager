namespace AreaManager.UI
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button tempAreasButton;
        private System.Windows.Forms.Button workspaceAreasButton;
        private System.Windows.Forms.Button addOdToShapesButton;
        private System.Windows.Forms.Button addRtfInfoButton;
        private System.Windows.Forms.Label headerLabel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tempAreasButton = new System.Windows.Forms.Button();
            this.workspaceAreasButton = new System.Windows.Forms.Button();
            this.addOdToShapesButton = new System.Windows.Forms.Button();
            this.addRtfInfoButton = new System.Windows.Forms.Button();
            this.headerLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tempAreasButton
            // 
            this.tempAreasButton.Location = new System.Drawing.Point(24, 58);
            this.tempAreasButton.Name = "tempAreasButton";
            this.tempAreasButton.Size = new System.Drawing.Size(234, 35);
            this.tempAreasButton.TabIndex = 0;
            this.tempAreasButton.Text = "Generate Temporary Areas Table";
            this.tempAreasButton.UseVisualStyleBackColor = true;
            this.tempAreasButton.Click += new System.EventHandler(this.tempAreasButton_Click);
            // 
            // workspaceAreasButton
            // 
            this.workspaceAreasButton.Location = new System.Drawing.Point(24, 109);
            this.workspaceAreasButton.Name = "workspaceAreasButton";
            this.workspaceAreasButton.Size = new System.Drawing.Size(234, 35);
            this.workspaceAreasButton.TabIndex = 1;
            this.workspaceAreasButton.Text = "Generate Crown Area Usage Table";
            this.workspaceAreasButton.UseVisualStyleBackColor = true;
            this.workspaceAreasButton.Click += new System.EventHandler(this.workspaceAreasButton_Click);
            // 
            // addOdToShapesButton
            // 
            this.addOdToShapesButton.Location = new System.Drawing.Point(24, 160);
            this.addOdToShapesButton.Name = "addOdToShapesButton";
            this.addOdToShapesButton.Size = new System.Drawing.Size(234, 35);
            this.addOdToShapesButton.TabIndex = 2;
            this.addOdToShapesButton.Text = "Add OD To Shapes";
            this.addOdToShapesButton.UseVisualStyleBackColor = true;
            this.addOdToShapesButton.Click += new System.EventHandler(this.addOdToShapesButton_Click);
            // 
            // addRtfInfoButton
            // 
            this.addRtfInfoButton.Location = new System.Drawing.Point(24, 211);
            this.addRtfInfoButton.Name = "addRtfInfoButton";
            this.addRtfInfoButton.Size = new System.Drawing.Size(234, 35);
            this.addRtfInfoButton.TabIndex = 3;
            this.addRtfInfoButton.Text = "Add RTF Info";
            this.addRtfInfoButton.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.addRtfInfoButton.UseVisualStyleBackColor = true;
            this.addRtfInfoButton.Click += new System.EventHandler(this.addRtfInfoButton_Click);
            // 
            // headerLabel
            // 
            this.headerLabel.AutoSize = true;
            this.headerLabel.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.headerLabel.Location = new System.Drawing.Point(20, 20);
            this.headerLabel.Name = "headerLabel";
            this.headerLabel.Size = new System.Drawing.Size(139, 19);
            this.headerLabel.TabIndex = 2;
            this.headerLabel.Text = "Area Manager Tools";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.headerLabel);
            this.Controls.Add(this.addRtfInfoButton);
            this.Controls.Add(this.addOdToShapesButton);
            this.Controls.Add(this.workspaceAreasButton);
            this.Controls.Add(this.tempAreasButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Area Manager";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
