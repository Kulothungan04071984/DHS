
namespace DHS_Test
{
    partial class Export_Import_Excel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.importExcelFileToolStripMenuItem = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // importExcelFileToolStripMenuItem
            // 
            this.importExcelFileToolStripMenuItem.Location = new System.Drawing.Point(210, 47);
            this.importExcelFileToolStripMenuItem.Name = "importExcelFileToolStripMenuItem";
            this.importExcelFileToolStripMenuItem.Size = new System.Drawing.Size(75, 23);
            this.importExcelFileToolStripMenuItem.TabIndex = 0;
            this.importExcelFileToolStripMenuItem.Text = "Import";
            this.importExcelFileToolStripMenuItem.UseVisualStyleBackColor = true;
            this.importExcelFileToolStripMenuItem.Click += new System.EventHandler(this.importExcelFileToolStripMenuItem_Click);
            // 
            // Export_Import_Excel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.importExcelFileToolStripMenuItem);
            this.Name = "Export_Import_Excel";
            this.Text = "Export_Import_Excel";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button importExcelFileToolStripMenuItem;
    }
}