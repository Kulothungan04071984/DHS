
namespace DHS_Test
{
    partial class Import
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
            FileUpload = new System.Windows.Forms.Button();
            dtView = new System.Windows.Forms.DataGridView();
            txtFileName = new System.Windows.Forms.TextBox();
            cmbSheet = new System.Windows.Forms.ComboBox();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            btnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)dtView).BeginInit();
            SuspendLayout();
            // 
            // FileUpload
            // 
            FileUpload.Location = new System.Drawing.Point(722, 386);
            FileUpload.Name = "FileUpload";
            FileUpload.Size = new System.Drawing.Size(46, 23);
            FileUpload.TabIndex = 0;
            FileUpload.Text = "....";
            FileUpload.UseVisualStyleBackColor = true;
            FileUpload.Click += FileUpload_Click;
            // 
            // dtView
            // 
            dtView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dtView.Location = new System.Drawing.Point(13, 23);
            dtView.Name = "dtView";
            dtView.RowTemplate.Height = 25;
            dtView.Size = new System.Drawing.Size(775, 359);
            dtView.TabIndex = 1;
            // 
            // txtFileName
            // 
            txtFileName.Location = new System.Drawing.Point(116, 387);
            txtFileName.Name = "txtFileName";
            txtFileName.Size = new System.Drawing.Size(591, 23);
            txtFileName.TabIndex = 2;
            // 
            // cmbSheet
            // 
            cmbSheet.FormattingEnabled = true;
            cmbSheet.Location = new System.Drawing.Point(116, 416);
            cmbSheet.Name = "cmbSheet";
            cmbSheet.Size = new System.Drawing.Size(209, 23);
            cmbSheet.TabIndex = 3;
            cmbSheet.SelectedIndexChanged += cmbSheet_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(68, 419);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(42, 15);
            label1.TabIndex = 4;
            label1.Text = "Sheet :";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(44, 390);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(66, 15);
            label2.TabIndex = 5;
            label2.Text = "File Name :";
            // 
            // btnExport
            // 
            btnExport.Location = new System.Drawing.Point(400, 419);
            btnExport.Name = "btnExport";
            btnExport.Size = new System.Drawing.Size(55, 23);
            btnExport.TabIndex = 6;
            btnExport.Text = "Export";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // Import
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(800, 450);
            Controls.Add(btnExport);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(cmbSheet);
            Controls.Add(txtFileName);
            Controls.Add(dtView);
            Controls.Add(FileUpload);
            Name = "Import";
            Text = "Import";
            ((System.ComponentModel.ISupportInitialize)dtView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button FileUpload;
        private System.Windows.Forms.DataGridView dtView;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.ComboBox cmbSheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnExport;
    }
}