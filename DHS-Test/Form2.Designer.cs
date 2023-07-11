
namespace DHS_Test
{
    partial class Form2
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
            this.DGVQuotecell = new System.Windows.Forms.DataGridView();
            this.btnImport = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.lblSumSilicon = new System.Windows.Forms.Label();
            this.lblSumQutecell = new System.Windows.Forms.Label();
            this.lblSumLeast = new System.Windows.Forms.Label();
            this.lblSiliconQuotecell = new System.Windows.Forms.Label();
            this.lblLeastQuotecell = new System.Windows.Forms.Label();
            this.lblMatchingQuotecellSilicon = new System.Windows.Forms.Label();
            this.lblMatchingLeastQuetecell = new System.Windows.Forms.Label();
            this.btnExcelExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DGVQuotecell)).BeginInit();
            this.SuspendLayout();
            // 
            // DGVQuotecell
            // 
            this.DGVQuotecell.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVQuotecell.Location = new System.Drawing.Point(148, 12);
            this.DGVQuotecell.Name = "DGVQuotecell";
            this.DGVQuotecell.RowTemplate.Height = 25;
            this.DGVQuotecell.Size = new System.Drawing.Size(976, 326);
            this.DGVQuotecell.TabIndex = 0;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(1034, 377);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(166, 389);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "Sum Of Silicon :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(157, 418);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 15);
            this.label3.TabIndex = 5;
            this.label3.Text = "Sum Of Qutecell :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(148, 449);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(109, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "Sum of LeastValue :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(375, 363);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(199, 15);
            this.label5.TabIndex = 7;
            this.label5.Text = "LeastValue And Quotecell Compare :";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(395, 393);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(179, 15);
            this.label6.TabIndex = 8;
            this.label6.Text = "Silicon And Quotecell Compare :";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(381, 418);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(194, 15);
            this.label7.TabIndex = 9;
            this.label7.Text = "Matching In Quotecell And Silicon :";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(375, 449);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(200, 15);
            this.label8.TabIndex = 10;
            this.label8.Text = "Matching LeastValue And Quetecell :";
            // 
            // lblSumSilicon
            // 
            this.lblSumSilicon.AutoSize = true;
            this.lblSumSilicon.Location = new System.Drawing.Point(270, 362);
            this.lblSumSilicon.Name = "lblSumSilicon";
            this.lblSumSilicon.Size = new System.Drawing.Size(0, 15);
            this.lblSumSilicon.TabIndex = 11;
            // 
            // lblSumQutecell
            // 
            this.lblSumQutecell.AutoSize = true;
            this.lblSumQutecell.Location = new System.Drawing.Point(270, 389);
            this.lblSumQutecell.Name = "lblSumQutecell";
            this.lblSumQutecell.Size = new System.Drawing.Size(0, 15);
            this.lblSumQutecell.TabIndex = 12;
            // 
            // lblSumLeast
            // 
            this.lblSumLeast.AutoSize = true;
            this.lblSumLeast.Location = new System.Drawing.Point(270, 418);
            this.lblSumLeast.Name = "lblSumLeast";
            this.lblSumLeast.Size = new System.Drawing.Size(0, 15);
            this.lblSumLeast.TabIndex = 13;
            // 
            // lblSiliconQuotecell
            // 
            this.lblSiliconQuotecell.AutoSize = true;
            this.lblSiliconQuotecell.Location = new System.Drawing.Point(574, 389);
            this.lblSiliconQuotecell.Name = "lblSiliconQuotecell";
            this.lblSiliconQuotecell.Size = new System.Drawing.Size(0, 15);
            this.lblSiliconQuotecell.TabIndex = 14;
            // 
            // lblLeastQuotecell
            // 
            this.lblLeastQuotecell.AutoSize = true;
            this.lblLeastQuotecell.Location = new System.Drawing.Point(574, 363);
            this.lblLeastQuotecell.Name = "lblLeastQuotecell";
            this.lblLeastQuotecell.Size = new System.Drawing.Size(0, 15);
            this.lblLeastQuotecell.TabIndex = 15;
            // 
            // lblMatchingQuotecellSilicon
            // 
            this.lblMatchingQuotecellSilicon.AutoSize = true;
            this.lblMatchingQuotecellSilicon.Location = new System.Drawing.Point(574, 418);
            this.lblMatchingQuotecellSilicon.Name = "lblMatchingQuotecellSilicon";
            this.lblMatchingQuotecellSilicon.Size = new System.Drawing.Size(0, 15);
            this.lblMatchingQuotecellSilicon.TabIndex = 16;
            // 
            // lblMatchingLeastQuetecell
            // 
            this.lblMatchingLeastQuetecell.AutoSize = true;
            this.lblMatchingLeastQuetecell.Location = new System.Drawing.Point(574, 449);
            this.lblMatchingLeastQuetecell.Name = "lblMatchingLeastQuetecell";
            this.lblMatchingLeastQuetecell.Size = new System.Drawing.Size(0, 15);
            this.lblMatchingLeastQuetecell.TabIndex = 17;
            // 
            // btnExcelExport
            // 
            this.btnExcelExport.Location = new System.Drawing.Point(901, 377);
            this.btnExcelExport.Name = "btnExcelExport";
            this.btnExcelExport.Size = new System.Drawing.Size(101, 23);
            this.btnExcelExport.TabIndex = 18;
            this.btnExcelExport.Text = "Export Excel";
            this.btnExcelExport.UseVisualStyleBackColor = true;
            this.btnExcelExport.Click += new System.EventHandler(this.btnExcelExport_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 593);
            this.Controls.Add(this.btnExcelExport);
            this.Controls.Add(this.lblMatchingLeastQuetecell);
            this.Controls.Add(this.lblMatchingQuotecellSilicon);
            this.Controls.Add(this.lblLeastQuotecell);
            this.Controls.Add(this.lblSiliconQuotecell);
            this.Controls.Add(this.lblSumLeast);
            this.Controls.Add(this.lblSumQutecell);
            this.Controls.Add(this.lblSumSilicon);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.DGVQuotecell);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGVQuotecell)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DGVQuotecell;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblSumSilicon;
        private System.Windows.Forms.Label lblSumQutecell;
        private System.Windows.Forms.Label lblSumLeast;
        private System.Windows.Forms.Label lblSiliconQuotecell;
        private System.Windows.Forms.Label lblLeastQuotecell;
        private System.Windows.Forms.Label lblMatchingQuotecellSilicon;
        private System.Windows.Forms.Label lblMatchingLeastQuetecell;
        private System.Windows.Forms.Button btnExcelExport;
    }
}