using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DHS_Test
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                string connectionString = "Server=10.10.112.91;Database=Quotecell;uid=sa;pwd=Syrma@123;";
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.ShowDialog();
                string filePath = openFileDialog.FileName;

                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(filePath);

                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;
                Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                string insertQuery = "INSERT INTO Material (MPN, MOQ, [Description],Silicon,Quotecell,Octopart,Oemsecretes,[Least]) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8)";
                SqlCommand command = new SqlCommand(insertQuery, connection);

                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    command.Parameters.Clear();
                    command.Parameters.AddWithValue("@Value1", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 1]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value2", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 2]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value3", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 3]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value4", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 4]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value5", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 5]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value6", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 6]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value7", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 7]).Value2.ToString());
                    command.Parameters.AddWithValue("@Value8", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 8]).Value2.ToString());
                    

                    command.ExecuteNonQuery();
                }
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           



        }
    }
}
