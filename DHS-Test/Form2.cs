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
           
        }
       

        public List<string> getLeastValue(string Silicon, string Quotecell, string Octopart,string Oemsecretes)
        {
            // Dictionary<string, decimal> obj = new Dictionary<string, decimal>();
            List<string> obj = new List<string>();
            try
            {
                string vendor = string.Empty;
                
                var sili = Convert.ToDecimal(Silicon=="0" ? int.MaxValue : Silicon);
                var Quot = Convert.ToDecimal(Quotecell == "0" ? int.MaxValue : Quotecell);
                var Oct = Convert.ToDecimal(Octopart == "0" ? int.MaxValue : Octopart);
                var Oem = Convert.ToDecimal(Oemsecretes == "0" ? int.MaxValue : Oemsecretes);
               
                decimal[] val = {sili, Quot, Oct, Oem };
                var fmin = (from value in val select value).Min();
                var IP = val.findIndex(fmin);
                if (IP == 0) { vendor = "Silicon"; } else if (IP == 1) { vendor = "Quotecell"; } else if (IP==2 ) { vendor = "Octopart"; } else if (IP == 3) { vendor = "Oemsecretes"; }
                obj.Add(fmin.ToString());
                obj.Add(vendor);
                return obj;

            }
            catch(Exception ex)
            {
                return obj;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtRowCount.Text.ToString() != string.Empty)
                {
                    OldDataUpdate();
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

                    //  Microsoft.Office.Interop.Excel.AutoFilter filter = range.AutoFilter;
                    Microsoft.Office.Interop.Excel.Range vRange = range.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible);

                   // string insertQuery = "INSERT INTO Material (MPN, MOQ, [Description],Silicon,Quotecell,Octopart,Oemsecretes,[Least],VendorName,IsActive) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8, @Value9,@Value10)";
                    SqlCommand command = new SqlCommand("pro_InsertExcelValues", connection);
                    command.CommandType = CommandType.StoredProcedure;
                    var count = Convert.ToInt32(txtRowCount.Text.ToString());
                    for (int row = 2; row <= count; row++)
                    {
                        var checkEmpty = ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 1]).Value2.ToString();
                        if (checkEmpty == null || checkEmpty.Length == 0)
                            return;
                        command.Parameters.Clear();
                        command.Parameters.AddWithValue("@MPN", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 1]).Value2.ToString());
                        command.Parameters.AddWithValue("@MOQ", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 2]).Value2.ToString());
                        command.Parameters.AddWithValue("@Description", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 3]).Value2.ToString());
                        command.Parameters.AddWithValue("@Silicon", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 4]).Value2.ToString());
                        command.Parameters.AddWithValue("@Quotecell", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 5]).Value2.ToString());
                        command.Parameters.AddWithValue("@Octopart", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 6]).Value2.ToString());
                        command.Parameters.AddWithValue("@Oemsecretes", ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 7]).Value2.ToString());


                        var LeastValue = getLeastValue(
                         ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 4]).Value2.ToString()
                        , ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 5]).Value2.ToString()
                        , ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 6]).Value2.ToString()
                        , ((Microsoft.Office.Interop.Excel.Range)range.Cells[row, 7]).Value2.ToString());

                        command.Parameters.AddWithValue("@Least", LeastValue[0].ToString());
                        command.Parameters.AddWithValue("@VendorName", LeastValue[1].ToString());
                       // command.Parameters.AddWithValue("@Value10", 1);


                        command.ExecuteNonQuery();
                    }
                    workbook.Close();
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    SqlCommand cmd = new SqlCommand("pro_getExcelValues", connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    //System.Data.DataTable dt = new System.Data.DataTable();
                    //da.Fill(dt);
                    DGVQuotecell.DataSource = ds.Tables[0];
                    if(ds.Tables[1].Rows.Count > 0)
                    {
                        lblSumSilicon.Text = ds.Tables[1].Rows[0]["silicon"].ToString();
                        lblSumQutecell.Text= ds.Tables[1].Rows[0]["Quotecell"].ToString();
                        lblSumLeast.Text= ds.Tables[1].Rows[0]["LeastValue"].ToString();
                        lblLeastQuotecell.Text= ds.Tables[1].Rows[0]["LeastANDQuotecelCompare"].ToString();
                        lblMatchingQuotecellSilicon.Text = Convert.ToString(Math.Round(Convert.ToDecimal(ds.Tables[1].Rows[0]["MatchingINQuotecellAndSilicon"].ToString()), 2)) + " %";
                        lblMatchingLeastQuetecell.Text=Convert.ToString(Math.Round(Convert.ToDecimal(ds.Tables[1].Rows[0]["MatchingLeastAndQuetecell"].ToString()),2)) + " %";
                        lblSiliconQuotecell.Text= ds.Tables[1].Rows[0]["SiliconANDQuotecelCompare"].ToString();
                    }

                    MessageBox.Show("Successfully Insert");
                }
                else
                    MessageBox.Show("Please Insert Valid Row Count");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }

        private void OldDataUpdate()
        {
            string connectionString = "Server=10.10.112.91;Database=Quotecell;uid=sa;pwd=Syrma@123;";
            SqlConnection connectionNew = new SqlConnection(connectionString);
            SqlCommand cmdUpdate = new SqlCommand("pro_UpdateMaterial", connectionNew);
            cmdUpdate.CommandType = CommandType.StoredProcedure;
            connectionNew.Open();
            cmdUpdate.ExecuteNonQuery();
            connectionNew.Close();
        }

        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            DGVQuotecell.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DGVQuotecell.MultiSelect = true;
            DGVQuotecell.SelectAll();

            DataObject dataObj = DGVQuotecell.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);

            //Open an excel instance and paste the copied data
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            //cmdSelecionar.Enabled = true;
        }
    }
    public static class Extensions
    {
        public static int findIndex<T>(this T[] array, T item)
        {
            return array
                .Select((element, index) => new { element, index })
                .FirstOrDefault(x => x.element.Equals(item))?.index ?? -1;
        }
    }
}
