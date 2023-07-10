using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Z.Dapper.Plus;

namespace DHS_Test
{
    public partial class Import : Form
    {
        private List<DataTable> tableCollection;

        public Import()
        {
            InitializeComponent();
            tableCollection = new List<DataTable>();
        }

        private void cmbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cmbSheet.SelectedIndex];
            //dtView.DataSource = dt;
           
            if (dt != null)
            {
                List<Quotecell> quotes = new List<Quotecell>();
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    Quotecell quotecell = new Quotecell();
                    quotecell.MPN = dt.Rows[i]["MPN"].ToString();
                    quotecell.MOQ = dt.Rows[i]["MOQ"].ToString() ;
                    quotecell.Description = dt.Rows[i]["[Description]"].ToString();
                    quotecell.Silicon = Convert.ToDouble(dt.Rows[i]["Silicon"].ToString());
                    quotecell.unitpriceQuotecell = Convert.ToDouble(dt.Rows[i]["Quotecell"].ToString());
                    quotecell.unitpriceoctopart = Convert.ToDouble(string.IsNullOrEmpty(dt.Rows[i]["Octopart"].ToString())? "0.00" : dt.Rows[i]["Octopart"].ToString());
                    quotecell.unitpriceOEMsectrets = Convert.ToDouble(string.IsNullOrEmpty(dt.Rows[i]["Oemsecretes"].ToString()) ? "0.00" : dt.Rows[i]["Oemsecretes"].ToString());
                    quotecell.Listoutofvendors = Convert.ToDouble(dt.Rows[i]["Least"].ToString());
                    quotecell.vendor = dt.Rows[i]["LeastVentor"].ToString();
                    quotes.Add(quotecell);
                }
                dtView.DataSource = quotes;
                
            }

        }

        private void FileUpload_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = openFileDialog.FileName;

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(openFileDialog.FileName, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        Sheets sheets = workbookPart.Workbook.Sheets;

                        tableCollection.Clear();
                        cmbSheet.Items.Clear();

                        foreach (Sheet sheet in sheets)
                        {
                            DataTable dt = new DataTable();
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                if (row.RowIndex.Value == 1)
                                {
                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        dt.Columns.Add(GetCellValue(cell, workbookPart));
                                    }
                                }
                                else
                                {
                                    DataRow dataRow = dt.NewRow();
                                    int columnIndex = 0;

                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        dataRow[columnIndex] = GetCellValue(cell, workbookPart);
                                        columnIndex++;
                                    }

                                    dt.Rows.Add(dataRow);
                                }
                            }

                            tableCollection.Add(dt);
                            cmbSheet.Items.Add(sheet.Name);
                        }
                    }
                }
            }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType == null)
            {
                return cell.InnerText;
            }

            string value = cell.InnerText;

            if (cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringId = int.Parse(value);
                SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (sharedStringPart != null)
                {
                    value = sharedStringPart.SharedStringTable.ElementAt(sharedStringId).InnerText;
                }
            }

            return value;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                string connstring = "Server=10.10.112.91;Database=Quotecell;uid=sa;pwd=Syrma@123";
                DapperPlusManager.Entity<Quotecell>().Table("Material");
                
                List<Quotecell> quote = dtView.DataSource as List<Quotecell>;
                if (quote != null)
                {
                    using (IDbConnection db = new SqlConnection(connstring))
                    {
                        db.BulkInsert(quote);
                    }
                }
                MessageBox.Show("finish");
                // Create a SqlConnection object and connect to the SQL database.
                //SqlConnection connection = new SqlConnection("Data Source=localhost;Initial Catalog=MyDatabase;Integrated Security=True");
                //connection.Open();

                //// Create a SqlBulkCopy object and set the destination table to the table you want to import the data to.
                //SqlBulkCopy bulkCopy = new SqlBulkCopy(connection);
                //bulkCopy.DestinationTableName = "MyTable";

                //// Create a FileStream object and open the Excel file.
                //FileStream fs = new FileStream("MyExcelFile.xlsx", FileMode.Open);

                //// Use the SqlBulkCopy object to read the data from the Excel file and insert it into the SQL table.
                ////bulkCopy.ReadFromStream(fs);

                //// Close the FileStream object.
                //fs.Close();

                //// Close the SqlConnection object.
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FileUpload_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = openFileDialog.FileName;

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(openFileDialog.FileName, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        Sheets sheets = workbookPart.Workbook.Sheets;

                        tableCollection.Clear();
                        cmbSheet.Items.Clear();

                        foreach (Sheet sheet in sheets)
                        {
                            DataTable dt = new DataTable();
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                if (row.RowIndex.Value == 1)
                                {
                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        dt.Columns.Add(GetCellValue(cell, workbookPart));
                                    }
                                }
                                else
                                {
                                    DataRow dataRow = dt.NewRow();
                                    int columnIndex = 0;

                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        dataRow[columnIndex] = GetCellValue(cell, workbookPart);
                                        columnIndex++;
                                    }

                                    dt.Rows.Add(dataRow);
                                }
                            }

                            tableCollection.Add(dt);
                            cmbSheet.Items.Add(sheet.Name);
                        }
                    }
                }
            }
        }

        private void btnExport_Click_1(object sender, EventArgs e)
        {
            try
            {
                string connstring = "Server=10.10.112.91;Database=Quotecell;uid=sa;pwd=Syrma@123";
                DapperPlusManager.Entity<Quotecell>().Table("Material");
                List<Quotecell> quote = dtView.DataSource as List<Quotecell>;
                if (quote != null)
                {
                    using (IDbConnection db = new SqlConnection(connstring))
                    {
                        db.BulkInsert(quote);
                    }
                }
                MessageBox.Show("finish");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbSheet_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cmbSheet.SelectedIndex];
            //dtView.DataSource = dt;

            if (dt != null)
            {
                List<Quotecell> quotes = new List<Quotecell>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Quotecell quotecell = new Quotecell();
                    quotecell.MPN = dt.Rows[i]["MPN"].ToString();
                    quotecell.MOQ = dt.Rows[i]["MOQ"].ToString();
                    quotecell.Description = dt.Rows[i]["[Description]"].ToString();
                    quotecell.Silicon = Convert.ToDouble(dt.Rows[i]["Silicon"].ToString());
                    quotecell.unitpriceQuotecell = Convert.ToDouble(dt.Rows[i]["Quotecell"].ToString());
                    quotecell.unitpriceoctopart = Convert.ToDouble(string.IsNullOrEmpty(dt.Rows[i]["Octopart"].ToString()) ? "0.00" : dt.Rows[i]["Octopart"].ToString());
                    quotecell.unitpriceOEMsectrets = Convert.ToDouble(string.IsNullOrEmpty(dt.Rows[i]["Oemsecretes"].ToString()) ? "0.00" : dt.Rows[i]["Oemsecretes"].ToString());
                    quotecell.Listoutofvendors = Convert.ToDouble(dt.Rows[i]["[Least]"].ToString());
                    quotecell.vendor = dt.Rows[i]["LeastVentor"].ToString();
                    quotes.Add(quotecell);
                }
                dtView.DataSource = quotes;
            }


        }
    }
}
