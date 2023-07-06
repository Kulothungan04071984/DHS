using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace DHS_Test
{
    public partial class Export_Import_Excel : Form
    {
        public Export_Import_Excel()
        {
            InitializeComponent();
        }

        private void importExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //  var filePath = string.Empty;  
            //OpenFileDialog OpenFile = new OpenFileDialog();  

            //OpenFile.Filter = "Excel Files|*.xl*"; //Filter for excel file  
            //OpenFile.Title = "Select your Test Score file";  
            //OpenFile.FilterIndex = 2;  //Don't know what it mean  
            //OpenFile.RestoreDirectory = true;  

            //if (OpenFile.ShowDialog() == DialogResult.OK)  
            //{  
            //    //Get the path of specified file  
            //    filePath = OpenFile.FileName;  
            //}  



            //Excel.Application xlApp = new Excel.Application();  

            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);  

            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];  



            //SqlBulkCopy oSqlBulk = null;  


            //OleDbConnection myExcelConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" + filePath + "; Extended Properties = 'Excel 8.0; HDR = Yes; IMEX = 1'");  



            //try  
            //{  
            //    myExcelConn.Open();  

            //    OleDbCommand ObjCmd = new OleDbCommand("SELECT ExamID, StudentID, Score FROM [" + xlWorksheet.Name + "$]", myExcelConn);  

            //    // READ THE DATA EXTRACTED FROM THE EXCEL FILE.  
            //    OleDbDataReader objBulkReader = null;  
            //    objBulkReader = ObjCmd.ExecuteReader();  


            //    /*  
            //    OleDbCommand objOleDB =  
            //        new OleDbCommand("SELECT ExamID, StudentID, Score FROM xlWorksheet", myExcelConn);  

            //    // READ THE DATA EXTRACTED FROM THE EXCEL FILE.  
            //    OleDbDataReader objBulkReader = null;  
            //    objBulkReader = objOleDB.ExecuteReader();  
            //    */  


            //    xlWorkbook.Close(false);  
            //    xlApp.Quit();  


            //    // SET THE CONNECTION STRING.  

            //    using (SqlConnection con = new SqlConnection(GlobalVariables.ConnectionString))  
            //    {  
            //        con.Open();  

            //        // FINALLY, LOAD DATA INTO THE DATABASE TABLE.  
            //        oSqlBulk = new SqlBulkCopy(con);  
            //        oSqlBulk.DestinationTableName = "Test_score"; // TABLE NAME.  
            //        oSqlBulk.WriteToServer(objBulkReader);  
            //    }  

            //    MessageBox.Show(GlobalVariables.Username + ": Your data imported sucessfully.");  

            //}  


            //catch (Exception ex)  
            //{  

            //    MessageBox.Show(ex.Message);  

            //}  

            //finally  
            //{  
            //        // CLEAR.  
            //        oSqlBulk.Close();  
            //        oSqlBulk = null;  
            //        myExcelConn.Close();  
            //        myExcelConn = null;  
            //}  

            var filePath = string.Empty;
            OpenFileDialog OpenFile = new OpenFileDialog();

            OpenFile.Filter = "Excel Files|*.xlsx*"; //Filter for excel file  
            OpenFile.Title = "Select your Test Score file";
            OpenFile.FilterIndex = 2;  //Don't know what it mean  
            OpenFile.RestoreDirectory = true;

            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file  
                filePath = OpenFile.FileName;
            }
            DataTable table = new DataTable();
            //string odbcstr = "Dsn=my sql;description=SQL SERVER;trusted_connection=Yes;app=Microsoft® Visual Studio®;wsid=LANGUAGE-TEAM";
            string odbcstr = ConfigurationManager.AppSettings["conn"].ToString();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                // Now you can get data from each sheet by its index or its "name"  
                table = dataSet.Tables[0];

                //...  
            }
            try
            {
                using (OdbcConnection con = new OdbcConnection(odbcstr))
                {
                    con.Open();
                    foreach (DataRow item in table.Rows)
                    {
                        Console.WriteLine(item[0]);
                        Console.WriteLine(item[1]);
                        Console.WriteLine(item[2]);

                        string sqlinsert = string.Format("Insert into Material(MPN,MOQ,Description,[Min Price ($) Silicon Expert],[Unit price($)(Quotecell)],[Unit price($)(Octopart)],[Unit price($)(Oemsecretes)],[Least out of vendors]) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7]);
                        OdbcCommand command = new OdbcCommand(sqlinsert, con);
                        command.ExecuteNonQuery();

                    }
                    con.Close();


                }

                MessageBox.Show(": Your data imported sucessfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());


            }
        }
    }
}
