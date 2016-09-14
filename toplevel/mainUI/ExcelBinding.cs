using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace mainUI
{
    public class ExcelBinding
    {
        private string FileNameSearch;
        private string SheetNameSearch;        
        private int ResultPosX;
        private int TesterPosX;
        private int VerPosX;
        private int DatePosX;
        private int RemarkPosX;
        private EN_FILE_STATE fs;

        #region Public Definition
        public ExcelBinding() 
        { 
            Console.WriteLine("Construct ExcelBinding Class");
            init();
        }

        public string SendFileName { set { FileNameSearch = value.ToString(); } }
        public string SendSheetName { set { SheetNameSearch = value.ToString(); } }
        public DataTable SaveAsResult { set { saveAs_data(value); } }       
        public DataTable GetResult { get { return get_data(); } }
        public List<String> GetSheetName { get { return get_sheet_name(); } }
        //public DataSet GetResultAllSheets { get { return get_data_all_sheets(); } } // Will be updated upon request
        public Boolean EmptyTable { get { return init(); } }
        #endregion

        #region Private Definition

        private enum EN_FILE_STATE
        {
            FILE_STATE_EMPTY,
            FILE_STATE_LOAD,
            FILE_STATE_SAVEAS,
            FILE_STATE_MAX
        }

        private Boolean init()
        {
            FileNameSearch = "";
            SheetNameSearch = "";
            ResultPosX = 0;
            TesterPosX = 0;
            VerPosX = 0;
            DatePosX = 0;
            RemarkPosX = 0;
            fs = EN_FILE_STATE.FILE_STATE_EMPTY;

            return true;
        }

        private List<string> get_sheet_name()
        {
            Console.WriteLine("ExcelBinding:get_sheet_name >> File Name : " + FileNameSearch);
            //String strExcelConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source="+FileNameSearch+";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
            String strExcelConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNameSearch + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 XML;HDR=No;IMEX=1\"";
            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();
            List<string> SheetNameList = new List<string>();
            DataTable dtSheet = new DataTable();

            cmdExcel.Connection = connExcel;
            try
            {
                connExcel.Open();
                dtSheet = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtSheet.Rows.Count > 1)
                {
                    foreach (DataRow row in dtSheet.Rows)
                    {
                        SheetNameList.Add(row["TABLE_NAME"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
            }

            return SheetNameList;
        }

        private DataTable get_data()
        {
            Console.WriteLine("ExcelBinding:get_data >> File Name : " + FileNameSearch);
            //String strExcelConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+FileNameSearch+"';Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
            String strExcelConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNameSearch + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";
            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();
            DataTable dtContent = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            cmdExcel.Connection = connExcel;            
            
            try
            {
                Console.WriteLine("get_data >> Sheet Name : " + SheetNameSearch);
                cmdExcel.CommandText = "select * from [" + SheetNameSearch + "]";

                connExcel.Open();            
                da.SelectCommand = cmdExcel;
                da.Fill(dtContent);                            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connExcel.Close();                
                connExcel.Dispose();
                da.Dispose();
            }

            dtContent = modify_data(dtContent);

            fs = EN_FILE_STATE.FILE_STATE_LOAD;
            return dtContent;
        }

        public Boolean save_data(List<string> NoToSave, List<string> ResultToSave, List<string> VerToSave, List<string> TesterToSave, List<string> DateToSave, List<string> RemarkToSave)
        {
            Console.WriteLine("ExcelBinding:save_data >> File Name : " + FileNameSearch);
            //String strExcelConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source="+FileNameSearch+";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
            String strExcelConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNameSearch + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 XML;HDR=No\"";
            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();
            cmdExcel.Connection = connExcel;

            char[] MyChar = { '$', '\'' };
            string SheetNameTrim = SheetNameSearch.Trim(MyChar);
            string cmdStr = "";

            try
            {              
                Console.WriteLine("save_data >> Sheet Name : " + SheetNameTrim);

                connExcel.Open();

                //int i = 1;
                bool flag = true;
                for(int j = 0; j< ResultToSave.Count; j++)
                {
                    switch(fs)
                    {
                        case EN_FILE_STATE.FILE_STATE_LOAD:
                            cmdStr = "UPDATE [" + SheetNameTrim + "$] SET F" + ResultPosX + " = '" + ResultToSave[j].ToString() + "', F" + VerPosX + " = '" + VerToSave[j].ToString() + "', F" + TesterPosX + " = '" + TesterToSave[j].ToString() + "', F" + DatePosX + " = '" + DateToSave[j].ToString() + "', F" + RemarkPosX + " = '" + RemarkToSave[j].ToString() + "' where F2 = '" + NoToSave[j].ToString() + "'";
                            //i++;
                            break;
                        case EN_FILE_STATE.FILE_STATE_SAVEAS:
                            cmdStr = "UPDATE [" + SheetNameTrim + "] SET F" + (ResultPosX - 1) + " = '" + ResultToSave[j].ToString() + "', F" + (VerPosX - 1) + " = '" + VerToSave[j].ToString() + "', F" + (TesterPosX - 1) + " = '" + TesterToSave[j].ToString() + "', F" + (DatePosX - 1) + " = '" + DateToSave[j].ToString() + "', F" + (RemarkPosX - 1) + " = '" + RemarkToSave[j].ToString() + "' where F1 = '" + NoToSave[j].ToString() + "'";
                            //i++;
                            break;
                        default:
                            cmdStr = "";
                            flag = false;
                            break;
                    }

                    if(flag)
                    {                        
                        cmdExcel.CommandText = cmdStr;
                        cmdExcel.ExecuteNonQuery();
                    }                    
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
            }
            return true;
        }

        private Boolean saveAs_data(DataTable dtContentToSave)
        {
            Console.WriteLine("ExcelBinding:saveAs_data : " + FileNameSearch);
            String strExcelConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNameSearch + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 XML;HDR=No\"";

            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel1 = new OleDbCommand();
            OleDbCommand cmdExcel2 = new OleDbCommand();
            OleDbCommand cmdExcel3 = new OleDbCommand();

            cmdExcel1.Connection = connExcel;
            cmdExcel2.Connection = connExcel;
            cmdExcel3.Connection = connExcel;

            // Create table (cmdExcel1)
            // Remove header (cmdExcel2)
            // Insert data (cmdExcel3
            char[] MyChar = { '$', '\'' };
            string SheetNameTrim = SheetNameSearch.Trim(MyChar);

            cmdExcel1.CommandText = "CREATE TABLE [" + SheetNameTrim + "] (";
            cmdExcel2.CommandText = "UPDATE [" + SheetNameTrim + "] SET ";
            cmdExcel3.CommandText = "INSERT INTO [" + SheetNameTrim + "] VALUES (";

            for (int i = 0; i < dtContentToSave.Columns.Count - 1; i++)
            {
                cmdExcel1.CommandText += "F" + i + " varchar(30)";
                cmdExcel2.CommandText += "F" + (i+1) + " = \"\"";
                cmdExcel3.CommandText += "@F" + i;
                cmdExcel3.Parameters.Add(new OleDbParameter("@F" + i, OleDbType.VarChar));
                if (i != dtContentToSave.Columns.Count - 2)
                {
                    cmdExcel1.CommandText += ", ";
                    cmdExcel2.CommandText += ", ";
                    cmdExcel3.CommandText += ", ";
                }
            }
            cmdExcel1.CommandText += " );";
            cmdExcel2.CommandText += ";";
            cmdExcel3.CommandText += ")";
            
            try
            {
                connExcel.Open();
                cmdExcel1.ExecuteNonQuery();
                cmdExcel2.ExecuteNonQuery();
                for (int i = 0; i < dtContentToSave.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dtContentToSave.Columns.Count - 1; j++)
                    {
                        cmdExcel3.Parameters["@F" + j].Value = dtContentToSave.Rows[i][j].ToString();
                    }
                    cmdExcel3.ExecuteNonQuery();
                }
                connExcel.Close();
      
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmdExcel1.Dispose();
                cmdExcel2.Dispose();
                connExcel.Dispose();
            }

            fs = EN_FILE_STATE.FILE_STATE_SAVEAS;
            return true;
        }      

        private DataTable modify_data(DataTable dtContent)
        {
            bool colZeroStart = false;
            // Delete unnecessary rows  
            foreach (DataRow dr1 in dtContent.Rows)
            {
                Regex regex = new Regex(@"no|No|NO"); // Set reg pattern
                Match match1 = regex.Match(dr1[0].ToString()); // Find reg pattern in 1st column
                Match match2 = regex.Match(dr1[1].ToString()); // Find reg pattern in 2nd column
                if (!(match1.Success || match2.Success)) 
                {
                    if (match1.Success)
                        colZeroStart = true;

                    // Delete row if cannot find reg pattern in 1st or 2nd column
                    dr1.Delete();
                }
                else
                {
                    // Stop delete operation after finding reg pattern
                    break;
                }
            }
            dtContent.AcceptChanges();

            if(dtContent.Rows.Count < 1)
            {
                MessageBox.Show("Please check test sheet format");
                return dtContent;
            }

            int tempResultPosX = 1;
            int tempTesterPosX = 1;
            int tempVerPosX = 1;
            int tempDatePosX = 1;
            int tempRemarkPosX = 1;
            // Check if first column empty
            if (!colZeroStart)
            {
                dtContent.Columns.RemoveAt(0);
                tempResultPosX = tempResultPosX + 1; // Add result pos x when deleting empty column
                tempTesterPosX = tempTesterPosX + 1; // Add tester pos x when deleting empty column
                tempVerPosX = tempVerPosX + 1; // Add ver pos x when deleting empty column
                tempDatePosX = tempDatePosX + 1; // Add date pos x when deleting empty column
                tempRemarkPosX = tempRemarkPosX + 1; // Add remark pos x when deleting empty column
            }       
            
            ResultPosX = 0;
            TesterPosX = 0;
            VerPosX = 0;
            DatePosX = 0;
            RemarkPosX = 0;
            for (int i = 0; i < dtContent.Columns.Count; i++)
            {
                if (dtContent.Rows[0][i].ToString() == "Result")
                {
                    ResultPosX = tempResultPosX + i;
                }
                else if (dtContent.Rows[0][i].ToString() == "Tester")
                {
                    TesterPosX = tempTesterPosX + i;
                }
                else if (dtContent.Rows[0][i].ToString() == "Ver." || dtContent.Rows[0][i].ToString() == "Ver")
                {
                    VerPosX = tempVerPosX + i;
                }
                else if (dtContent.Rows[0][i].ToString() == "Date" || dtContent.Rows[0][i].ToString() == "Test Date")
                {
                    DatePosX = tempDatePosX + i;
                }
                else if (dtContent.Rows[0][i].ToString() == "Remark" || dtContent.Rows[0][i].ToString() == "Remarks")
                {
                    RemarkPosX = tempRemarkPosX + i;
                }                

                bool[] needRemoval = new bool[dtContent.Rows.Count];
                for (int j = 0; j < dtContent.Rows.Count; j++)
                {
                    var content = dtContent.Rows[j][i].ToString();
                    needRemoval[j] = content.Equals(string.Empty);
                }

                if (needRemoval.All(x => x))
                {
                    // Only delete a columns if all row inside it are empty
                    dtContent.Columns.RemoveAt(i);
                    tempResultPosX = tempResultPosX + 1; // Add result pos x when deleting empty column
                    tempTesterPosX = tempTesterPosX + 1; // Add tester pos x when deleting empty column
                    tempVerPosX = tempVerPosX + 1; // Add ver pos x when deleting empty column
                    tempDatePosX = tempDatePosX + 1; // Add date pos x when deleting empty column
                    tempRemarkPosX = tempRemarkPosX + 1; // Add remark pos x when deleting empty column
                }
            }

            return dtContent;
        }

        [Obsolete("Not use anymore. Will be updated upon request")]
        private DataSet get_data_all_sheets()
        {
            Console.WriteLine("ExcelBinding:get_data_all_sheets >> File Name : " + FileNameSearch);
            //String strExcelConn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source="+FileNameSearch+";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\"";
            String strExcelConn = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileNameSearch + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 XML;HDR=No;IMEX=1\"";
            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();            
            DataTable dtSheet = new DataTable();
            DataTable dtContent = new DataTable();
            DataSet dsContent = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter();

            cmdExcel.Connection = connExcel;

            try
            {
                connExcel.Open();
                dtSheet = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Get sheet name and get data inside the sheet
                SheetNameSearch = "";
                for (int i = 0; i < dtSheet.Rows.Count; i++)
                {                    
                    SheetNameSearch = dtSheet.Rows[i]["TABLE_NAME"].ToString();
                    //Console.WriteLine("get_data_all_sheets >> Sheet Name : " + SheetNameSearch);
                    cmdExcel.CommandText = "select * from [" + SheetNameSearch + "]";

                    dtContent = new DataTable(SheetNameSearch);
                    da.SelectCommand = cmdExcel;
                    da.Fill(dtContent);

                    // Add all tables inside dataset
                    dsContent.Tables.Add(dtContent);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                connExcel.Close();
                da.Dispose();
                connExcel.Dispose();
            }

            return dsContent;
        }

        [Obsolete("Not use anymore. Will be updated upon request")]
        private void copy_merge_cells(DataTable dt)
        {
            // Handle merge cells
            int i = 2;
            foreach (DataRow dr2 in dt.Rows)
            {
                int k = 0;     

                foreach (DataColumn dc1 in dt.Columns)
                {
                    if ((dt.Rows[i][k].ToString() == "") )
                    {
                        dt.Rows[i][k] = dt.Rows[i - 1][k];                
                    }        
                    k++;
                }

                if (i == dt.Rows.Count - 1)
                {
                    break;
                }
                else
                {
                    i++;
                }               
            }
        }
        #endregion
    }        
}
