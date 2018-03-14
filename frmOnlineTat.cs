using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Threading;
using TransHangupHandler.Class;
using System.Data.OleDb;
using System.Globalization;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections;
using NameSpaceExceptionHandler;
using System.Text.RegularExpressions;




namespace TransHangupHandler
{
    public partial class frmOnlineTat : Form
    {
        //----------------------------------------------------Online TAT-----------------------------------
       #region Variables
            DataSet dsMain = null;
            DataTable dtHCA= null;
            
        
            BusinessLogic BusinessLogic = new BusinessLogic();
            string sourcePath = ConfigurationSettings.AppSettings["SourcePath"].ToString();
            string targetPath = ConfigurationSettings.AppSettings["TargetPath"].ToString();
               
        

            string pathToWriteLog = string.Empty;
            
            //collection to hold files which are not processed.
            ArrayList FileWhichHitException;
            
            //To avoid listview flickering
            public enum ListViewExtendedStyles
            {
                /// <summary>
                /// LVS_EX_GRIDLINES
                /// </summary>
                GridLines = 0x00000001,
                /// <summary>
                /// LVS_EX_SUBITEMIMAGES
                /// </summary>
                SubItemImages = 0x00000002,
                /// <summary>
                /// LVS_EX_CHECKBOXES
                /// </summary>
                CheckBoxes = 0x00000004,
                /// <summary>
                /// LVS_EX_TRACKSELECT
                /// </summary>
                TrackSelect = 0x00000008,
                /// <summary>
                /// LVS_EX_HEADERDRAGDROP
                /// </summary>
                HeaderDragDrop = 0x00000010,
                /// <summary>
                /// LVS_EX_FULLROWSELECT
                /// </summary>
                FullRowSelect = 0x00000020,
                /// <summary>
                /// LVS_EX_ONECLICKACTIVATE
                /// </summary>
                OneClickActivate = 0x00000040,
                /// <summary>
                /// LVS_EX_TWOCLICKACTIVATE
                /// </summary>
                TwoClickActivate = 0x00000080,
                /// <summary>
                /// LVS_EX_FLATSB
                /// </summary>
                FlatsB = 0x00000100,
                /// <summary>
                /// LVS_EX_REGIONAL
                /// </summary>
                Regional = 0x00000200,
                /// <summary>
                /// LVS_EX_INFOTIP
                /// </summary>
                InfoTip = 0x00000400,
                /// <summary>
                /// LVS_EX_UNDERLINEHOT
                /// </summary>
                UnderlineHot = 0x00000800,
                /// <summary>
                /// LVS_EX_UNDERLINECOLD
                /// </summary>
                UnderlineCold = 0x00001000,
                /// <summary>
                /// LVS_EX_MULTIWORKAREAS
                /// </summary>
                MultilWorkAreas = 0x00002000,
                /// <summary>
                /// LVS_EX_LABELTIP
                /// </summary>
                LabelTip = 0x00004000,
                /// <summary>
                /// LVS_EX_BORDERSELECT
                /// </summary>
                BorderSelect = 0x00008000,
                /// <summary>
                /// LVS_EX_DOUBLEBUFFER
                /// </summary>
                DoubleBuffer = 0x00010000,
                /// <summary>
                /// LVS_EX_HIDELABELS
                /// </summary>
                HideLabels = 0x00020000,
                /// <summary>
                /// LVS_EX_SINGLEROW
                /// </summary>
                SingleRow = 0x00040000,
                /// <summary>
                /// LVS_EX_SNAPTOGRID
                /// </summary>
                SnapToGrid = 0x00080000,
                /// <summary>
                /// LVS_EX_SIMPLESELECT
                /// </summary>
                SimpleSelect = 0x00100000
            }
            public enum ListViewMessages
            {
                First = 0x1000,
                SetExtendedStyle = (First + 54),
                GetExtendedStyle = (First + 55),
            }

            //---------Hangup Handler Varibles-------------------
            string FileExtension = string.Empty;
            int TotalRecordsInDtable;
            int TotalColumnsInDtable;
            int iProcessedFiles;
            string AccountName = string.Empty;
            //To get the exact location of the client or account
            string SITESpecific = string.Empty;
            DataSet dsTransHanupExcelContent = new DataSet();



            public class LstCollection
            {
                public int iType;
                public ListViewItem oItem;
            }
            public enum lstViewTpe
            {
                TwoOne6 = 1,
                HangUP = 2
            }
            //---------MASTER DATA Varibles-------------------

            DateTime FromDate;
            DateTime ToDate;
            String ExcelFile = string.Empty;


        #endregion                     
       #region Methods

          public frmOnlineTat()
        {
            InitializeComponent();

        }   
        private void frmOnlineTat_Load(object sender, EventArgs e)
        {
            try
            {   //To avoid Listview Flickering 
                //To pass the listview control name to the method in the Listview Helper class
                //And make sure above two ENUMS are used.
                ListViewHelper.EnableDoubleBuffer(lsvOnlineTat);
                ListViewHelper.EnableDoubleBuffer(lsvContent);

                //SS Module does not includes a Master data table
                //tabControlMain.TabPages.Remove(tabMasterData);

                LoadReport();
                showStatus("Ready.", true);
                showStatus("Ready.", true, "forHangup");
                Control.CheckForIllegalCrossThreadCalls = false;
                Thread oThread = new Thread(new ThreadStart(ProcessFiles));
                oThread.Start();

            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }      
        }

        private void LoadReport()
        {
            try
            {
                DataSet dsHangup = BusinessLogic.TransHangUpReport(DateTime.Now.Month, DateTime.Now.Year);
                if (dsHangup.Tables[0].Rows[0]["client"].ToString() == "dictaphone")
                {
                    lblDateDicta.Text = dsHangup.Tables[0].Rows[0]["Date"].ToString() + "    (  No of Files " + dsHangup.Tables[0].Rows[0]["Total_files"].ToString() + "  )";
                    lblDateEsc.Text = dsHangup.Tables[0].Rows[1]["Date"].ToString() + "    (  No of Files " + dsHangup.Tables[0].Rows[1]["Total_files"].ToString() + "  )";
                }
                else
                {
                    lblDateEsc.Text = dsHangup.Tables[0].Rows[0]["Date"].ToString() + "    (  No of Files " + dsHangup.Tables[0].Rows[0]["Total_files"].ToString() + "  )";
                    lblDateDicta.Text = dsHangup.Tables[0].Rows[1]["Date"].ToString() + "    (  No of Files " + dsHangup.Tables[0].Rows[1]["Total_files"].ToString() + "  )";
                }
                lblTotalJobsToBe.Text = dsHangup.Tables[1].Rows.Count.ToString() + " Jobs to be Reconciled";
                int iRowNumber = 0;
                lsvtobereconciled.Items.Clear();
                lsvtobereconciled.BeginUpdate();
                foreach (DataRow drow in dsHangup.Tables[1].Rows)
                {

                    iRowNumber++;
                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                    lvi.SubItems.Add(drow["voice_file_id"].ToString());
                    lvi.SubItems.Add(drow["file_date"].ToString());
                    lvi.SubItems.Add(drow["client_name"].ToString());
                    lsvtobereconciled.Items.Add(lvi);

                }
                lsvtobereconciled.EndUpdate();
                Reset_ListViewColumn(lsvtobereconciled);
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }      
        }

        /// <summary>
        /// This sets the width for the list view control based on header content or field content
        /// </summary>
        /// <param name="oList"></param>
        public static void Reset_ListViewColumn(ListView oList)
        {
            try
            {
                foreach (ColumnHeader oColumn in oList.Columns)
                {
                    oColumn.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                    int iWidth = oColumn.Width;
                    oColumn.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                    if (iWidth > oColumn.Width)
                        oColumn.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Resizing listview columns " + Environment.NewLine + ex.ToString());
            }
        }
        //For OnlineTAT
        public void showStatus(string sMessage, bool bSuccessOrError)
        {
            
            //To change colour if its a error
            if (!bSuccessOrError)
                statusLblOnline.ForeColor = System.Drawing.Color.Red;
            else
                statusLblOnline.ForeColor = System.Drawing.Color.Black;
            
            //To show content
            statusLblOnline.Text = sMessage.ToString();            
        }
        //For Hangup        
        public void showStatus(string sMessage, bool bSuccessOrError,string forHangUp)
        {

            //To change colour if its a error
            if (!bSuccessOrError)
                statusLabelHangUp .ForeColor = System.Drawing.Color.Red;
            else
                statusLabelHangUp.ForeColor = System.Drawing.Color.Black;

            //To show content
            statusLabelHangUp.Text = sMessage.ToString();
        }
        //For MasterReport        
        public void showStatus(string sMessage, bool bSuccessOrError,byte bforMasterData)
        {

            //To change colour if its a error
            if (!bSuccessOrError)
                statusLableMasterData.ForeColor = System.Drawing.Color.Red;
            else
                statusLableMasterData.ForeColor = System.Drawing.Color.Black;
            //To show content
            statusLableMasterData.Text = sMessage.ToString();
        }

        private void ProcessFiles()
        {
            try
            {
                do
                {
                    try
                    {
                        lsvOnlineTat.Items.Clear();
                        pathToWriteLog = "";
                        
                        BrowseFiles();
                    }
                    catch (Exception ex)
                    {
                        ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                    }
                    finally 
                    {   
                        DeleteSourceFile(); 
                    }
                    //To re run after a minute
                    Thread.Sleep((1000 * 300));
                } while (true);
            }
            catch(Exception ex) 
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }
        private void BrowseFiles()
        {
            try
            {
                            
                DirectoryInfo directoryInfo = new DirectoryInfo(sourcePath);
                string newSubFolder = string.Empty;
                if (Directory.Exists(sourcePath))
                {
                    
                    string[] files = System.IO.Directory.GetFiles(sourcePath);
                    if (files.Length <= 0)
                    {
                        showStatus("Directory is empty, feed me with files  " + sourcePath + "", true);
                        return;
                    }
                   
                    if (!Directory.Exists(targetPath))
                        Directory.CreateDirectory(targetPath);
                    showStatus("Loading files in directory...", true); 
                    var shortDate = String.Format("{0:m}", DateTime.Now);
                    string newfolderName = shortDate;
                    string newFolder = Path.Combine(targetPath, newfolderName);
                    if (!Directory.Exists(newFolder))
                        Directory.CreateDirectory(newFolder);
                    
                    string[] newfolderfiles = System.IO.Directory.GetDirectories(newFolder);
                    int count = newfolderfiles.Length;
                    string newSubFolderName = string.Empty;
                    newSubFolderName = Convert.ToInt32(count + 1).ToString();
                    newSubFolder = Path.Combine(newFolder, newSubFolderName);
                    if (!Directory.Exists(newSubFolder))
                        Directory.CreateDirectory(newSubFolder);

                    dsMain = new DataSet();
                    

                    foreach (string file in files)
                    {
                        string fileName = Path.GetFileName(file);
                        string sourceFile = Path.Combine(sourcePath, fileName);
                        string destFile = Path.Combine(newSubFolder, fileName);
                        dsMain.Tables.Add(CsvFileToDatatable(sourceFile, true));
                        File.Copy(sourceFile, destFile, true);
                        pathToWriteLog = destFile;
                    }


                    // To update the content in dataset to List View
                    showStatus("Completed loading files in directory.", true);
                    LoadListView(dsMain);
                }
                else
                {   //To create directory
                    DialogResult createDirectory = MessageBox.Show("Directory does not exists would you like to create?" + Environment.NewLine + "Required path " + sourcePath + "", "Directory not found", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error);
                    if (createDirectory == DialogResult.Yes)
                        if (!Directory.Exists(Path.GetDirectoryName(sourcePath.ToString())))
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(sourcePath.ToString()));
                            if (Directory.Exists(Path.GetDirectoryName(sourcePath.ToString())))
                                showStatus("Directory created sucessfully please use it to process online files. " + sourcePath, true);
                            else
                                showStatus("Failed creating directory", false);
                        }
                }
            }
            catch (Exception ex)
            {   
                if(ex.Message.Contains("The network path was not found"))
                    showStatus(sourcePath + " is not accessible.",false);

                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
            finally 
            { }
        }

        public int CountStringOccurrences(string text, string pattern)
        {
            // Loop through all instances of the string 'text'.
            int count = 0;
            int i = 0;
            while ((i = text.IndexOf(pattern, i)) != -1)
            {
                i += pattern.Length;
                count++;
            }
            return count;
        }

        public int ConvertToSeconds(string sMinutes)
        {
            int iPerfectSeconds = 0;
            try
            {

                if (sMinutes.Trim().Length == 0)
                    sMinutes = "0";


                if (sMinutes.Contains("."))
                {
                    string[] sValue = sMinutes.Split('.');
                    iPerfectSeconds = Convert.ToInt32(sValue[0]) * 60 + Convert.ToInt32(sValue[1]);
                }
                else
                {
                    iPerfectSeconds = Convert.ToInt32(sMinutes) * 60;
                }
                return iPerfectSeconds;
            }
            catch (Exception ex)
            {
                throw new Exception("Error in converting minutes to seconds " + Environment.NewLine + ex.ToString());
            }
        }
        public int UpdateLines ()
        {
            try
            {   
                
                int result = 0;
                result = dgvGetPullData.Rows.Count;
                result = dgvGetPullData.Columns.Count;
          

                result = 0;
                foreach (DataGridViewRow dr in dgvGetPullData.Rows)
                {
                    string minutes = string.Empty;
                    string Job_ID = string.Empty;
                    object oDuration = null;
                    int WorkType = 0;
                    decimal Linecount = 0;
                    int AccountID= 0;
                    int ProductionID = 0;
                    string submission_time=string.Empty;
                    DateTime Dsubmission_time=DateTime.Now;
                    dr.DefaultCellStyle.BackColor = Color.SkyBlue;
                    dgvGetPullData.CurrentCell = dr.Cells[0];
                    foreach (DataGridViewColumn dc in dgvGetPullData.Columns)
                    {
                        try
                        {
                            if (dc.HeaderText == "Duration")
                            {
                                int count = CountStringOccurrences(dr.Cells[dc.Index].Value.ToString(), ":");
                                if (count > 1)
                                    oDuration = (DateTime.Parse(dr.Cells[dc.Index].Value.ToString()).Hour * 60 * 60) + (DateTime.Parse(dr.Cells[dc.Index].Value.ToString()).Minute * 60) + DateTime.Parse(dr.Cells[dc.Index].Value.ToString()).Second;
                                else
                                {
                                    if (dr.Cells[dc.Index].Value.ToString().Length == 3)
                                        minutes = dr.Cells[dc.Index].Value.ToString().TrimStart();
                                    else if (dr.Cells[dc.Index].Value.ToString().Replace(":", ".").Length > 5)
                                        minutes = dr.Cells[dc.Index].Value.ToString().Replace(":", ".").Substring(0, 5);
                                    else
                                        minutes = dr.Cells[dc.Index].Value.ToString();

                                    if(minutes.Substring(0, 1).Equals(":"))
                                        oDuration = minutes.Replace(":","");
                                    else
                                    oDuration = ConvertToSeconds(minutes.Replace(':','.'));
                                }

                                    
                               
                            }

                            if (dc.HeaderText == "Dictation ID")
                            {
                                Job_ID = dr.Cells[dc.Index].Value.ToString();
                            }
                            if (dc.HeaderText == "Work Type")
                            {
                                if ((dr.Cells[dc.Index].Value.ToString() == "Cancelled Dictation (992)") || (dr.Cells[dc.Index].Value.ToString() == "No Dictation (999)"))
                                    WorkType = 1;
                                else
                                    WorkType = 0;

                            }

                            if (dc.HeaderText == "Net LC")
                            {
                                Linecount =Convert.ToDecimal(dr.Cells[dc.Index].Value);
                            }
                            if (dc.HeaderText == "eScriptionist")
                            {
                                if (dr.Cells[dc.Index].Value.ToString().Split(',').GetValue(1).ToString().Contains(" RND"))
                                    ProductionID = Convert.ToInt32(dr.Cells[dc.Index].Value.ToString().Split(',').GetValue(1).ToString().Replace(" RND", ""));
                            }
                            if (dc.HeaderText == "Transcription Date")
                            {
                                submission_time = dr.Cells[dc.Index].Value.ToString().Replace(" EDT","");
                                submission_time = submission_time.ToString().Replace(" EST", "");
                                submission_time = submission_time.ToString().Replace(" PDT", "");
                                submission_time = submission_time.ToString().Replace(" PST", "");
                                submission_time = submission_time.ToString().Replace(" CDT", "");
                                Dsubmission_time = Convert.ToDateTime(submission_time.ToString().Replace(" CST", ""));
                            }

                            if (dc.HeaderText == "Entity")
                            {
                                if ((dr.Cells[dc.Index].Value.ToString() == "Detroit Receiving") || (dr.Cells[dc.Index].Value.ToString() == "Harper University") || (dr.Cells[dc.Index].Value.ToString() == "Sinai-Grace Hospital") || (dr.Cells[dc.Index].Value.ToString() == "Hutzel Hospital"))
                                    AccountID = 314;
                                else if ((dr.Cells[dc.Index].Value.ToString() == "COCMMC - Memorial") || (dr.Cells[dc.Index].Value.ToString() == "COCPOP - Pasadena") || (dr.Cells[dc.Index].Value.ToString() == "COCTAC - Tampa Comm"))
                                    AccountID = 607;
                                else if (dr.Cells[dc.Index].Value.ToString() == "Centinela Hospital")
                                    AccountID = 602;
                                else if ((dr.Cells[dc.Index].Value.ToString() == "Willowbrook - AWBW") || (dr.Cells[dc.Index].Value.ToString() == "Main Methodist - AMMM") || (dr.Cells[dc.Index].Value.ToString() == "Sugar Land - ASLL"))
                                    AccountID = 601;
                                else if ((dr.Cells[dc.Index].Value.ToString() == "Riverside"))
                                    AccountID = 141;
                                else
                                    AccountID = 51;
                            }
                            
                        }
                        catch (Exception ex)
                        {
                            result = -1;
                        }
                    }

                    result = BusinessLogic.ProcessEscriptionlines(Job_ID, Convert.ToInt32(oDuration), Linecount, AccountID, WorkType, Dsubmission_time, ProductionID);
                    Application.DoEvents();
                    showStatus(Job_ID + "File Processing.....", true);
                }
                return result;
                showStatus("Done.", true);
            }
            catch(Exception ex)
            {
                MessageBox.Show("You are trying to process an invaild Excel. Contact Software.");
            return -1;
            }

        }
        public DataTable CsvFileToDatatable(string path, bool IsFirstRowHeader)
        {
            string header = "No";
            string sql = string.Empty;
            DataTable dataTable = null;
            string pathOnly = string.Empty;
            string fileName = string.Empty;
            try
            {
                pathOnly = Path.GetDirectoryName(path);
                fileName = Path.GetFileName(path);
                sql = @"SELECT * FROM [" + fileName + "]";
                if (IsFirstRowHeader)
                {
                    header = "Yes";
                }
                using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly + ";Extended Properties=\"Text;HDR=" + header + "\""))
                {
                    using (OleDbCommand command = new OleDbCommand(sql, connection))
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            dataTable = new DataTable();
                            dataTable.TableName = fileName.ToString();
                            dataTable.Locale = CultureInfo.CurrentCulture;
                            adapter.Fill(dataTable);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                showStatus(ex.Message.ToString(), false);
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
            finally
            {
            }
            return dataTable;
        }

        public DataTable ExcelFileToDatatable(string path, bool IsFirstRowHeader)
        {
            string header = "No";
            string sql = string.Empty;
            DataTable dataTable = null;
            string pathOnly = string.Empty;
            string fileName = string.Empty;
            try
            {
                pathOnly = Path.GetDirectoryName(path);
                fileName = Path.GetFileName(path);
                string SheetName ="";
                
                if (IsFirstRowHeader)
                {
                    header = "Yes";
                }


                using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=" + header + ";IMEX=1\""))
                {
                    connection.Open();
                    DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (dr["Table_name"].ToString().Contains("Sheet"))
                            continue;
                        else
                            SheetName = dr["Table_name"].ToString();
                    }
                    sql = @"SELECT * FROM [" + SheetName + "]";
                    using (OleDbCommand command = new OleDbCommand(sql, connection))
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            dataTable = new DataTable();
                            dataTable.TableName = SheetName;
                            dataTable.Locale = CultureInfo.CurrentCulture;
                            adapter.Fill(dataTable);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                showStatus(ex.Message.ToString(), false);
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
            finally
            {
            }
            return dataTable;
        }        
        /// To update the content in dataset to List View
        public void LoadListView(DataSet dsMain) 
            {
               //collection to hold files whihc are not processed.
               FileWhichHitException = new ArrayList();
                try 
                { 
                        showStatus("Loading files...", true);
                        int iRowNumber = 0;
                                          
                        foreach (DataTable dtTable in dsMain.Tables)
                        {
                            //Try-Catch inside foreach to continue the loop even after hitting an expection
                            //Reason Added: if the required column is not in excel/csv file then it vil throw error and stop processing other files in directory. 
                            try
                            {
                                    Application.DoEvents();                                              
                                    System.Data.DataTable uniqueCols;
                                    if (dtTable.Columns.Contains("WT (Text)") && dtTable.Columns.Contains("Field 1 (Text)"))
                                    {
                                        uniqueCols = dtTable.DefaultView.ToTable(true, "Job Id", "Target Time", "Len (MM:SS)", "Flags", "Field 1 (Text)", "WT (Text)", "D# Site", "Date D# Closed");
                                        if (uniqueCols.Rows.Count > 0)
                                        {
                                            try
                                            {

                                                lsvOnlineTat.BeginUpdate();
                                                foreach (DataRow drow in uniqueCols.Rows)
                                                {
                                                    
                                                    iRowNumber++; 
                                                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                                                    lvi.SubItems.Add(drow["Job Id"].ToString());
                                                    lvi.SubItems.Add(drow["Target Time"].ToString());
                                                    lvi.SubItems.Add(drow["Len (MM:SS)"].ToString());
                                                    lvi.SubItems.Add(drow["Flags"].ToString());
                                                    if ((drow["Field 1 (Text)"].ToString().Length <= 0) && (drow["WT (Text)"].ToString().Length <= 0))
                                                        lvi.SubItems.Add(drow["Field 1 (Text)"].ToString());
                                                    else if ((drow["Field 1 (Text)"].ToString().Length <= 0) && (drow["WT (Text)"].ToString().Length > 0))
                                                        lvi.SubItems.Add(drow["WT (Text)"].ToString());
                                                    else
                                                        lvi.SubItems.Add(drow["Field 1 (Text)"].ToString());
                                                    lvi.SubItems.Add(drow["D# Site"].ToString());
                                                    lvi.SubItems.Add(drow["Date D# Closed"].ToString());
                                                    lsvOnlineTat.Items.Add(lvi);
                                                    
                                                }
                                            }

                                            catch { }
                                            finally { lsvOnlineTat.EndUpdate(); }
                                        }
                                    }
                                    //Excel Column which has "Field 1 (Text)"
                                    else if (dtTable.Columns.Contains("Field 1 (Text)"))
                                    {
                                        uniqueCols = dtTable.DefaultView.ToTable(true, "Job Id", "Target Time", "Len (MM:SS)", "Flags", "Field 1 (Text)", "D# Site", "Date D# Closed");
                                        if (uniqueCols.Rows.Count > 0)
                                        {
                                            try
                                            {

                                                lsvOnlineTat.BeginUpdate();
                                                foreach (DataRow drow in uniqueCols.Rows)
                                                {
                                                    iRowNumber++; 
                                                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                                                    lvi.SubItems.Add(drow["Job Id"].ToString());
                                                    lvi.SubItems.Add(drow["Target Time"].ToString());
                                                    lvi.SubItems.Add(drow["Len (MM:SS)"].ToString());
                                                    lvi.SubItems.Add(drow["Flags"].ToString());
                                                    lvi.SubItems.Add(drow["Field 1 (Text)"].ToString());
                                                    lvi.SubItems.Add(drow["D# Site"].ToString());
                                                    lvi.SubItems.Add(drow["Date D# Closed"].ToString());
                                                    lsvOnlineTat.Items.Add(lvi);
                                                    

                                                }
                                            }
                                            catch { }
                                            finally { lsvOnlineTat.EndUpdate(); }
                                        }
                                    }
                                    //Column which has "Split Job Link"
                                    else if (dtTable.Columns.Contains("Split Job Link"))
                                    {

                                        uniqueCols = dtTable.DefaultView.ToTable(true, "Job Id", "Target Time", "Len (MM:SS)", "Flags", "T# Site (Text)", "Date D# Closed", "Split Job Link");
                                        if (uniqueCols.Rows.Count > 0)
                                        {
                                            try
                                            {

                                                lsvOnlineTat.BeginUpdate();
                                                foreach (DataRow drow in uniqueCols.Rows)
                                                {
                                                    iRowNumber++;
                                                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                                                    lvi.SubItems.Add(drow["Job Id"].ToString());
                                                    lvi.SubItems.Add(drow["Target Time"].ToString());
                                                    lvi.SubItems.Add(drow["Len (MM:SS)"].ToString());
                                                    lvi.SubItems.Add(drow["Flags"].ToString());
                                                    lvi.SubItems.Add(drow["Split Job Link"].ToString());
                                                    lvi.SubItems.Add("15");
                                                    lvi.SubItems.Add(drow["Date D# Closed"].ToString());
                                                    lsvOnlineTat.Items.Add(lvi);

                                                }
                                            }
                                            catch { }
                                            finally { lsvOnlineTat.EndUpdate(); }
                                        }
                                    }
                                    //Column which has "Date Closed field"
                                    else if (dtTable.Columns.Contains("Date D# Closed"))
                                    {

                                        uniqueCols = dtTable.DefaultView.ToTable(true, "Job Id", "Target Time", "Len (MM:SS)", "Flags",  "D# Site", "Date D# Closed");
                                        if (uniqueCols.Rows.Count > 0)
                                        {
                                            try
                                            {

                                                lsvOnlineTat.BeginUpdate();
                                                foreach (DataRow drow in uniqueCols.Rows)
                                                {
                                                    iRowNumber++;
                                                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                                                    lvi.SubItems.Add(drow["Job Id"].ToString());
                                                    lvi.SubItems.Add(drow["Target Time"].ToString());
                                                    lvi.SubItems.Add(drow["Len (MM:SS)"].ToString());
                                                    lvi.SubItems.Add(drow["Flags"].ToString());
                                                    lvi.SubItems.Add(string.Empty);
                                                    lvi.SubItems.Add(drow["D# Site"].ToString());
                                                    lvi.SubItems.Add(drow["Date D# Closed"].ToString());
                                                    lsvOnlineTat.Items.Add(lvi);

                                                }
                                            }
                                            catch { }
                                            finally { lsvOnlineTat.EndUpdate(); }
                                        }
                                    }
                                 

                                    else 
                                    {
                                        uniqueCols = dtTable.DefaultView.ToTable(true, "Job Id", "Target Time", "Len (MM:SS)", "Flags", "WT (Text)", "D# Site", "Date D# Created");
                                        if (uniqueCols.Rows.Count > 0)
                                        {
                                            try
                                            {
                                                lsvOnlineTat.BeginUpdate();
                                                foreach (DataRow drow in uniqueCols.Rows)
                                                {
                                                    
                                                    iRowNumber++; 
                                                    ListViewItem lvi = new ListViewItem(iRowNumber.ToString());
                                                    lvi.SubItems.Add(drow["Job Id"].ToString());
                                                    lvi.SubItems.Add(drow["Target Time"].ToString());
                                                    lvi.SubItems.Add(drow["Len (MM:SS)"].ToString());
                                                    lvi.SubItems.Add(drow["Flags"].ToString());
                                                    
                                                    if (drow["WT (Text)"].ToString().Length > 0)
                                                        lvi.SubItems.Add(drow["WT (Text)"].ToString());
                                                    else
                                                        lvi.SubItems.Add(drow["--Nil Work Type Texting--"].ToString());

                                                    lvi.SubItems.Add(drow["D# Site"].ToString());
                                                    lvi.SubItems.Add(drow["Date D# Created"].ToString());
                                                   lsvOnlineTat.Items.Add(lvi);
                                                  
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                                            }
                                            finally { lsvOnlineTat.EndUpdate(); }
                                        }
                                    }
                                    Application.DoEvents();

                                }
                         catch(Exception ex)
                        {

                            if (ex.ToString().Contains("does not belong to"))
                             //ex.Message.Split('\'')[3]  --> this thing will add file name in array.
                                FileWhichHitException.Add(ex.Message.Split('\'')[3]);
                            ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");

                         }
                        }
                        
                        showStatus("Completed loading files...", true);
                        // To process the content in the List View to Data base
                        UpdateDB();
                }
                catch (Exception ex)
                {
                    ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                }
                finally
                {
                    
                }
            }
    
        /// To process the content in the List View to Data base
        private void UpdateDB()
        {
            try
            {
                showStatus("Begin processing the files to database...", true);
                foreach (ListViewItem item in lsvOnlineTat.Items)
                {
                    try
                    {
                        Application.DoEvents();
                        string SJob_Id = string.Empty;
                        DateTime STarget_Time;
                        string SLength = string.Empty;
                        string SVR_Flag = string.Empty;
                        string S_Flag = string.Empty;
                        string SWorkType = string.Empty;
                        string SAccount = string.Empty;
                        DateTime SFile_Date;


                        SJob_Id = item.SubItems[1].Text;
                        STarget_Time = Convert.ToDateTime(item.SubItems[2].Text);
                        SLength = item.SubItems[3].Text;
                        S_Flag = item.SubItems[4].Text;
                        SWorkType = item.SubItems[5].Text;
                        SAccount = item.SubItems[6].Text;
                        SFile_Date = Convert.ToDateTime(item.SubItems[7].Text);



                        DateTime DtTarget_Time = STarget_Time.AddHours(9);
                        DateTime DtSource_Time = SFile_Date.AddHours(9);


                        if (S_Flag.Contains("XSD"))
                            SVR_Flag = "1";
                        else
                            SVR_Flag = "0";


                        float fLength = 0.0F;
                        float.TryParse(DateTime.Parse(SLength).Hour + "." + DateTime.Parse(SLength).Minute, out fLength);

                        showStatus("Processing " + SJob_Id + " into database...", true);
                        //Insert
                        if (SAccount == "15")
                        {
                            if (SWorkType != "0")
                                BusinessLogic.ProcessToDataBase(SJob_Id, DtTarget_Time, fLength, SVR_Flag, S_Flag, SWorkType, SAccount, DtSource_Time, 1);
                            else
                                BusinessLogic.ProcessToDataBase(SJob_Id, DtTarget_Time, fLength, SVR_Flag, S_Flag, SWorkType, SAccount, DtSource_Time, 0);
                        }
                        else
                            BusinessLogic.ProcessToDataBase(SJob_Id, DtTarget_Time, fLength, SVR_Flag, S_Flag, SWorkType, SAccount, DtSource_Time, 0);

                        item.BackColor = Color.Green;
                        item.ForeColor = Color.White;
                        item.EnsureVisible();

                    }
                    catch (Exception ex)
                    {
                        item.BackColor = Color.Red;
                        item.ForeColor = Color.White;
                        item.EnsureVisible();
                        ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                    }
                }

                //To get all the names in a arraylist
                string SFileListFromArrayList = string.Empty;
                for (int i = 0; i < FileWhichHitException.Count; i++)
                    SFileListFromArrayList = SFileListFromArrayList + FileWhichHitException[i].ToString() + " , ";
                if (FileWhichHitException.Count > 0)
                {
                    showStatus("Files which are not processed " + SFileListFromArrayList + "", false);
                    //To maintain a list of unprocesed files.
                    ExceptionHandler.HandleException("Files which are not processed " + SFileListFromArrayList + "", Environment.MachineName, Environment.UserName, "");
                }
                else if (FileWhichHitException.Count <= 0)
                    showStatus("Finished processing the files in directory.", true);

            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
            finally
            {
                DeleteSourceFile(); 
            }

        }
        private void DeleteSourceFile()
        {
            try
            {
                dsMain = new DataSet();
                string[] files = System.IO.Directory.GetFiles(sourcePath);
                foreach (string file in files)
                {
                    string fileName = Path.GetFileName(file);
                    string sourceFile = Path.Combine(sourcePath, fileName);
                    File.Delete(sourceFile);
                }
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }

        
       #endregion


        //----------------------------------------------------Hangup Handler-----------------------------------
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                errorProvider1.Clear();
                showStatus("Ready.", true, "HanupUp");
                lsvContent.Items.Clear();
                lblHangUpAccountName.Text = "Account Name";
                lbl216FileCount.Text = "0";
                lsvHanupFiles.Items.Clear();
                lblHanupFileCount.Text = "0";
                lsv216Files.Items.Clear();

                if (txtFilePath.Text.Length == 0)
                {
                    if (BrowseFileDialog.ShowDialog() == DialogResult.OK)

                    {
                        txtFilePath.Text = BrowseFileDialog.FileName;
                        FileExtension = Path.GetExtension("" + txtFilePath.Text + "");
                    }
                    else
                    { 
                        
                    }
                }
                else if (txtFilePath.Text.Length > 0)
                {
                    txtFilePath.Text = "";
                    txtFilePath.Clear();
                }
            }
            catch(Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }
        private void btnSubmit_Click(object sender, EventArgs e)
        {
                
            showStatus("Loading...", true, "");
            try
            {  
               //Check if the email id is empty or valid
                if (string.IsNullOrEmpty(txtFilePath.Text))
                {
                    errorProvider1.SetError(txtFilePath, "Invalid file.");
                    showStatus("Invalid file.", true, "");
                    return;
                }
                else
                {
                    Control.CheckForIllegalCrossThreadCalls = false;
                    Thread oThread = new Thread(new ThreadStart(LoadExcelValuesIntoDataSet));
                    oThread.Start();
                }

               
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }
        //1. To get excel content into a seperate dataset
        public void LoadExcelValuesIntoDataSet()
        {
            try
            {
                string Excelpath = txtFilePath.Text.ToString();
                dsTransHanupExcelContent = new DataSet();
                DataTable dtTranHanupTemp;
                string FileName = Path.GetFileName(Excelpath);
                string Extension = Path.GetExtension(Excelpath);
                string FilePath = Path.GetFullPath(Excelpath);

                string conStr = "";
                switch (Extension)
                {
                    case ".xls": //Excel 97-03
                        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ Excelpath + ";Extended Properties='Excel 8.0;HDR=Yes'";
                            
                        break;
                    case ".xlsx": //Excel 07
                        conStr = "";
                        break;
                }
              
                conStr = String.Format(conStr, FilePath, 1);
                OleDbConnection connExcel = new OleDbConnection(conStr);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                DataTable dtExcelContent = new DataTable();
                cmdExcel.Connection = connExcel;
                connExcel.Open();
                DataTable dtExcelSchema = null;                
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                //Iasis and MKMG will be having only one sheet but Emts will have number of sheets so to process everything in same dataset.
                int iSheetCount = dtExcelSchema.Rows.Count;

                for (int i = 1; i <= iSheetCount; i++)
                {
                    dtTranHanupTemp = new DataTable();
                    dtTranHanupTemp.TableName = dtExcelSchema.Rows[i - 1]["TABLE_NAME"].ToString();
                    //Get the name Sheets
                    string SheetName = dtExcelSchema.Rows[i-1]["TABLE_NAME"].ToString();
                    connExcel.Close();
                    //Read Data from First Sheet
                    connExcel.Open();
                    cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                    oda.SelectCommand = cmdExcel;
                    oda.Fill(dtTranHanupTemp);
                    dsTransHanupExcelContent.Tables.Add(dtTranHanupTemp);
                    connExcel.Close();
                }


                //Pass collected database files to listview function
                LoadListViewTransHangUp(dsTransHanupExcelContent);
                
            }
            catch (Exception ex)
            {
                showStatus(ex.Message.ToString(), false, "");
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
          
        }
        //2. Load the content in dataset to Listview
        public void LoadListViewTransHangUp(DataSet dsExcelContent)
        {

            try
            {
                DataTable dtExcelContent = dsExcelContent.Tables[0];
                TotalRecordsInDtable = dtExcelContent.Rows.Count;
                TotalColumnsInDtable = dtExcelContent.Columns.Count;    
               

                iProcessedFiles = 0;

                showStatus("Loading datas from excel...", true, "");
                if (TotalRecordsInDtable <= 0)
                {
                    showStatus("No datas to be processed.", true, "");
                    return;
                }
                else
                {
                    //Total no of columns in IASIS and MKMG is  less than 30
                    if (dtExcelContent.Columns.Count < 30)
                    {
                        //FOR IASIS and MKMG REPORTS ONLY
                        //Stored account name is displayed in Label.
                        SITESpecific = Convert.ToString(dtExcelContent.Rows[2]["F4"]);

                        if (dtExcelContent.Rows[0]["F4"].ToString().Contains("Mount Kisko"))
                            lblHangUpAccountName.Text = AccountName = "MKMG";
                        else if (dtExcelContent.Rows[0]["F4"].ToString().Contains("IASIS"))
                            lblHangUpAccountName.Text = AccountName = "IASIS";
                        else
                        {   showStatus("You are trying to process a invalid excel.", true, "");
                            return;                         
                        }

                       
                        lsvContent.BeginUpdate();
                        lsvContent.Items.Clear();
                        foreach (DataRow drow in dtExcelContent.Select("F6 is not null"))
                        {
                            
                            if (drow["F3"].ToString() == "" || drow["F3"].ToString() == "Worktype")
                            {
                                continue;
                            }
                            string sVfileId = drow["F3"].ToString();//  F10
                            string sMtId = drow["F13"].ToString();//    F4
                            DateTime dtsubmission = new DateTime(Convert.ToDateTime(drow["F14"]).Year, Convert.ToDateTime(drow["F14"]).Month, Convert.ToDateTime(drow["F14"]).Day, Convert.ToInt32(drow["F15"]), Convert.ToInt32(drow["F16"]), 0);

                            string Production_id = drow["F13"].ToString().Split('(').GetValue(1).ToString().Replace(")","");
                            ListViewItem lvi = new ListViewItem();
                            lvi.Text = drow["F4"].ToString(); //        WORK TYPE
                            lvi.SubItems.Add(drow["F13"].ToString());// TRANSCRIPTIONIST
                            lvi.SubItems.Add(drow["F3"].ToString());//  JOB ID
                            lvi.SubItems.Add(drow["F10"].ToString());// TYPE
                            lvi.SubItems.Add(drow["F5"].ToString());//  DICTATION MINUTES
                            lvi.SubItems.Add(drow["F6"].ToString());//  CUSTOMER LINES                            
                            lvi.SubItems.Add(string.Empty);
                            lvi.SubItems.Add(dtsubmission.ToString());//  SUBMISSION TIME
                            lvi.SubItems.Add(Production_id.ToString());//  SUBMISSION TIME
                            lvi.BackColor = Color.Green;
                            lvi.ForeColor = Color.White;
                            lsvContent.Items.Add(lvi);
                            iProcessedFiles++;

                            //To update the status in the label lblStatusUpdate and show progress.
                            showStatus("Processing File : " + sVfileId + " Account Name : " + AccountName +
                                                     " User ID : " + sMtId.ToUpper() +
                                                    " Total Files : " + TotalRecordsInDtable.ToString() +
                                                     " Processed Files : " + iProcessedFiles.ToString(), true, "");
                            lsvContent.EndUpdate();
                        }
                    }
                    else
                    {
                        //---------------------------------------------------------------
                        //REMOVE ONCE COMPLETED
                        return;

                        //---------------------------------------------------------------
                        //FOR EMTS REPORTS ONLY

                        iProcessedFiles = 0;
                       
                        
                        lsvContent.BeginUpdate();
                        lsvContent.Items.Clear();

                        //To Store the account from the Excel. //Stored account name is displayed in Label.
                        lblHangUpAccountName.Text = SITESpecific = AccountName = "EMTS";

                        foreach (DataRow drow in dtExcelContent.Select("Line Count not in null"))
                                {
                                 //To get the minutes and to handle empty fields
                                 string minutes = string.Empty;
                                 if (drow["Duration"].ToString().Length == 3)
                                     minutes = drow["Duration"].ToString().TrimStart();
                                 else if (drow["Duration"].ToString().Length > 4)
                                     minutes = Convert.ToString(drow["Duration"].ToString().Substring(10, 5).TrimStart());
                                 else
                                     minutes = "";
                                                    

                                    
                                    if(minutes.Contains(':'))
                                    minutes = minutes.Replace(':','.');
                                    string sVfileId = drow["Dictation Id"].ToString();
                                    string sMtId = drow["eScriptionist"].ToString();
                                    ListViewItem lvi = new ListViewItem();
                                    lvi.Text = drow["Work Type"].ToString();            // WORK TYPE
                                    lvi.SubItems.Add(sMtId.ToString());                 // TRANSCRIPTIONIST
                                    lvi.SubItems.Add(drow["Dictation Id"].ToString());  //  JOB ID
                                    lvi.SubItems.Add(string.Empty);                     //  TYPE
                                    lvi.SubItems.Add(minutes);         //Duration
                                    lvi.SubItems.Add(drow["Line Count"].ToString());    //  CUSTOMER LINES
                                    lvi.SubItems.Add(string.Empty);
                                    lvi.BackColor = Color.Green;
                                    lvi.ForeColor = Color.White;                            
                                    lsvContent.Items.Add(lvi);
                                    iProcessedFiles++;

                                   //To update the status in the label lblStatusUpdate and show progress.
                                    showStatus("Processing File : " + sVfileId + " Account Name : " + AccountName +
                                                     " User ID : " + sMtId.ToUpper() +
                                                    " Total Files : " + TotalRecordsInDtable.ToString() +
                                                     " Processed Files : " + iProcessedFiles.ToString(), true, "");

                                    lsvContent.EndUpdate();

                                }
                        }
                }
               
            }
            catch (Exception ex)
            {
                showStatus(ex.Message.ToString(), false, "TransHanup");
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
           processListViewItemsToDataBase(lsvContent.Items.Count);
           
        }
        //3. To process the listview row by row to database
        public void processListViewItemsToDataBase(int iNumberOfItemInListView)
        {
            try
            {
                iProcessedFiles = 0;
              
                lbl216FileCount.Text = "0";
                string AccountID = string.Empty;
                showStatus("Processing to database...", true, "");
                foreach (ListViewItem lvi in lsvContent.Items)
                {
                    try
                    {   
                        //Getting site id for IASIS & MKMG

                        //------------------------------------------------------------PENDING------------
                            //emts,vanguard
                        //-------------------------------------------------------------------------------
                        switch (SITESpecific.ToLower())
                        {
                            case "emts": //EMTS
                                AccountID = "-2";
                                break;                          
                            case "mount kisco medical group - site 2": //MKMG
                                AccountID = "8201";
                                break;
                            case "mount kisco- site2-caremount-idc3": //MKMG
                                AccountID = "8201";
                                break;
                            case "mount kisco - site 2 - caremount": //MKMG
                                AccountID = "8201";
                                break;
                            case "iasis st lukes med ctr": //IASIS - St Lukes
                                AccountID = "2001";
                                break;
                            case "iasis behavioral health slmcbh": //IASIS - BHC
                                AccountID = "2004";
                                break;
                            case "iasis memorial hospital": //IASIS - Memorial
                                AccountID = "7807";
                                break;
                            case "iasis palms of pasadena": //IASIS - Pasadena
                                AccountID = "7809";
                                break;
                            case "iasis town & country hospital": //IASIS - Town and Country
                                AccountID = "7815";
                                break;
                            case "iasis glenwood regional med ctr": //IASIS - GLENWOOD
                                AccountID = "2022";
                                break;
                            case "iasis mountain vista med ctr": //IASIS - MountainVista
                                AccountID = "2020";
                                break;
                            case "iasis odessa regional med ctr": //IASIS - Odessa
                                AccountID = "2006";
                                break;
                            case "iasis medical ctr of se texas": //IASIS - SouthEastTexas
                                AccountID = "2010";
                                break;
                            case "iasis tempe st lukes med ctr": //IASIS - TEMPE
                                AccountID = "2002";
                                break;
                            case "iasis southwest general hospital": //IASIS - SouthWestGeneral
                                AccountID = "2079";
                                break;
                        }


                        //To set the colour at the time of processing
                        lvi.BackColor = Color.Orange;
                        lvi.ForeColor = Color.White;
                        lvi.SubItems[6].Text = "Processing...";
                        lvi.EnsureVisible();
                        iProcessedFiles++;
                        showStatus("Comparing File : " + lvi.SubItems[2].Text + " Account Name : " + AccountName + " User ID : " + lvi.SubItems[1].Text.ToUpper() + " Total Files : " + iNumberOfItemInListView.ToString() +
                                             " Processed Files : " + iProcessedFiles.ToString(),true,"");



                        //To compare it in database values one by one. 
                        //ALItems to store compared item from the DB.

                          ArrayList ALItems;
                          if (AccountName != "EMTS")
                          {
                              //----------------------------------------------IASIS NAD MKMG--------------------------------------
                              ALItems = compareLinesToDataBase(lvi.SubItems[2].Text, AccountID, lvi.SubItems[3].Text, lvi.SubItems[1].Text, lvi.SubItems[4].Text, lvi.Text, lvi.SubItems[5].Text,Convert.ToDateTime(lvi.SubItems[7].Text),Convert.ToInt32(lvi.SubItems[8].Text));
                              foreach (object Item in ALItems)
                              {
                                  LstCollection oCustomListViewItem = (LstCollection)Item;
                                  if (oCustomListViewItem.iType == (int)lstViewTpe.TwoOne6)
                                  {
                                      //216 here


                                      lsv216Files.Items.Add(oCustomListViewItem.oItem);
                                      lvi.BackColor = Color.Brown;
                                      lvi.ForeColor = Color.White;
                                      oCustomListViewItem.oItem.BackColor = Color.Brown;
                                      oCustomListViewItem.oItem.ForeColor = Color.White;
                                      lbl216FileCount.Text = Convert.ToString(lsv216Files.Items.Count);
                                  }
                                  else if (oCustomListViewItem.iType == (int)lstViewTpe.HangUP)
                                  {
                                      //Hang up here


                                      lvi.BackColor = Color.Violet;
                                      lvi.ForeColor = Color.White;
                                      lsvHanupFiles.Items.Add(oCustomListViewItem.oItem);
                                      oCustomListViewItem.oItem.BackColor = Color.Violet;
                                      oCustomListViewItem.oItem.ForeColor = Color.White;
                                      lblHanupFileCount.Text = Convert.ToString(lsvHanupFiles.Items.Count);
                                  }
                              }

                          }
                
                       

                          lvi.SubItems[6].Text = "Completed";

                    }
                    catch (Exception ex)
                    {
                        ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                    }

                }
                showStatus("Completed.Total Files processed" + iProcessedFiles.ToString() + " Total 216 files found " + lsv216Files.Items.Count.ToString() + " Total Hangup files found " + lsvHanupFiles.Items.Count.ToString(), true, "");                       
                
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
 
        }
        //3.a. To compare the listview item with  dataset and process them to database and the return value will be highlighted in listview
        public ArrayList compareLinesToDataBase(string Job_id, string Account_Id, string VR_status, string user_id, string Length, string WorkType, string Client_lines, DateTime dtsubmission, int production_id)
        {
            try
            {
                int option_flag = 0;
                int Is_split= 0;
                float cLines = 0.0F;
                string dbC_lines = string.Empty;

                 
                //VR Flag
                 int VR_Flag;
                 if (VR_status == "Edited")
                     VR_status = "Speech";

                 if (VR_status == "Non-Speech")
                     VR_Flag = 0;
                 else
                     VR_Flag = 1;

                //For Length
                 float len = 0.0F;
                 float.TryParse(Length, out len);
                //Split
                 if ((len == null) || (len <1))
                     Is_split=1;
                 else
                     Is_split = 0;

                //Work Type
                 if (WorkType != null)
                 {
                     if ((WorkType.ToLower().Contains("999")) ||(WorkType.ToLower().Contains("9999"))  || (WorkType.ToLower().Contains("8888")) || (WorkType.ToLower().Contains("dead job")) || (WorkType.ToLower().Contains("no dictation")) || (WorkType.ToLower().Contains("trash")))
                         option_flag = 1;
                     else
                         option_flag = 0;
                 }
                 else
                     option_flag = 0;

                //Client Lines
                if (Account_Id == "8201" || Account_Id == "1214")
                    dbC_lines = Client_lines;
                else
                {
                    float.TryParse(Client_lines, out cLines);
                    dbC_lines = (cLines / 65).ToString();

                }       
                    
                //Processing to servcie 
                DataSet tempds = BusinessLogic.ProcessToDataBaseTransHangUp(Job_id, Account_Id, VR_Flag, VR_status, len, WorkType, Convert.ToDouble(dbC_lines), dtsubmission, option_flag, Is_split, production_id);
                ArrayList oItmList = new ArrayList();
                //Check if trans field is updated for this file
                if (Convert.ToBoolean(tempds.Tables[0].Rows[0]["is_Trans_updated"]))
                {
                    LstCollection oItem216 = new LstCollection();
                    oItem216.iType = (int)lstViewTpe.TwoOne6;
                    if (Account_Id == "8201")
                    {
                        oItem216.oItem = LoadListViewFor216(Job_id, VR_status, user_id, cLines.ToString(), tempds.Tables[0].Rows[0]["file_lines"].ToString());   
                        oItmList.Add(oItem216);
                    }
                    else
                    {
                        oItem216.oItem = LoadListViewFor216(Job_id, VR_status, user_id, (cLines / 65).ToString(), tempds.Tables[0].Rows[0]["file_lines"].ToString());   
                        oItmList.Add(oItem216);
                    }

                }
                //Check if hangup field is updated for this file
                if (Convert.ToBoolean(tempds.Tables[0].Rows[0]["is_hangup_update"]))
                {
                    LstCollection oItemHangUp = new LstCollection();
                    oItemHangUp.iType = (int)lstViewTpe.HangUP;

                    if (Account_Id == "8201")
                    {
                        oItemHangUp.oItem = LoadListViewForHangUP(Job_id, VR_status, user_id, cLines.ToString(), tempds.Tables[0].Rows[0]["file_lines"].ToString());
                        oItmList.Add(oItemHangUp);
                    }
                    else
                    {
                        oItemHangUp.oItem = LoadListViewForHangUP(Job_id, VR_status, user_id, (cLines / 65).ToString(), tempds.Tables[0].Rows[0]["file_lines"].ToString());
                        oItmList.Add(oItemHangUp);
                    }

                }


                return oItmList;
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                return null;

            }
            finally
            {


            }

        }
        //3.a.a seperate listview to show 216 files
        public ListViewItem LoadListViewFor216(string Job_id, string VR_status, string user_id, string Client_lines, string rnd_lines)
        {
            try
            {
                // this.BeginUpdate();
                string DisplayUser = string.Empty;
                ListViewItem lvi = new ListViewItem();
                lvi.Text = Job_id;                
                if (VR_status == "1")
                {
                    DisplayUser = "SPEECH";

                }
                else
                {
                    DisplayUser = "NON-SPEECH";
                }
                lvi.SubItems.Add(DisplayUser);                
                lvi.SubItems.Add(user_id);
                lvi.SubItems.Add(Client_lines);
                lvi.SubItems.Add(rnd_lines);

                lvi.BackColor = Color.White;
                lvi.ForeColor = Color.Black;
                return lvi;

            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                return null;
            }
            finally
            {
                Application.DoEvents();
            }
        }
        //3.a.b seperate listview to show HangUp files
        public ListViewItem LoadListViewForHangUP(string Job_id, string VR_status,string user_id, string Client_lines, string rnd_lines)
        {
            try
            {
                string DisplayUser = string.Empty;
                ListViewItem lvi = new ListViewItem();
                lvi.Text = Job_id;

                
                if (VR_status == "1")
                {
                    DisplayUser = "SPEECH";

                }
                else
                {
                    DisplayUser = "NON-SPEECH";
                }
                lvi.SubItems.Add(DisplayUser);
                lvi.SubItems.Add(user_id);
                lvi.SubItems.Add(Client_lines);
                lvi.SubItems.Add(rnd_lines);

                lvi.BackColor = Color.White;
                lvi.ForeColor = Color.Black;

                return lvi;

            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
                return null;
            }
            finally
            {
                Application.DoEvents();
            }
        }
        //----------------------------------------------------End Hangup Handler-----------------------------------

        //----------------------------------------------------Master Data----------------------------------------------------------------------
    
      /*  private void btnMasterReportSubmit_Click(object sender, EventArgs e)
        {
            try 
            {
                showStatus("Loading...", true, 0);
                Thread myThread = new Thread(new ThreadStart(MasterSetDataToExcel));
                myThread.Start();
                myThread.IsBackground = true;


            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }

        public void FetchRecordFromDB()
        {
            try 
            {
            

            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }

        public void MasterSetDataToExcel()
        {

            try
            {  
                FromDate =  Convert.ToDateTime(dtPickerFrom.Value);
                ToDate = Convert.ToDateTime(dtPickerTo.Value);

                //Check if directory exists
                if (!Directory.Exists(MasterReport))
                    Directory.CreateDirectory(MasterReport);

                //To create folder in DATE
                var shortDate = String.Format("{0:m}", DateTime.Now);
                string subFolder= shortDate;
                string ReportPath = Path.Combine(MasterReport, subFolder);
                if (!Directory.Exists(ReportPath))
                    Directory.CreateDirectory(ReportPath);

                DataSet dsMasterDataContent = new DataSet();
                dsMasterDataContent = BusinessLogic.getMasterDataReport(FromDate, ToDate);
                DataTable dt= new DataTable();
                dt=dsMasterDataContent.Tables[0];
                CreateCSVFile(dt, ReportPath);
                

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Logon failure"))
                    showStatus(MasterReport + " is not accessible.", false,0);
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }

        public void CreateCSVFile(DataTable dtDataTablesList, string strFilePath)
        {
            try 
            {
                showStatus("Exporting to CSV...", true, 0);
                string sCsvFileName = Convert.ToString(Directory.GetFiles(strFilePath).Length + 1) + ".csv";
                strFilePath = Path.Combine(strFilePath, sCsvFileName);
                
                //Create the CSV file
                StreamWriter sw = new StreamWriter(strFilePath, false);
                //Write the headers.
                int iColCount = dtDataTablesList.Columns.Count;

                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write(dtDataTablesList.Columns[i]);
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);

                //Write all the rows.

                foreach (DataRow dr in dtDataTablesList.Rows)
                {
                    for (int i = 0; i < iColCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write(dr[i].ToString());
                        }
                        if (i < iColCount - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
                showStatus("Done exporting.", true, 0);
                showStatus("Report exported to  " + strFilePath, true, 0);
            
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
            
        }*/
        // -----------------------------------------------End Master Data----------------------------------------------------------------------

        private void BtnPaste_Click(object sender, EventArgs e)
        {
            PasteToGrid();
            UpdateLines();
            dgvGetPullData.Rows.Clear();
            dgvGetPullData.Columns.Clear();

        }

        private void PasteToGrid()
        {
            try
            {
                dgvGetPullData.DataSource = null;
                dgvGetPullData.AllowUserToAddRows = true;


                DataObject o = (DataObject)Clipboard.GetDataObject();
                int myRowIndex = dgvGetPullData.Rows.Count - 1;
                myRowIndex = 0;

                if (o.GetDataPresent(DataFormats.Text))
                {
                    if (dgvGetPullData.RowCount > 0)
                        dgvGetPullData.Rows.Clear();

                    if (dgvGetPullData.ColumnCount > 0)
                        dgvGetPullData.Columns.Clear();


                    bool columnsAdded = false;
                    string[] pastedRows = Regex.Split(o.GetData(DataFormats.Text).ToString().TrimEnd("\r\n".ToCharArray()), "\r\n");
                    foreach (string pastedRow in pastedRows)
                    {
                        //string[] pastedRowCells = pastedRow.Split(new char[] { '\t' });
                        string[] pastedRowCells = pastedRow.Split('\t');

                        if (!columnsAdded)
                        {
                            for (int i = 0; i < pastedRowCells.Length; i++)
                            {
                                dgvGetPullData.Columns.Add("col" + i, pastedRowCells[i]);
                            }
                            columnsAdded = true;
                            continue;
                        }
                        dgvGetPullData.Rows.Add(string.Empty);
                        for (int i = 0; i < pastedRowCells.Length; i++)
                        {
                            string sString = pastedRowCells[i].ToString();
                            dgvGetPullData.Rows[myRowIndex].Cells[i].Value = sString;
                        }
                        myRowIndex++;
                        dgvGetPullData.CurrentCell = dgvGetPullData.Rows[myRowIndex].Cells[0];
                        System.Windows.Forms.Application.DoEvents();

                    }
                }
                dgvGetPullData.AllowUserToAddRows = false;
                
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex.ToString(), Environment.UserName, Environment.MachineName, "");
            }
        }

        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlMain.SelectedTab.Text == "Report")
                LoadReport();
        }     
    }
}
