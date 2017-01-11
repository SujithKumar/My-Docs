using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace MasterData_EXE
{
    public partial class frmMasterData : Form
    {
        #region "Declaration"

        public frmMasterData()
        {
            InitializeComponent();
        }

        int iRowCount = 0;
        string sJobID = string.Empty;
        int iTranscriptionID = 0;
        DataSet _dsJob_NTS = new DataSet();
        DataSet _dsTouch_NTS = new DataSet();
        DataSet _dsInvoice_NTS = new DataSet();

        #endregion "Declaration"

        #region "Methods"

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

        public DataTable Clean_Dataset(DataSet i_dSourceTable)
        {
            DataTable dtGroup1 = new DataTable();
            dtGroup1.Columns.Add("voice_file_id").ToString();
            dtGroup1.Columns.Add("file_minutes").ToString();
            dtGroup1.Columns.Add("location_name").ToString();
            dtGroup1.Columns.Add("comby").ToString();
            dtGroup1.Columns.Add("file_date").ToString();
            dtGroup1.Columns.Add("transcription_id").ToString();
            dtGroup1.Columns.Add("Count").ToString();
            dtGroup1.Columns.Add("Del").ToString();

            try
            {
                //IDENTIFY DUPLICATES
                DataView dv = new DataView(i_dSourceTable.Tables[0]);

                //GETTING DISTINCT VALUES FOR GROUP COLUMN
                DataTable dtGroup = dv.ToTable(true, new string[] { "voice_file_id", "file_minutes", "location_name", "comby", "file_date", "transcription_id" });

                //ADDING COLUMN FOR THE ROW COUNT
                dtGroup.Columns.Add("Count", typeof(int));

                //LOOPING THRU DISTINCT VALUES FOR THE GROUP, COUNTING
                foreach (DataRow dr in dtGroup.Rows)
                {
                    dr["Count"] = i_dSourceTable.Tables[0].Compute("Count(" + "comby" + ")", "comby" + " = '" + dr["comby"] + "'");
                }

                //RETURNING GROUPED/COUNTED RESULT
                int iRowCountDup = 0;
                string sJob, sMinutes, sTrans, SCount, sDate, sFileMinutes, sLocationName;
                foreach (DataRow _drRow in dtGroup.Rows)
                {
                    sJob = _drRow["voice_file_id"].ToString();
                    sFileMinutes = _drRow["file_minutes"].ToString();
                    sLocationName = _drRow["location_name"].ToString();
                    sMinutes = _drRow["comby"].ToString();
                    sDate = _drRow["file_date"].ToString();
                    sTrans = _drRow["transcription_id"].ToString();
                    SCount = _drRow["Count"].ToString();

                    dtGroup1.Rows.Add(sJob);
                    dtGroup1.Rows[iRowCountDup]["file_minutes"] = sFileMinutes.ToString();
                    dtGroup1.Rows[iRowCountDup]["location_name"] = sLocationName.ToString();
                    dtGroup1.Rows[iRowCountDup]["comby"] = sMinutes.ToString();
                    dtGroup1.Rows[iRowCountDup]["file_date"] = sDate.ToString();
                    dtGroup1.Rows[iRowCountDup]["Count"] = SCount.ToString();
                    dtGroup1.Rows[iRowCountDup]["transcription_id"] = sTrans.ToString();
                    iRowCountDup++;
                }
                dtGroup1.DefaultView.Sort = "voice_file_id,file_date asc";
                dtGroup1 = dtGroup1.DefaultView.ToTable(true);
                return dtGroup1;
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
                ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
                return null;
            }
        }


        private void BindDataset_NTS()
        {
            DataTable dtNew = new DataTable();
            try
            {
                lbxMsb.Items.Clear();
                #region "Create Datatable"

                lbxMsb.Items.Add("Binding dataset...");

                dtNew.Columns.Add("voice_file_id").ToString();
                dtNew.Columns.Add("file_minutes").ToString();
                dtNew.Columns.Add("location_name").ToString();
                dtNew.Columns.Add("rnd_content").ToString();
                dtNew.Columns.Add("branch").ToString();

                dtNew.Columns.Add("file_lines").ToString();
                dtNew.Columns.Add("touch_one_status").ToString();
                dtNew.Columns.Add("touch_one_by").ToString();
                dtNew.Columns.Add("touch_one_lines").ToString();
                dtNew.Columns.Add("touch_one_date").ToString();

                dtNew.Columns.Add("touch_two_status").ToString();
                dtNew.Columns.Add("touch_two_by").ToString();
                dtNew.Columns.Add("touch_two_lines").ToString();
                dtNew.Columns.Add("touch_two_date").ToString();

                dtNew.Columns.Add("touch_three_status").ToString();
                dtNew.Columns.Add("touch_three_by").ToString();
                dtNew.Columns.Add("touch_three_lines").ToString();
                dtNew.Columns.Add("touch_three_date").ToString();

                dtNew.Columns.Add("touch_four_status").ToString();
                dtNew.Columns.Add("touch_four_by").ToString();
                dtNew.Columns.Add("touch_four_lines").ToString();
                dtNew.Columns.Add("touch_four_date").ToString();

                dtNew.Columns.Add("touch_five_status").ToString();
                dtNew.Columns.Add("touch_five_by").ToString();
                dtNew.Columns.Add("touch_five_lines").ToString();
                dtNew.Columns.Add("touch_five_date").ToString();

                dtNew.Columns.Add("expected_tat").ToString();
                dtNew.Columns.Add("tat_dsp").ToString();
                dtNew.Columns.Add("dsp_typed").ToString();
                dtNew.Columns.Add("ndsp_typed").ToString();
                dtNew.Columns.Add("dsp_edited").ToString();
                dtNew.Columns.Add("ndsp_edited").ToString();
                dtNew.Columns.Add("all_lines").ToString();
                dtNew.Columns.Add("client_content").ToString();
                dtNew.Columns.Add("tat_slabs").ToString();
                dtNew.Columns.Add("money_value").ToString();
                dtNew.Columns.Add("no_of_touches").ToString();
                dtNew.Columns.Add("salary_lines").ToString();
                dtNew.Columns.Add("belongs_to_month").ToString();
                dtNew.Columns.Add("belongs_to_year").ToString();
                dtNew.Columns.Add("imported_date").ToString();

                #endregion "Create Datatable"

                //lblMsg.Text = "Fetching Data...";
                lbxMsb.Items.Add("Fetching Data...");

                _dsJob_NTS = Class.ClsDA.GETJOBID(dtpDate.Value.Month, dtpDate.Value.Year);
                _dsTouch_NTS = Class.ClsDA.GETJOBID(dtpDate.Value.Month, dtpDate.Value.Year, null);
                _dsInvoice_NTS = Class.ClsDA.GET_INVOICE(dtpDate.Value.Month, dtpDate.Value.Year);

                DataTable _dtJob_NTS = new DataTable();

                //lblMsg.Text = "Transferring Data...";
                lbxMsb.Items.Add("Transferring Data...");

                foreach (DataColumn dcc in _dsJob_NTS.Tables[0].Columns)
                {
                    lbxMsb.Items.Add("Adding Columns" + dcc.ColumnName);
                    _dtJob_NTS.Columns.Add(dcc.ColumnName, dcc.DataType);
                }

                foreach (DataRow dcc in _dsJob_NTS.Tables[0].Rows)
                {
                    _dtJob_NTS.Rows.Add(dcc.ItemArray);
                }

                string iTra = string.Empty;
                int iRR = 0;
                string sActualTransID = string.Empty;
                string sReplaceTransID = string.Empty;
                string sStrayJobIDs = string.Empty;

                int iI = 0;
                try
                {
                    //TO MARK DUPLICATE ROWS WITH Y
                    //COMBY IS SORTED ASCENDING

                    for (iI = 0; iI < _dtJob_NTS.Rows.Count - 1; iI++)
                    {
                        lbxMsb.Items.Add("Unique :" + _dtJob_NTS.Rows[iI]["comby"].ToString());
                        if (iTra == _dtJob_NTS.Rows[iI]["comby"].ToString()) //FIRST ROW WILL NOT MATCH WITH EMPTY
                        {
                            _dtJob_NTS.Rows[iI]["Del"] = "Y"; //DUPLICATE COMBI ROW FOUND AND MARKED WITH Y
                            sActualTransID = _dtJob_NTS.Rows[iI - 1][4].ToString();
                            sReplaceTransID = _dtJob_NTS.Rows[iI][4].ToString();
                            //UPDATE ALL DUPLICATE TOUCHES WITH ACTUAL TRANACTION ID IN 
                            DataRow drTemp = null;
                            try
                            {
                                drTemp = _dsTouch_NTS.Tables[0].Select(" transcription_id = " + sReplaceTransID)[0];
                                drTemp["transcription_id"] = sActualTransID;
                            }
                            catch (System.IndexOutOfRangeException ex)
                            {
                                //MASTER ENTRY WITHOUT TOUCH ENTRY
                                //DO NOTHING
                                sStrayJobIDs = sStrayJobIDs + "," + _dtJob_NTS.Rows[iI][0].ToString();
                            }
                            iRR++;
                            lbxMsb.TopIndex = lbxMsb.Items.Count - 1;
                            lbxMsb.ForeColor = System.Drawing.Color.Tomato;
                        }
                        else
                        {
                            iTra = _dtJob_NTS.Rows[iI]["comby"].ToString(); //FIRST ROW WILL ALWAYS COME HERE
                            lbxMsb.TopIndex = lbxMsb.Items.Count - 1;
                            lbxMsb.ForeColor = System.Drawing.Color.Tomato;
                        }
                    }                    
                }
                catch (Exception ex)
                {
                    lblMsg.Text = ex.ToString();
                    MessageBox.Show(ex.ToString());
                }

                for (int i = _dtJob_NTS.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = _dtJob_NTS.Rows[i];
                    if (dr["Del"].ToString() == "Y")
                    {
                        lbxMsb.Items.Add("Deleted " + dr["voice_file_id"]).ToString();
                        dr.Delete();
                        lbxMsb.TopIndex = lbxMsb.Items.Count - 1;
                        lbxMsb.ForeColor = System.Drawing.Color.Teal;
                    }
                }

                lblMsg.Text = "Looping Data..";
                //iRowCount = 0;
                foreach (DataRow _drRow in _dtJob_NTS.Rows)
                {
                    lbxMsb.Items.Add("Adding transaction data to " + _drRow["voice_file_id"]).ToString();
                    sJobID = _drRow["voice_file_id"].ToString();
                    iTranscriptionID = Convert.ToInt32(_drRow["transcription_id"]);
                    string sMinutes = _drRow["file_minutes"].ToString();
                    string sLocation = _drRow["location_name"].ToString();
                    string sCurrentDate = System.DateTime.Now.ToString();

                    dtNew.Rows.Add(sJobID);
                    dtNew.Rows[iRowCount]["imported_date"] = Format_Date(sCurrentDate).ToString();
                    dtNew.Rows[iRowCount]["file_minutes"] = sMinutes.ToString();
                    dtNew.Rows[iRowCount]["location_name"] = sLocation.ToString();

                    string sFileLines = string.Empty;
                    string sFile_status = string.Empty;
                    string sTouchBy = string.Empty;
                    string sTouchLines = string.Empty;
                    string sTouchDate = string.Empty;
                    string sFilter = " transcription_id=" + iTranscriptionID + "";
                    string sRndContent = string.Empty;

                    int iRowsTrans = 0;
                    decimal dTotalSalaryLines = 0;
                    foreach (DataRow _dr in _dsTouch_NTS.Tables[0].Select(sFilter))
                    {
                        int iTouch = 0;
                        sFileLines = _dr["Rnd_lines"].ToString();
                        sRndContent = _dr["rnd_content"].ToString();
                        string sBranch = _dr["branch"].ToString();
                        dtNew.Rows[iRowCount]["file_lines"] = sFileLines.ToString();
                        dtNew.Rows[iRowCount]["rnd_content"] = sRndContent.ToString();
                        dtNew.Rows[iRowCount]["branch"] = sBranch.ToString();

                        sFile_status = _dr["file_status_description"].ToString();
                        sTouchBy = _dr["emp_name"].ToString();
                        sTouchLines = _dr["salary_lines"].ToString();
                        sTouchDate = _dr["submitted_time"].ToString();

                        if (iRowsTrans == 0)
                        {
                            iTouch = 1;
                            dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                            dtNew.Rows[iRowCount]["touch_one_status"] = sFile_status.ToString();
                            dtNew.Rows[iRowCount]["touch_one_by"] = sTouchBy.ToString();
                            dtNew.Rows[iRowCount]["touch_one_lines"] = sTouchLines.ToString();
                            dtNew.Rows[iRowCount]["touch_one_date"] = Format_Date(sTouchDate);
                            dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                            dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                        }
                        else if (iRowsTrans == 1)
                        {
                            iTouch = 2;
                            dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                            dtNew.Rows[iRowCount]["touch_two_status"] = sFile_status.ToString();
                            dtNew.Rows[iRowCount]["touch_two_by"] = sTouchBy.ToString();
                            dtNew.Rows[iRowCount]["touch_two_lines"] = sTouchLines.ToString();
                            dtNew.Rows[iRowCount]["touch_two_date"] = Format_Date(sTouchDate);
                            dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                            dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                        }
                        else if (iRowsTrans == 2)
                        {
                            iTouch = 3;
                            dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                            dtNew.Rows[iRowCount]["touch_three_status"] = sFile_status.ToString();
                            dtNew.Rows[iRowCount]["touch_three_by"] = sTouchBy.ToString();
                            dtNew.Rows[iRowCount]["touch_three_lines"] = sTouchLines.ToString();
                            dtNew.Rows[iRowCount]["touch_three_date"] = Format_Date(sTouchDate);
                            dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                            dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                        }
                        else if (iRowsTrans == 3)
                        {
                            iTouch = 4;
                            dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                            dtNew.Rows[iRowCount]["touch_four_status"] = sFile_status.ToString();
                            dtNew.Rows[iRowCount]["touch_four_by"] = sTouchBy.ToString();
                            dtNew.Rows[iRowCount]["touch_four_lines"] = sTouchLines.ToString();
                            dtNew.Rows[iRowCount]["touch_four_date"] = Format_Date(sTouchDate);
                            dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                            dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                        }
                        else if (iRowsTrans == 4)
                        {
                            iTouch = 5;
                            dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                            dtNew.Rows[iRowCount]["touch_five_status"] = sFile_status.ToString();
                            dtNew.Rows[iRowCount]["touch_five_by"] = sTouchBy.ToString();
                            dtNew.Rows[iRowCount]["touch_five_lines"] = sTouchLines.ToString();
                            dtNew.Rows[iRowCount]["touch_five_date"] = Format_Date(sTouchDate);
                            dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                            dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                        }
                        iRowsTrans++;
                        lbxMsb.TopIndex = lbxMsb.Items.Count - 1;
                        lbxMsb.ForeColor = System.Drawing.Color.Teal;
                    }

                    if ((_dsInvoice_NTS != null) && (_dsInvoice_NTS.Tables[0].Rows.Count > 1))
                    {
                        int iRowInvoice = 0;
                        string sFilter_I = " JOB_ID='" + sJobID + "'";
                        foreach (DataRow _drI in _dsInvoice_NTS.Tables[0].Select(sFilter_I.ToString()))
                        {

                            string sExpectedTat = string.Empty;
                            string sTatDsp = string.Empty;
                            string sDspTyped = string.Empty;
                            string sNdspTyped = string.Empty;
                            string sDspEdited = string.Empty;
                            string sNdspEdited = string.Empty;
                            string sAllLines = string.Empty;
                            string sClientType = string.Empty;
                            string sTatSlabs = string.Empty;
                            string sMoney = string.Empty;
                            string sBelongsToMonth = string.Empty;
                            string sBelongsToYear = string.Empty;

                            sExpectedTat = _drI["EXPECTED_TAT"].ToString();
                            sTatDsp = _drI["TAT"].ToString();
                            sDspTyped = _drI["DSP_TYPED"].ToString();
                            sNdspTyped = _drI["NONDSP_TYPED"].ToString();
                            sDspEdited = _drI["DSP_EDITED"].ToString();
                            sNdspEdited = _drI["NONDSP_EDITED"].ToString();
                            sAllLines = _drI["ALL_LINES"].ToString();
                            sClientType = _drI["TYPE"].ToString();
                            sTatSlabs = _drI["TAT_SLABS"].ToString();
                            sMoney = _drI["Money_value"].ToString();
                            sBelongsToMonth = _drI["belongs_to_month"].ToString();
                            sBelongsToYear = _drI["belongs_to_year"].ToString();


                            dtNew.Rows[iRowCount]["expected_tat"] = sExpectedTat.ToString();
                            dtNew.Rows[iRowCount]["tat_dsp"] = sTatDsp.ToString();
                            dtNew.Rows[iRowCount]["dsp_typed"] = sDspTyped.ToString();
                            dtNew.Rows[iRowCount]["ndsp_typed"] = sNdspTyped.ToString();
                            dtNew.Rows[iRowCount]["dsp_edited"] = sDspEdited.ToString();
                            dtNew.Rows[iRowCount]["ndsp_edited"] = sNdspEdited.ToString();
                            dtNew.Rows[iRowCount]["all_lines"] = sAllLines.ToString();
                            dtNew.Rows[iRowCount]["client_content"] = sClientType.ToString();
                            dtNew.Rows[iRowCount]["tat_slabs"] = sTatSlabs.ToString();
                            dtNew.Rows[iRowCount]["money_value"] = sMoney.ToString();
                            dtNew.Rows[iRowCount]["belongs_to_month"] = sBelongsToMonth.ToString();
                            dtNew.Rows[iRowCount]["belongs_to_year"] = sBelongsToYear.ToString();
                            dtNew.AcceptChanges();
                            iRowInvoice++;
                        }
                        lbxMsb.TopIndex = lbxMsb.Items.Count - 1;
                        lbxMsb.ForeColor = System.Drawing.Color.Teal;
                    }
                    iRowCount++;
                }
                string sFileDirectory = "c:\\MasterData\\";
                string sFileDirectoryOne = "c:\\MasterData\\DoNotOpen\\";
                string sFileName = "Final MasterData" + System.DateTime.Now.ToString("yyyyMMddhhmmss") + ".csv";
                string sFileNameNew = "Final MasterData" + dtpDate.Value.Month + ".csv";
                if (!Directory.Exists(sFileDirectory))
                {
                    Directory.CreateDirectory(sFileDirectory);
                    CreateCSVFile(dtNew, sFileDirectory + sFileName);

                    if (!Directory.Exists(sFileDirectoryOne))
                    {
                        Directory.CreateDirectory(sFileDirectoryOne);
                        CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                        lblMsg.Text = "Exported Data...";
                    }
                }
                else
                {
                    CreateCSVFile(dtNew, sFileDirectory + sFileName);

                    if (!Directory.Exists(sFileDirectoryOne))
                    {
                        Directory.CreateDirectory(sFileDirectoryOne);
                        CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                        lblMsg.Text = "Exported Data...";
                    }
                    else
                    {
                        File.Delete(sFileDirectoryOne + sFileNameNew);
                        CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                        lblMsg.Text = "Exported Data...";
                    }
                }
                //LoadListView(dtNew);
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
                ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
            }
            finally
            {
                Thread.Sleep(1000);
                LoadInsert(dtNew);
            }
        }

        public void CreateCSVFile(DataTable dt, string strFilePath)
        {
            try
            {
                StreamWriter sw = new StreamWriter(strFilePath, false);
                int iColCount = dt.Columns.Count;
                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write(dt.Columns[i]);
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
                foreach (DataRow dr in dt.Rows)
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
                lblMsg.Text = "CSV Exported...";
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
                ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
            }
        }

        private void LoadListView(DataTable dtList)
        {
            try
            {
                //#region "Columns"

                //lvMasterData.Columns.Add("#");
                //lvMasterData.Columns.Add("voice_file_id");
                //lvMasterData.Columns.Add("file_minutes");
                //lvMasterData.Columns.Add("location_name");
                //lvMasterData.Columns.Add("rnd_content");
                //lvMasterData.Columns.Add("file_lines");
                //lvMasterData.Columns.Add("touch_one_status");
                //lvMasterData.Columns.Add("touch_one_by");
                //lvMasterData.Columns.Add("touch_one_lines");
                //lvMasterData.Columns.Add("touch_one_date");
                //lvMasterData.Columns.Add("touch_two_status");
                //lvMasterData.Columns.Add("touch_two_by");
                //lvMasterData.Columns.Add("touch_two_lines");
                //lvMasterData.Columns.Add("touch_two_date");
                //lvMasterData.Columns.Add("touch_three_status");
                //lvMasterData.Columns.Add("touch_three_by");
                //lvMasterData.Columns.Add("touch_three_lines");
                //lvMasterData.Columns.Add("touch_three_date");
                //lvMasterData.Columns.Add("touch_four_status");
                //lvMasterData.Columns.Add("touch_four_by");
                //lvMasterData.Columns.Add("touch_four_lines");
                //lvMasterData.Columns.Add("touch_four_date");
                //lvMasterData.Columns.Add("touch_five_status");
                //lvMasterData.Columns.Add("touch_five_by");
                //lvMasterData.Columns.Add("touch_five_lines");
                //lvMasterData.Columns.Add("touch_five_date");
                //lvMasterData.Columns.Add("expected_tat");
                //lvMasterData.Columns.Add("tat_dsp");
                //lvMasterData.Columns.Add("dsp_typed");
                //lvMasterData.Columns.Add("ndsp_typed");
                //lvMasterData.Columns.Add("dsp_edited");
                //lvMasterData.Columns.Add("ndsp_edited");
                //lvMasterData.Columns.Add("all_lines");
                //lvMasterData.Columns.Add("client_content");

                //lvMasterData.Items.Clear();
                //int iRowCount = 0;

                //foreach (DataRow _drRow in dtList.Select())
                //    lvMasterData.Items.Add(new ListViewMasterData(_drRow, iRowCount++));

                //Reset_ListViewColumn(lvMasterData);

                //#endregion "Columns"
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
                ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
            }
        }

        private string Format_Date(string sDate)
        {
            try
            {
                string sDateFor = Convert.ToDateTime(sDate).ToString("yyyy-MM-dd HH:MM:ss");
                return sDateFor;
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
                ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
                return null;
            }
        }

        DataSet _dsJob_Clinic = new DataSet();
        DataSet _dsTouch_Clinic = new DataSet();
        DataSet _dsInvoice_Clinic = new DataSet();
        private void BindDataset_Clinics()
        {
            do
            {
                try
                {
                    #region "Create Datatable"

                    DataTable dtNew = new DataTable();
                    dtNew.Columns.Add("voice_file_id").ToString();
                    dtNew.Columns.Add("file_minutes").ToString();
                    dtNew.Columns.Add("location_name").ToString();
                    dtNew.Columns.Add("rnd_content").ToString();
                    dtNew.Columns.Add("branch").ToString();

                    dtNew.Columns.Add("file_lines").ToString();
                    dtNew.Columns.Add("touch_one_status").ToString();
                    dtNew.Columns.Add("touch_one_by").ToString();
                    dtNew.Columns.Add("touch_one_lines").ToString();
                    dtNew.Columns.Add("touch_one_date").ToString();

                    dtNew.Columns.Add("touch_two_status").ToString();
                    dtNew.Columns.Add("touch_two_by").ToString();
                    dtNew.Columns.Add("touch_two_lines").ToString();
                    dtNew.Columns.Add("touch_two_date").ToString();

                    dtNew.Columns.Add("touch_three_status").ToString();
                    dtNew.Columns.Add("touch_three_by").ToString();
                    dtNew.Columns.Add("touch_three_lines").ToString();
                    dtNew.Columns.Add("touch_three_date").ToString();

                    dtNew.Columns.Add("touch_four_status").ToString();
                    dtNew.Columns.Add("touch_four_by").ToString();
                    dtNew.Columns.Add("touch_four_lines").ToString();
                    dtNew.Columns.Add("touch_four_date").ToString();

                    dtNew.Columns.Add("touch_five_status").ToString();
                    dtNew.Columns.Add("touch_five_by").ToString();
                    dtNew.Columns.Add("touch_five_lines").ToString();
                    dtNew.Columns.Add("touch_five_date").ToString();

                    dtNew.Columns.Add("expected_tat").ToString();
                    dtNew.Columns.Add("tat_dsp").ToString();
                    dtNew.Columns.Add("dsp_typed").ToString();
                    dtNew.Columns.Add("ndsp_typed").ToString();
                    dtNew.Columns.Add("dsp_edited").ToString();
                    dtNew.Columns.Add("ndsp_edited").ToString();
                    dtNew.Columns.Add("all_lines").ToString();
                    dtNew.Columns.Add("client_content").ToString();
                    dtNew.Columns.Add("tat_slabs").ToString();
                    dtNew.Columns.Add("money_value").ToString();
                    dtNew.Columns.Add("no_of_touches").ToString();
                    dtNew.Columns.Add("salary_lines").ToString();
                    dtNew.Columns.Add("belongs_to_month").ToString();
                    dtNew.Columns.Add("belongs_to_year").ToString();
                    dtNew.Columns.Add("imported_date").ToString();

                    #endregion "Create Datatable"

                    lblMsg.Text = "Fetching Data...";

                    _dsJob_Clinic = Class.ClsDA.GETJOBID_Clinic(dtpDate.Value.Month, dtpDate.Value.Year);
                    _dsTouch_Clinic = Class.ClsDA.GETJOBID_Clinic(dtpDate.Value.Month, dtpDate.Value.Year, null);
                    _dsInvoice_Clinic = Class.ClsDA.GET_INVOICE(dtpDate.Value.Month, dtpDate.Value.Year);

                    DataTable _dtJob_Clinics = new DataTable();

                    lblMsg.Text = "Transferring Data...";

                    foreach (DataColumn dcc in _dsJob_Clinic.Tables[0].Columns)
                    {
                        _dtJob_Clinics.Columns.Add(dcc.ColumnName, dcc.DataType);
                    }

                    foreach (DataRow dcc in _dsJob_Clinic.Tables[0].Rows)
                    {
                        _dtJob_Clinics.Rows.Add(dcc.ItemArray);
                    }

                    string iTra = string.Empty;
                    int iRR = 0;
                    string sActualTransID = string.Empty;
                    string sReplaceTransID = string.Empty;
                    string sStrayJobIDs = string.Empty;

                    int iI = 0;
                    try
                    {
                        //TO MARK DUPLICATE ROWS WITH Y
                        //COMBY IS SORTED ASCENDING

                        for (iI = 0; iI < _dtJob_Clinics.Rows.Count - 1; iI++)
                        {
                            if (iTra == _dtJob_Clinics.Rows[iI]["comby"].ToString()) //FIRST ROW WILL NOT MATCH WITH EMPTY
                            {
                                _dtJob_Clinics.Rows[iI]["Del"] = "Y"; //DUPLICATE COMBI ROW FOUND AND MARKED WITH Y
                                sActualTransID = _dtJob_Clinics.Rows[iI - 1][4].ToString();
                                sReplaceTransID = _dtJob_Clinics.Rows[iI][4].ToString();
                                //UPDATE ALL DUPLICATE TOUCHES WITH ACTUAL TRANACTION ID IN 
                                DataRow drTemp = null;
                                try
                                {
                                    drTemp = _dsTouch_Clinic.Tables[0].Select(" transcription_id = " + sReplaceTransID)[0];
                                    drTemp["transcription_id"] = sActualTransID;
                                }
                                catch (System.IndexOutOfRangeException ex)
                                {
                                    //MASTER ENTRY WITHOUT TOUCH ENTRY
                                    //DO NOTHING
                                    sStrayJobIDs = sStrayJobIDs + "," + _dtJob_Clinics.Rows[iI][0].ToString();
                                }
                                iRR++;
                            }
                            else
                            {
                                iTra = _dtJob_Clinics.Rows[iI]["comby"].ToString(); //FIRST ROW WILL ALWAYS COME HERE
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lblMsg.Text = ex.ToString();
                        MessageBox.Show(ex.ToString());
                    }

                    for (int i = _dtJob_Clinics.Rows.Count - 1; i >= 0; i--)
                    {
                        DataRow dr = _dtJob_Clinics.Rows[i];
                        if (dr["Del"].ToString() == "Y")
                            dr.Delete();
                    }
                    lblMsg.Text = "Looping Data..";
                    iRowCount = 0;
                    foreach (DataRow _drRow in _dtJob_Clinics.Rows)
                    {
                        sJobID = _drRow["voice_file_id"].ToString();
                        iTranscriptionID = Convert.ToInt32(_drRow["transcription_id"]);
                        string sMinutes = _drRow["file_minutes"].ToString();
                        string sLocation = _drRow["location_name"].ToString();
                        string sCurrentDate = System.DateTime.Now.ToString();

                        dtNew.Rows.Add(sJobID);
                        dtNew.Rows[iRowCount]["imported_date"] = Format_Date(sCurrentDate).ToString();
                        dtNew.Rows[iRowCount]["file_minutes"] = sMinutes.ToString();
                        dtNew.Rows[iRowCount]["location_name"] = sLocation.ToString();

                        string sFileLines = string.Empty;
                        string sFile_status = string.Empty;
                        string sTouchBy = string.Empty;
                        string sTouchLines = string.Empty;
                        string sTouchDate = string.Empty;
                        string sFilter = " transcription_id=" + iTranscriptionID + "";
                        string sRndContent = string.Empty;

                        int iRowsTrans = 0;
                        decimal dTotalSalaryLines = 0;
                        foreach (DataRow _dr in _dsTouch_Clinic.Tables[0].Select(sFilter))
                        {
                            int iTouch = 0;
                            sFileLines = _dr["Rnd_lines"].ToString();
                            sRndContent = _dr["rnd_content"].ToString();
                            string sBranch = _dr["branch"].ToString();
                            dtNew.Rows[iRowCount]["file_lines"] = sFileLines.ToString();
                            dtNew.Rows[iRowCount]["rnd_content"] = sRndContent.ToString();
                            dtNew.Rows[iRowCount]["branch"] = sBranch.ToString();

                            sFile_status = _dr["file_status_description"].ToString();
                            sTouchBy = _dr["emp_name"].ToString();
                            sTouchLines = _dr["salary_lines"].ToString();
                            sTouchDate = _dr["submitted_time"].ToString();

                            if (iRowsTrans == 0)
                            {
                                iTouch = 1;
                                dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                                dtNew.Rows[iRowCount]["touch_one_status"] = sFile_status.ToString();
                                dtNew.Rows[iRowCount]["touch_one_by"] = sTouchBy.ToString();
                                dtNew.Rows[iRowCount]["touch_one_lines"] = sTouchLines.ToString();
                                dtNew.Rows[iRowCount]["touch_one_date"] = Format_Date(sTouchDate);
                                dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                                dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                            }
                            else if (iRowsTrans == 1)
                            {
                                iTouch = 2;
                                dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                                dtNew.Rows[iRowCount]["touch_two_status"] = sFile_status.ToString();
                                dtNew.Rows[iRowCount]["touch_two_by"] = sTouchBy.ToString();
                                dtNew.Rows[iRowCount]["touch_two_lines"] = sTouchLines.ToString();
                                dtNew.Rows[iRowCount]["touch_two_date"] = Format_Date(sTouchDate);
                                dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                                dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                            }
                            else if (iRowsTrans == 2)
                            {
                                iTouch = 3;
                                dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                                dtNew.Rows[iRowCount]["touch_three_status"] = sFile_status.ToString();
                                dtNew.Rows[iRowCount]["touch_three_by"] = sTouchBy.ToString();
                                dtNew.Rows[iRowCount]["touch_three_lines"] = sTouchLines.ToString();
                                dtNew.Rows[iRowCount]["touch_three_date"] = Format_Date(sTouchDate);
                                dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                                dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                            }
                            else if (iRowsTrans == 3)
                            {
                                iTouch = 4;
                                dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                                dtNew.Rows[iRowCount]["touch_four_status"] = sFile_status.ToString();
                                dtNew.Rows[iRowCount]["touch_four_by"] = sTouchBy.ToString();
                                dtNew.Rows[iRowCount]["touch_four_lines"] = sTouchLines.ToString();
                                dtNew.Rows[iRowCount]["touch_four_date"] = Format_Date(sTouchDate);
                                dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                                dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                            }
                            else if (iRowsTrans == 4)
                            {
                                iTouch = 5;
                                dTotalSalaryLines += Convert.ToDecimal(sTouchLines);
                                dtNew.Rows[iRowCount]["touch_five_status"] = sFile_status.ToString();
                                dtNew.Rows[iRowCount]["touch_five_by"] = sTouchBy.ToString();
                                dtNew.Rows[iRowCount]["touch_five_lines"] = sTouchLines.ToString();
                                dtNew.Rows[iRowCount]["touch_five_date"] = Format_Date(sTouchDate);
                                dtNew.Rows[iRowCount]["no_of_touches"] = iTouch.ToString();
                                dtNew.Rows[iRowCount]["salary_lines"] = dTotalSalaryLines.ToString();
                            }
                            iRowsTrans++;
                        }

                        int iRowInvoice = 0;
                        string sFilter_I = " JOB_ID='" + sJobID + "'";
                        foreach (DataRow _drI in _dsInvoice_Clinic.Tables[0].Select(sFilter_I.ToString()))
                        {
                            string sExpectedTat = string.Empty;
                            string sTatDsp = string.Empty;
                            string sDspTyped = string.Empty;
                            string sNdspTyped = string.Empty;
                            string sDspEdited = string.Empty;
                            string sNdspEdited = string.Empty;
                            string sAllLines = string.Empty;
                            string sClientType = string.Empty;
                            string sTatSlabs = string.Empty;
                            string sMoney = string.Empty;
                            string sBelongsToMonth = string.Empty;
                            string sBelongsToYear = string.Empty;

                            sExpectedTat = _drI["EXPECTED_TAT"].ToString();
                            sTatDsp = _drI["TAT_DSP"].ToString();
                            sDspTyped = _drI["TYPED"].ToString();
                            sNdspTyped = _drI["NONDSP_TYPED"].ToString();
                            sDspEdited = _drI["DSP_EDITED"].ToString();
                            sNdspEdited = _drI["NONDSP_EDITED"].ToString();
                            sAllLines = _drI["ALL_LINES"].ToString();
                            sClientType = _drI["TYPE"].ToString();
                            sTatSlabs = _drI["TAT_SLABS"].ToString();
                            sMoney = _drI["Money_value"].ToString();
                            sBelongsToMonth = _drI["belongs_to_month"].ToString();
                            sBelongsToYear = _drI["belongs_to_year"].ToString();

                            dtNew.Rows[iRowCount]["expected_tat"] = sExpectedTat.ToString();
                            dtNew.Rows[iRowCount]["tat_dsp"] = sTatDsp.ToString();
                            dtNew.Rows[iRowCount]["dsp_typed"] = sDspTyped.ToString();
                            dtNew.Rows[iRowCount]["ndsp_typed"] = sNdspTyped.ToString();
                            dtNew.Rows[iRowCount]["dsp_edited"] = sDspEdited.ToString();
                            dtNew.Rows[iRowCount]["ndsp_edited"] = sNdspEdited.ToString();
                            dtNew.Rows[iRowCount]["all_lines"] = sAllLines.ToString();
                            dtNew.Rows[iRowCount]["client_content"] = sClientType.ToString();
                            dtNew.Rows[iRowCount]["tat_slabs"] = sTatSlabs.ToString();
                            dtNew.Rows[iRowCount]["money_value"] = sMoney.ToString();
                            dtNew.Rows[iRowCount]["belongs_to_month"] = sBelongsToMonth.ToString();
                            dtNew.Rows[iRowCount]["belongs_to_year"] = sBelongsToYear.ToString();
                            iRowInvoice++;
                        }
                        iRowCount++;
                    }
                    string sFileDirectory = "c:\\MasterData\\";
                    string sFileDirectoryOne = "c:\\MasterData\\DoNotOpen\\";
                    string sFileName = "Final MasterData" + System.DateTime.Now.ToString("yyyyMMddhhmmss") + ".csv";
                    string sFileNameNew = "Final MasterData" + dtpDate.Value.Month + ".csv";
                    if (!Directory.Exists(sFileDirectory))
                    {
                        Directory.CreateDirectory(sFileDirectory);
                        CreateCSVFile(dtNew, sFileDirectory + sFileName);

                        if (!Directory.Exists(sFileDirectoryOne))
                        {
                            Directory.CreateDirectory(sFileDirectoryOne);
                            CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                            lblMsg.Text = "Exported Data...";
                        }
                    }
                    else
                    {
                        CreateCSVFile(dtNew, sFileDirectory + sFileName);

                        if (!Directory.Exists(sFileDirectoryOne))
                        {
                            Directory.CreateDirectory(sFileDirectoryOne);
                            CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                            lblMsg.Text = "Exported Data...";
                        }
                        else
                        {
                            File.Delete(sFileDirectoryOne + sFileNameNew);
                            CreateCSVFile(dtNew, sFileDirectoryOne + sFileNameNew);
                            lblMsg.Text = "Exported Data...";
                        }
                    }
                    //LoadListView(dtNew);
                }
                catch (Exception ex)
                {
                    lblMsg.Text = ex.ToString();
                    ExceptionHandler.HandleException(ex.ToString(), Environment.MachineName, "");
                }
                finally
                {
                    Thread.Sleep((1000 * 60) * 30);
                }
            } while (true);
        }

        #endregion "Methods"                       

        #region "Class"

        public class ListViewMasterData : ListViewItem
        {
            string stouch_one_by = string.Empty;
            string stouch_two_by = string.Empty;
            string stouch_three_by = string.Empty;
            string stouch_four_by = string.Empty;
            string stouch_five_by = string.Empty;

            public ListViewMasterData(DataRow _drRow, int iRows)
                : base()
            {
                stouch_one_by = _drRow["touch_one_by"].ToString();
                stouch_two_by = _drRow["touch_two_by"].ToString();
                stouch_three_by = _drRow["touch_three_by"].ToString();
                stouch_four_by = _drRow["touch_four_by"].ToString();
                stouch_five_by = _drRow["touch_five_by"].ToString();

                this.Text = iRows.ToString();
                this.SubItems.Add(_drRow["voice_file_id"].ToString());
                this.SubItems.Add(_drRow["file_minutes"].ToString());
                this.SubItems.Add(_drRow["location_name"].ToString());
                this.SubItems.Add(_drRow["rnd_content"].ToString());
                this.SubItems.Add(_drRow["file_lines"].ToString());
                this.SubItems.Add(_drRow["touch_one_status"].ToString());
                this.SubItems.Add(_drRow["touch_one_by"].ToString());
                this.SubItems.Add(_drRow["touch_one_lines"].ToString());
                this.SubItems.Add(_drRow["touch_one_date"].ToString());
                this.SubItems.Add(_drRow["touch_two_status"].ToString());
                this.SubItems.Add(_drRow["touch_two_by"].ToString());
                this.SubItems.Add(_drRow["touch_two_lines"].ToString());
                this.SubItems.Add(_drRow["touch_two_date"].ToString());
                this.SubItems.Add(_drRow["touch_three_status"].ToString());
                this.SubItems.Add(_drRow["touch_three_by"].ToString());
                this.SubItems.Add(_drRow["touch_three_lines"].ToString());
                this.SubItems.Add(_drRow["touch_three_date"].ToString());
                this.SubItems.Add(_drRow["touch_four_status"].ToString());
                this.SubItems.Add(_drRow["touch_four_by"].ToString());
                this.SubItems.Add(_drRow["touch_four_lines"].ToString());
                this.SubItems.Add(_drRow["touch_four_date"].ToString());
                this.SubItems.Add(_drRow["touch_five_status"].ToString());
                this.SubItems.Add(_drRow["touch_five_by"].ToString());
                this.SubItems.Add(_drRow["touch_five_lines"].ToString());
                this.SubItems.Add(_drRow["touch_five_date"].ToString());
                this.SubItems.Add(_drRow["expected_tat"].ToString());
                this.SubItems.Add(_drRow["tat_dsp"].ToString());
                this.SubItems.Add(_drRow["dsp_typed"].ToString());
                this.SubItems.Add(_drRow["ndsp_typed"].ToString());
                this.SubItems.Add(_drRow["dsp_edited"].ToString());
                this.SubItems.Add(_drRow["ndsp_edited"].ToString());
                this.SubItems.Add(_drRow["all_lines"].ToString());
                this.SubItems.Add(_drRow["client_content"].ToString());

                if ((stouch_one_by != "") && (stouch_two_by == "") && (stouch_three_by == "") && (stouch_four_by == "") && (stouch_five_by == ""))
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#f48fe5");
                else if ((stouch_one_by != "") && (stouch_two_by != "") && (stouch_three_by == "") && (stouch_four_by == "") && (stouch_five_by == ""))
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#fc775a");
                else if ((stouch_one_by != "") && (stouch_two_by != "") && (stouch_three_by != "") && (stouch_four_by == "") && (stouch_five_by == ""))
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#fcd55a");
                else if ((stouch_one_by != "") && (stouch_two_by != "") && (stouch_three_by != "") && (stouch_four_by != "") && (stouch_five_by == ""))
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#70c5fd");
                else if ((stouch_one_by != "") && (stouch_two_by != "") && (stouch_three_by != "") && (stouch_four_by != "") && (stouch_five_by != ""))
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#70c5fd");
            }
        }

        #endregion "Class"

        #region "Events"

        private void frmMasterData_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            lblMsg.Text = "Started...";            
            Thread tThread = new Thread(BindDataset_NTS);
            tThread.Start();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            
        }

        private void LoadInsert(DataTable dtTable)
        {
            try
            {
                //string sFileDirectoryOne = "c:\\MasterData\\DoNotOpen\\";
                //string sFileNameNew = "Final MasterData" + dtpDate.Value.Month + ".csv";
                //Class.MySqlBulkImport.ImportToDB(sFileDirectoryOne + sFileNameNew);
                
                List<Class.ClsCommom.Insert_Details> oInsert = new List<Class.ClsCommom.Insert_Details>();

                foreach (DataRow _drInsertRow in dtTable.Rows)
                {
                    Class.ClsCommom.Insert_Details oInsertData = new Class.ClsCommom.Insert_Details();
                    oInsertData.V_voice_file_id = _drInsertRow["voice_file_id"].ToString();
                }                    

                lblMsg.Text = "Data Imported...";
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
            }
        }

        private void btnClinics_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            Thread tThread = new Thread(BindDataset_Clinics);
            tThread.Start();
        }

        private void btnDuplicate_Click(object sender, EventArgs e)
        {
            frmDuplicateFix frmDup = new frmDuplicateFix();
            frmDup.Show();
        }

        #endregion "Events"
    }
}
