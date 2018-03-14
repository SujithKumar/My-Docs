using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Globalization;
using System.Collections;
using System.Net.Mail;

namespace eAllocation
{
    public partial class frmAllocation : Form
    {

        #region " DEFAULTS "
        public frmAllocation()
        {
            InitializeComponent();
            InitializeEvent(BusinessLogic.oProgressEvent);
            InitializeEvent(BusinessLogic.oMessageEvent);
        }
        public void InitializeEvent(ProgressEventClass PEC)
        {
            PEC.ProgressChanging += new ProgressEventClass.ProgressEventHander(ShowProgress);
        }

        /// <summary>
        /// This method is to show the progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void ShowProgress(object sender, ProgressEventArgs e)
        {
            pgrProgress.Style = e.oStyle;
            pgrProgress.Minimum = e.iMinValue;
            pgrProgress.Maximum = e.iMaxValue;
            pgrProgress.Value = e.iValue;
            pgrProgress.Visible = e.bVisible;
            pgrProgress.Refresh();
        }

        public void InitializeEvent(MessageEventClass MEC)
        {
            MEC.MessageThrown += new MessageEventClass.MessageEventHandler(ShowStatusMessage);
        }

        /// <summary>
        /// This method is to show the status message 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void ShowStatusMessage(object sender, MessageEventArgsClass e)
        {
            lblStatusMessage.Text = e.sMessage;
            lblStatusMessage.Refresh();
        }

        #endregion

        #region " VARIABLES "
        private DataTable dt;
        private int iTotalFiles;
        private double iTotalMins;
        private int iTatHours = 0;
        private int WORK_PLATFORM = 0;
        public string sAccountID;
        public string ProductionId;
        public Boolean chlStatus = false;
        private DataTable dtManual = new DataTable();
        BusinessLogic BusinessLogic = new BusinessLogic();
        DataSet _dsVolume = new DataSet();
        string sEditVoice = null;
        private string sBranch_ID = "-1";
        private string sDesgination_ID = "1";
        private string sPTAG_ID = string.Empty;
        private int iCurrent = 0;
        private int iOverall = 0;
        private int CLIENT_TYPE_ID = 0;
        DataSet _dsHourlyWiseReport = new DataSet();

        private int ICUSTOM_REMOVAL = 0;
        #endregion

        #region " CLASSES "

        /// <summary>
        /// SET CUSTOMIZED ME ALLOT FILES
        /// </summary>
        public class MyListItem_QuickAllot_ME_TAT : ListViewItem
        {
            public string SDURATION = string.Empty;
            public int ICLIENT_ID = 0;
            public int IDOCTOR_ID = 0;
            public string SSTATUS = string.Empty;
            public string SEMP_NAME = string.Empty;
            public string SUSER_ID = string.Empty;
            public int ITRANSCRIPTION_ID = 0;
            public int IPRODUCTION_ID = 0;
            public string SVOICE_FILE_ID = string.Empty;
            public string ALLOTED_PTAG_ID = string.Empty;
            public string USERID;
            public string EMP_NAME;
            public int ISFILE_OPEN = 0;

            public MyListItem_QuickAllot_ME_TAT(DataRow dr, int iRowCount)
                : base()
            {
                Text = iRowCount.ToString();
                SubItems.Add(dr["client_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["file_date"].ToString());
                SubItems.Add(dr["file_minutes"].ToString());
                SubItems.Add(dr["TAT"].ToString());
                SubItems.Add(dr["Bal_tat"].ToString());
                SubItems.Add(dr["alloted_PTag_Id"].ToString());
                SubItems.Add(dr["alloted_to"].ToString());
                SubItems.Add(dr["alloted_date"].ToString());

                SubItems.Add(dr["trans_by"].ToString());
                SubItems.Add(dr["trans_id"].ToString());
                SubItems.Add(dr["ted_by"].ToString());
                SubItems.Add(dr["ted_id"].ToString());

                SDURATION = dr["Duration"].ToString();
                ICLIENT_ID = Convert.ToInt32(dr["client_id"].ToString());
                IDOCTOR_ID = Convert.ToInt32(dr["doctor_id"].ToString());
                SEMP_NAME = dr["alloted_to"].ToString();
                SUSER_ID = dr["alloted_id"].ToString();
                ITRANSCRIPTION_ID = Convert.ToInt32(dr["transcription_id"].ToString());
                IPRODUCTION_ID = Convert.ToInt32(dr["transcription_id"].ToString());
                SVOICE_FILE_ID = dr["voice_file_id"].ToString();
                ISFILE_OPEN = Convert.ToInt32(dr["is_fileopen"].ToString());
                ALLOTED_PTAG_ID=dr["alloted_PTag_Id"].ToString();

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (ISFILE_OPEN == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF5733");
            }

            public string OFFLINE_P_EMPNAME
            {
                set
                {
                    this.SubItems[10].Text = value;
                    EMP_NAME = value;
                }
                get
                {
                    this.SubItems[10].Text = EMP_NAME;
                    return EMP_NAME;
                }
            }

            public string OFFLINE_P_USERID
            {
                set
                {
                    USERID = value;
                }
                get
                {
                    return USERID;
                }
            }
        }

        /// <summary>
        /// SET CUSTOMIZED ME ALLOT FILES
        /// </summary>
        public class MyListItem_QuickAllot_MT_TAT : ListViewItem
        {
            public string SDURATION = string.Empty;
            public int ICLIENT_ID = 0;
            public int IDOCTOR_ID = 0;
            public string SSTATUS = string.Empty;
            public string SEMP_NAME = string.Empty;
            public string SUSER_ID = string.Empty;
            public int ITRANSCRIPTION_ID = 0;
            public int IPRODUCTION_ID = 0;
            public string SVOICE_FILE_ID = string.Empty;
            public string USERID;
            public string EMP_NAME;
            public int ISFILE_OPEN = 0;

            public MyListItem_QuickAllot_MT_TAT(DataRow dr, int iRowCount)
                : base()
            {
                Text = iRowCount.ToString();
                SubItems.Add(dr["client_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["file_date"].ToString());
                SubItems.Add(dr["file_minutes"].ToString());
                SubItems.Add(dr["TAT"].ToString());
                SubItems.Add(dr["Bal_tat"].ToString());
                SubItems.Add(dr["alloted_PTag_Id"].ToString());
                SubItems.Add(dr["alloted_to"].ToString());
                SubItems.Add(dr["alloted_date"].ToString());

                SDURATION = dr["Duration"].ToString();
                ICLIENT_ID = Convert.ToInt32(dr["client_id"].ToString());
                IDOCTOR_ID = Convert.ToInt32(dr["doctor_id"].ToString());
                SEMP_NAME = dr["alloted_to"].ToString();
                SUSER_ID = dr["alloted_id"].ToString();
                ITRANSCRIPTION_ID = Convert.ToInt32(dr["transcription_id"].ToString());
                //IPRODUCTION_ID = Convert.ToInt32(dr["transcription_id"].ToString());
                SVOICE_FILE_ID = dr["voice_file_id"].ToString();
                ISFILE_OPEN = Convert.ToInt32(dr["is_fileopen"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (ISFILE_OPEN == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF5733");
            }

            public string OFFLINE_P_EMPNAME
            {
                set
                {
                    this.SubItems[10].Text = value;
                    EMP_NAME = value;
                }
                get
                {
                    this.SubItems[10].Text = EMP_NAME;
                    return EMP_NAME;
                }
            }

            public string OFFLINE_P_USERID
            {
                set
                {
                    USERID = value;
                }
                get
                {
                    return USERID;
                }
            }
        }

        public class MyListItem_EmpConsolidation : ListViewItem
        {
            public MyListItem_EmpConsolidation(DataRow dr)
            {
                Text = dr["branch_name"].ToString();
                SubItems.Add(dr["HT_User"].ToString());
                SubItems.Add(dr["Inhouse_User"].ToString());
                SubItems.Add(dr["TED_Inhouse"].ToString());
                SubItems.Add(dr["TED_HT"].ToString());
            }
        }

        public class MyListItem_DoctorwiseTotal : ListViewItem
        {
            public MyListItem_DoctorwiseTotal(DataRow dr)
            {
                Text = dr["doctor_full_name"].ToString();
                SubItems.Add(dr["Tot_files"].ToString());
                SubItems.Add(dr["Alloted_files"].ToString());
                SubItems.Add(dr["Bal_Files"].ToString());
                SubItems.Add(dr["Tot_Mins"].ToString());
                SubItems.Add(dr["Tot_Allot_Mins"].ToString());
                SubItems.Add(dr["Tot_Bal"].ToString());
            }
        }

        /// <summary>
        /// WEEKLY PROCESSED MINUTES
        /// </summary>
        public class MyListItem_WeeklyProcessedmins : ListViewItem
        {
            public MyListItem_WeeklyProcessedmins(DataRow dr, int iRowCount)
            {
                string sType = string.Empty;

                //Text = iRowCount.ToString();
                Text = dr["week_start"].ToString();
                SubItems.Add(dr["week_end"].ToString());
                SubItems.Add(dr["client_type"].ToString());
                SubItems.Add(dr["filecount"].ToString());
                SubItems.Add(dr["file_minutes"].ToString());

                sType = dr["client_type"].ToString();

                if (sType == "NTS")
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#A70AB0");
                    this.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                }
                else
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#12F5F5");
                }
            }

            public MyListItem_WeeklyProcessedmins(string iTot_Files, string Tot_Minust)
            {
                Text = string.Empty;
                SubItems.Add(string.Empty);
                SubItems.Add("Total");
                SubItems.Add(iTot_Files);
                SubItems.Add(Tot_Minust);
            }
        }

        public class Mylistitem_AutoAllocationUsers : ListViewItem
        {
            public int iPriorityID;
            public int iStatus;

            public int ICLIENTID;
            public string SLOCATIONID;
            public int IDOCTORID;
            public int IPRODUCTION_ID;
            public Mylistitem_AutoAllocationUsers(DataRow dr, int rowcount)
            {
                SubItems.Add(rowcount.ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["priority"].ToString());
                SubItems.Add(dr["user_name"].ToString());
                SubItems.Add(dr["autoallocationFiles"].ToString());

                iPriorityID = Convert.ToInt32(dr["priority_id"]);
                ICLIENTID = Convert.ToInt32(dr["client_id"]);
                SLOCATIONID = dr["location_id"].ToString();
                IDOCTORID = Convert.ToInt32(dr["doctor_id"]);

                iStatus = Convert.ToInt32(dr["is_active"]);
                IPRODUCTION_ID = Convert.ToInt32(dr["production_id"].ToString());

                //if (rowcount % 2 == 1)
                //    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                //else
                //    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        //USED TO LOAD HUNDRED PERCENT DOCTOR'S REVIEW LOCATION LIST
        public class ListItemHundred_Location : ListViewItem
        {
            public string LOCATION_ID;
            public ListItemHundred_Location(DataRow dr)
            {
                Text = dr["location_id"].ToString();
                SubItems.Add(dr["location_name"].ToString());

                LOCATION_ID = dr["location_id"].ToString();
            }
        }

        //USED TO LOAD HUNDRED PERCENT DOCTOR'S REVIEW LOCATION LIST
        public class ListItemHundred_Doctor : ListViewItem
        {
            public int DOCTOR_ID;
            public ListItemHundred_Doctor(DataRow dr)
            {
                Text = dr["doctor_id"].ToString();
                SubItems.Add(dr["doctor_full_name"].ToString());

                DOCTOR_ID = Convert.ToInt32(dr["doctor_id"].ToString());
            }
        }

        public class ListView_Complaints : ListViewItem
        {
            public ListView_Complaints(DataRow dr, int iRowCount)
            {
                Text = dr["complaint_date"].ToString();
                SubItems.Add(dr["Complaint"].ToString().Replace(',', ' '));
                SubItems.Add(dr["location_name"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }



        //USED TO LOAD HUNDRED PERCENT DOCTOR'S REVIEW LOCATION LIST
        public class ListItemHundred_Review : ListViewItem
        {
            public int IDOCTOR_ID;
            public ListItemHundred_Review(DataRow dr, int iRowCount)
            {
                Text = iRowCount.ToString();
                SubItems.Add(dr["location_id"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());

                IDOCTOR_ID = Convert.ToInt32(dr["doctor_id"].ToString());
            }
        }
        /// <summary>
        /// USED TO DISPLAY MT MET LIST
        /// </summary>
        public class ListItemMTMETList : ListViewItem
        {
            public int IPRODUCTION_ID;

            public ListItemMTMETList(DataRow dr, int iRowCount)
            {
                Text = iRowCount.ToString();
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr["Inserted_Date"].ToString());

                IPRODUCTION_ID = Convert.ToInt32(dr["production_id"].ToString());
            }
        }

        public class ListItemNightshift : ListViewItem
        {
            public int IBRANCH_ID;
            public int IBATCH_ID;
            public int IPRODUCTION_ID;

            public ListItemNightshift(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR].ToString());
                SubItems.Add(dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());

                IBRANCH_ID = Convert.ToInt32(dr[Framework.BRANCH.FIELD_BATCH_BRANCHID_INT].ToString());
                IBATCH_ID = Convert.ToInt32(dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT].ToString());
                IPRODUCTION_ID = Convert.ToInt32(dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        /// <summary>
        /// GET CUSTOMIZED INCENTIVE MINS
        /// </summary>
        public class ListItemIncentive_Mins : ListViewItem
        {
            public int IPRODUCTION_ID;
            public decimal DINCENTIVE_AMOUNT;
            public int IMINS_DONE;

            public ListItemIncentive_Mins(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["ID"].ToString());
                SubItems.Add(dr["Name"].ToString());
                SubItems.Add(dr["Target"].ToString());
                SubItems.Add(dr["File_Mins"].ToString());
                SubItems.Add(dr["Converted_Mins"].ToString());
                SubItems.Add(dr["30to39"].ToString());
                SubItems.Add(dr["40to49"].ToString());
                SubItems.Add(dr["50to59"].ToString());
                SubItems.Add(dr["Above60"].ToString());
                SubItems.Add(dr["Amount"].ToString());

                //IPRODUCTION_ID = Convert.ToInt32(dr["production_id"].ToString());
                //DINCENTIVE_AMOUNT = Convert.ToDecimal(dr["Incentive_Amount"].ToString());
                //IMINS_DONE = Convert.ToInt32(dr["mins_done"].ToString());
            }

            public ListItemIncentive_Mins(string iRow, string sId, string sName, string iTarget, string sMins, string sConvMins, string s30Mins, string s40Mins, string s50Mins, string s60Mins, string iAmount)
            {
                this.Text = iRow.ToString();
                SubItems.Add(sId);
                SubItems.Add(sName);
                SubItems.Add(iTarget.ToString());
                SubItems.Add(sMins);
                SubItems.Add(sConvMins);
                SubItems.Add(s30Mins);
                SubItems.Add(s40Mins);
                SubItems.Add(s50Mins);
                SubItems.Add(s60Mins);
                SubItems.Add(iAmount.ToString());
            }

            public ListItemIncentive_Mins(string iRow, string sMsss, string sMsg)
            {
                this.Text = iRow.ToString();
                SubItems.Add(string.Empty);
                SubItems.Add(sMsg);

                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// CAPACITY LEAVE DETAILS
        /// </summary>
        public class ListItemCapacity_Leave : ListViewItem
        {
            public int IPRODUCTION_ID;
            public string SEMP_NAME;

            public ListItemCapacity_Leave(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr["target_mins"].ToString());

                IPRODUCTION_ID = Convert.ToInt32(dr["production_id"].ToString());
                SEMP_NAME = dr["emp_full_name"].ToString();
            }
        }

        /// <summary>
        /// GET CUTOMIZED USERS LIST
        /// </summary>
        public class ListItem_UserName : ListViewItem
        {
            public int IPRODUCTION_ID;
            public ListItem_UserName(int iProductionid, string sUsername)
            {
                Text = sUsername.Trim();
                IPRODUCTION_ID = iProductionid;
            }
        }

        /// <summary>
        /// GET CUTOMIZED CAPACITY LIST
        /// </summary>
        public class ListItem_MTME_CurrentdateCapacity : ListViewItem
        {
            public int IPRODUCTION_ID = 0;

            public ListItem_MTME_CurrentdateCapacity(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                SubItems.Add(dr["target_mins"].ToString());
                SubItems.Add(dr["file_minutes"].ToString());
                SubItems.Add(dr["Converted_minutes"].ToString());

                IPRODUCTION_ID = Convert.ToInt32(dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

            }
            public ListItem_MTME_CurrentdateCapacity(string sCount, string sTag, string sName, string starget, string sTotMins, string sTot_ConvMins)
            {
                Text = string.Empty;
                SubItems.Add(string.Empty);
                SubItems.Add("Total");
                SubItems.Add(starget);
                SubItems.Add(sTotMins);
                SubItems.Add(sTot_ConvMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }

        }


        /// <summary>
        /// GET CUTOMIZED CAPACITY LIST
        /// </summary>
        public class ListItem_MTMECapacity : ListViewItem
        {
            public int IPRODUCTION_ID = 0;
            public ListItem_MTMECapacity(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                SubItems.Add(dr["target_mins"].ToString());

                IPRODUCTION_ID = Convert.ToInt32(dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
            }
            public ListItem_MTMECapacity(string sCount, string sTag, string sName, string starget)
            {
                Text = string.Empty;
                SubItems.Add(string.Empty);
                SubItems.Add("Total");
                SubItems.Add(starget);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }

        }


        /// <summary>
        /// GET CUSTOMIZED EMPLOYEE LIST
        /// </summary>
        public class ListItem_Customized_Employee : ListViewItem
        {
            public int IBRANCH_ID;
            public int IBATCH_ID;
            public int IPRODUCTION_ID;
            public int IGROUP_ID;
            public ListItem_Customized_Employee(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR].ToString());
                SubItems.Add(dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                SubItems.Add(dr["group_name"].ToString());

                IBRANCH_ID = Convert.ToInt32(dr[Framework.BRANCH.FIELD_BATCH_BRANCHID_INT].ToString());
                IBATCH_ID = Convert.ToInt32(dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT].ToString());
                IPRODUCTION_ID = Convert.ToInt32(dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                IGROUP_ID = Convert.ToInt32(dr["group_id"].ToString());
            }
        }

        /// <summary>
        /// GET TARGET DETAILS
        /// </summary>
        public class ListItem_TargetDetails : ListViewItem
        {
            public string SBAL_MINS;
            public string SIS_HT;
            public string SIS_NIGHT;

            public ListItem_TargetDetails(DataRow dr, int iRow)
            {
                Text = dr["submitted_time"].ToString();
                SubItems.Add(dr["Tot_files"].ToString());
                SubItems.Add(dr["target_mins"].ToString());
                SubItems.Add(dr["Achieved_Mins"].ToString());
                SubItems.Add(dr["Bal_mins"].ToString());
                SubItems.Add(dr["Completed_Percentage"].ToString());

                SBAL_MINS = dr["Bal_mins"].ToString();
                SIS_HT = dr["is_ht_user"].ToString();
                SIS_NIGHT = dr["isNightShift"].ToString();

                if (iRow % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (SBAL_MINS.Contains("-"))
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }
            }

            public ListItem_TargetDetails(int iRow, string sDate, string sTotfiles, string sTarget, string sAchieved_mins, string sBal_mins, string sComplete_mins, string sAchievedLines, string SDetails)
            {
                this.Name = SDetails.ToString();
                Text = sDate.ToString();
                SubItems.Add(sTotfiles.Trim());
                SubItems.Add(sTarget.Trim());
                SubItems.Add(sAchieved_mins.Trim());
                SubItems.Add(sBal_mins.Trim());
                SubItems.Add(sComplete_mins.Trim());
                SubItems.Add(sAchievedLines.Trim());

                if (iRow % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (sBal_mins.Contains("-"))
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR EMPLOYEE TARGET LIST
        /// </summary>
        public class ListItem_MTMEList : ListViewItem
        {
            public int IPRODUCTIONID;
            public string SEMP_NAME;
            public ListItem_MTMEList(DataRow dr, int i)
            {
                Text = dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString();
                SubItems.Add(dr["emp_full_name"].ToString());

                IPRODUCTIONID = Convert.ToInt32(dr["production_id"].ToString());
                SEMP_NAME = dr["emp_full_name"].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

        }

        /// <summary>
        /// CUSTOM CLASS FOR TRANSCRIBED ONLY ONLINE DETAILS
        /// </summary>
        public class MyTransonly_online : ListViewItem
        {
            public MyTransonly_online(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                SubItems.Add(dr["Minutes"].ToString());
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIBED_DATE_DTIME].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public MyTransonly_online(string sRowcount, string sClient, string sLocation, string sDoctor, string sVoice, string sFiledate, string sMins, string sPtag, string sEmpname, string sTransdate)
            {
                Text = string.Empty;
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(sFiledate);
                SubItems.Add(sMins);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR TRANSCRIBED ONLY OFFLINE DETAILS
        /// </summary>
        public class MyTransonly_Offline : ListViewItem
        {
            public MyTransonly_Offline(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                SubItems.Add(dr["Minutes"].ToString());
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIBED_DATE_DTIME].ToString());
                SubItems.Add(dr["Alloted_for"].ToString());
                SubItems.Add(dr["Alloted_Name"].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_ALLOTED_DATE_DTIME].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public MyTransonly_Offline(string sRowcount, string sClient, string sLocation, string sDoctor, string sVoice, string sFiledate, string sMins, string sPtag, string sEmpname, string sTransdate, string sAllotfor, string sAllotname, string sAllotdate)
            {
                Text = string.Empty;
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(sFiledate);
                SubItems.Add(sMins);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);
                SubItems.Add(string.Empty);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR REALLOCATION INFORMATION
        /// </summary>
        public class MyReallocation : ListViewItem
        {
            public int IPRODUCTION_ID;
            public int ITRANSCRIPTION_ID;
            public string SPTAG_ID;
            public string SREPORT_NAME;
            public MyReallocation(DataRow dr, int iRowcount)
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                SubItems.Add(dr["user_name"].ToString());
                SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                IPRODUCTION_ID = Convert.ToInt32(dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                ITRANSCRIPTION_ID = Convert.ToInt32(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString());
                SPTAG_ID = dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString();
                SREPORT_NAME = dr[Framework.MAINTRANSCRIPTION.FIELD_REPORT_NAME_STR].ToString();
            }
        }

        /// <summary>
        /// Custom class for Allocation Information
        /// </summary>
        public class MyAllocatioFile : ListViewItem
        {
            public int TRANSCRIPTIONID;
            public int CLIENTID;
            public int DOCTORID;
            public double MINUTES;
            public string STATUS;
            public string USERID;
            public string EMP_NAME;

            public MyAllocatioFile(DataRow dr, int i)
            {
                this.Name = dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString();
                this.Text = dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());
                this.SubItems.Add(dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIBED_BY_STR].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_EDIT_BY_STR].ToString());

                STATUS = dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString();
                EMP_NAME = "";

                TRANSCRIPTIONID = Convert.ToInt32(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString());
                CLIENTID = Convert.ToInt32(dr[Framework.CLIENT.FIELD_CLIENT_ID_BINT].ToString());
                DOCTORID = Convert.ToInt32(dr[Framework.DOCTOR.FIELD_DOCTOR_ID_BINT].ToString());
                MINUTES = Convert.ToDouble(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public string P_STATUS
            {
                set
                {
                    this.SubItems[8].Text = value;
                    STATUS = value;
                    if (value.ToString() == "Allotted") //ENUM
                        this.BackColor = Color.Orange;
                    else if (value.ToString() == "Ready")
                    {
                        this.BackColor = Color.White;
                    }
                }
                get
                {
                    this.SubItems[8].Text = STATUS;
                    return STATUS;
                }
            }

            public string P_EMPNAME
            {
                set
                {
                    this.SubItems[9].Text = value;
                    EMP_NAME = value;
                }
                get
                {
                    this.SubItems[9].Text = EMP_NAME;
                    return EMP_NAME;
                }
            }

            public string P_USERID
            {
                set
                {
                    USERID = value;
                }
                get
                {
                    return USERID;
                }
            }

            public double P_FILEMINS
            {
                set
                {
                    this.SubItems[4].Text = value.ToString();
                }
                get
                {
                    return Convert.ToDouble(this.SubItems[4].Text);
                }
            }

            public string P_VOICE_FILE_NAME
            {
                set
                {
                    this.SubItems[3].Text = value.ToString();
                }
                get
                {
                    return this.SubItems[3].Text;
                }
            }
        }

        private class ListItem_Employee : ListViewItem
        {
            public string PRODUCTION_EMPLOYEE_ID;
            public string PRODUCTION_EMPLOYEE_NAME;

            public ListItem_Employee(DataRow _dr, int i)
            {
                this.Text = _dr["emp_name"].ToString();
                PRODUCTION_EMPLOYEE_ID = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();
                PRODUCTION_EMPLOYEE_NAME = _dr["emp_name"].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_OnlineAllocationStatus : ListViewItem
        {
            public string sTranscriptionID;
            public string sVoiceFileID;
            public string sMinutes;

            public ListItem_OnlineAllocationStatus(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr["client_name"].ToString();
                this.SubItems.Add(_dr["location_name"].ToString());
                this.SubItems.Add(_dr["doctor_full_name"].ToString());
                this.SubItems.Add(_dr["voice_file_id"].ToString());
                this.SubItems.Add(_dr["file_date"].ToString());
                this.SubItems.Add(_dr["Transcibed_by"].ToString());
                this.SubItems.Add(_dr["Transcibed_Time"].ToString());
                this.SubItems.Add(_dr["Transcibed_Status"].ToString());
                this.SubItems.Add(_dr["TED_by"].ToString());
                this.SubItems.Add(_dr["TED_Time"].ToString());
                this.SubItems.Add(_dr["TED_Status"].ToString());
                this.SubItems.Add(_dr["NDSP_by"].ToString());
                this.SubItems.Add(_dr["NDSP_Time"].ToString());
                this.SubItems.Add(_dr["NDSP_Status"].ToString());
                this.SubItems.Add(_dr["Edit_by"].ToString());
                this.SubItems.Add(_dr["Edit_Time"].ToString());
                this.SubItems.Add(_dr["Edit_Status"].ToString());
                this.SubItems.Add(_dr["QC_by"].ToString());
                this.SubItems.Add(_dr["QC_Time"].ToString());
                this.SubItems.Add(_dr["QC_Status"].ToString());

                //sTranscriptionID = _dr["transcription_id"].ToString();
                sVoiceFileID = _dr["voice_file_id"].ToString();
                //sMinutes = _dr["file_minutes"].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_OfflineAllocationStatus : ListViewItem
        {
            public string sTranscriptionID;
            public string sVoiceFileID;
            public string sMinutes;

            public ListItem_OfflineAllocationStatus(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr["client_name"].ToString();
                this.SubItems.Add(_dr["location_name"].ToString());
                this.SubItems.Add(_dr["doctor_full_name"].ToString());
                this.SubItems.Add(_dr["voice_file_id"].ToString());
                this.SubItems.Add(_dr["file_minutes"].ToString());
                this.SubItems.Add(_dr["file_date"].ToString());
                this.SubItems.Add(_dr["file_status_description"].ToString());
                this.SubItems.Add(_dr["alloted_for"].ToString());
                this.SubItems.Add(_dr["alloted_date"].ToString());
                this.SubItems.Add(_dr["transcribed_by"].ToString());
                this.SubItems.Add(_dr["transcribed_date"].ToString());
                this.SubItems.Add(_dr["edit_by"].ToString());
                this.SubItems.Add(_dr["edit_date"].ToString());
                this.SubItems.Add(_dr["hold_review_by"].ToString());
                this.SubItems.Add(_dr["hold_review_date"].ToString());

                sTranscriptionID = _dr["transcription_id"].ToString();
                sVoiceFileID = _dr["voice_file_id"].ToString();
                sMinutes = _dr["file_minutes"].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        private class ListItem_LoginEmployees : ListViewItem
        {
            public string EMP_PRODUCTION_ID;

            public ListItem_LoginEmployees(DataRow _dr, int i)
            {
                this.Text = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr["AllottedFile"].ToString());
                this.SubItems.Add(_dr["Allottedmins"].ToString());
                this.SubItems.Add(_dr["Achievedfile"].ToString());
                this.SubItems.Add(_dr["Achievedmins"].ToString());
                this.SubItems.Add(_dr["Totalfile"].ToString());
                this.SubItems.Add(_dr["Totalmins"].ToString());

                EMP_PRODUCTION_ID = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        private class ListItem_AllotedFilesForUsers : ListViewItem
        {
            public int TRANSCRIPTION_ID;
            public string sVoiceFile_ID;

            public ListItem_AllotedFilesForUsers(DataRow _dr, int i)
            {
                this.SubItems.Add(_dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());

                TRANSCRIPTION_ID = Convert.ToInt32(_dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT]);
                sVoiceFile_ID = _dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        private class ListItem_EmployeeList : ListViewItem
        {
            public string PRODUCTION_ID;

            public ListItem_EmployeeList(DataRow _dr, int i)
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_ID].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_DESIGNATION].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

                PRODUCTION_ID = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_LeaveDetails : ListViewItem
        {
            public int DAY_ID;

            public ListItem_LeaveDetails(DataRow _dr, int i)
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr["Achieved_Minutes"].ToString());
                this.SubItems.Add(_dr[Framework.ME_DAY.DAY_STATUS].ToString());
                this.SubItems.Add(_dr[Framework.ME_DAY.SS_COMMENT].ToString());
                this.SubItems.Add(_dr[Framework.ME_DAY.POCESSED_DATE].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                DAY_ID = Convert.ToInt32(_dr[Framework.ME_DAY.DAY_ID]);
            }
        }

        public class ListItem_UserAllocationProfile : ListViewItem
        {
            public ListItem_UserAllocationProfile(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_drRow[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_drRow["no_of_files"].ToString());
                this.SubItems.Add(_drRow["allocation_set_date"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_LargeMinutes : ListViewItem
        {
            public int MINUTES;
            public string VOICE_FILE_ID;

            public ListItem_LargeMinutes(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_drRow["TES_minutes"].ToString());
                this.SubItems.Add(_drRow["User_Entered"].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_LINES_DECIMAL].ToString());
                this.SubItems.Add(_drRow["Trans"].ToString());
                this.SubItems.Add(_drRow["Edit"].ToString());
                this.SubItems.Add(_drRow["Hold"].ToString());

                MINUTES = Convert.ToInt32(_drRow["TES_Seconds"]);
                VOICE_FILE_ID = _drRow["voice_file_id"].ToString();

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (MINUTES >= 1200)
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }

            }
        }

        public class ListItem_LineCountDeatils : ListViewItem
        {
            public ListItem_LineCountDeatils(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_drRow[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_REPORT_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_CONVERTED_MINUTES_DOUBLE].ToString().Replace(":", "."));
                this.SubItems.Add(_drRow[Framework.TRANSCRIPTIONTRANSACTION.FIELD_SUBMITTED_TIME].ToString());
                this.SubItems.Add(_drRow[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());
                this.SubItems.Add(_drRow[Framework.TRANSCRIPTIONTRANSACTION.FIELD_CONVERTED_LINES_DECIMAL].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_LineCountDeatils(string sMessage, string sTotMins, string dTotal)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sTotMins);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(dTotal.ToString());

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// tab the list view items
        /// </summary>
        public class MylsvManualDetails : ListViewItem
        {
            DataRow dr;
            public int DOCTOR_ID;
            public string EXTENSION;
            public int TRANS_ID;
            public string JOBID;
            public string DICTATION_PATH;
            public decimal DURATION;
            public decimal SIZE;
            public DateTime DOWNLOAD_DATE;
            public int TAT;
            public string STR_STATUS = string.Empty;

            public MylsvManualDetails(DataRow dr, int i)
            {
                this.Text = i.ToString();
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(dr["dictation_path"].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_SIZE_BINT].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());
                STR_STATUS = "Ready";
                this.SubItems.Add(STR_STATUS);

                TRANS_ID = i;
                JOBID = dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();
                DOCTOR_ID = Convert.ToInt32(dr[Framework.MAINTRANSCRIPTION.FIELD_DOCTOR_ID_BINT]);
                EXTENSION = dr[Framework.MAINTRANSCRIPTION.FIELD_EXTENSION_STR].ToString();
                DICTATION_PATH = dr["dictation_path"].ToString();
                DURATION = Convert.ToDecimal(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                SIZE = Convert.ToDecimal(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_SIZE_BINT].ToString());
                TAT = Convert.ToInt32(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public string _STATUS
            {
                get
                {
                    STR_STATUS = this.SubItems[6].Text;
                    return STR_STATUS;
                }
                set
                {
                    STR_STATUS = value;
                    this.SubItems[6].Text = STR_STATUS;
                }
            }
        }

        /// <summary>
        /// Load the list view items
        /// </summary>
        public class Mylsvextractdetails : ListViewItem
        {
            DataRow dr;
            public Mylsvextractdetails(DataRow dr, int iRowCount)
            {
                this.Text = dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString();
                this.SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(dr["voicefilename"].ToString());
                this.SubItems.Add(dr["fileminutes"].ToString());
                this.SubItems.Add(dr["filesize"].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_EmployeeTotalMinutes : ListViewItem
        {
            public ListItem_EmployeeTotalMinutes(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr["Totfiles"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_CONVERTED_LINES_DECIMAL].ToString());

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_EmployeeHourly_Log : ListViewItem
        {
            public ListItem_EmployeeHourly_Log(DataRow _dr, int i)
                : base()
            {

                this.Name = _dr["details"].ToString();
                this.Text = _dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR].ToString());
                //this.SubItems.Add(_dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());
                this.SubItems.Add(_dr["Totfiles"].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr["Tot_Mins"].ToString().Replace(":", "."));
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString().Replace(":", "."));
                this.SubItems.Add(_dr["Bal_mins"].ToString());
                this.SubItems.Add(_dr["Completed_Percentage"].ToString());
                this.SubItems.Add(_dr["Linecount"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_CONVERTED_LINES_DECIMAL].ToString());
                this.SubItems.Add(_dr["location_names"].ToString());

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_TransEdit : ListViewItem
        {
            public ListItem_TransEdit(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.REPORT_TRANSACTION.FIELD_REPORT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_dr["file_minutes"].ToString().Replace(":", "."));
                this.SubItems.Add(_dr["Converted_minutes"].ToString());
                this.SubItems.Add(_dr["file_lines"].ToString());
                this.SubItems.Add(_dr["converted_lines"].ToString());
                this.SubItems.Add(_dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_TransEdit(string sMessage, string sTotMins, string sTotConvMins, string dTotal, string dTotConvLines)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sTotMins);
                this.SubItems.Add(sTotConvMins);
                this.SubItems.Add(dTotal);
                this.SubItems.Add(dTotConvLines);
                this.SubItems.Add(string.Empty);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class ListItem_TransEdit_Clinics : ListViewItem
        {
            public ListItem_TransEdit_Clinics(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.REPORT_TRANSACTION.FIELD_REPORT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_dr["file_minutes"].ToString().Replace(":", "."));
                this.SubItems.Add(_dr["Converted_minutes"].ToString());
                this.SubItems.Add(_dr["file_lines"].ToString());
                this.SubItems.Add(_dr["converted_lines"].ToString());
                this.SubItems.Add(_dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_TransEdit_Clinics(string sMessage, string sTotMins, string sTotConvMins, string dTotal, string dTotConvLines)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sTotMins);
                this.SubItems.Add(sTotConvMins);
                this.SubItems.Add(dTotal);
                this.SubItems.Add(dTotConvLines);
                this.SubItems.Add(string.Empty);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR ACCOUNT WISE FILE DETAILS
        /// </summary>
        public class ListItem_AccountWiseInfo : ListViewItem
        {
            public int ICLIENT_ID = 0;
            public string SLOCATION_ID = string.Empty;

            public ListItem_AccountWiseInfo(DataRow _dr, int i)
                : base()
            {
                //this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                //this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                //this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());

                this.Text = _dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_dr["Totfiles"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString().Replace(":", "."));

                ICLIENT_ID = Convert.ToInt32(_dr["client_id"].ToString());
                SLOCATION_ID = _dr["location_id"].ToString();

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");



            }

            public ListItem_AccountWiseInfo(string sMessage, string sTotMins)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sTotMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR ACCOUNT WISE FILE DETAILS
        /// </summary>
        public class ListItem_Offline_AccountWiseInfo : ListViewItem
        {
            public ListItem_Offline_AccountWiseInfo(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_dr["Totfiles"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString().Replace(":", "."));

                if (i % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_Offline_AccountWiseInfo(string sMessage, string sTotMins)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sTotMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// GET BACKLOCK MINUTES DONE
        /// </summary>
        public class ListItem_Offline_AccountWiseInfo_BackLock : ListViewItem
        {
            public ListItem_Offline_AccountWiseInfo_BackLock(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr["BackLock_Files"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString().Replace(":", "."));

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_Offline_AccountWiseInfo_BackLock(string sMessage, string sTotMins)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sTotMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }


        /// <summary>
        /// CLASS TO LOAD ACCOUNT WISE INFO BRANCH WISE
        /// </summary>
        public class ListItem_AccountWiseInfo_BranchWise : ListViewItem
        {
            public ListItem_AccountWiseInfo_BranchWise(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr["Totfiles"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString().Replace(":", "."));
                this.SubItems.Add(_dr["FileLines"].ToString());
            }

            public ListItem_AccountWiseInfo_BranchWise(string sMessage, string sTotMins, string sTotLines)
                : base()
            {
                this.Text = sMessage;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sTotMins);
                this.SubItems.Add(sTotLines);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        private string sGetDuration(int Seconds)
        {
            string sDuration = string.Empty;

            object oTotalDuration = null;
            oTotalDuration = Seconds;
            if (oTotalDuration != null)
            {
                int iTotalSeconds = Convert.ToInt32(oTotalDuration);
                int iMinutes = 0;
                int iSeconds = 0;
                if (iTotalSeconds > 0)
                {
                    iMinutes = iTotalSeconds / 60;
                    iSeconds = iTotalSeconds % 60;
                }
                sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
            }
            return sDuration;
        }

        private string sGetDuration_Neg(int Seconds)
        {
            string sDuration = string.Empty;

            object oTotalDuration = null;
            oTotalDuration = Seconds;
            if (oTotalDuration != null)
            {
                int iTotalSeconds = Convert.ToInt32(oTotalDuration);
                int iMinutes = 0;
                int iSeconds = 0;

                iMinutes = iTotalSeconds / 60;
                iSeconds = iTotalSeconds % 60;

                sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
            }
            return sDuration;
        }

        /// <summary>
        /// CLADD TO LOAD EMPLOYEE BATCH
        /// </summary>
        public class ListItem_EmpBatchName : ListViewItem
        {
            public int iBatchID;
            public ListItem_EmpBatchName(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT].ToString();
                this.SubItems.Add(_dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                iBatchID = Convert.ToInt32(_dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT]);
            }
        }

        /// <summary>
        /// CUSTOM CLASS FOR TARGET
        /// </summary>
        public class ListItem_Target : ListViewItem
        {
            public int iTargetProdID;
            public int iTarget;
            public int iTragetID;

            public ListItem_Target(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_DATE_DTIME].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                iTargetProdID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                iTarget = Convert.ToInt32(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                iTragetID = Convert.ToInt32(_dr[Framework.TARGET.FIELD_TARGET_ID_BINT].ToString());
            }
        }

        public int iMissingEntryCount = 1;
        public int iChangeMinutesCount = 1;
        public int iChangeLines = 1;
        public int iChangeShift = 1;
        public int iConvertTransEdit = 1;
        public int iChangeFilesStatus = 1;

        /// <summary>
        /// CUSTOM CLASS FOR DISCREPANCY
        /// </summary>
        public class ListItem_Discrepancy : ListViewItem
        {
            public int iAccountID;
            public int iDiscrepancyID;
            public int iDiscrepancyMasterID;
            public string sVoiceFileID;
            public string sMinutes;
            public string sLines;
            public int iPorductionID;
            public int iFileStatusID;
            public DateTime dDateEntered;
            public DateTime dSubmitted_time;

            public ListItem_Discrepancy(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_NAME].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_ENTERED_DATE].ToString());
                this.SubItems.Add(_dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_LINES_DECIMAL].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_COMMENT].ToString());

                iDiscrepancyID = Convert.ToInt32(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_TRANSACTION_ID].ToString());
                iDiscrepancyMasterID = Convert.ToInt32(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString());
                sVoiceFileID = _dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();
                sMinutes = _dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString();
                iPorductionID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                sLines = _dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_LINES_DECIMAL].ToString();
                iFileStatusID = Convert.ToInt32(_dr[Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT].ToString());
                dDateEntered = Convert.ToDateTime(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_ENTERED_DATE]);
                iAccountID = Convert.ToInt32(_dr[Framework.CLIENT.FIELD_CLIENT_ID_BINT]);
                dSubmitted_time = Convert.ToDateTime(_dr["file_submitted_time"].ToString());

                if (iDiscrepancyMasterID == 1)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#19aeff");
                }
                else if (iDiscrepancyMasterID == 2)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#ff4141");
                }
                else if (iDiscrepancyMasterID == 3)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#ffff3e");
                }
                else if (iDiscrepancyMasterID == 4)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#b88100");
                }
                else if (iDiscrepancyMasterID == 5)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#9ade00");
                }
                else if (iDiscrepancyMasterID == 6)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#d76cff");
                }
            }
        }

        public class Listitem_MissingFileDetails : ListViewItem
        {
            public int iFileStatusID;

            public Listitem_MissingFileDetails(DataRow _Dr, int iCount)
                : base()
            {
                this.Text = _Dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_Dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_Dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_Dr[Framework.REPORT_TRANSACTION.FIELD_REPORT_NAME_STR].ToString());
                this.SubItems.Add(_Dr[Framework.REPORT_TRANSACTION.FIELD_FILE_LINES_DECIMAL].ToString());
                this.SubItems.Add(_Dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_Dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());
                this.SubItems.Add(_Dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIBED_BY_STR].ToString());
                this.SubItems.Add(_Dr[Framework.MAINTRANSCRIPTION.FIELD_EDIT_BY_STR].ToString());
                this.SubItems.Add(_Dr[Framework.MAINTRANSCRIPTION.FIELD_HOLD_REVIEW_BY_STR].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_InclusiveLines : ListViewItem
        {
            public ListItem_InclusiveLines(DataRow _Dr, int iCount)
                : base()
            {
                this.Text = _Dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_LINES.INCLUSIVE_LINE_DATE].ToString());
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_LINES.INCLUSIVE_LI_LINES].ToString());
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_LINES.INCLUSIVE_LINE_COMMENT].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_InclusiveMinutes : ListViewItem
        {
            public ListItem_InclusiveMinutes(DataRow _Dr, int iCount)
                : base()
            {
                this.Text = _Dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_MINUTES.INCLUSIVE_MINUTES_DATE].ToString());
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_MINUTES.INCLUSIVE_MINUTES_ID].ToString());
                this.SubItems.Add(_Dr[Framework.INCLUSIVE_MINUTES.INCLUSIVE_MINUTES_COMMENT].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#5544778");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#9888999");
            }
        }

        public class ListItem_Account : ListViewItem
        {
            public int iClientID;
            public string sClietnName;

            public ListItem_Account(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                iClientID = Convert.ToInt32(_dr[Framework.CLIENT.FIELD_CLIENT_ID_BINT].ToString());
                sClietnName = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
            }
        }

        public class Listitem_AccountIncentive : ListViewItem
        {
            public Listitem_AccountIncentive(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = _dr[Framework.ACCOUNT_INCENTIVE.FIELD_INCENTIVE_MONTH].ToString();
                this.SubItems.Add(_dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.ACCOUNT_INCENTIVE.FIELD_INCENTIVE_RATE].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class Listitem_AllEmployees : ListViewItem
        {
            public int iProductionID;

            public Listitem_AllEmployees(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_ID].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());

                iProductionID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
            }
        }

        public class Listitem_AllEmployees_List : ListViewItem
        {
            public int iProductionID;
            public bool isActive;

            public Listitem_AllEmployees_List(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = iRowCount.ToString();
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr["branch_name"].ToString());
                this.SubItems.Add(_dr["is_ht"].ToString());
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_ID].ToString());
                this.SubItems.Add(_dr["Active_Status"].ToString());
                this.SubItems.Add(_dr["dictaphone_id"].ToString());
                this.SubItems.Add(_dr["escription_id"].ToString());
                this.SubItems.Add(_dr["Work_platform"].ToString());

                if (!Convert.ToBoolean(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEES_ISACTIVE_BIT]))
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#A72D2D");
                    this.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FCF8F8");
                }
                else
                {
                    if (iRowCount % 2 == 1)
                        this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                    else
                        this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
                }



                iProductionID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                isActive = Convert.ToBoolean(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEES_ISACTIVE_BIT]);
            }
        }

        public class Listitem_Category : ListViewItem
        {
            public int iCategoryID;

            public Listitem_Category(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_CATEGORY].ToString();

                iCategoryID = Convert.ToInt32(_dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_ID].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class Listitem_NightShift_Marked : ListViewItem
        {
            public int IPRODUCTION_ID;
            public int ICATEGORY_ID;
            public int INIGHTSHIFT_TRANS_ID;

            public Listitem_NightShift_Marked(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_ID].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                this.SubItems.Add(_dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_CATEGORY].ToString());
                this.SubItems.Add(_dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_ALLOWANCE_PAISE].ToString());

                IPRODUCTION_ID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                ICATEGORY_ID = Convert.ToInt32(_dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_ID].ToString());
                INIGHTSHIFT_TRANS_ID = Convert.ToInt32(_dr["night_shift_trans_id"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_LineCountFileItem : ListViewItem
        {

            public int TRANSCRIPTION_TRANS_ID;
            public string FINAL_REPORT_NAME;
            public string FINAL_REPORT_STATUS;
            public string DOCTOR_NAME;
            public string LOCATION_NAME;
            public string STATUS_FOLDER_NAME;
            public string FINAL_STATUS_FOLDER_NAME;
            public int IS_NIGHTSHIFT;
            public int IS_HOLDLINE;
            public int FILE_STATUS_ID;
            public int REPORT_TRANSACTION_ID;
            public string sVOICE_FILE_ID;

            public ListItem_LineCountFileItem(DataRow _drFileInfo, int iRowCount)
                : base()
            {
                this.Name = _drFileInfo["" + Framework.REPORT_TRANSACTION.FIELD_REPORT_NAME_STR + ""].ToString();
                this.Text = (iRowCount + 1).ToString();
                this.SubItems.Add(_drFileInfo["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_REPORT_NAME_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME + ""].ToString());
                this.SubItems.Add(_drFileInfo[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_drFileInfo["Converted_minutes"].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.REPORT_TRANSACTION.FIELD_FILE_LINES_DECIMAL + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_CONVERTED_LINES_DECIMAL + ""].ToString());

                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_SUBMITTED_TIME + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_EVALUATED_DATE_DTIME + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.TEMPLATE.FIELD_TEMPLATE_DESCRIPTION_STR + ""].ToString());

                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_ACCURACY_DECIMAL + ""].ToString());
                this.SubItems.Add(_drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_IS_GRADED_BINT + ""].ToString());

                this.SubItems.Add(_drFileInfo["Transcibed_by"].ToString());
                this.SubItems.Add(_drFileInfo["Trans_Blank"].ToString());
                this.SubItems.Add(_drFileInfo["TED_by"].ToString());
                this.SubItems.Add(_drFileInfo["TED_Blank"].ToString());
                this.SubItems.Add(_drFileInfo["NDSP_by"].ToString());
                this.SubItems.Add(_drFileInfo["NDSP_Blank"].ToString());
                this.SubItems.Add(_drFileInfo["Edit_by"].ToString());
                this.SubItems.Add(_drFileInfo["Edit_Blank"].ToString());
                this.SubItems.Add(_drFileInfo["QC_by"].ToString());
                this.SubItems.Add(_drFileInfo["QC_Blank"].ToString());

                this.SubItems.Add(_drFileInfo["isNightShift"].ToString());
                this.SubItems.Add(_drFileInfo["is_Holdline"].ToString());

                this.SubItems.Add(_drFileInfo["report_transaction_id"].ToString());
                this.SubItems.Add(_drFileInfo["voice_file_id"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = Color.LightGoldenrodYellow;
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                FINAL_REPORT_NAME = _drFileInfo["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_REPORT_NAME_STR + ""].ToString();
                FINAL_REPORT_STATUS = _drFileInfo["" + Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR + ""].ToString();
                LOCATION_NAME = _drFileInfo["" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + ""].ToString();
                DOCTOR_NAME = _drFileInfo["" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + ""].ToString();
                IS_NIGHTSHIFT = Convert.ToInt32(_drFileInfo["isNightShift"].ToString());
                FILE_STATUS_ID = Convert.ToInt32(_drFileInfo["file_status_id"].ToString());
                IS_HOLDLINE = Convert.ToInt32(_drFileInfo["is_Holdline"].ToString());
                REPORT_TRANSACTION_ID = Convert.ToInt32(_drFileInfo["report_transaction_id"].ToString());
                sVOICE_FILE_ID = _drFileInfo["voice_file_id"].ToString();

                if (IS_NIGHTSHIFT == 1)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFBF");
                }

                if (IS_HOLDLINE == 1)
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }
            }

            public ListItem_LineCountFileItem(string sMessage)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);


                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }

            public ListItem_LineCountFileItem(string sMessage, string Minutes, string ConvertedMinutes, string sLines, string sConvertedLines, string sErrorPoint, string sAccuracy)
                : base()
            {
                //Minutes
                //var timeMins = Math.Round(TimeSpan.FromSeconds(Convert.ToDouble(Minutes)).TotalMinutes, 2);


                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(Minutes.ToString());
                this.SubItems.Add(ConvertedMinutes.ToString());
                this.SubItems.Add(sLines);
                this.SubItems.Add(sConvertedLines);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sErrorPoint);
                this.SubItems.Add(sAccuracy);
                this.SubItems.Add(string.Empty);


                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }


            public int P_REPORT_TRANSACTION
            {
                get
                {
                    return Convert.ToInt32(this.SubItems[18].Text.ToString());
                }
            }

            public int P_FILE_STATUS_ID
            {
                get
                {
                    return Convert.ToInt32(FILE_STATUS_ID);
                }
            }



            public int P_VOICE_ID
            {
                get
                {
                    return Convert.ToInt32(this.SubItems[19].Text.ToString());
                }
            }

            public string P_FILE_MINS
            {
                get
                {
                    return this.SubItems[6].Text.ToString();
                }
            }

            public string P_CLIENT
            {
                get
                {
                    return this.SubItems[1].Text.ToString();
                }
            }

            public decimal P_LINES
            {
                get
                {
                    return Convert.ToDecimal(this.SubItems[8].Text.ToString());
                }
            }

            public int P_STATUS
            {
                get
                {
                    return Convert.ToInt32(this.SubItems[12].Text.ToString());
                }
            }

            public bool P_IS_GRADED
            {
                get
                {
                    return (this.SubItems[16].Text.ToLower().Equals("yes"));
                }
            }

            public bool P_IS_COMPLETED
            {
                get
                {
                    return (this.SubItems[10].Text.ToLower().Equals("completed"));
                }
            }

            public DateTime P_SUBMITTED_DATE
            {
                get
                {
                    return Convert.ToDateTime(this.SubItems[5].Text);
                }
            }

            public string P_REPORT_NAME
            {
                get
                {
                    return this.SubItems[4].Text.ToString();
                }
            }

            public string P_GROUP_NAME
            {
                get
                {
                    return this.SubItems[1].Text.ToString();
                }
            }

            public string P_LOCATION_NAME
            {
                get
                {
                    return this.SubItems[2].Text.ToString();
                }
            }

            public string P_DOCTOR_NAME
            {
                get
                {
                    return this.SubItems[3].Text.ToString();
                }
            }

            public string P_SUBMISSION_STATUS
            {
                get
                {
                    return this.SubItems[9].Text.ToString();
                }
            }

            ~ListItem_LineCountFileItem() { }
        }

        public class ListSummery : ListViewItem
        {
            public ListSummery(DataRow dr)
                : base()
            {

                //var timeMins = TimeSpan.FromSeconds(Convert.ToDouble(dr["Tot_mins"].ToString())).TotalMinutes;

                this.Text = dr["file_status_description"].ToString();
                this.SubItems.Add(dr["Totfiles"].ToString());
                this.SubItems.Add(dr["Tot_mins"].ToString());
                this.SubItems.Add(dr["Linecount"].ToString());
                this.SubItems.Add(dr["Converted_Linecount"].ToString());
                this.SubItems.Add(dr["NightShift_Linecount"].ToString());
                this.SubItems.Add(dr["Sunday_Shift_Lines"].ToString());
                this.SubItems.Add(dr["Extra_Support_Lines"].ToString());
                this.SubItems.Add(dr["HoldPercentage"].ToString());
                this.SubItems.Add(dr["Accuracy"].ToString());
                this.SubItems.Add(dr["Incentive_Lines"].ToString());
                this.SubItems.Add(dr["Sunday_Shift_Allowance"].ToString());
                this.SubItems.Add(dr["Extra_Support_Allowance"].ToString());

                this.SubItems.Add(dr["LinecountSalary"].ToString());
                this.SubItems.Add(dr["Nightshift_Allowance"].ToString());
                this.SubItems.Add(dr["Puntuality_Incentive"].ToString());
                this.SubItems.Add(dr["Total_ConvertedLines_New"].ToString());
                this.SubItems.Add(dr["ApproxSal"].ToString());
            }

            public ListSummery(string sTotal, string Totfiles, string TotMins, string TotLines, string ConLines, string NightLines, string SundayShiftLines, string ExtraSupportLines, string HoldPer, string sAccuracy, string IncentiveLines, string SundayShiftAllowance, string ExtraSupportAllowance, string sLinecountSal, string sNightAllow, string sPunctualityIncen, string TotalConvertedLines, string sAccount, string ApproxSal)
                : base()
            {
                this.Text = "TOTAL";
                this.SubItems.Add(Totfiles);
                this.SubItems.Add(TotMins);
                this.SubItems.Add(TotLines);
                this.SubItems.Add(ConLines);
                this.SubItems.Add(sAccount);
                this.SubItems.Add(NightLines);
                this.SubItems.Add(SundayShiftLines);
                this.SubItems.Add(ExtraSupportLines);
                this.SubItems.Add(HoldPer);
                this.SubItems.Add(sAccuracy);
                this.SubItems.Add(IncentiveLines);
                this.SubItems.Add(SundayShiftAllowance);
                this.SubItems.Add(ExtraSupportAllowance);
                this.SubItems.Add(sLinecountSal);
                this.SubItems.Add(sNightAllow);
                this.SubItems.Add(sPunctualityIncen);
                this.SubItems.Add(TotalConvertedLines);
                this.SubItems.Add(ApproxSal);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class ListItem_Discrepancy_Report : ListViewItem
        {
            public int iIsResolved;
            public DateTime DFILE_SUBMITTED_TIME;

            public ListItem_Discrepancy_Report(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_NAME].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_ENTERED_DATE].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_LINES_DECIMAL].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_COMMENT].ToString());

                iIsResolved = Convert.ToInt32(_dr[Framework.MASTER_DISCREPANCY.DISCREPANCY_IS_RESOLVED].ToString());
                DFILE_SUBMITTED_TIME = Convert.ToDateTime(_dr["file_submitted_time"].ToString());

                if (iIsResolved == 1)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#F7FE2E");
                }
                else if (iIsResolved == 0)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#FE9A2E");
                }
            }
        }

        public class ListIte_MappingDeatils : ListViewItem
        {
            public int PRODUCTION_ID;

            public ListIte_MappingDeatils(DataRow _dr, int i)
                : base()
            {
                int AllottedFile = Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(1));
                string AllottedMin = _dr["alloted"].ToString().Split('-').GetValue(0).ToString();
                int AchivedFile = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(1));
                string AchivedMins = _dr["achived"].ToString().Split('-').GetValue(0).ToString();
                int TotFile = AchivedFile + AllottedFile;
                int Totsecs = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(2)) + Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(2));
                string TotMins = sGetDuration(Totsecs);
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());

                this.Text = _dr["emp_id"].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(AllottedFile.ToString());
                this.SubItems.Add(AllottedMin);
                this.SubItems.Add(AchivedFile.ToString());
                this.SubItems.Add(AchivedMins);
                this.SubItems.Add(TotFile.ToString());
                this.SubItems.Add(TotMins);

                PRODUCTION_ID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            private string sGetDuration(int Seconds)
            {
                string sDuration = string.Empty;

                object oTotalDuration = null;
                oTotalDuration = Seconds;
                if (oTotalDuration != null)
                {
                    int iTotalSeconds = Convert.ToInt32(oTotalDuration);
                    int iMinutes = 0;
                    int iSeconds = 0;
                    if (iTotalSeconds > 0)
                    {
                        iMinutes = iTotalSeconds / 60;
                        iSeconds = iTotalSeconds % 60;
                    }
                    sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
                }
                return sDuration;
            }
        }

        public class ListItem_OfflineAccountVolume : ListViewItem
        {
            public int client_id = 0;
            public int TotSec = 0;
            public ListItem_OfflineAccountVolume(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr["Tot_downloaded_file"].ToString());
                this.SubItems.Add(_dr["Tot_minutes"].ToString());


                client_id = Convert.ToInt32(_dr["Client_id"]);
                TotSec = Convert.ToInt32(_dr["Tot_sec"]);

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_OfflineAccountVolume(string TotFiles, string sTotMins)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(TotFiles);
                this.SubItems.Add(sTotMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class ListItem_OfflineLocationVolume : ListViewItem
        {
            public int client_id = 0;
            public string location_id = string.Empty;
            public int Doctor_id = 0;
            public ListItem_OfflineLocationVolume(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_dr["Tot_downloaded_file"].ToString());
                this.SubItems.Add(_dr["Processed_files"].ToString());
                this.SubItems.Add(_dr["Pending_files"].ToString());
                this.SubItems.Add(_dr["Tot_minutes"].ToString());
                this.SubItems.Add(_dr["FileLines"].ToString());

                client_id = Convert.ToInt32(_dr["Client_id"]);
                location_id = _dr["location_id"].ToString();

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_OfflineDoctorVolume : ListViewItem
        {
            public int Doctor_id = 0;
            public string location_id = string.Empty;
            public ListItem_OfflineDoctorVolume(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString();
                this.SubItems.Add(_dr["Tot_downloaded_file"].ToString());
                this.SubItems.Add(_dr["Processed_files"].ToString());
                this.SubItems.Add(_dr["Pending_files"].ToString());
                this.SubItems.Add(_dr["Tot_minutes"].ToString());
                this.SubItems.Add(_dr["FileLines"].ToString());

                Doctor_id = Convert.ToInt32(_dr["Doctor_id"]);
                location_id = _dr["location_id"].ToString();

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_OfflinePendingFiles : ListViewItem
        {
            public string location_id = string.Empty;
            public ListItem_OfflinePendingFiles(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();
                this.SubItems.Add(_dr["minutes"].ToString());
                this.SubItems.Add(_dr["TAT"].ToString());
                this.SubItems.Add(_dr["Bal_Tat"].ToString());
                this.SubItems.Add(_dr["file_lines"].ToString());
                this.SubItems.Add(_dr["alloted_for"].ToString());
                this.SubItems.Add(_dr["Trans"].ToString());
                this.SubItems.Add(_dr["Edit"].ToString());
                this.SubItems.Add(_dr["Hold"].ToString());
                this.SubItems.Add(_dr["file_status_description"].ToString());

                location_id = _dr["location_id"].ToString();

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_AccountVolume : ListViewItem
        {
            public ListItem_AccountVolume(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_dr["TotFiles"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_dr["FileLines"].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_AccountVolume(string sMessage, string sTotMins, string sTotLines)
                : base()
            {
                this.Text = sMessage;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sTotMins);
                this.SubItems.Add(sTotLines);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class ListItem_TatPercentage : ListViewItem
        {
            public decimal ONTAT;

            public ListItem_TatPercentage(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = _dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_dr["ClientTat"].ToString());
                this.SubItems.Add(_dr["TotalFiles"].ToString());
                this.SubItems.Add(_dr["OnTat"].ToString());
                this.SubItems.Add(_dr["OffTat"].ToString());
                this.SubItems.Add(_dr["OnTatPerc"].ToString());
                this.SubItems.Add(_dr["OffTatPerc"].ToString());

                this.SubItems.Add(_dr["Trans_Late"].ToString());
                this.SubItems.Add(_dr["Edit_Late"].ToString());
                this.SubItems.Add(_dr["Hold_Late"].ToString());
                this.SubItems.Add(_dr["Delivered_Late"].ToString());

                ONTAT = Convert.ToDecimal(_dr["ToCompute"]);

                if (Convert.ToInt32(ONTAT) == 100.00)
                {
                    this.BackColor = System.Drawing.Color.Green;
                    this.ForeColor = System.Drawing.Color.White;
                }
                else if ((Convert.ToInt32(ONTAT) >= 90.00) && ((Convert.ToInt32(ONTAT) <= 99.00)))
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#ff7d66");
                    this.ForeColor = System.Drawing.Color.Black;
                }
                else if ((Convert.ToInt32(ONTAT) >= 50.00) && ((Convert.ToInt32(ONTAT) <= 89.00)))
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#F7FE2E");
                    this.ForeColor = System.Drawing.Color.Black;
                }
                else
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#2ECCFA");
                    this.ForeColor = System.Drawing.Color.Black;
                }
            }

            public ListItem_TatPercentage(string sMessage, string sAvgOnTAT, string sAvgOffTAT)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(sMessage);
                this.SubItems.Add(sAvgOnTAT);
                this.SubItems.Add(sAvgOffTAT);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class ListItem_HoldPercentage : ListViewItem
        {
            public ListItem_HoldPercentage(DataRow _dr, int iRowCount)
                : base()
            {
                this.Text = _dr["HoldPercentage_IASIS"].ToString();

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_IdleSummary : ListViewItem
        {
            public int iProductionID;

            public int iTotalAllowed;
            public int iTotalIdle;

            public ListItem_IdleSummary(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow["user_name"].ToString();
                this.SubItems.Add(_drRow["Total_Idle"].ToString());
                this.SubItems.Add(_drRow["_TotalIdleAllowed"].ToString());

                iProductionID = Convert.ToInt32(_drRow["production_id"].ToString());
                iTotalIdle = Convert.ToInt32(_drRow["Total_Idle"].ToString());
                iTotalAllowed = Convert.ToInt32(_drRow["_TotalIdleAllowed"].ToString());

                if (iTotalIdle > iTotalAllowed)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF0000");
                    this.ForeColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
                }
                else
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
                }
            }
        }

        public class ListItem_IdleDateWise : ListViewItem
        {
            public ListItem_IdleDateWise(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow["emp_full_name"].ToString();
                this.SubItems.Add(_drRow["idle_start_time"].ToString());
                this.SubItems.Add(_drRow["idle_end_time"].ToString());
                this.SubItems.Add(_drRow["TotalIdle"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public ListItem_IdleDateWise(string TotalMinutes)
                : base()
            {
                this.Text = string.Empty;
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(string.Empty);
                this.SubItems.Add(TotalMinutes);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        /// <summary>
        /// GET CUSTOMIZED INCENTIVE MINS
        /// </summary>
        public class ListItemBlankCount : ListViewItem
        {
            public int IPRODUCTION_IDMT;
            public int IPRODUCTION_IDME;
            public int IPRODUCTION_IDQC;

            public int MT_Blank;
            public int ME_Blank;
            public int QC_Blank;


            public ListItemBlankCount(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["client_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["file_date"].ToString());

                SubItems.Add(dr["Transcibed_by"].ToString());
                SubItems.Add(dr["Trans_Lines"].ToString());
                SubItems.Add(dr["Transcibed_Status"].ToString());
                SubItems.Add(dr["Transcibed_Time"].ToString());
                if (dr["Trans_Blank"].ToString() == "-1")
                    SubItems.Add("NA");
                else
                    SubItems.Add(dr["Trans_Blank"].ToString());

                SubItems.Add(dr["TED_by"].ToString());
                SubItems.Add(dr["TED_Lines"].ToString());
                SubItems.Add(dr["TED_Status"].ToString());
                SubItems.Add(dr["TED_Time"].ToString());
                if (dr["TED_Blank"].ToString() == "-1")
                    SubItems.Add("NA");
                else
                    SubItems.Add(dr["TED_Blank"].ToString());


                SubItems.Add(dr["NDSP_by"].ToString());
                SubItems.Add(dr["NDSP_Lines"].ToString());
                SubItems.Add(dr["NDSP_Status"].ToString());
                SubItems.Add(dr["NDSP_Time"].ToString());
                if (dr["NDSP_Blank"].ToString() == "-1")
                    SubItems.Add("NA");
                else
                    SubItems.Add(dr["NDSP_Blank"].ToString());

                SubItems.Add(dr["Edit_by"].ToString());
                SubItems.Add(dr["Edit_Lines"].ToString());
                SubItems.Add(dr["Edit_Status"].ToString());
                SubItems.Add(dr["Edit_Time"].ToString());
                if (dr["Edit_Blank"].ToString() == "-1")
                    SubItems.Add("NA");
                else
                    SubItems.Add(dr["Edit_Blank"].ToString());

                SubItems.Add(dr["QC_by"].ToString());
                SubItems.Add(dr["QC_Lines"].ToString());
                SubItems.Add(dr["QC_Status"].ToString());
                SubItems.Add(dr["QC_Time"].ToString());
                if (dr["QC_Blank"].ToString() == "-1")
                    SubItems.Add("NA");
                else
                    SubItems.Add(dr["QC_Blank"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

        }

        public class ListItem_Multiple_Entries : ListViewItem
        {
            public string TranscriptionID = string.Empty;

            public ListItem_Multiple_Entries(DataRow _dr, int iCount)
                : base()
            {
                this.Text = _dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.REPORT_TRANSACTION.FIELD_IS_SUBMITTED_TIME].ToString());

                TranscriptionID = _dr["transcription_id"].ToString();

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItemConBlankCount : ListViewItem
        {
            public ListItemConBlankCount(DataRow dr, int iRowcount)
                : base()
            {
                this.Text = iRowcount.ToString();
                this.SubItems.Add(dr["location_name"].ToString());
                this.SubItems.Add(dr["Trans_Count"].ToString());
                this.SubItems.Add(dr["Trans_Lines"].ToString());
                this.SubItems.Add(dr["Trans_Blank"].ToString());
                this.SubItems.Add(dr["Trans_Percentage"].ToString() + "%");
                this.SubItems.Add(dr["TED_Count"].ToString());
                this.SubItems.Add(dr["TED_Lines"].ToString());
                this.SubItems.Add(dr["TED_Blank"].ToString());
                this.SubItems.Add(dr["TED_Percentage"].ToString() + "%");
                this.SubItems.Add(dr["NDSP_Count"].ToString());
                this.SubItems.Add(dr["NDSP_Lines"].ToString());
                this.SubItems.Add(dr["NDSP_Blank"].ToString());
                this.SubItems.Add(dr["NDSP_Percentage"].ToString() + "%");
                this.SubItems.Add(dr["Edit_Count"].ToString());
                this.SubItems.Add(dr["Edit_Lines"].ToString());
                this.SubItems.Add(dr["Edit_Blank"].ToString());
                this.SubItems.Add(dr["Edit_Percentage"].ToString() + "%");

                this.SubItems.Add(dr["Total_Files"].ToString());
                this.SubItems.Add(dr["Files_With_Blank"].ToString());
                this.SubItems.Add(dr["Files_Without_Blank"].ToString());
                this.SubItems.Add(dr["Pended_files"].ToString());
                this.SubItems.Add(dr["QC_Count"].ToString());
                this.SubItems.Add(dr["QC_Lines"].ToString());
                this.SubItems.Add(dr["QC_Blank"].ToString());
                this.SubItems.Add(dr["QC_Percentage"].ToString() + "%");

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
            public ListItemConBlankCount(string MTCount, string MTLines, string MTBlanks, string MTPercentage,
                string TedCount, string TedLines, string TedBlanks, string TedPercentage,
                string NDSPCount, string NDSPLines, string NDSpBlanks, string NDSPPercentage,
                string MECount, string MELines, string MEBlanks, string MEPercentage,
                string Total_Files, string Pended_files, string Files_With_Blank, string Files_Without_Blank,
                string QCCount, string QCLines, string QCBlanks, string QCPercentage)
                : base()
            {

                this.Text = "";
                this.SubItems.Add(" Total ");
                this.SubItems.Add(MTCount.Length > 0 ? MTCount : "0");
                this.SubItems.Add(MTLines.Length > 0 ? MTLines : "0");
                this.SubItems.Add(MTBlanks.Length > 0 ? MTBlanks : "0");
                this.SubItems.Add((MTPercentage.Length > 0 ? MTPercentage : "0") + "%");
                this.SubItems.Add(TedCount.Length > 0 ? TedCount : "0");
                this.SubItems.Add(TedLines.Length > 0 ? TedLines : "0");
                this.SubItems.Add(TedBlanks.Length > 0 ? TedBlanks : "0");
                this.SubItems.Add((TedPercentage.Length > 0 ? TedPercentage : "0") + "%");
                this.SubItems.Add(NDSPCount.Length > 0 ? NDSPCount : "0");
                this.SubItems.Add(NDSPLines.Length > 0 ? NDSPLines : "0");
                this.SubItems.Add(NDSpBlanks.Length > 0 ? NDSpBlanks : "0");
                this.SubItems.Add((NDSPPercentage.Length > 0 ? NDSPPercentage : "0") + "%");
                this.SubItems.Add(MECount.Length > 0 ? MECount : "0");
                this.SubItems.Add(MELines.Length > 0 ? MELines : "0");
                this.SubItems.Add(MEBlanks.Length > 0 ? MEBlanks : "0");
                this.SubItems.Add((MEPercentage.Length > 0 ? MEPercentage : "0") + "%");
                this.SubItems.Add(Total_Files.Length > 0 ? Total_Files : "0");
                this.SubItems.Add(Pended_files.Length > 0 ? Pended_files : "0");
                this.SubItems.Add(Files_With_Blank.Length > 0 ? Files_With_Blank : "0");
                this.SubItems.Add(Files_Without_Blank.Length > 0 ? Files_Without_Blank : "0");
                this.SubItems.Add(QCCount.Length > 0 ? QCCount : "0");
                this.SubItems.Add(QCLines.Length > 0 ? QCLines : "0");
                this.SubItems.Add(QCBlanks.Length > 0 ? QCBlanks : "0");
                this.SubItems.Add((QCPercentage.Length > 0 ? QCPercentage : "0") + "%");

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }


        public class ListItem_ProductivityDeatils : ListViewItem
        {
            public ListItem_ProductivityDeatils(DataRow _dr, int iCount, int sNo)
                : base()
            {
                this.Text = _dr["hr_id"].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr["internal_mail"].ToString());
                this.SubItems.Add(_dr["ptag_id"].ToString());
                this.SubItems.Add(_dr["Total_Files"].ToString());
                this.SubItems.Add(_dr["TotalLines"].ToString());

                if (iCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_Marquee : ListViewItem
        {
            public int iSTATUS;
            public int iSCROLLID;
            public string sScrollText;
            public int iAdded_by;
            public string slogin_name;
            public string sCP_Number;
            public ListItem_Marquee(DataRow _dr, int sNo)
                : base()
            {
                this.Text = sNo.ToString();

                this.SubItems.Add(_dr["Scroll_text"].ToString());
                this.SubItems.Add(_dr["login_name"].ToString());

                iSTATUS = Convert.ToInt32(_dr["Is_Active"]);
                iSCROLLID = Convert.ToInt32(_dr["Scroll_id"]);
                sScrollText = _dr["Scroll_text"].ToString();
                iAdded_by = Convert.ToInt32(_dr["Added_by"]);
                slogin_name = _dr["login_name"].ToString();
                sCP_Number = _dr["CP_NUMBER"].ToString();
            }
        }

        public class MyListItemUserwise : ListViewItem
        {
            public MyListItemUserwise(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["user_name"].ToString());
                SubItems.Add(dr["Details"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

        }

        public class MyListVoiceLogItem : ListViewItem
        {
            public MyListVoiceLogItem(DataRow drlog, int iRowCount)
            {
                this.Name = drlog["hosp_id"].ToString();
                this.Text = drlog["host_name"].ToString().ToUpper();
                this.SubItems.Add(drlog["DAY1"].ToString());
                this.SubItems.Add(drlog["Day1Minutes"].ToString());
                this.SubItems.Add(drlog["DAY2"].ToString());
                this.SubItems.Add(drlog["Day2Minutes"].ToString());
                this.SubItems.Add(drlog["DAY3"].ToString());
                this.SubItems.Add(drlog["Day3Minutes"].ToString());
                this.SubItems.Add(drlog["DAY4"].ToString());
                this.SubItems.Add(drlog["Day4Minutes"].ToString());
                this.SubItems.Add(drlog["DAY5"].ToString());
                this.SubItems.Add(drlog["Day5Minutes"].ToString());
                this.SubItems.Add(drlog["DAY6"].ToString());
                this.SubItems.Add(drlog["Day6Minutes"].ToString());
                this.SubItems.Add(drlog["DAY7"].ToString());
                this.SubItems.Add(drlog["Day7Minutes"].ToString());
                this.SubItems.Add(drlog["DAY8"].ToString());
                this.SubItems.Add(drlog["Day8Minutes"].ToString());
                this.SubItems.Add(drlog["DAY9"].ToString());
                this.SubItems.Add(drlog["Day9Minutes"].ToString());
                this.SubItems.Add(drlog["DAY10"].ToString());
                this.SubItems.Add(drlog["Day10Minutes"].ToString());
                this.SubItems.Add(drlog["DAY11"].ToString());
                this.SubItems.Add(drlog["Day11Minutes"].ToString());
                this.SubItems.Add(drlog["DAY12"].ToString());
                this.SubItems.Add(drlog["Day12Minutes"].ToString());
                this.SubItems.Add(drlog["DAY13"].ToString());
                this.SubItems.Add(drlog["Day13Minutes"].ToString());
                this.SubItems.Add(drlog["DAY14"].ToString());
                this.SubItems.Add(drlog["Day14Minutes"].ToString());
                this.SubItems.Add(drlog["DAY15"].ToString());
                this.SubItems.Add(drlog["Day15Minutes"].ToString());
                this.SubItems.Add(drlog["DAY16"].ToString());
                this.SubItems.Add(drlog["Day16Minutes"].ToString());
                this.SubItems.Add(drlog["DAY17"].ToString());
                this.SubItems.Add(drlog["Day17Minutes"].ToString());
                this.SubItems.Add(drlog["DAY18"].ToString());
                this.SubItems.Add(drlog["Day18Minutes"].ToString());
                this.SubItems.Add(drlog["DAY19"].ToString());
                this.SubItems.Add(drlog["Day19Minutes"].ToString());
                this.SubItems.Add(drlog["DAY20"].ToString());
                this.SubItems.Add(drlog["Day20Minutes"].ToString());
                this.SubItems.Add(drlog["DAY21"].ToString());
                this.SubItems.Add(drlog["Day21Minutes"].ToString());
                this.SubItems.Add(drlog["DAY22"].ToString());
                this.SubItems.Add(drlog["Day22Minutes"].ToString());
                this.SubItems.Add(drlog["DAY23"].ToString());
                this.SubItems.Add(drlog["Day23Minutes"].ToString());
                this.SubItems.Add(drlog["DAY24"].ToString());
                this.SubItems.Add(drlog["Day24Minutes"].ToString());
                this.SubItems.Add(drlog["DAY25"].ToString());
                this.SubItems.Add(drlog["Day25Minutes"].ToString());
                this.SubItems.Add(drlog["DAY26"].ToString());
                this.SubItems.Add(drlog["Day26Minutes"].ToString());
                this.SubItems.Add(drlog["DAY27"].ToString());
                this.SubItems.Add(drlog["Day27Minutes"].ToString());
                this.SubItems.Add(drlog["DAY28"].ToString());
                this.SubItems.Add(drlog["Day28Minutes"].ToString());
                this.SubItems.Add(drlog["DAY29"].ToString());
                this.SubItems.Add(drlog["Day29Minutes"].ToString());
                this.SubItems.Add(drlog["DAY30"].ToString());
                this.SubItems.Add(drlog["Day30Minutes"].ToString());
                this.SubItems.Add(drlog["DAY31"].ToString());
                this.SubItems.Add(drlog["Day31Minutes"].ToString());
                int Total_files = Convert.ToInt32(drlog["DAY1"]) + Convert.ToInt32(drlog["DAY1"]) + Convert.ToInt32(drlog["DAY2"]) + Convert.ToInt32(drlog["DAY3"]) +
                    Convert.ToInt32(drlog["DAY4"]) + Convert.ToInt32(drlog["DAY5"]) + Convert.ToInt32(drlog["DAY6"]) + Convert.ToInt32(drlog["DAY7"]) +
                    Convert.ToInt32(drlog["DAY8"]) + Convert.ToInt32(drlog["DAY9"]) + Convert.ToInt32(drlog["DAY10"]) + Convert.ToInt32(drlog["DAY11"]) +
                    Convert.ToInt32(drlog["DAY12"]) + Convert.ToInt32(drlog["DAY13"]) + Convert.ToInt32(drlog["DAY14"]) + Convert.ToInt32(drlog["DAY15"]) +
                    Convert.ToInt32(drlog["DAY16"]) + Convert.ToInt32(drlog["DAY17"]) + Convert.ToInt32(drlog["DAY18"]) + Convert.ToInt32(drlog["DAY19"]) +
                    Convert.ToInt32(drlog["DAY20"]) + Convert.ToInt32(drlog["DAY21"]) + Convert.ToInt32(drlog["DAY22"]) + Convert.ToInt32(drlog["DAY23"]) +
                    Convert.ToInt32(drlog["DAY24"]) + Convert.ToInt32(drlog["DAY25"]) + Convert.ToInt32(drlog["DAY26"]) + Convert.ToInt32(drlog["DAY27"]) +
                    Convert.ToInt32(drlog["DAY28"]) + Convert.ToInt32(drlog["DAY29"]) + Convert.ToInt32(drlog["DAY30"]) + Convert.ToInt32(drlog["DAY31"]);
                Color ForeColor = SystemColors.ButtonShadow;
                Color BackColor = SystemColors.MenuHighlight;
                Font font = null;
                this.SubItems.Add(Total_files.ToString(), ForeColor, BackColor, font);

                int Total_Minutes = Convert.ToInt32(drlog["Day1Seconds"]) + Convert.ToInt32(drlog["Day2Seconds"]) + Convert.ToInt32(drlog["Day3Seconds"]) +
                    Convert.ToInt32(drlog["Day4Seconds"]) + Convert.ToInt32(drlog["Day5Seconds"]) + Convert.ToInt32(drlog["Day6Seconds"]) + Convert.ToInt32(drlog["Day7Seconds"]) +
                    Convert.ToInt32(drlog["Day8Seconds"]) + Convert.ToInt32(drlog["Day9Seconds"]) + Convert.ToInt32(drlog["Day10Seconds"]) + Convert.ToInt32(drlog["Day11Seconds"]) +
                    Convert.ToInt32(drlog["Day12Seconds"]) + Convert.ToInt32(drlog["Day13Seconds"]) + Convert.ToInt32(drlog["Day14Seconds"]) + Convert.ToInt32(drlog["Day15Seconds"]) +
                    Convert.ToInt32(drlog["Day16Seconds"]) + Convert.ToInt32(drlog["Day17Seconds"]) + Convert.ToInt32(drlog["Day18Seconds"]) + Convert.ToInt32(drlog["Day19Seconds"]) +
                    Convert.ToInt32(drlog["Day20Seconds"]) + Convert.ToInt32(drlog["Day21Seconds"]) + Convert.ToInt32(drlog["Day22Seconds"]) + Convert.ToInt32(drlog["Day23Seconds"]) +
                    Convert.ToInt32(drlog["Day24Seconds"]) + Convert.ToInt32(drlog["Day25Seconds"]) + Convert.ToInt32(drlog["Day26Seconds"]) + Convert.ToInt32(drlog["Day27Seconds"]) +
                    Convert.ToInt32(drlog["Day28Seconds"]) + Convert.ToInt32(drlog["Day29Seconds"]) + Convert.ToInt32(drlog["Day30Seconds"]) + Convert.ToInt32(drlog["Day31Seconds"]);
                this.SubItems.Add(GiveMinutes(Total_Minutes.ToString()), ForeColor, BackColor, font);

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public MyListVoiceLogItem(
                string Day1FileCount, string Day1TotalMinues,
                string Day2FileCount, string Day2TotalMinues,
                string Day3FileCount, string Day3TotalMinues,
                string Day4FileCount, string Day4TotalMinues,
                string Day5FileCount, string Day5TotalMinues,
                string Day6FileCount, string Day6TotalMinues,
                string Day7FileCount, string Day7TotalMinues,
                string Day8FileCount, string Day8TotalMinues,
                string Day9FileCount, string Day9TotalMinues,
                string Day10FileCount, string Day10TotalMinues,
                string Day11FileCount, string Day11TotalMinues,
                string Day12FileCount, string Day12TotalMinues,
                string Day13FileCount, string Day13TotalMinues,
                string Day14FileCount, string Day14TotalMinues,
                string Day15FileCount, string Day15TotalMinues,
                string Day16FileCount, string Day16TotalMinues,
                string Day17FileCount, string Day17TotalMinues,
                string Day18FileCount, string Day18TotalMinues,
                string Day19FileCount, string Day19TotalMinues,
                string Day20FileCount, string Day20TotalMinues,
                string Day21FileCount, string Day21TotalMinues,
                string Day22FileCount, string Day22TotalMinues,
                string Day23FileCount, string Day23TotalMinues,
                string Day24FileCount, string Day24TotalMinues,
                string Day25FileCount, string Day25TotalMinues,
                string Day26FileCount, string Day26TotalMinues,
                string Day27FileCount, string Day27TotalMinues,
                string Day28FileCount, string Day28TotalMinues,
                string Day29FileCount, string Day29TotalMinues,
                string Day30FileCount, string Day30TotalMinues,
                string Day31FileCount, string Day31TotalMinues)

                : base()
            {
                this.Text = " Total ";
                this.SubItems.Add(Day1FileCount);
                this.SubItems.Add(GiveMinutes(Day1TotalMinues));
                this.SubItems.Add(Day2FileCount);
                this.SubItems.Add(GiveMinutes(Day2TotalMinues));
                this.SubItems.Add(Day3FileCount);
                this.SubItems.Add(GiveMinutes(Day3TotalMinues));
                this.SubItems.Add(Day4FileCount);
                this.SubItems.Add(GiveMinutes(Day4TotalMinues));
                this.SubItems.Add(Day5FileCount);
                this.SubItems.Add(GiveMinutes(Day5TotalMinues));
                this.SubItems.Add(Day6FileCount);
                this.SubItems.Add(GiveMinutes(Day6TotalMinues));
                this.SubItems.Add(Day7FileCount);
                this.SubItems.Add(GiveMinutes(Day7TotalMinues));
                this.SubItems.Add(Day8FileCount);
                this.SubItems.Add(GiveMinutes(Day8TotalMinues));
                this.SubItems.Add(Day9FileCount);
                this.SubItems.Add(GiveMinutes(Day9TotalMinues));
                this.SubItems.Add(Day10FileCount);
                this.SubItems.Add(GiveMinutes(Day10TotalMinues));
                this.SubItems.Add(Day11FileCount);
                this.SubItems.Add(GiveMinutes(Day11TotalMinues));
                this.SubItems.Add(Day12FileCount);
                this.SubItems.Add(GiveMinutes(Day12TotalMinues));
                this.SubItems.Add(Day13FileCount);
                this.SubItems.Add(GiveMinutes(Day13TotalMinues));
                this.SubItems.Add(Day14FileCount);
                this.SubItems.Add(GiveMinutes(Day14TotalMinues));
                this.SubItems.Add(Day15FileCount);
                this.SubItems.Add(GiveMinutes(Day15TotalMinues));
                this.SubItems.Add(Day16FileCount);
                this.SubItems.Add(GiveMinutes(Day16TotalMinues));
                this.SubItems.Add(Day17FileCount);
                this.SubItems.Add(GiveMinutes(Day17TotalMinues));
                this.SubItems.Add(Day18FileCount);
                this.SubItems.Add(GiveMinutes(Day18TotalMinues));
                this.SubItems.Add(Day19FileCount);
                this.SubItems.Add(GiveMinutes(Day19TotalMinues));
                this.SubItems.Add(Day20FileCount);
                this.SubItems.Add(GiveMinutes(Day20TotalMinues));
                this.SubItems.Add(Day21FileCount);
                this.SubItems.Add(GiveMinutes(Day21TotalMinues));
                this.SubItems.Add(Day22FileCount);
                this.SubItems.Add(GiveMinutes(Day22TotalMinues));
                this.SubItems.Add(Day23FileCount);
                this.SubItems.Add(GiveMinutes(Day23TotalMinues));
                this.SubItems.Add(Day24FileCount);
                this.SubItems.Add(GiveMinutes(Day24TotalMinues));
                this.SubItems.Add(Day25FileCount);
                this.SubItems.Add(GiveMinutes(Day25TotalMinues));
                this.SubItems.Add(Day26FileCount);
                this.SubItems.Add(GiveMinutes(Day26TotalMinues));
                this.SubItems.Add(Day27FileCount);
                this.SubItems.Add(GiveMinutes(Day27TotalMinues));
                this.SubItems.Add(Day28FileCount);
                this.SubItems.Add(GiveMinutes(Day28TotalMinues));
                this.SubItems.Add(Day29FileCount);
                this.SubItems.Add(GiveMinutes(Day29TotalMinues));
                this.SubItems.Add(Day30FileCount);
                this.SubItems.Add(GiveMinutes(Day30TotalMinues));
                this.SubItems.Add(Day31FileCount);
                this.SubItems.Add(GiveMinutes(Day31TotalMinues));

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class  listItem_ConsolidatedMonth: ListViewItem
        {
            public listItem_ConsolidatedMonth(DataRow _dr, int i)
                : base()
            {
                this.Name = _dr["details"].ToString();
                this.Text = _dr["date"].ToString();
                this.SubItems.Add(_dr["target_mins"].ToString());
                this.SubItems.Add(_dr["Tot_mins"].ToString());
                this.SubItems.Add(_dr["total_lines"].ToString());
                this.SubItems.Add(_dr["actual_shift"].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");                
            }
        }

        public class listItem_ConsolidatedYear : ListViewItem
        {
            public listItem_ConsolidatedYear(DataRow _dr, int i)
                : base()
            {
                this.Name = _dr["details"].ToString();
                this.Text = _dr["month"].ToString();
                this.SubItems.Add(_dr["target_mins"].ToString());
                this.SubItems.Add(_dr["target_lines"].ToString());
                this.SubItems.Add(_dr["Tot_mins"].ToString());
                this.SubItems.Add(_dr["total_lines"].ToString());
                this.SubItems.Add(_dr["accuracy"].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }
        


        #endregion "CLASS"

        #region " METHODS "

        #region " BPM WEEKLY MINS "

        #endregion " BPM WEEKLY MINS "


        /// <summary>
        /// LOAD ALL MTS LIST
        /// </summary>
        private void Load_MEFiles_List()
        {
            try
            {
                int iRows = 1;
                DataSet dsMEFiles_List = BusinessLogic.WS_Allocation.Get_Editedfiles_List( Convert.ToDateTime(dtp_TATMonitor_MEFromdate.Value), Convert.ToDateTime(dtp_TATMonitor_METodate.Value));
                if (dsMEFiles_List != null)
                {
                    if (dsMEFiles_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_TATMonitor_ME.Items.Clear();
                        foreach (DataRow dr in dsMEFiles_List.Tables[0].Rows)
                            lsv_TATMonitor_ME.Items.Add(new MyListItem_QuickAllot_ME_TAT(dr, iRows++));

                        BusinessLogic.Reset_ListViewColumn(lsv_TATMonitor_ME);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        /// <summary>
        /// LOAD ALL MTS LIST
        /// </summary>
        private void Load_MTFiles_List()
        {
            try
            {
                int iRows = 1;
                DataSet dsMTFiles_List = BusinessLogic.WS_Allocation.Get_Transfiles_List(Convert.ToDateTime(dtp_TATMonitor_Fromdate.Value), Convert.ToDateTime(dtp_TATMonitor_Todate.Value));
                if (dsMTFiles_List != null)
                {
                    if (dsMTFiles_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_Quick_MT_TAT.Items.Clear();
                        foreach (DataRow dr in dsMTFiles_List.Tables[0].Rows)
                            lsv_Quick_MT_TAT.Items.Add(new MyListItem_QuickAllot_MT_TAT(dr, iRows++));

                        BusinessLogic.Reset_ListViewColumn(lsv_Quick_MT_TAT);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        /// <summary>
        /// Load Employee Names list to allocate the files
        /// </summary>
        private void Load_Emp_List()
        {
            try
            {
                double dTotmins = 0;
                int iTotalFiles = 0;
                MyListItem_QuickAllot_MT_TAT oItem = (MyListItem_QuickAllot_MT_TAT)lsv_Quick_MT_TAT.SelectedItems[0];
                foreach (MyListItem_QuickAllot_MT_TAT oFile in lsv_Quick_MT_TAT.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.SDURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.ICLIENT_ID), Convert.ToInt32(oItem.IDOCTOR_ID), iTotalFiles, Convert.ToInt32(dTotmins), 1, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyListItem_QuickAllot_MT_TAT oFile in lsv_Quick_MT_TAT.SelectedItems)
                    {
                        oFile.SSTATUS = "Allotted";
                        oFile.SEMP_NAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.SUSER_ID = BusinessLogic.ALLOTEDUSERID;

                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;


                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.ITRANSCRIPTION_ID, oFile.SUSER_ID, DateTime.Now, oFile.SVOICE_FILE_ID, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_Quick_MT_TAT.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
            finally
            {
                Load_MTFiles_List();
            }
        }

        /// <summary>
        /// EMP CONSOLIDATAION
        /// </summary>
        private void Load_Emp_Consolidation()
        {
            try
            {
                DataSet dsConsolidation = BusinessLogic.WS_Allocation.Get_All_Employees_Consolidation(Convert.ToInt32(cmb_Emp_Branch.SelectedValue), Convert.ToInt32(cmb_Emp_Workplatform.SelectedValue), Convert.ToInt32(cmb_Emp_Batch.SelectedValue));
                if (dsConsolidation != null)
                {
                    if (dsConsolidation.Tables[0].Rows.Count > 0)
                    {
                        lsv_Emp_Consolidated.Items.Clear();
                        foreach (DataRow dr in dsConsolidation.Tables[0].Rows)
                            lsv_Emp_Consolidated.Items.Add(new MyListItem_EmpConsolidation(dr));

                        BusinessLogic.Reset_ListViewColumn(lsv_Emp_Consolidated);
                    }
                    else
                        lsv_Emp_Consolidated.Items.Clear();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Work_Platform()
        {
            try
            {
                DataTable dtEmp = new DataTable();
                dtEmp.Columns.Add("slno");
                dtEmp.Columns.Add("Name");

                DataRow dr = dtEmp.NewRow();
                dr["slno"] = -1;
                dr["Name"] = "--Select--";
                dtEmp.Rows.InsertAt(dr, 0);

                DataRow dr1 = dtEmp.NewRow();
                dr1["slno"] = 1;
                dr1["Name"] = "NTS";
                dtEmp.Rows.InsertAt(dr1, 1);

                DataRow dr2 = dtEmp.NewRow();
                dr2["slno"] = 2;
                dr2["Name"] = "clinics";
                dtEmp.Rows.InsertAt(dr2, 2);

                cmb_Emp_Workplatform.DataSource = dtEmp;
                cmb_Emp_Workplatform.ValueMember = "slno";
                cmb_Emp_Workplatform.DisplayMember = "Name";
                cmb_Emp_Workplatform.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// GET DOCTOR WISE FILE COUNT
        /// </summary>
        private void Load_Doctorwise_filecount()
        {
            try
            {
                DataSet dsDoctor = BusinessLogic.WS_Allocation.Get_AutoAllocation_Doctorwise_Total(cmbLocation.SelectedValue.ToString(), Convert.ToDateTime(dtp_Filedate.Value), Convert.ToDateTime(dtp_FileTodate.Value));
                if (dsDoctor != null)
                {
                    if (dsDoctor.Tables[0].Rows.Count > 0)
                    {
                        lsvAuto_doctorfiles.Items.Clear();
                        foreach (DataRow dr in dsDoctor.Tables[0].Rows)
                            lsvAuto_doctorfiles.Items.Add(new MyListItem_DoctorwiseTotal(dr));

                        BusinessLogic.Reset_ListViewColumn(lsvAuto_doctorfiles);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// GET WEEKLY PROCESSED MINUTES
        /// </summary>
        /// 
        private void Load_Weekly_Mins()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring data..!");
                BusinessLogic.oProgressEvent.Start(true);
                int iTotal_Files = 0;
                int iTotal_Minutes = 0;
                DateTime dStart_date = DateTime.Now;
                List<BusinessLogic.WEEKLY_PROCESSED_MINS> oWeek = new List<BusinessLogic.WEEKLY_PROCESSED_MINS>();

                int iRows = 1;
                DataSet dsWeekly = BusinessLogic.WS_Allocation.Get_Weekly_Processed_mins(Convert.ToDateTime(dtp_Week_Fromdate.Value), Convert.ToDateTime(dtp_Week_Todate.Value));
                if (dsWeekly != null)
                {
                    if (dsWeekly.Tables[0].Rows.Count > 0)
                    {

                        foreach (DataRow dr in dsWeekly.Tables[0].Rows)
                        {
                            oWeek.Add(new BusinessLogic.WEEKLY_PROCESSED_MINS(iRows, Convert.ToDateTime(dr["week_start"].ToString()), Convert.ToDateTime(dr["week_end"].ToString()), Convert.ToInt32(dr["filecount"].ToString()), dr["Tot_minutes"].ToString()));
                            iRows += iRows;
                        }

                        lsv_WeeklyProcessedMins.Items.Clear();

                        foreach (DataRow dr in dsWeekly.Tables[0].Rows)
                        {
                            lsv_WeeklyProcessedMins.Items.Add(new MyListItem_WeeklyProcessedmins(dr, iRows++));

                            if (dStart_date != Convert.ToDateTime(dr["week_start"].ToString()))
                            {
                                dStart_date = Convert.ToDateTime(dr["week_start"].ToString());
                                //var Tot_files = (from c in oWeek group c by dStart_date into g select new { totfiles = g.Sum(x => x.TOT_FILES) });                                
                                var Tot_files = oWeek.GroupBy(i => i.START_DATE).Select(i => new
                                {
                                    date = i.Key,
                                    no_of_files = i.Where(j => j.START_DATE == dStart_date).Sum(k => k.TOT_FILES),
                                    no_of_minutes = i.Where(j => j.START_DATE == dStart_date).Sum(k => Convert.ToInt32(k.FILE_MINUTES)),
                                });

                                var Tot_Mins = (from c in oWeek group c by dStart_date into g select new { Tot_minutes = g.Sum(x => Convert.ToDecimal(x.FILE_MINUTES)) });

                                foreach (var totfiles in Tot_files)
                                {
                                    iTotal_Files += Convert.ToInt32(totfiles.no_of_files);
                                    iTotal_Minutes += Convert.ToInt32(totfiles.no_of_minutes);
                                }
                            }
                            else
                            {
                                lsv_WeeklyProcessedMins.Items.Add(new MyListItem_WeeklyProcessedmins(iTotal_Files.ToString(), sGetDuration(iTotal_Minutes)));
                                iTotal_Files = 0;
                                iTotal_Minutes = 0;
                            }
                        }
                        BusinessLogic.Reset_ListViewColumn(lsv_WeeklyProcessedMins);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done..!");
                BusinessLogic.oProgressEvent.Start(true);
            }
        }

        private void Load_Location_Type(int iClienttypeid)
        {
            try
            {
                DataSet dsLocation = BusinessLogic.WS_Allocation.Get_Location_Type(iClienttypeid);
                if (dsLocation != null)
                {
                    if (dsLocation.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = dsLocation.Tables[0].NewRow();
                        dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR] = " -- Select -- ";
                        dr[Framework.LOCATION.FIELD_LOCATION_ID_STR] = "-1";
                        dsLocation.Tables[0].Rows.InsertAt(dr, 0);

                        cmb_Hourly_Location.DataSource = dsLocation.Tables[0];
                        cmb_Hourly_Location.DisplayMember = Framework.LOCATION.FIELD_LOCATION_NAME_STR;
                        cmb_Hourly_Location.ValueMember = Framework.LOCATION.FIELD_LOCATION_ID_STR;
                        cmb_Hourly_Location.SelectedIndex = 0;
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// TL'S ACCOUNT WISE FILE STATUS
        /// </summary>
        private void Load_Acc_TL_File_Status()
        {
            try
            {
                DataSet dsFile_Status = BusinessLogic.WS_Allocation.Get_Acc_TL_File_Status();
                
                if (dsFile_Status != null)
                {
                    if (dsFile_Status.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = dsFile_Status.Tables[0].NewRow();
                        dr[Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT] = 0;
                        dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR] = "-- Select --";
                        dsFile_Status.Tables[0].Rows.InsertAt(dr, 0);

                        cmb_AccTL_Status.DataSource = dsFile_Status.Tables[0];
                        cmb_AccTL_Status.ValueMember = Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT;
                        cmb_AccTL_Status.DisplayMember = Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR;
                        cmb_AccTL_Status.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Acc_File_status()
        {
            try
            {
                DataSet dsFile_Status = BusinessLogic.WS_Allocation.Get_Acc_File_Status();
                if (dsFile_Status != null)
                {
                    if (dsFile_Status.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = dsFile_Status.Tables[0].NewRow();
                        dr[Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT] = 0;
                        dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR] = "-- Select --";
                        dsFile_Status.Tables[0].Rows.InsertAt(dr, 0);

                        cmb_Acc_Status.DataSource = dsFile_Status.Tables[0];
                        cmb_Acc_Status.ValueMember = Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT;
                        cmb_Acc_Status.DisplayMember = Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR;
                        cmb_Acc_Status.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Customized_group()
        {
            try
            {
                DataSet dsGroup = BusinessLogic.WS_Allocation.GET_CUSTOMIZED_GROUP_LIST();
                if (dsGroup != null)
                {
                    if (dsGroup.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = dsGroup.Tables[0].NewRow();
                        dr["group_id"] = "0";
                        dr["group_name"] = "-- Select --";
                        dsGroup.Tables[0].Rows.InsertAt(dr, 0);


                        cmb_Customize_Group.DataSource = dsGroup.Tables[0];
                        cmb_Customize_Group.ValueMember = "group_id";
                        cmb_Customize_Group.DisplayMember = "group_name";
                        cmb_Customize_Group.SelectedIndex = 0;

                        cmb_hourly_group.DataSource = dsGroup.Tables[0];
                        cmb_hourly_group.ValueMember = "group_id";
                        cmb_hourly_group.DisplayMember = "group_name";
                        cmb_hourly_group.SelectedIndex = 0;
                    }
                    else
                    {
                        cmb_Customize_Group.DataSource = null;
                        cmb_hourly_group.DataSource = null;
                    }

                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_HundredPercent_Doctors_List()
        {
            try
            {
                int iRowcount = 1;
                DataSet dsDoctor = BusinessLogic.WS_Allocation.Get_HundredPercent_Doctors_List();
                if (dsDoctor != null)
                {
                    if (dsDoctor.Tables[0].Rows.Count > 0)
                    {
                        lsv_Hundred_Percent.Items.Clear();
                        foreach (DataRow dr in dsDoctor.Tables[0].Rows)
                            lsv_Hundred_Percent.Items.Add(new ListItemHundred_Review(dr, iRowcount++));

                        BusinessLogic.Reset_ListViewColumn(lsv_Hundred_Percent);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Location_List()
        {
            try
            {
                DataSet dsLocation = BusinessLogic.WS_Allocation.Get_Location_List();
                if (dsLocation != null)
                {
                    if (dsLocation.Tables[0].Rows.Count > 0)
                    {
                        lsv_Hund_Location.Items.Clear();
                        foreach (DataRow dr in dsLocation.Tables[0].Rows)
                            lsv_Hund_Location.Items.Add(new ListItemHundred_Location(dr));

                        BusinessLogic.Reset_ListViewColumn(lsv_Hund_Location);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Doctor_List(string sLocation_id)
        {
            try
            {
                DataSet dsDoctor = BusinessLogic.WS_Allocation.Get_Locationwise_Doctor(sLocation_id);
                if (dsDoctor != null)
                {
                    if (dsDoctor.Tables[0].Rows.Count > 0)
                    {
                        lsv_Hund_Doctor.Items.Clear();
                        foreach (DataRow dr in dsDoctor.Tables[0].Rows)
                            lsv_Hund_Doctor.Items.Add(new ListItemHundred_Doctor(dr));

                        BusinessLogic.Reset_ListViewColumn(lsv_Hund_Doctor);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }

        /// <summary>
        /// REMOVE NIGHT SHIFT ALLWANCE
        /// </summary>
        private void Remove_MTMET_Users()
        {
            try
            {
                foreach (ListItemMTMETList oItem in lsv_MTMETOverall.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_MTMET_Inactive(Convert.ToInt32(oItem.IPRODUCTION_ID));

                    if (iUpdate > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Removed Successfully..!");
                        Load_MTMETList(1);
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Can't Remove..!");
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_MTMETList(int iOption_id)
        {
            try
            {
                int iRowCount = 1;
                DataSet dsMT_List = BusinessLogic.WS_Allocation.Get_MtMET_List(Convert.ToDateTime(dtp_MTMETDate.Value), iOption_id);
                if (dsMT_List != null)
                {
                    if (dsMT_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_MTMETOverall.Items.Clear();
                        foreach (DataRow dr in dsMT_List.Tables[0].Rows)
                        {
                            lsv_MTMETOverall.Items.Add(new ListItemMTMETList(dr, iRowCount++));
                        }
                        BusinessLogic.Reset_ListViewColumn(lsv_MTMETOverall);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD MT LIST IN THE COMBO BOX
        /// </summary>
        private void Load_MT_List()
        {
            try
            {
                DataSet dsMT_List = BusinessLogic.WS_Allocation.Get_MtList();
                if (dsMT_List != null)
                {
                    if (dsMT_List.Tables[0].Rows.Count > 0)
                    {
                        cmb_MTMET_Name.DataSource = dsMT_List.Tables[0];
                        cmb_MTMET_Name.DisplayMember = "emp_full_name";
                        cmb_MTMET_Name.ValueMember = "production_id";
                        cmb_MTMET_Name.SelectedIndex = 0;
                    }
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// REMOVE NIGHT SHIFT ALLWANCE
        /// </summary>
        private void Remove_Nightshift_Users()
        {
            try
            {
                foreach (ListItemNightshift oItem in lsv_Nightshift_Employee.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Nightshift_Users_Remove(Convert.ToInt32(oItem.IPRODUCTION_ID), Convert.ToDateTime(dtp_Nightshift_User.Value));

                    if (iUpdate > 0)
                    {
                        Load_Nightshift_User();
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD NIGHT SHIFT USERS LIST
        /// </summary>
        private void Load_Nightshift_User()
        {
            try
            {
                int iRowcount = 1;
                DataSet dsEmp_List = BusinessLogic.WS_Allocation.Get_Nightshift_Users(Convert.ToDateTime(dtp_Nightshift_User.Value));

                if (dsEmp_List != null)
                {
                    if (dsEmp_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_Nightshift_Employee.Items.Clear();
                        foreach (DataRow dr in dsEmp_List.Tables[0].Rows)
                        {
                            lsv_Nightshift_Employee.Items.Add(new ListItemNightshift(dr, iRowcount));
                            iRowcount++;
                        }
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Nightshift_Employee);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// SAVE NIGHT SHIFT USERS DETAILS
        /// </summary>
        private void Save_Nightshift_Users_List()
        {
            try
            {
                //Validation
                if (cmb_Nightshift_Branch.Text.Trim() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the branch Name");
                    cmb_Nightshift_Branch.Focus();
                    return;
                }

                //ListItemNightshift
                int iResult = BusinessLogic.WS_Allocation.Set_Nightshift_Users(Convert.ToInt32(cmb_Nightshift_EmpName.SelectedValue), Convert.ToDateTime(dtp_Nightshift_User.Value));

                if (iResult > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Added Sucessfully");
                    Load_Nightshift_User();
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Already Added");
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// REMOVE NIGHT SHIFT ALLWANCE
        /// </summary>
        private void Remove_Nightshift_Allowance()
        {
            try
            {
                foreach (Listitem_NightShift_Marked oItem in lvNightShiftMarked.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Remove_Nightshift_Allowance(Convert.ToInt32(oItem.INIGHTSHIFT_TRANS_ID), Convert.ToInt32(oItem.IPRODUCTION_ID), Convert.ToInt32(BusinessLogic.SPRODUCTIONID), Convert.ToInt32(oItem.ICATEGORY_ID));

                    if (iUpdate > 0)
                    {
                        Load_Night_Shift_Marking();
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        /// <summary>
        /// UPDATE NIGHT SHIFT ALLOOWANCE
        /// </summary>
        private void Update_Nightshift_Allowance()
        {
            try
            {
                foreach (Listitem_NightShift_Marked oItem in lvNightShiftMarked.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Change_Nightshift_Category(Convert.ToInt32(cmb_Category.SelectedValue), Convert.ToInt32(oItem.INIGHTSHIFT_TRANS_ID), Convert.ToInt32(oItem.IPRODUCTION_ID));

                    if (iUpdate > 0)
                    {
                        Load_Night_Shift_Marking();
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD CATEGORY DETAILS
        /// </summary>
        private void Load_Category_Details()
        {
            try
            {
                DataSet _dsCategory = BusinessLogic.WS_Allocation.Get_NightShift_Category();
                if (_dsCategory != null)
                {
                    if (_dsCategory.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = _dsCategory.Tables[0].NewRow();
                        dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_ID] = 0;
                        dr[Framework.NIGHT_SHIFT_CATEGORY.FIELD_NIGHT_SHIFT_CATEGORY] = "-- All --";
                        _dsCategory.Tables[0].Rows.InsertAt(dr, 0);

                        cmb_Category.DataSource = _dsCategory.Tables[0];
                        cmb_Category.DisplayMember = "night_shift_category";
                        cmb_Category.ValueMember = "night_shift_inc_id";
                        cmb_Category.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Incentive_View_New()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data.....");
                DateTime dtInc_Date = DateTime.Now;
                DataSet dsIncentive_Dates = BusinessLogic.WS_Allocation.Get_Incentive_dates(Convert.ToDateTime(dtp_incentive_fromdate.Value), Convert.ToDateTime(dtp_incentive_todate.Value));
                DataSet dsIncentive_Shift = BusinessLogic.WS_Allocation.Get_Incentive_shift_V2();
                if (dsIncentive_Dates != null)
                {
                    if (dsIncentive_Dates.Tables[0].Rows.Count > 0)
                    {
                        lsv_Incentive.Items.Clear();
                        foreach (DataRow drDate in dsIncentive_Dates.Tables[0].Rows)
                        {
                            DataSet dsIncentive = BusinessLogic.WS_Allocation.Get_Incentive_Calculation_New(Convert.ToDateTime(drDate["incentive_date"].ToString()), Convert.ToDateTime(drDate["incentive_date"].ToString()), (chb_Night.Checked == true ? 1 : 0));
                            if (dsIncentive != null)
                            {
                                if (dsIncentive.Tables[0].Rows.Count > 0)
                                {
                                    int iRowcount = 1;
                                    DataTable dt = new DataTable();
                                    dt.Columns.Add("ID");
                                    dt.Columns.Add("Name");
                                    dt.Columns.Add("Target");
                                    dt.Columns.Add("File_Mins");
                                    dt.Columns.Add("Converted_Mins");
                                    dt.Columns.Add("30to39");
                                    dt.Columns.Add("40to49");
                                    dt.Columns.Add("50to59");
                                    dt.Columns.Add("Above60");
                                    dt.Columns.Add("Amount");
                                    dt.Columns.Add("incentive_date");

                                    foreach (DataRow dr in dsIncentive.Tables[0].Rows)
                                    {
                                        BusinessLogic.oMessageEvent.Start("Fetching datas!");
                                        foreach (DataRow dr_Shift in dsIncentive_Shift.Tables[0].Rows)
                                        {
                                            if ((Convert.ToDecimal((dr["Mins_Diff"].ToString().Replace(":", "."))) >= Convert.ToDecimal(dr_Shift["extraMinutesFrom"].ToString())) &&
                                                            (Convert.ToDecimal((dr["Mins_Diff"].ToString().Replace(":", "."))) < Convert.ToDecimal(dr_Shift["extraMinutesTo"].ToString())))
                                            {
                                                BusinessLogic.oMessageEvent.Start("Calculate the incentive amount");
                                                DataRow drInc = dt.NewRow();
                                                drInc["ID"] = dr["ptag_id"].ToString();
                                                drInc["Name"] = dr["emp_full_name"].ToString();
                                                drInc["Target"] = dr["target_mins"].ToString();
                                                drInc["File_Mins"] = dr["File_Mins"].ToString();
                                                drInc["Converted_Mins"] = dr["Converted_Mins"].ToString();
                                                drInc["30to39"] = dr_Shift["30to39"].ToString();
                                                drInc["40to49"] = dr_Shift["40to49"].ToString();
                                                drInc["50to59"] = dr_Shift["50to59"].ToString();
                                                drInc["Above60"] = dr_Shift["Above60"].ToString();
                                                drInc["Amount"] = dr_Shift["incentiveAmount"].ToString();
                                                drInc["incentive_date"] = drDate["incentive_date"].ToString();
                                                dt.Rows.Add(drInc);
                                            }
                                        }
                                        BusinessLogic.oMessageEvent.Start("Done...");
                                        BusinessLogic.Reset_ListViewColumn(lsv_Incentive);
                                    }
                                    foreach (DataRow _dr in dt.Rows)
                                    {
                                        if (dtInc_Date != Convert.ToDateTime(_dr["incentive_date"].ToString()))
                                        {
                                            BusinessLogic.oMessageEvent.Start("Adding Records");
                                            dtInc_Date = Convert.ToDateTime(_dr["incentive_date"].ToString());
                                            lsv_Incentive.Items.Add(new ListItemIncentive_Mins(string.Empty, string.Empty, "Incentive Date : " + dtInc_Date.ToString()));
                                        }
                                        lsv_Incentive.Items.Add(new ListItemIncentive_Mins(_dr, iRowcount++));

                                    }
                                }
                                else
                                {
                                    lsv_Incentive.Items.Clear();
                                    BusinessLogic.oMessageEvent.Start("No Data found..");
                                }
                            }
                            else
                            {
                                lsv_Incentive.Items.Clear();
                                BusinessLogic.oMessageEvent.Start("No Data found..");
                            }
                            BusinessLogic.oMessageEvent.Start("Done.");
                        }
                    }
                    else
                    {
                        lsv_Incentive.Items.Clear();
                        BusinessLogic.oMessageEvent.Start("No Data found..");
                    }

                }
                else
                {
                    lsv_Incentive.Items.Clear();
                    BusinessLogic.oMessageEvent.Start("No Data found..");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Incentive_View2()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..");
                DateTime dtInc_Date = DateTime.Now;
                DataSet dsIncentive = BusinessLogic.WS_Allocation.Get_Incentive_Calculation_V2(Convert.ToDateTime(dtp_incentive_fromdate.Value), Convert.ToDateTime(dtp_incentive_todate.Value), (chb_Night.Checked == true ? 1 : 0));
                //DataSet dsIncentive_Shift = BusinessLogic.WS_Allocation.Get_Incentive_Shift();
                DataSet dsIncentive_Shift = BusinessLogic.WS_Allocation.Get_Incentive_shift_V2();
                if (dsIncentive != null)
                {
                    if (dsIncentive.Tables[0].Rows.Count > 0)
                    {
                        lsv_Incentive.Items.Clear();
                        int iRowcount = 1;
                        DataTable dt = new DataTable();
                        dt.Columns.Add("ID");
                        dt.Columns.Add("Name");
                        dt.Columns.Add("Target");
                        dt.Columns.Add("File_Mins");
                        dt.Columns.Add("Converted_Mins");
                        dt.Columns.Add("30to39");
                        dt.Columns.Add("40to49");
                        dt.Columns.Add("50to59");
                        dt.Columns.Add("Above60");
                        dt.Columns.Add("Amount");
                        dt.Columns.Add("incentive_date");

                        foreach (DataRow dr in dsIncentive.Tables[0].Rows)
                        {
                            BusinessLogic.oMessageEvent.Start("Fetching datas!");
                            foreach (DataRow dr_Shift in dsIncentive_Shift.Tables[0].Rows)
                            {
                                if ((Convert.ToDecimal((dr["Mins_Different"].ToString().Replace(":", "."))) >= Convert.ToDecimal(dr_Shift["extraMinutesFrom"].ToString())) &&
                                                (Convert.ToDecimal((dr["Mins_Different"].ToString().Replace(":", "."))) < Convert.ToDecimal(dr_Shift["extraMinutesTo"].ToString())))
                                {
                                    BusinessLogic.oMessageEvent.Start("Calculate the incentive amount");
                                    DataRow drInc = dt.NewRow();
                                    drInc["ID"] = dr["ptag_id"].ToString();
                                    drInc["Name"] = dr["emp_full_name"].ToString();
                                    drInc["Target"] = dr["target_mins"].ToString();
                                    drInc["File_Mins"] = dr["File_Mins"].ToString();
                                    drInc["Converted_Mins"] = dr["Converted_Mins"].ToString();
                                    drInc["30to39"] = dr_Shift["30to39"].ToString();
                                    drInc["40to49"] = dr_Shift["40to49"].ToString();
                                    drInc["50to59"] = dr_Shift["50to59"].ToString();
                                    drInc["Above60"] = dr_Shift["Above60"].ToString();
                                    drInc["Amount"] = dr_Shift["incentiveAmount"].ToString();
                                    drInc["incentive_date"] = dr["incentive_date"].ToString();
                                    dt.Rows.Add(drInc);
                                }
                            }
                        }

                        foreach (DataRow _dr in dt.Rows)
                        {
                            if (dtInc_Date != Convert.ToDateTime(_dr["incentive_date"].ToString()))
                            {
                                BusinessLogic.oMessageEvent.Start("Adding Records");
                                dtInc_Date = Convert.ToDateTime(_dr["incentive_date"].ToString());
                                lsv_Incentive.Items.Add(new ListItemIncentive_Mins(string.Empty, string.Empty, "Incentive Date : " + dtInc_Date.ToString()));
                            }
                            lsv_Incentive.Items.Add(new ListItemIncentive_Mins(_dr, iRowcount++));

                        }

                        BusinessLogic.oMessageEvent.Start("Done...");
                        BusinessLogic.Reset_ListViewColumn(lsv_Incentive);
                    }
                    else
                        BusinessLogic.oMessageEvent.Start("No Record Found");
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Incentive_View()
        {
            try
            {
                //DataTable dtView = BusinessLogic.WS_Allocation.Get_Incentive_amount(Convert.ToInt32(cmb_incentive_branch.SelectedValue), dtp_incentive_fromdate.Value, dtp_incentive_todate.Value);
                //if (dtView.Rows.Count > 0)
                //{
                //    int iRowcount = 1;
                //    lsv_Incentive.Items.Clear();

                //    foreach (DataRow dr in dtView.Rows)
                //    {
                //        lsv_Incentive.Items.Add(new ListItemIncentive_Mins(dr, iRowcount++));
                //    }
                //    BusinessLogic.Reset_ListViewColumn(lsv_Incentive);
                //}

                // ------------------------- Altered on Mar 2016 ---------------------------------------
                //DataSet dsUser = BusinessLogic.WS_Allocation.Get_Usernamelist();
                //DataSet dsIncentive = BusinessLogic.WS_Allocation.Get_Incentive_Details();
                //DataSet dsIncentive_Shift = BusinessLogic.WS_Allocation.Get_Incentive_Shift();
                //DateTime dIn_Date = DateTime.Now;

                //List<BusinessLogic.INCENTIVE_DATE> oIncentive = new List<BusinessLogic.INCENTIVE_DATE>();
                //List<BusinessLogic.INCENTIVE_TARGET> oTarget = new List<BusinessLogic.INCENTIVE_TARGET>();


                //int iMonth =  dtp_incentive_fromdate.Value.Month;
                //foreach(DataRow dr in dsIncentive.Tables[0].Rows)
                //{
                //    oIncentive.Add(new BusinessLogic.INCENTIVE_DATE(Convert.ToDateTime(dr["incentive_date"].ToString()),  dr["location_id"].ToString(), dr["shift_type"].ToString()));
                //}

                //var vIncentive_date = (from c in oIncentive where ((Convert.ToDateTime(c.DINCENTIVE_DATE) >= Convert.ToDateTime(dtp_incentive_fromdate.Value) && Convert.ToDateTime(c.DINCENTIVE_DATE) <= Convert.ToDateTime(dtp_incentive_todate.Value))) select c).Distinct();

                ////foreach(var in_date in vIncentive_date)
                ////    lsv_Incentive.Columns.Add(Convert.ToDateTime(in_date.DINCENTIVE_DATE).ToString());

                //if (dsUser != null)
                //{
                //    if (dsUser.Tables[0].Rows.Count > 0)
                //    {
                //        int iRowcount = 1;
                //        lsv_Incentive.Items.Clear();

                //        foreach (DataRow dr in dsUser.Tables[0].Rows)
                //        {
                //            foreach (var in_date in vIncentive_date)
                //            {
                //                DataSet dsTarget = BusinessLogic.WS_Allocation.Get_Incentive_Calculation(Convert.ToDateTime(in_date.DINCENTIVE_DATE), Convert.ToInt32(dr["production_id"].ToString()), in_date.LOCATION_ID.ToString());

                //                if (dsTarget.Tables[0].Rows.Count > 0)
                //                {
                //                    foreach (DataRow _dr in dsTarget.Tables[0].Rows)
                //                    {
                //                        if (Convert.ToInt32(_dr["Mins_Diff"].ToString()) > 0)
                //                        {
                //                            foreach (DataRow drInc_Target in dsIncentive_Shift.Tables[0].Rows)
                //                            {
                //                                if ((Convert.ToDecimal((_dr["Mins_Different"].ToString().Replace(":", "."))) >= Convert.ToDecimal(drInc_Target["extraMinutesFrom"].ToString())) &&
                //                                (Convert.ToDecimal((_dr["Mins_Different"].ToString().Replace(":", "."))) < Convert.ToDecimal(drInc_Target["extraMinutesTo"].ToString())))
                //                                {
                //                                    if (Convert.ToDateTime(in_date.DINCENTIVE_DATE) != dIn_Date)
                //                                    {
                //                                        dIn_Date = Convert.ToDateTime(in_date.DINCENTIVE_DATE);
                //                                        lsv_Incentive.Items.Add(new ListItemIncentive_Mins(string.Empty, string.Empty, "Incentive Date : " + dIn_Date.ToString()));
                //                                    }   

                //                                    lsv_Incentive.Items.Add(new ListItemIncentive_Mins(iRowcount.ToString(), dr["ptag_id"].ToString(), dr["emp_full_name"].ToString(),
                //                                        _dr["target_mins"].ToString(), _dr["File_Mins"].ToString(), drInc_Target["1to30"].ToString(),
                //                                        drInc_Target["31to40"].ToString(), drInc_Target["41to50"].ToString(), drInc_Target["Above50"].ToString(),
                //                                        drInc_Target["incentiveAmount"].ToString()));

                //                                    iRowcount++;
                //                                }
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //        BusinessLogic.Reset_ListViewColumn(lsv_Incentive);
                //    }
                //}

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// GET INCENTIVE MINS
        /// </summary>
        private void Load_Incentive_Mins()
        {
            try
            {
                int isNight = 0;
                if (chb_Night.Checked == true)
                    isNight = 1;

                string sFilter = " Incentive_Amount > 0 ";
                DataTable dtIncentive_Mins = BusinessLogic.WS_Allocation.Get_Incentive_Mins(Convert.ToInt32(cmb_incentive_branch.SelectedValue), dtp_incentive_fromdate.Value, dtp_incentive_todate.Value, isNight);
                if (dtIncentive_Mins.Rows.Count > 0)
                {
                    int iRowcount = 1;
                    lsv_Incentive.Items.Clear();

                    foreach (DataRow dr in dtIncentive_Mins.Select(sFilter))
                    {
                        lsv_Incentive.Items.Add(new ListItemIncentive_Mins(dr, iRowcount++));
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Incentive);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// GET LEAVE DETAILS 
        /// </summary>
        /// <param name="dFromdate"></param>
        /// <param name="dTodate"></param>
        /// <param name="iBatch_id"></param>
        /// <param name="iBranch_Id"></param>
        private void Load_Capacity_Leave(DateTime dFromdate, DateTime dTodate, int iBatch_id, int iBranch_Id)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Initializing .......");
                int iRowcount = 1;
                DataSet dsLeave = BusinessLogic.WS_Allocation.Get_Capacity_Leavedetails(dFromdate, dTodate, iBatch_id, iBranch_Id);
                if (dsLeave != null)
                {
                    if (dsLeave.Tables[0].Rows.Count > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Loading Datas");
                        lsv_Capacity_Leave.Items.Clear();
                        foreach (DataRow dr in dsLeave.Tables[0].Rows)
                        {
                            lsv_Capacity_Leave.Items.Add(new ListItemCapacity_Leave(dr, iRowcount++));
                        }
                        BusinessLogic.Reset_ListViewColumn(lsv_Capacity_Leave);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD CAPACITY PERCENTAGE
        /// </summary>
        /// <param name="dFromdate"></param>
        /// <param name="dTodate"></param>
        private void Load_Capacity_Percentage(DateTime dFromdate, DateTime dTodate)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Initializing .......");
                DataSet dsCapacity_Per = BusinessLogic.WS_Allocation.Get_Capacity_Percentage(dFromdate, dTodate);
                if (dsCapacity_Per != null)
                {
                    BusinessLogic.oMessageEvent.Start("Loading Datas");
                    if (dsCapacity_Per.Tables[0].Rows.Count > 0)
                    {
                        lsv_CapacityPercentage.Items.Clear();
                        foreach (DataRow dr in dsCapacity_Per.Tables[0].Rows)
                        {
                            ListViewItem oList = null;

                            // ------------------------------- ON LINE
                            oList = new ListViewItem();
                            oList.Name = "OFFLINE";
                            oList.Text = "OFFLINE";
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00008B");
                            oList.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            oList.ForeColor = System.Drawing.ColorTranslator.FromHtml("#F8F8FF");
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT Mins Completed";
                            oList.Text = "Mins Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_off_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT TARGET";
                            oList.Text = "MT TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_off_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT PRECENTAGE";
                            oList.Text = "MT PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_Off_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME Mins Completed";
                            oList.Text = "ME Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_off_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME TARGET";
                            oList.Text = "ME TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_off_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME PRECENTAGE";
                            oList.Text = "ME PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_Off_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED Mins Completed";
                            oList.Text = "TED Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_off_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED TARGET";
                            oList.Text = "TED TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_off_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED PRECENTAGE";
                            oList.Text = "TED PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_Off_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FFFF");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = string.Empty;
                            oList.Text = string.Empty;
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            // ------------------------------- ON LINE

                            oList = new ListViewItem();
                            oList.Name = "ON LINE";
                            oList.Text = "ON LINE";
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFE0");
                            oList.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT Mins Completed";
                            oList.Text = "Mins Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_ON_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT TARGET";
                            oList.Text = "MT TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_ON_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "MT PRECENTAGE";
                            oList.Text = "MT PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["MT_ON_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME Mins Completed";
                            oList.Text = "ME Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_ON_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME TARGET";
                            oList.Text = "ME TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_ON_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "ME PRECENTAGE";
                            oList.Text = "ME PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["ME_ON_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED Mins Completed";
                            oList.Text = "TED Completed";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_ON_MinsDone"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED TARGET";
                            oList.Text = "TED TARGET";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_ON_Target"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);

                            oList = new ListViewItem();
                            oList.Name = "TED PRECENTAGE";
                            oList.Text = "TED PRECENTAGE";
                            oList.SubItems.Add(dsCapacity_Per.Tables[0].Rows[0]["TED_ON_Cap_Per"].ToString());
                            oList.BackColor = System.Drawing.ColorTranslator.FromHtml("#00FF00");
                            oList.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            lsv_CapacityPercentage.Items.Add(oList);
                            BusinessLogic.Reset_ListViewColumn(lsv_CapacityPercentage);

                        }
                        BusinessLogic.oMessageEvent.Start("Done!...");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD CAPACITY USERS LIST
        /// </summary>
        /// <param name="iBranchId"></param>
        /// <param name="iBatchId"></param>
        private void Load_Overall_Capacity_UserList(int iBranchId, int iBatchId, DateTime dFromdate, DateTime dTodate, int iWorktype)
        {
            try
            {
                int iRowcount = 1;
                int Tot_Mins = 0;
                ListItem_MTMECapacity oItem;
                DataSet dsCapacity = BusinessLogic.WS_Allocation.Get_Overall_Capacity_UserList(iBranchId, iBatchId, dFromdate, dTodate, iWorktype);
                if (dsCapacity != null)
                {
                    if (dsCapacity.Tables[0].Rows.Count > 0)
                    {
                        lsv_Capacity.Items.Clear();
                        foreach (DataRow dr in dsCapacity.Tables[0].Rows)
                        {
                            lsv_Capacity.Items.Add(new ListItem_MTMECapacity(dr, iRowcount++));
                            Tot_Mins += Convert.ToInt32(dr["target_mins"].ToString());
                        }
                        oItem = new ListItem_MTMECapacity(string.Empty, string.Empty, string.Empty, Tot_Mins.ToString());
                        lsv_Capacity.Items.Add(oItem);
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Capacity);
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }

        /// <summary>
        /// LOAD USER NAMES LIST
        /// </summary>
        /// <param name="iBatch_id"></param>
        /// <param name="iWorktype"></param>
        private void Load_Capacity_Usernames(int iBatch_id, int iWorktype)
        {
            try
            {
                DataSet dsUsername = BusinessLogic.WS_Allocation.Get_Capacity_Username(iBatch_id, iWorktype);
                if (dsUsername != null)
                {
                    if (dsUsername.Tables[0].Rows.Count > 0)
                    {
                        cmb_AddCap_Name.DataSource = dsUsername.Tables[0];
                        cmb_AddCap_Name.DisplayMember = "Username";
                        cmb_AddCap_Name.ValueMember = "production_id";
                        cmb_AddCap_Name.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD CAPACITY USERS LIST
        /// </summary>
        /// <param name="iBranchId"></param>
        /// <param name="iBatchId"></param>
        private void Load_Capacity_UserList(int iBranchId, int iBatchId, DateTime dFromdate, DateTime dTodate, int iWorktype)
        {
            try
            {
                int iRowcount = 1;
                int Tot_Mins = 0;
                int Tot_mins_done = 0;
                int tot_conv = 0;

                int Tot_Mins_MKMG = 0;
                int Tot_mins_done_MKMG = 0;
                int tot_conv_MKMG = 0;

                int Tot_Mins_Night = 0;
                int Tot_mins_done_Night = 0;
                int tot_conv_Night = 0;

                string sPresent_Client = string.Empty;
                string sCurrent_Client = string.Empty;
                string sMins_Done = string.Empty;
                string sConv_Mins = string.Empty;

                string sMins_Done_MKMG = string.Empty;
                string sConv_Mins_MKMG = string.Empty;

                string sMins_Done_Night = string.Empty;
                string sConv_Mins_Night = string.Empty;

                ListItem_MTME_CurrentdateCapacity oItem;
                DataSet dsCapacity = BusinessLogic.WS_Allocation.Get_UserList(iBranchId, iBatchId, dFromdate, dTodate, iWorktype);
                DataSet dsNightShift = BusinessLogic.WS_Allocation.Get_UserList_Nightshift(iBranchId, iBatchId, dFromdate, dTodate, iWorktype);

                if (dsCapacity != null)
                {
                    if (dsCapacity.Tables[0].Rows.Count > 0)
                    {
                        lsv_Currentdate_Capacity.Items.Clear();
                        lsv_Iasis_capcity.Items.Clear();
                        lsv_Mkmg_capcity.Items.Clear();
                        lsv_Nightshift_Capacity.Items.Clear();
                        foreach (DataRow dr in dsCapacity.Tables[0].Rows)
                        {
                            //Radhika
                            if (Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue) == 1)
                            {
                                pnl_Current_cap_offline.Visible = true;
                                pnl_Current_Online_All.Visible = false;
                                lsv_Currentdate_Capacity.Items.Add(new ListItem_MTME_CurrentdateCapacity(dr, iRowcount++));
                                Tot_Mins += Convert.ToInt32(dr["target_mins"].ToString());
                                Tot_mins_done += Convert.ToInt32(dr["file_minutes_Total"].ToString());
                                tot_conv += Convert.ToInt32(dr["Converted_minutes_Total"].ToString());

                                int iInsert = BusinessLogic.WS_Allocation.Set_Capacity(Convert.ToInt32(dr["production_id"].ToString()), Convert.ToDateTime(dFromdate));
                            }
                            else
                            {
                                //pnl_Current_cap_offline.Visible = false;
                                //pnl_Current_Online_All.Visible = true;

                                if (dr["client_name"].ToString().Contains("IASIS"))
                                {
                                    //lsv_Currentdate_Capacity.Items.Add(new ListItem_MTME_CurrentdateCapacity(dr, iRowcount++));
                                    lsv_Iasis_capcity.Items.Add(new ListItem_MTME_CurrentdateCapacity(dr, iRowcount++));

                                    Tot_Mins += Convert.ToInt32(dr["target_mins"].ToString());
                                    Tot_mins_done += Convert.ToInt32(dr["file_minutes_Total"].ToString());
                                    tot_conv += Convert.ToInt32(dr["Converted_minutes_Total"].ToString());

                                    int iInsert = BusinessLogic.WS_Allocation.Set_Capacity(Convert.ToInt32(dr["production_id"].ToString()), Convert.ToDateTime(dFromdate));
                                }
                                else if (dr["client_name"].ToString().Contains("MKMG"))
                                {
                                    //lsv_Currentdate_Capacity.Items.Add(new ListItem_MTME_CurrentdateCapacity(dr, iRowcount++));
                                    lsv_Mkmg_capcity.Items.Add(new ListItem_MTME_CurrentdateCapacity(dr, iRowcount++));

                                    Tot_Mins_MKMG += Convert.ToInt32(dr["target_mins"].ToString());
                                    Tot_mins_done_MKMG += Convert.ToInt32(dr["file_minutes_Total"].ToString());
                                    tot_conv_MKMG += Convert.ToInt32(dr["Converted_minutes_Total"].ToString());

                                    int iInsert = BusinessLogic.WS_Allocation.Set_Capacity(Convert.ToInt32(dr["production_id"].ToString()), Convert.ToDateTime(dFromdate));
                                }
                            }
                        }

                        //Rad

                        if (dsNightShift != null)
                        {
                            if (dsNightShift.Tables[0].Rows.Count > 0)
                            {
                                foreach (DataRow drNight in dsNightShift.Tables[0].Rows)
                                {
                                    lsv_Nightshift_Capacity.Items.Add(new ListItem_MTME_CurrentdateCapacity(drNight, iRowcount++));
                                    Tot_Mins_Night += Convert.ToInt32(drNight["target_mins"].ToString());
                                    Tot_mins_done_Night += Convert.ToInt32(drNight["file_minutes_Total"].ToString());
                                    tot_conv_Night += Convert.ToInt32(drNight["Converted_minutes_Total"].ToString());

                                    int iInsert = BusinessLogic.WS_Allocation.Set_Capacity(Convert.ToInt32(drNight["production_id"].ToString()), Convert.ToDateTime(dFromdate));
                                }
                            }
                        }

                        if (Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue) == 1)
                        {
                            sMins_Done = sGetDuration(Tot_mins_done);
                            sConv_Mins = sGetDuration(tot_conv);

                            oItem = new ListItem_MTME_CurrentdateCapacity(string.Empty, string.Empty, string.Empty, Tot_Mins.ToString(), sMins_Done.ToString(), sConv_Mins.ToString());
                            lsv_Currentdate_Capacity.Items.Add(oItem);
                        }
                        else
                        {
                            sMins_Done = sGetDuration(Tot_mins_done);
                            sConv_Mins = sGetDuration(tot_conv);

                            if (sMins_Done != "")
                            {
                                oItem = new ListItem_MTME_CurrentdateCapacity(string.Empty, string.Empty, string.Empty, Tot_Mins.ToString(), sMins_Done.ToString(), sConv_Mins.ToString());
                                lsv_Iasis_capcity.Items.Add(oItem);
                            }

                            if (Tot_mins_done_MKMG != 0)
                            {
                                sMins_Done_MKMG = sGetDuration(Tot_mins_done_MKMG);
                                sConv_Mins_MKMG = sGetDuration(tot_conv_MKMG);

                                oItem = new ListItem_MTME_CurrentdateCapacity(string.Empty, string.Empty, string.Empty, Tot_Mins_MKMG.ToString(), sMins_Done_MKMG.ToString(), sConv_Mins_MKMG.ToString());
                                lsv_Mkmg_capcity.Items.Add(oItem);
                            }

                            if (Tot_mins_done_Night != 0)
                            {
                                sMins_Done_Night = sGetDuration(Tot_mins_done_Night);
                                sConv_Mins_Night = sGetDuration(tot_conv_Night);

                                oItem = new ListItem_MTME_CurrentdateCapacity(string.Empty, string.Empty, string.Empty, Tot_Mins_Night.ToString(), sMins_Done_Night.ToString(), sConv_Mins_Night.ToString());
                                lsv_Nightshift_Capacity.Items.Add(oItem);
                            }
                        }
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Currentdate_Capacity);
                    BusinessLogic.Reset_ListViewColumn(lsv_Mkmg_capcity);
                    BusinessLogic.Reset_ListViewColumn(lsv_Iasis_capcity);
                    BusinessLogic.Reset_ListViewColumn(lsv_Nightshift_Capacity);
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }

        /// <summary>
        /// GET CLIENT TYPE DETAILS
        /// </summary>
        private void Load_Clienttype()
        {
            try
            {
                DataSet dsClient = BusinessLogic.WS_Allocation.Get_ClientType();
                if (dsClient != null)
                {
                    if (dsClient.Tables[0].Rows.Count > 0)
                    {
                        cbxHourPlatform.DataSource = dsClient.Tables[0].DefaultView;
                        cbxHourPlatform.DisplayMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR;
                        cbxHourPlatform.ValueMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT;
                        //cbxHourPlatform.SelectedIndex = 0;

                        cmb_Capacity_Worktype.DataSource = dsClient.Tables[0].DefaultView;
                        cmb_Capacity_Worktype.DisplayMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR;
                        cmb_Capacity_Worktype.ValueMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT;
                        //cmb_Capacity_Worktype.SelectedIndex = 0;

                        cmb_AddCap_Worktype.DataSource = dsClient.Tables[0].DefaultView;
                        cmb_AddCap_Worktype.DisplayMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR;
                        cmb_AddCap_Worktype.ValueMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT;
                        //cmb_AddCap_Worktype.SelectedIndex = 0;

                        cmb_MTTrack_Workplatform.DataSource = dsClient.Tables[0].DefaultView;
                        cmb_MTTrack_Workplatform.DisplayMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR;
                        cmb_MTTrack_Workplatform.ValueMember = Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT;
                        //cmb_MTTrack_Workplatform.SelectedIndex = 0;

                    }
                    else
                    {
                        cbxHourPlatform.DataSource = null;
                        cbxHourPlatform.DataSource = null;
                    }

                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Customized_Employee_List()
        {
            try
            {
                lsv_Customized_Employee.Items.Clear();
                int iRowcount = 1;
                string sGroupID = string.Empty;
                if (cmb_Customize_Group.SelectedIndex == 0)
                    sGroupID = "-1";
                else
                    sGroupID = cmb_Customize_Group.SelectedValue.ToString();

                DataSet dsEmp_List = BusinessLogic.WS_Allocation.Get_Customized_EmployeeList_New(Convert.ToInt32(sGroupID));
                if (dsEmp_List != null)
                {
                    if (dsEmp_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_Customized_Employee.Items.Clear();
                        foreach (DataRow dr in dsEmp_List.Tables[0].Rows)
                        {
                            lsv_Customized_Employee.Items.Add(new ListItem_Customized_Employee(dr, iRowcount));
                            iRowcount++;
                        }
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Customized_Employee);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// Get Customized Removal Entry
        /// </summary>
        private void Load_Customized_Employee_Removal_List()
        {
            try
            {
                int iRowcount = 1;
                DataSet dsEmp_List = BusinessLogic.WS_Allocation.Get_Customized_EmployeeList_Removal();
                if (dsEmp_List != null)
                {
                    if (dsEmp_List.Tables[0].Rows.Count > 0)
                    {
                        lsv_Customized_Remove_Employee.Items.Clear();
                        foreach (DataRow dr in dsEmp_List.Tables[0].Rows)
                        {
                            lsv_Customized_Remove_Employee.Items.Add(new ListItem_Customized_Employee(dr, iRowcount));
                            iRowcount++;
                        }
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Customized_Remove_Employee);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// TRANS ONLY OFFLINE
        /// </summary>
        private void Get_TransOnly_Offline()
        {
            try
            {
                int dTotMins = 0;
                BusinessLogic.oMessageEvent.Start("Requesting Database...");
                DataTable dtTransonly_Offline = BusinessLogic.WS_Allocation.Get_Transcribed_only_Offline(Convert.ToDateTime(dtp_transonly_offline_fromdate.Value), Convert.ToDateTime(dtp_transonly_offline_Todate.Value));
                MyTransonly_Offline oListItem;
                if (dtTransonly_Offline != null)
                {
                    if (dtTransonly_Offline.Rows.Count > 0)
                    {
                        int i = 1;
                        lsv_transonly_offline.Items.Clear();
                        foreach (DataRow dr in dtTransonly_Offline.Rows)
                        {
                            lsv_transonly_offline.Items.Add(new MyTransonly_Offline(dr, i));
                            i++;

                            if (Convert.ToInt32(dr["file_minutes"]) != 0)
                            {
                                if (dr["file_minutes"].ToString().Contains('.'))
                                    dTotMins += Convert.ToInt32(dr["file_minutes"].ToString().Split('.').GetValue(0));
                                else
                                    dTotMins += Convert.ToInt32(dr["file_minutes"].ToString());
                            }
                        }
                        string oMins = sGetDuration(dTotMins);

                        oListItem = new MyTransonly_Offline(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "Total Transcribed Minutes : ", oMins.ToString(), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        lsv_transonly_offline.Items.Add(oListItem);

                        BusinessLogic.Reset_ListViewColumn(lsv_transonly_offline);
                    }
                    BusinessLogic.oMessageEvent.Start("Loaded Dictations...");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// GET TRANSCRIBED ONLY ONLINE DETAILS
        /// </summary>
        private void Get_TransOnly_Online()
        {
            try
            {
                int dTotMins = 0;
                BusinessLogic.oMessageEvent.Start("Requesting Database...");
                DataTable dtTransonly_Online = BusinessLogic.WS_Allocation.Get_Transcribed_only_Online(Convert.ToDateTime(dtp_online_transonly_Fromdate.Value), Convert.ToDateTime(dtp_online_transonly_Todate.Value));
                MyTransonly_online oListItem;
                if (dtTransonly_Online != null)
                {
                    if (dtTransonly_Online.Rows.Count > 0)
                    {
                        int i = 1;
                        lsv_Online_Transonly.Items.Clear();
                        foreach (DataRow dr in dtTransonly_Online.Rows)
                        {
                            lsv_Online_Transonly.Items.Add(new MyTransonly_online(dr, i));
                            i++;

                            if (Convert.ToInt32(dr["file_minutes"]) != 0)
                            {
                                if (dr["file_minutes"].ToString().Contains('.'))
                                    dTotMins += Convert.ToInt32(dr["file_minutes"].ToString().Split('.').GetValue(0));
                                else
                                    dTotMins += Convert.ToInt32(dr["file_minutes"].ToString());
                            }
                        }
                        string oMins = sGetDuration(dTotMins);

                        oListItem = new MyTransonly_online(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "Total Transcribed Minutes : ", oMins.ToString(), string.Empty, string.Empty, string.Empty);
                        lsv_Online_Transonly.Items.Add(oListItem);

                        BusinessLogic.Reset_ListViewColumn(lsv_Online_Transonly);
                    }
                    BusinessLogic.oMessageEvent.Start("Loaded Dictations...");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD THE REALLOTED LIST FROM THE DATABASE
        /// </summary>
        private void Get_Realloted_List()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Requesting Database...");
                DataTable dtRe_Allocation = BusinessLogic.WS_Allocation.Get_Reallocatingfiles(txt_Reallot_Voice.Text.Trim(), Convert.ToDateTime(dtpFrom_Date.Value)).Tables[0];
                if (dtRe_Allocation != null)
                {
                    if (dtRe_Allocation.Rows.Count > 0)
                    {
                        int i = 1;
                        lsv_Reallocation.Items.Clear();
                        foreach (DataRow dr in dtRe_Allocation.Rows)
                        {
                            lsv_Reallocation.Items.Add(new MyReallocation(dr, i));
                            i++;
                        }
                        BusinessLogic.Reset_ListViewColumn(lsv_Reallocation);
                    }
                    BusinessLogic.oMessageEvent.Start("Loaded Dictations...");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        /// <summary>
        /// LOAD THE ALLOCATION FILES LIST
        /// </summary>
        /// 
        private void GetAllocationedetails()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Requesting Database...");
                if (txtEditingVoice.Text != "")
                    sEditVoice = txtEditingVoice.Text.ToString();
                DataTable dtAllocation = BusinessLogic.WS_Allocation.Get_AllocationList(1, Convert.ToDateTime(dtpAll_From.Text).ToString("yyyy/MM/dd"), Convert.ToDateTime(dtpAllTo.Text).ToString("yyyy/MM/dd"), sEditVoice).Tables[0];
                if (dtAllocation != null)
                {
                    if (dtAllocation.Rows.Count > 0)
                    {
                        int i = 1;
                        lsvFileDetails.Items.Clear();
                        foreach (DataRow dr in dtAllocation.Rows)
                        {
                            lsvFileDetails.Items.Add(new MyAllocatioFile(dr, i));
                            i++;
                        }
                        BusinessLogic.Reset_ListViewColumn(lsvFileDetails);
                    }
                    BusinessLogic.oMessageEvent.Start("Loaded Dictations...");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                txtEditingVoice.Text = "";
            }
        }

        /// <summary>
        /// LOAD TAB PAGES
        /// </summary>
        /// <param name="Index"></param>
        private void Load_TabPage(int Index)
        {
            try
            {
                Load_Clienttype();
                if (tabControlMain.SelectedTab.Name == "tabPageEmployee")
                {
                    Load_Branch();
                    GetEmployeeList();
                }

                if (tabAllocation.SelectedTab.Name == "tabFileDetails")
                {
                    GetAllocationedetails();
                }
                else if (tabAllocation.SelectedTab.Name == "tabDeAllocation")
                {
                    Load_Batch_Employee();
                    //Load_Login_Employee();
                }
                else if (tabAllocation.SelectedTab.Name == "tabAttendance")
                {
                    Load_Batch_Employee();
                    Load_Leave_List();
                }
                else if ((tabAllocation.SelectedTab.Name == "tabOnlineAllocation") || (tabAllocation.SelectedTab.Name == "tabPReports"))
                {
                    Load_Account();
                }
                else if (tabAllocation.SelectedTab.Name == "tabPReports")
                {
                    Load_Clienttype();
                    Load_Branch();
                    Load_Batchs();
                    Load_Customized_group();
                    Load_Batch_Employee();
                    Load_Location_Type(CLIENT_TYPE_ID);
                    cbxFromHours.SelectedIndex = 0;
                    cbxToHours.SelectedIndex = 0;
                    cbxFromHours.DropDownStyle = ComboBoxStyle.DropDownList;
                    cbxToHours.DropDownStyle = ComboBoxStyle.DropDownList;

                    if (tabReports.SelectedTab.Name == "tabUserFileMinutes")
                        LoadHourlyReport();
                }
                else if (tabAllocation.SelectedTab.Name == "tabDiscrepancy")
                    Load_Discrepancy();
                else if (tabAllocation.SelectedTab.Name == "tbCustomizeEntry")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    Load_Employee_Full_name("1", "1");
                    Load_Customized_Employee_List();
                    Load_Customized_group();
                    Load_Location_Type(CLIENT_TYPE_ID);
                    Load_Customized_Employee_Removal_List();
                }

                else if (tabAllocation.SelectedTab.Name == "tbCapacity")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    Load_Clienttype();
                }
                else if (tabAllocation.SelectedTab.Name == "tbNDSPDetails")
                {
                    Load_Designation();
                }
                else if (tabAllocation.SelectedTab.Name == "tbNightshiftusers")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    Load_Nightshift_User();
                }

                else if ((tabOffline.SelectedTab.Name == "tabTLFileTrack"))
                {
                    Load_Batch_Employee();
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD DOCTOR
        /// </summary>
        /// <param name="iAccountID"></param>
        /// <param name="sLocation"></param>
        private void Load_Doctor(int iAccountID, string sLocation)
        {
            try
            {
                DataSet _dsDoctor = new DataSet();
                _dsDoctor = BusinessLogic.WS_Allocation.Get_Doctor(iAccountID, sLocation);

                if (tabOffline.SelectedTab.Name == "tabOffPriorityMap")
                {
                    lbDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                    lbDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                    lbDoctor.DataSource = _dsDoctor.Tables[0];
                }
                else if ((TabExtraction.SelectedTab.Name == "tabPageManual") || (TabExtraction.SelectedTab.Name == "tabPageExtraction"))
                {
                    cboExtractDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                    cboExtractDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                    cboExtractDoctor.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboExtractDoctor.DataSource = _dsDoctor.Tables[0];

                    cboManualDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                    cboManualDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                    cboManualDoctor.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboManualDoctor.DataSource = _dsDoctor.Tables[0];
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD ACCOUNT
        /// </summary>
        private void Load_Account()
        {
            try
            {
                DataSet _dsAccount = new DataSet();
                _dsAccount = BusinessLogic.WS_Allocation.Get_ClientName();

                if (tabOffline.SelectedTab.Name == "tabOffPriorityMap")
                {
                    lbClient.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                    lbClient.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                    lbClient.DataSource = _dsAccount.Tables[0];
                }
                else if ((TabExtraction.SelectedTab.Name == "tabPageManual") || (TabExtraction.SelectedTab.Name == "tabPageExtraction"))
                {
                    cboManualAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                    cboManualAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                    cboManualAccount.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboManualAccount.DataSource = _dsAccount.Tables[1];

                    cboExtractAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                    cboExtractAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                    cboExtractAccount.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboExtractAccount.DataSource = _dsAccount.Tables[1];
                }
                else if (TabExtraction.SelectedTab.Name == "tabOnlineEntry")
                {
                    cmbEntryAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                    cmbEntryAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                    cmbEntryAccount.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmbEntryAccount.DataSource = _dsAccount.Tables[1];
                }
                if ((tabControlMain.SelectedTab.Name == "tabPageOffline") && (tabOffline.SelectedTab.Name == "tbAllocationPriority"))
                {
                    cmbClient.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                    cmbClient.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                    cmbClient.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmbClient.DataSource = _dsAccount.Tables[0];
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD LOCATION
        /// </summary>
        /// <param name="sAccountID"></param>
        private void Load_Location(int sAccountID)
        {
            try
            {
                DataSet _dsLocation = new DataSet();
                _dsLocation = BusinessLogic.WS_Allocation.Get_Location(Convert.ToInt32(sAccountID));

                if (tabOffline.SelectedTab.Name == "tabOffPriorityMap")
                {
                    lbLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                    lbLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                    lbLocation.DataSource = _dsLocation.Tables[0];
                }
                else if ((tabControlMain.SelectedTab.Name == "Offline") || (tbConOfReport.SelectedTab.Name == "tbpTAT"))
                {
                    cmbLocationTat.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                    cmbLocationTat.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                    cmbLocationTat.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmbLocationTat.DataSource = _dsLocation.Tables[0];
                }

                else if ((TabExtraction.SelectedTab.Name == "tabPageManual") || (TabExtraction.SelectedTab.Name == "tabPageExtraction"))
                {
                    cboManualLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                    cboManualLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                    cboManualLocation.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboManualLocation.DataSource = _dsLocation.Tables[0];

                    cboExtractLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                    cboExtractLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                    cboExtractLocation.DropDownStyle = ComboBoxStyle.DropDownList;
                    cboExtractLocation.DataSource = _dsLocation.Tables[0];
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD PRIORITY MAPPING
        /// </summary>
        private void Load_Priority_Mapping()
        {
            try
            {
                DataSet _dsPriority = new DataSet();
                _dsPriority = BusinessLogic.WS_Allocation.Get_UserAllocation_Priority();

                if (_dsPriority == null)
                {
                    BusinessLogic.oMessageEvent.Start("No Records to display");
                    return;
                }

                lsvProfiles.Items.Clear();
                int iRowCount = 1;

                foreach (DataRow _dr in _dsPriority.Tables[1].Select())
                    lsvProfiles.Items.Add(new ListItem_UserAllocationProfile(_dr, ++iRowCount));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsvProfiles);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD ALLOTED FILES 
        /// </summary>
        private void Load_AllotedFilesForEmployees()
        {
            try
            {
                DataTable _dsAllotedFiles = new DataTable();
                //ListItem_LoginEmployees oCurrentLogin = (ListItem_LoginEmployees)lsvLoginEmoloyees.SelectedItems[0];
                MyDeallocate_Offline_EmployeeList oCurrentLogin = (MyDeallocate_Offline_EmployeeList)lsvLoginEmoloyees.SelectedItems[0];
                _dsAllotedFiles = BusinessLogic.WS_Allocation.Get_AllocationDetails(oCurrentLogin.EMP_PRODUCTION_ID, 2);

                lsvAllotedFiles.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drAllFiles in _dsAllotedFiles.Select())
                    lsvAllotedFiles.Items.Add(new ListItem_AllotedFilesForUsers(_drAllFiles, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsvAllotedFiles);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD LOGGED IN EMPLOYEES
        /// </summary>
        private void Load_Login_Employee()
        {
            try
            {
                DataTable _dsLoginEmp = new DataTable();
                _dsLoginEmp = BusinessLogic.WS_Allocation.Get_EmployeeDetails(-1, -1, -1);

                lsvLoginEmoloyees.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drLogingEmp in _dsLoginEmp.Select())
                    lsvLoginEmoloyees.Items.Add(new ListItem_LoginEmployees(_drLogingEmp, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsvLoginEmoloyees);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD EMPOYEE LIST
        /// </summary>
        private void Load_Employee_List()
        {
            BusinessLogic.oMessageEvent.Start("Transferring Data..!");
            BusinessLogic.oProgressEvent.Start(true);
            try
            {
                DataSet _dsEmpList = new DataSet();

                if (tabAllocation.SelectedTab.Name == "tabAttendance")
                {
                    int Batch = ((ListItem_EmpBatchName)lvDesignation.SelectedItems[0]).iBatchID;
                    _dsEmpList = BusinessLogic.WS_Allocation.Get_Employee_List(Batch);

                    lsvEmployeeList.Items.Clear();
                    int iRowCount = 1;
                    foreach (DataRow _dr in _dsEmpList.Tables[0].Select())
                        lsvEmployeeList.Items.Add(new ListItem_EmployeeList(_dr, ++iRowCount));

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lsvEmployeeList);
                }
                else if (tabAllocation.SelectedTab.Name == "tabDeAllocation" && tabControlMain.SelectedTab.Text == "Online")
                {
                    int Batch = ((ListItem_EmpBatchName)lsvDeAllotBatch.SelectedItems[0]).iBatchID;
                    _dsEmpList = BusinessLogic.WS_Allocation.GET_ALLOTED_DETAILS_NEW(Batch);

                    lsvLoginEmoloyees.Items.Clear();
                    int iRowCount = 1;
                    foreach (DataRow _dr in _dsEmpList.Tables[0].Select())
                        lsvLoginEmoloyees.Items.Add(new MyDeallocate_Offline_EmployeeList(_dr, ++iRowCount));

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lsvLoginEmoloyees);
                }
                else if (tabOffline.SelectedTab.Name == "tablOffDeallot" && tabControlMain.SelectedTab.Text == "Offline")
                {
                    int Batch = ((ListItem_EmpBatchName)lsv_Offline_Deall_Designation.SelectedItems[0]).iBatchID;
                    _dsEmpList = BusinessLogic.WS_Allocation.GET_ALLOTED_DETAILS_NEW(Batch);

                    DataRow drRow = _dsEmpList.Tables[0].NewRow();
                    drRow["emp_id"] = 0;
                    drRow["emp_full_name"] = "ALL";
                    drRow["designation"] = "ALL";
                    drRow["target_mins"] = 0;
                    drRow["production_id"] = -1;
                    drRow["ptag_id"] = 0;
                    drRow["alloted"] = "0-0-0";
                    drRow["achived"] = "0-0-0";
                    _dsEmpList.Tables[0].Rows.InsertAt(drRow, 0);

                    lsv_Offline_Deall_EmplDetails.Items.Clear();
                    int iRowCount = 1;
                    foreach (DataRow _dr in _dsEmpList.Tables[0].Select())
                        lsv_Offline_Deall_EmplDetails.Items.Add(new MyDeallocate_Offline_EmployeeList(_dr, ++iRowCount));

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Deall_EmplDetails);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Ready");
            }
        }

        /// <summary>
        /// LOAD EMPOYEE LIST
        /// </summary>
        private void Load_Employee_List_NightShiftMark()
        {
            try
            {
                lvNEmployee.Items.Clear();
                DataSet _dsEmpList = new DataSet();
                int Batch = ((ListItem_EmpBatchName)lvNDesignation.SelectedItems[0]).iBatchID;
                _dsEmpList = BusinessLogic.WS_Allocation.Get_Employee_List(Batch);
                int iRowCount = 1;
                foreach (DataRow _dr in _dsEmpList.Tables[0].Select())
                    lvNEmployee.Items.Add(new ListItem_EmployeeList(_dr, ++iRowCount));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lvNEmployee);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD LEAVE LIST
        /// </summary>
        private void Load_Leave_List()
        {
            try
            {
                lsvLeaveList.Items.Clear();
                DateTime sDate = dtpAttendanceDate.Value;
                DataSet _dsLeaveList = new DataSet();
                _dsLeaveList = BusinessLogic.WS_Allocation.Set_Leave_List(1, sDate, "D", "C", 2);

                int iRowCount = 1;
                foreach (DataRow _dr in _dsLeaveList.Tables[0].Select())
                    lsvLeaveList.Items.Add(new ListItem_LeaveDetails(_dr, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsvLeaveList);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// Used to convert into minutes
        /// </summary>
        /// <param name="sFile"></param>
        /// <returns></returns>
        public static string Get_Minutes(string sFile)
        {
            try
            {
                TagLib.File file = TagLib.File.Create(sFile);
                int s_time = (int)file.Properties.Duration.TotalSeconds;

                return s_time.ToString();
            }
            catch
            {
                return "0";
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// TEMP DATATABLE CREATION
        /// </summary>
        private void Create_DataTable()
        {
            try
            {
                dtManual = new DataTable();
                dtManual.TableName = "Manual";
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_DOCTOR_ID_BINT, typeof(int));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR, typeof(string));
                dtManual.Columns.Add("Dictation_path", typeof(string));
                dtManual.Columns.Add("Client_Id", typeof(int));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_EXTENSION_STR, typeof(string));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE, typeof(decimal));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_FILE_SIZE_BINT, typeof(int));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_TAT_INT, typeof(decimal));
                dtManual.Columns.Add(Framework.MAINTRANSCRIPTION.FIELD_COMMENT_STR, typeof(string));
            }
            catch (Exception ex)
            {
                BusinessLogic.oMessageEvent.Start(ex.ToString());
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// EXTRACTION MODULE
        /// </summary>
        private void Extract()
        {
            try
            {
                //Validation
                if (Convert.ToInt32(cboManualAccount.SelectedValue) == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Select the client account");
                    cboManualAccount.Focus();
                    return;
                }

                if (Convert.ToInt32(cboManualDoctor.SelectedValue) == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Select the doctor");
                    cboManualDoctor.Focus();
                    return;
                }

                if (cboManualLocation.SelectedValue.ToString() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the location");
                    cboManualLocation.Focus();
                    return;
                }

                if (txt_Location.Text.Trim() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the folder path");
                    txt_Location.Focus();
                    return;
                }

                // Confirm the entries
                string sMessage = "Confirm the detials below " + Environment.NewLine +
                                  "CLIENT : " + cboManualAccount.Text.ToString() + Environment.NewLine +
                                  "DOCTOR : " + cboManualDoctor.Text.ToString() + Environment.NewLine +
                                  "FILE COUNT : " + lsvManual.Items.Count.ToString() + Environment.NewLine + "Are you sure to continue?";

                DialogResult oResult = MessageBox.Show(sMessage);

                if (oResult == DialogResult.OK)
                {
                    btnBrowse.Enabled = false;
                    btnExtract.Enabled = false;

                    BusinessLogic.oMessageEvent.Start("Initializing...");
                    BusinessLogic.oProgressEvent.Start(true);

                    int iDictationCount = 0;
                    int iInvalidExtension = 0;
                    int iTotalDictationCount = lsvExtractInfo.Items.Count;

                    string sDoctorId = cboManualDoctor.SelectedValue.ToString();
                    string sAccount = cboManualAccount.Text.Trim();
                    string sDoctorName = cboManualDoctor.Text.Trim();
                    string sLocationName = cboManualLocation.Text.Trim();

                    DataTable dtVoice = (DataTable)BusinessLogic.WS_Allocation.GetVoiceExtension().Tables[0];

                    //If the voice file extension is empty and set no voice file extions found
                    if ((dtVoice == null) || (dtVoice.Rows.Count == 0))
                    {
                        BusinessLogic.oMessageEvent.Start("No voice file extensions found.");
                        return;
                    }

                    //Get the server path
                    DataTable dtVoicePath = BusinessLogic.WS_Allocation.GetVoicePath(1).Tables[0];
                    string sVoicePath = ((dtVoicePath.Rows[0]["voice_path"].ToString()));

                    //Set voice path
                    BusinessLogic.oMessageEvent.Start("Check Network.");
                    string sVoiceServerPath = Path.Combine(sVoicePath, BusinessLogic.SERVER_DATE.ToString("yyyy") + "\\" + BusinessLogic.SERVER_DATE.ToString("MMMdd") + "\\" + sAccount.Trim() + "\\" + sLocationName.Trim() + "\\" + sDoctorName.Trim());

                    sVoiceServerPath = "\\\\" + sVoiceServerPath;

                    if (!Directory.Exists(sVoiceServerPath))
                        Directory.CreateDirectory(sVoiceServerPath);

                    //Save the record
                    foreach (ListViewItem oItem in lsvManual.Items)
                    {
                        iInvalidExtension = 0;
                        MylsvManualDetails oManual = (MylsvManualDetails)oItem;
                        BusinessLogic.oMessageEvent.Start("Initializing...");
                        BusinessLogic.oProgressEvent.Start(iDictationCount, 0, iTotalDictationCount);
                        Color oColor = oItem.BackColor;

                        oManual.EnsureVisible();
                        oManual.BackColor = Color.Orange;

                        string sDictationPath = oManual.DICTATION_PATH;

                        //Generate Voice file name
                        string sVoiceFileName = Path.GetFileName(oManual.DICTATION_PATH);

                        //Duration 
                        TimeSpan oTime = TimeSpan.FromSeconds(Convert.ToDouble(Get_Minutes(oManual.DICTATION_PATH)));
                        decimal seconds = Convert.ToDecimal(oTime.TotalSeconds);

                        //Invalid extension
                        foreach (DataRow drvoice in dtVoice.Rows)
                        {
                            string sVoiceextension = dtVoice.Rows[0]["voice_file_extension"].ToString();
                            string[] sVoice = sVoiceextension.Split(';');
                            foreach (string sV in sVoice)
                            {
                                if (sV.ToLower() == oManual.EXTENSION.Replace(".", "").ToLower())
                                {
                                    iInvalidExtension = 1;
                                    break;
                                }
                            }
                        }

                        if (iInvalidExtension == 1)
                        {
                            if (seconds > 0)
                            {
                                //Size
                                int iDictationSize = 0;
                                int iResult = 0;
                                FileInfo oFileInfo = new FileInfo(sDictationPath);
                                iDictationSize = ((int)oFileInfo.Length);

                                //Move the file to the path           sDownloadFileName.Replace(Path.GetExtension(sDownloadFileName), ".txt");                 
                                string VoiceFileFullPath = Path.Combine(sVoiceServerPath, sVoiceFileName);


                                DataTable dtDup = BusinessLogic.WS_Allocation.Get_VoiceCount(Path.GetFileNameWithoutExtension(sDictationPath));
                                int ijobcount = 0;
                                if (dtDup != null)
                                {
                                    if (dtDup.Rows.Count > 0)
                                    {
                                        ijobcount = Convert.ToInt32(dtDup.Rows[0]["jobcount"].ToString());
                                    }
                                }

                                bool check_voice = false;
                                foreach (DataRow drvoice in dtVoice.Rows)
                                {
                                    string sVoiceextension = dtVoice.Rows[0]["voice_file_extension"].ToString();
                                    string[] sVoice = sVoiceextension.Split(';');
                                    foreach (string sV in sVoice)
                                    {
                                        if (sV.ToLower() == oManual.EXTENSION.Replace(".", "").ToLower())
                                        {
                                            check_voice = true;
                                            break;
                                        }
                                    }
                                }

                                string CurDate = DateTime.Now.ToShortDateString();
                                DateTime cDate = Convert.ToDateTime(CurDate);

                                if (ijobcount == 0 && check_voice)
                                {
                                    // check the file move mode
                                    //Copying the dictation to the voice folder
                                    BusinessLogic.oMessageEvent.Start("Extracting... : " + sVoiceFileName);
                                    oManual._STATUS = "Extracting...";

                                    if (BusinessLogic.GET_FILE_TRANSFER_MODE(VoiceFileFullPath) == 1)
                                        UploadFile.Upload_DictationByStream(sDictationPath, VoiceFileFullPath, cboManualAccount.Text.Trim());
                                    else
                                    {
                                        UploadFile.Upload_DictationByStream(sDictationPath, VoiceFileFullPath, cboManualAccount.Text.Trim());
                                    }

                                    iResult = BusinessLogic.WS_Allocation.Set_DictationDetails(oManual.DOCTOR_ID, oManual.JOBID, oManual.EXTENSION, seconds, iDictationSize, iTatHours, 2);
                                    oManual._STATUS = "Completed";
                                    if (iResult == 1)
                                    {
                                        File.Delete(oManual.DICTATION_PATH);
                                    }
                                }
                                else if (ijobcount >= 0 && !check_voice)
                                    oManual._STATUS = "Invalid extension";
                                else
                                    // oItem.SubItems[6].Text = "Duplicate Entry";
                                    oManual._STATUS = "Duplicate Entry";
                            }
                            else
                            {
                                if (!File.Exists(oManual.DICTATION_PATH))
                                    oManual._STATUS = "File Not found";
                            }
                        }
                        else
                            oManual._STATUS = "Invalid Extension";
                    }
                    btnBrowse.Enabled = true;
                    BusinessLogic.oMessageEvent.Start("Done");
                    lsvManual.Items.Clear();
                    txt_Location.Enabled = true;
                    btnBrowse.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// Get the extraction details
        /// </summary>
        private void Get_Extractdetails()
        {
            try
            {
                DataTable dtFiles = BusinessLogic.WS_Allocation.Get_Extractlog_Online(Convert.ToDateTime(dtpFrom.Value.ToString("yyyy/MM/dd")), Convert.ToDateTime(dtpToDate.Value.ToString("yyyy/MM/dd")), Convert.ToInt32(cboExtractAccount.SelectedValue), Convert.ToInt32(cboExtractDoctor.SelectedValue), cboExtractLocation.SelectedValue.ToString());
                if (dtFiles != null)
                {
                    if (dtFiles.Rows.Count > 0)
                    {
                        int iRowCount = 0;
                        lsvExtractInfo.Items.Clear();
                        foreach (DataRow dr in dtFiles.Rows)
                        {
                            lsvExtractInfo.Items.Add(new Mylsvextractdetails(dr, iRowCount));
                        }

                        BusinessLogic.oMessageEvent.Start("Ready.");
                        BusinessLogic.Reset_ListViewColumn(lsvExtractInfo);
                    }
                    else
                    {
                        lsvExtractInfo.Items.Clear();
                    }
                }
                else
                {
                    lsvExtractInfo.Items.Clear();
                }
                foreach (ColumnHeader oColumn in lsvExtractInfo.Columns)
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
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD ALLOCATION STATUS
        /// </summary>
        private void Load_Allocation_Status()
        {
            BusinessLogic.oMessageEvent.Start("Transferring data..!");
            try
            {
                lvAllocationStatus.Items.Clear();
                DataSet _dsOnlineAllocation = new DataSet();
                object oFromDate = null;
                object oToDate = null;
                int iOptionID = 0;
                string sVoiceFile = string.Empty;
                if (txtVoice.Text == "")
                {
                    if (chkIncludeDate.Checked == true)
                    {
                        oFromDate = Convert.ToDateTime(dtpFromDate.Text).ToString("yyyy/MM/dd");
                        oToDate = Convert.ToDateTime(dtpTO.Text).ToString("yyyy/MM/dd");
                        iOptionID = 2;
                    }
                    else
                    {
                        oFromDate = null;
                        oToDate = null;
                        iOptionID = 1;
                    }
                }
                else
                {
                    iOptionID = 3;
                    oFromDate = null;
                    oToDate = null;
                    sVoiceFile = txtVoice.Text;
                }
                _dsOnlineAllocation = BusinessLogic.WS_Allocation.Get_Online_AllocationStatus_New(iOptionID, oFromDate, oToDate, sVoiceFile);
                txtVoice.Text = "";
                int iRowCount = 1;
                foreach (DataRow _drAll in _dsOnlineAllocation.Tables[0].Select())
                    lvAllocationStatus.Items.Add(new ListItem_OnlineAllocationStatus(_drAll, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lvAllocationStatus);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD BRANCH DETAILS
        /// </summary>
        private void Load_Branch()
        {
            try
            {
                DataSet dsBranch = BusinessLogic.WS_Allocation.Get_Branch_list();

                cmb_Emp_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                cmb_Emp_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                cmb_Emp_Branch.DataSource = dsBranch.Tables[0];
                cmb_Emp_Branch.DropDownStyle = ComboBoxStyle.DropDownList;

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tbOnline_Incentive"))
                {
                    cmb_incentive_branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_incentive_branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_incentive_branch.DataSource = dsBranch.Tables[0];
                    cmb_incentive_branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if (tabAllocation.SelectedTab.Name == "tbCapacity")
                {
                    cmb_Capacity_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Capacity_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Capacity_Branch.DataSource = dsBranch.Tables[0];
                    cmb_Capacity_Branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if (tabAllocation.SelectedTab.Name == "tbCustomizeEntry")
                {
                    cmb_Customized_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Customized_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Customized_Branch.DataSource = dsBranch.Tables[0];
                    cmb_Customized_Branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tabUserFileMinutes"))
                {
                    cbxUserBranch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cbxUserBranch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cbxUserBranch.DataSource = dsBranch.Tables[0];
                    cbxUserBranch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "HourlyReport"))
                {
                    DataRow dr = dsBranch.Tables[0].NewRow();
                    dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR] = " -- All -- ";
                    dr[Framework.BRANCH.FIELD_BATCH_BRANCHID_INT] = 0;
                    dsBranch.Tables[0].Rows.InsertAt(dr, 0);

                    cbxHourBranch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cbxHourBranch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cbxHourBranch.DataSource = dsBranch.Tables[0];
                    cbxHourBranch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tabLogSheet"))
                {
                    cbmLogBranch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cbmLogBranch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cbmLogBranch.DataSource = dsBranch.Tables[0];
                    cbmLogBranch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_Logsheethourly"))
                {
                    cmb_Log_branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Log_branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Log_branch.DataSource = dsBranch.Tables[0];
                    cmb_Log_branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tbTargetDetails"))
                {
                    cmb_Target_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Target_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Target_Branch.DataSource = dsBranch.Tables[0];
                    cmb_Target_Branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tbBlankReport"))
                {
                    cmb_Target_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Target_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Target_Branch.DataSource = dsBranch.Tables[0];
                    cmb_Target_Branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tbNightshiftusers"))
                {
                    DataRow dr = dsBranch.Tables[0].NewRow();
                    dr[Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR] = " -- All -- ";
                    dr[Framework.BRANCH.FIELD_BATCH_BRANCHID_INT] = 0;
                    dsBranch.Tables[0].Rows.InsertAt(dr, 0);

                    cmb_Nightshift_Branch.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmb_Nightshift_Branch.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmb_Nightshift_Branch.DataSource = dsBranch.Tables[0];
                    cmb_Nightshift_Branch.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_user_consolidated"))
                {
                    comboBox15.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    comboBox15.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    comboBox15.DataSource = dsBranch.Tables[0];
                    comboBox15.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabControlMain.SelectedTab.Name == "tbpMapping"))
                {
                    cmbBranchMapp.DisplayMember = Framework.BRANCH.FIELD_BATCH_BRANCHNAME_STR;
                    cmbBranchMapp.ValueMember = Framework.BRANCH.FIELD_BATCH_BRANCHID_INT;
                    cmbBranchMapp.DataSource = dsBranch.Tables[0];
                    cmbBranchMapp.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD BATCH
        /// </summary>
        private void Load_Batch_Employee()
        {
            try
            {
                DataSet dsBatch = new DataSet();
                dsBatch = BusinessLogic.WS_Allocation.Get_BatchDetails();

                cmb_Auto_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                cmb_Auto_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                cmb_Auto_Batch.DataSource = dsBatch.Tables[0];
                cmb_Auto_Batch.DropDownStyle = ComboBoxStyle.DropDownList;

                if (tabAllocation.SelectedTab.Name == "tbCapacity")
                {
                    cmb_Capacity_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Capacity_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Capacity_Batch.DataSource = dsBatch.Tables[0];
                    cmb_Capacity_Batch.DropDownStyle = ComboBoxStyle.DropDownList;

                    cmb_AddCap_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_AddCap_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_AddCap_Batch.DataSource = dsBatch.Tables[0];
                    cmb_AddCap_Batch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tbNightshiftusers"))
                {
                    cmb_Nightshift_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Nightshift_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Nightshift_Batch.DataSource = dsBatch.Tables[0];
                    cmb_Nightshift_Batch.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if (tabAllocation.SelectedTab.Name == "tbCustomizeEntry")
                {
                    cmb_Customized_Desig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Customized_Desig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Customized_Desig.DataSource = dsBatch.Tables[0];
                    cmb_Customized_Desig.DropDownStyle = ComboBoxStyle.DropDownList;
                }

                if ((tabAllocation.SelectedTab.Name == "tabOnlineAllocation") && (TabExtraction.SelectedTab.Name == "TabEmployee"))
                {
                    cboDesignation.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cboDesignation.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cboDesignation.DataSource = dsBatch.Tables[0];
                    cboDesignation.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if (tabAllocation.SelectedTab.Name == "tabDeAllocation")
                {
                    int iRowCount = 0;
                    lsvDeAllotBatch.Items.Clear();
                    foreach (DataRow _dr in dsBatch.Tables[0].Select())
                    {
                        lsvDeAllotBatch.Items.Add(new ListItem_EmpBatchName(_dr, iRowCount++));
                        BusinessLogic.oMessageEvent.Start("Ready.");
                        BusinessLogic.Reset_ListViewColumn(lsvDeAllotBatch);
                    }
                    lsvDeAllotBatch.Items[0].Selected = true;
                }
                if (tabAllocation.SelectedTab.Name == "tabAttendance")
                {
                    int iRowCount = 0;
                    lvDesignation.Items.Clear();
                    foreach (DataRow _dr in dsBatch.Tables[0].Select())
                    {
                        lvDesignation.Items.Add(new ListItem_EmpBatchName(_dr, iRowCount++));
                        BusinessLogic.oMessageEvent.Start("Ready.");
                        BusinessLogic.Reset_ListViewColumn(lvDesignation);
                    }
                    lvDesignation.Items[0].Selected = true;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tabUserFileMinutes"))
                {
                    cbxUserDesig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbxUserDesig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbxUserDesig.DataSource = dsBatch.Tables[0];
                    cbxUserDesig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "HourlyReport"))
                {
                    //DataRow dr = dsBatch.Tables[0].NewRow();
                    //dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR] = " -- All -- ";
                    //dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT] = 0;
                    //dsBatch.Tables[0].Rows.InsertAt(dr, 0);

                    cbxHourDesig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbxHourDesig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbxHourDesig.DataSource = dsBatch.Tables[0];
                    cbxHourDesig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tabLogSheet"))
                {
                    cbmLogBatch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbmLogBatch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbmLogBatch.DataSource = dsBatch.Tables[0];
                    cbmLogBatch.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_Logsheethourly"))
                {
                    cmb_Log_desig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Log_desig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Log_desig.DataSource = dsBatch.Tables[0];
                    cmb_Log_desig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tbTargetDetails"))
                {
                    dtp_Target_Fromdate.Value = DateTime.Now;
                    dtp_Target_Todate.Value = DateTime.Now;
                    cmb_Target_Desig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Target_Desig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Target_Desig.DataSource = dsBatch.Tables[0];
                    cmb_Target_Desig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_user_consolidated"))
                {
                    comboBox17.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    comboBox17.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    comboBox17.DataSource = dsBatch.Tables[0];
                    comboBox17.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if (TabExtraction.SelectedTab.Name == "tabOnlineEntry")
                {
                    cbxOnlineEntry.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbxOnlineEntry.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbxOnlineEntry.DataSource = dsBatch.Tables[0];
                    cbxOnlineEntry.DropDownStyle = ComboBoxStyle.DropDownList;

                    cbxInDesignation.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbxInDesignation.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbxInDesignation.DataSource = dsBatch.Tables[0];
                    cbxInDesignation.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabOnlineAllocation") && (TabExtraction.SelectedTab.Name == "tabTarget"))
                {
                    cmbTargetDisig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmbTargetDisig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmbTargetDisig.DataSource = dsBatch.Tables[0];
                    cmbTargetDisig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabAllocation.SelectedTab.Name == "tabOnlineAllocation") && (TabExtraction.SelectedTab.Name == "tabNightShift"))
                {
                    int iRowCount = 0;
                    lvNDesignation.Items.Clear();
                    foreach (DataRow _dr in dsBatch.Tables[0].Select())
                    {
                        lvNDesignation.Items.Add(new ListItem_EmpBatchName(_dr, iRowCount++));
                        BusinessLogic.oMessageEvent.Start("Ready.");
                        BusinessLogic.Reset_ListViewColumn(lvNDesignation);
                    }
                    lvNDesignation.Items[0].Selected = true;
                }
                if (tabOffline.SelectedTab.Name == "tablOffDeallot")
                {
                    int iRowCount = 0;
                    lsv_Offline_Deall_Designation.Items.Clear();
                    foreach (DataRow _dr in dsBatch.Tables[0].Select())
                    {
                        lsv_Offline_Deall_Designation.Items.Add(new ListItem_EmpBatchName(_dr, iRowCount++));
                        BusinessLogic.oMessageEvent.Start("Ready.");
                        BusinessLogic.Reset_ListViewColumn(lsv_Offline_Deall_Designation);
                    }
                    lsv_Offline_Deall_Designation.Items[0].Selected = true;
                }
                if ((tabControlMain.SelectedTab.Name == "tabPageOffline") && (tabOffline.SelectedTab.Name == "tabOfflineReport"))
                {
                    dtp_Offline_Fromdate.Value = DateTime.Now;
                    dtp_Offline_Todate.Value = DateTime.Now;
                    cmb_Offline_Desig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_Offline_Desig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_Offline_Desig.DataSource = dsBatch.Tables[0];
                    cmb_Offline_Desig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if ((tabControlMain.SelectedTab.Name == "tabPageOffline") && (tabOffline.SelectedTab.Name == "tbAllocationPriority"))
                {
                    cmbDesignation.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmbDesignation.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmbDesignation.DataSource = dsBatch.Tables[0];
                    cmbDesignation.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if ((tabOffline.SelectedTab.Name == "tabOfflineReport") && (tbConOfReport.SelectedTab.Name == "tbFilesDownloaded"))
                {
                    cbxReportDesig.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbxReportDesig.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbxReportDesig.DataSource = dsBatch.Tables[0];
                    cbxReportDesig.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if ((tabOffline.SelectedTab.Name == "tabTLFileTrack"))
                {
                    cmb_TL_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cmb_TL_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cmb_TL_Batch.DataSource = dsBatch.Tables[0];
                    cmb_TL_Batch.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// TAT PERCENTAGE FILES
        /// </summary>
        private void LoadTatPercentageFiles(string sYear, int sMonth, string sFromDate, string sToDate, int iOption)
        {
            try
            {
                decimal iAverageOnTat = 0;
                decimal iAverageOffTat = 0;
                DataSet _dsTatPercentage = new DataSet();
                _dsTatPercentage = BusinessLogic.WS_Allocation.Get_TAT_percentage(Convert.ToInt32(sYear), Convert.ToInt32(sMonth), Convert.ToDateTime(sFromDate), Convert.ToDateTime(sToDate), iOption);

                int iRowCount = 0;
                lvTatPercentage.Items.Clear();
                foreach (DataRow _dr in _dsTatPercentage.Tables[0].Select())
                {
                    lvTatPercentage.Items.Add(new ListItem_TatPercentage(_dr, iRowCount++));
                    iAverageOnTat += Convert.ToDecimal(_dr["ToCompute"].ToString());
                    iAverageOffTat += Convert.ToDecimal(_dr["ToComputeOff"].ToString());
                }
                //int iOn = Convert.ToInt32(iAverageOnTat / iRowCount);
                //int iOff = Convert.ToInt32(iAverageOffTat / iRowCount);

                decimal iOn = Convert.ToDecimal(iAverageOnTat / iRowCount);
                decimal iOff = Convert.ToDecimal(iAverageOffTat / iRowCount);
                ListItem_TatPercentage oListItem;
                oListItem = new ListItem_TatPercentage("Average TAT: ", Math.Round(iOn, 2).ToString() + "%", Math.Round(iOff, 2).ToString() + "%");
                lvTatPercentage.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lvTatPercentage);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// GET HOLD PERCENTAGE
        /// </summary>
        private void Load_Hold_Percentage(int sMonth, string iYear)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                lbliasis.Text = "";
                lblMkmg.Text = "";
                lblHCA.Text = "";
                DataSet _dsHoldPercentage = new DataSet();
                _dsHoldPercentage = BusinessLogic.WS_Allocation.Get_HoldPercentage(Convert.ToInt32(sMonth), Convert.ToInt32(iYear));

                lbliasis.Text = _dsHoldPercentage.Tables[0].Rows[0]["HoldPercentage_IASIS"].ToString() + " %";
                lblMkmg.Text = _dsHoldPercentage.Tables[1].Rows[0]["HoldPercentageMKMG"].ToString() + " %";
                lblHCA.Text = _dsHoldPercentage.Tables[2].Rows[0]["HoldPercentageHCA"].ToString() + " %";

                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD TOTAL EMPLOYEE ACHIEVEMENT
        /// </summary>
        private void LoadEmployeeAchievement()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");

                int iBatchId = 0;
                object oFromDate = null;
                object oTodate = null;

                iBatchId = Convert.ToInt32(cboDesignation.SelectedValue);
                oFromDate = Convert.ToDateTime(dtpEmployeeFromDate.Value).ToString("yyyy/MM/dd");
                oTodate = Convert.ToDateTime(dtpEmployeeFromDate.Value).ToString("yyyy/MM/dd");

                lvEmployeeAchievement.Items.Clear();
                DataSet dsEmployeeAchievement = new DataSet();

                dsEmployeeAchievement = BusinessLogic.WS_Allocation.Get_EmployeeTotalMinutes(iBatchId, oFromDate, oTodate);

                int iRowCount = 0;
                int iTotalSeconds = 0;
                foreach (DataRow drEmployee in dsEmployeeAchievement.Tables[0].Rows)
                {
                    lvEmployeeAchievement.Items.Add(new ListItem_EmployeeTotalMinutes(drEmployee, iRowCount++));
                    iTotalSeconds = iTotalSeconds + Convert.ToInt32(drEmployee["Minutes"].ToString());
                }

                string oMins = sGetDuration(iTotalSeconds);

                Int32 iTotalFiles = 0;
                decimal dConvertedLines = 0;
                foreach (ListViewItem o in this.lvEmployeeAchievement.Items)
                {
                    iTotalFiles = iTotalFiles + Convert.ToInt32(o.SubItems[2].Text);
                    dConvertedLines = dConvertedLines + Convert.ToDecimal(o.SubItems[4].Text);
                }

                ListViewItem lvTotal = new ListViewItem();
                lvTotal.SubItems.Add("Total");
                lvTotal.SubItems.Add(iTotalFiles.ToString());
                lvTotal.SubItems.Add(oMins);
                lvTotal.SubItems.Add(dConvertedLines.ToString());
                lvTotal.Font = new Font(lvTotal.Font, FontStyle.Bold);
                lvTotal.BackColor = System.Drawing.ColorTranslator.FromHtml("#622c2c");
                lvTotal.ForeColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                lvEmployeeAchievement.Items.Add(lvTotal);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done....!");
            }
        }

        /// <summary>
        /// LOAD EMPLOYEE FULL NAME
        /// </summary>
        /// <param name="BatchID"></param>
        private void Load_Employee_Full_name(string BatchID, string BranchID)
        {
            try
            {
                DataSet _dsEmployee = new DataSet();
                //_dsEmployee = BusinessLogic.WS_Allocation.Get_Desigwise_employees(Convert.ToInt32(BatchID));
                _dsEmployee = BusinessLogic.WS_Allocation.Get_branch_Desigwise_employees(Convert.ToInt32(BatchID), Convert.ToInt32(BranchID));

                if ((tabAllocation.SelectedTab.Name == "tbNightshiftusers"))
                {
                    cmb_Nightshift_EmpName.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cmb_Nightshift_EmpName.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cmb_Nightshift_EmpName.DataSource = _dsEmployee.Tables[0];
                    cmb_Nightshift_EmpName.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmb_Nightshift_EmpName.Tag = Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG;
                }

                if (tabAllocation.SelectedTab.Name == "tbCustomizeEntry")
                {
                    cmb_Customized_Employee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cmb_Customized_Employee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cmb_Customized_Employee.DataSource = _dsEmployee.Tables[0];
                    cmb_Customized_Employee.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmb_Customized_Employee.Tag = Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG;
                }

                if (TabExtraction.SelectedTab.Name == "tabOnlineEntry")
                {
                    cbxEntryEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cbxEntryEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cbxEntryEmployee.DataSource = _dsEmployee.Tables[0];
                    cbxEntryEmployee.DropDownStyle = ComboBoxStyle.DropDownList;

                    cbxInEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cbxInEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cbxInEmployee.DataSource = _dsEmployee.Tables[0];
                    cbxInEmployee.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if (TabExtraction.SelectedTab.Name == "tabTarget")
                {
                    if (BusinessLogic.USERNAME == "Admin-Trivandrum")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataRow drRow = _dsEmployee.Tables[0].NewRow();
                            drRow[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID] = 0;
                            drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME] = "--Select--";
                            _dsEmployee.Tables[0].Rows.InsertAt(drRow, 0);

                            DataTable _dtWithLinq1 = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                      where dr["branch_id"].ToString() == "3"
                                                      select dr).CopyToDataTable();

                            cmbEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbEmployee.DataSource = _dtWithLinq1;
                            cmbEmployee.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                    if (BusinessLogic.USERNAME == "Admin-Cochin")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataRow drRow = _dsEmployee.Tables[0].NewRow();
                            drRow[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID] = 0;
                            drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME] = "--Select--";
                            _dsEmployee.Tables[0].Rows.InsertAt(drRow, 0);

                            DataTable _dtWithLinq1 = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                      where dr["branch_id"].ToString() == "2"
                                                      select dr).CopyToDataTable();

                            cmbEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbEmployee.DataSource = _dtWithLinq1;
                            cmbEmployee.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                    if (BusinessLogic.USERNAME == "Admin-Pondichery")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataRow drRow = _dsEmployee.Tables[0].NewRow();
                            drRow[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID] = 0;
                            drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME] = "--Select--";
                            _dsEmployee.Tables[0].Rows.InsertAt(drRow, 0);

                            DataTable _dtWithLinq1 = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                      where dr["branch_id"].ToString() == "4"
                                                      select dr).CopyToDataTable();

                            cmbEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbEmployee.DataSource = _dtWithLinq1;
                            cmbEmployee.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                    else
                    {
                        DataRow drRow = _dsEmployee.Tables[0].NewRow();
                        drRow[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID] = 0;
                        drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME] = "--Select--";
                        _dsEmployee.Tables[0].Rows.InsertAt(drRow, 0);

                        cmbEmployee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                        cmbEmployee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                        cmbEmployee.DataSource = _dsEmployee.Tables[0];
                        cmbEmployee.DropDownStyle = ComboBoxStyle.DropDownList;
                    }

                    if (BusinessLogic.USERNAME == "Admin-Trivandrum")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                     where dr["branch_id"].ToString() == "3"
                                                     select dr).CopyToDataTable();

                            cmbLogUser.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbLogUser.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbLogUser.DataSource = _dtWithLinq;
                            cmbLogUser.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                    if (BusinessLogic.USERNAME == "Admin-Cochin")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                     where dr["branch_id"].ToString() == "2"
                                                     select dr).CopyToDataTable();

                            cmbLogUser.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbLogUser.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbLogUser.DataSource = _dtWithLinq;
                            cmbLogUser.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                    if (BusinessLogic.USERNAME == "Admin-Pondichery")
                    {
                        if (_dsEmployee.Tables[0].Rows.Count > 0)
                        {
                            DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                                     where dr["branch_id"].ToString() == "4"
                                                     select dr).CopyToDataTable();

                            cmbLogUser.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                            cmbLogUser.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                            cmbLogUser.DataSource = _dtWithLinq;
                            cmbLogUser.DropDownStyle = ComboBoxStyle.DropDownList;
                        }
                    }
                }
                else if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tabLogSheet"))
                {
                    cmbLogUser.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cmbLogUser.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cmbLogUser.DataSource = _dsEmployee.Tables[0];
                    cmbLogUser.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_Logsheethourly"))
                {
                    cmb_Log_Employee.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cmb_Log_Employee.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cmb_Log_Employee.DataSource = _dsEmployee.Tables[0];
                    cmb_Log_Employee.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if ((tabAllocation.SelectedTab.Name == "tabPReports") && (tabReports.SelectedTab.Name == "tab_user_consolidated"))
                {
                    comboBox16.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    comboBox16.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    comboBox16.DataSource = _dsEmployee.Tables[0];
                    comboBox16.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else
                {
                    cbxUserName.ValueMember = Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID;
                    cbxUserName.DisplayMember = Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME;
                    cbxUserName.DataSource = _dsEmployee.Tables[0];
                    cbxUserName.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// EXPORT TO EXCEL
        /// </summary>
        /// <param name="myList"></param>
        /// <param name="sFolderNAme"></param>
        /// <param name="sXLSName"></param>
        /// <returns></returns>
        private string ExportToExcel(ListView myList, string sFolderNAme, string sXLSName)
        {
            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            object oMissing = System.Reflection.Missing.Value;
            string sExcelFullFileName;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                BusinessLogic.oMessageEvent.Start("Exporting...");
                BusinessLogic.oProgressEvent.Start(true);

                if (!Directory.Exists(sFolderNAme))
                    Directory.CreateDirectory(sFolderNAme);

                if (!Directory.Exists(sFolderNAme))
                    return string.Empty;

                sExcelFullFileName = Path.Combine(sFolderNAme, sXLSName);

                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                wb = app.Workbooks.Add(1);
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                ws.Name = "Report";
                int i = 1;
                int i2 = 1;
                i = 1;

                foreach (ColumnHeader oHeader in myList.Columns)
                {
                    ws.Cells[i2, i] = oHeader.Text;
                    i++;
                }

                Microsoft.Office.Interop.Excel.Range all = app.get_Range("A1:IV1", Type.Missing);
                all.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon);
                all.Font.Size = 12;
                all.Font.Name = "Calibri";
                all.Font.Bold = true;
                app.ActiveWindow.DisplayGridlines = true;

                all = null;
                all = app.get_Range("A1:IV1000", Type.Missing);
                all.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous,
                Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium,
                Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon));

                i2 = 2;
                foreach (ListViewItem lvi in myList.Items)
                {
                    if (lvi.Text.Trim().Length <= 0)
                    {
                        string sRange = "A" + i2.ToString() + ":" + "E" + i2.ToString();
                        Microsoft.Office.Interop.Excel.Range all1 = app.get_Range(sRange, Type.Missing);
                        all1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        all1.Font.Size = 12;
                        all1.Font.Name = "Calibri";
                        all1.Font.Bold = true;
                        app.ActiveWindow.DisplayGridlines = true;

                        sRange = "F" + i2.ToString() + ":" + "IV" + i2.ToString();
                        all1 = app.get_Range(sRange, Type.Missing);
                        all1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        all1.Font.Size = 12;
                        all1.Font.Name = "Calibri";
                        all1.Font.Bold = true;
                        app.ActiveWindow.DisplayGridlines = true;

                        all1 = null;
                        all1 = app.get_Range("A1:IV1000", Type.Missing);
                        all1.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous,
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium,
                        Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon));
                    }

                    i = 1;
                    foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                    {
                        ws.Cells[i2, i] = lvs.Text;
                        i++;
                    }
                    i2++;
                }
                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    if (File.Exists(sExcelFullFileName))
                        File.Delete(sExcelFullFileName);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                }
                ws.Columns.AutoFit();
                ws.SaveAs(sExcelFullFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                wb.Save();
                wb.Close(true, sExcelFullFileName, oMissing);
                BusinessLogic.oMessageEvent.Start("Done");

                if (File.Exists(sExcelFullFileName))
                    return sExcelFullFileName;
                else
                    return string.Empty;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                return string.Empty;
            }
            finally
            {
                this.Cursor = Cursors.Default;
                BusinessLogic.oProgressEvent.Start(false);

                if (app != null)
                {
                    app.DisplayAlerts = true;
                    app.Quit();
                    Release(ws);
                    Release(ws);
                    Release(wb);
                    Release(wb);
                    Release(app);
                    ws = null;
                    ws = null;
                    wb = null;
                    wb = null;
                    app = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private static void Release(object obj)
        {
            try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj); }
            catch { }
        }

        /// <summary>
        /// LOAD HIGHER MINUTES
        /// </summary>
        private void Load_LargeMinutes()
        {
            try
            {
                lsvLargeMinutes.Items.Clear();
                BusinessLogic.oMessageEvent.Start("Transferring data..!");
                BusinessLogic.oProgressEvent.Start(true);
                string sFromDate = dtpLMfrom.Value.ToString("yyyy/MM/dd");
                string sToDate = dtpLMto.Value.ToString("yyyy/MM/dd");
                DataSet _dsLargeFiles = new DataSet();
                _dsLargeFiles = BusinessLogic.WS_Allocation.Get_Large_Minutes(sFromDate, sToDate);

                if (_dsLargeFiles == null)
                {
                    BusinessLogic.oMessageEvent.Start("No Records to display");
                    return;
                }
                int iRowCount = 1;

                foreach (DataRow _dr in _dsLargeFiles.Tables[0].Select())
                    lsvLargeMinutes.Items.Add(new ListItem_LargeMinutes(_dr, ++iRowCount));

                BusinessLogic.Reset_ListViewColumn(lsvLargeMinutes);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD TRANS EDIT FILES
        /// </summary>
        private void LoadTransEditFiles()
        {
            try
            {
                lsvTransEdit.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_TransEdit oListItem;

                string FromDate = Convert.ToDateTime(dtpTransEditFrom.Text).ToString("yyyy-MM-dd");
                string Todate = Convert.ToDateTime(dtpTransEditTo.Text).ToString("yyyy-MM-dd");

                DataTable _dtTransEdit = BusinessLogic.WS_Allocation.Get_TransEditDetails(Convert.ToDateTime(FromDate), Convert.ToDateTime(Todate));

                int iRowCount = 0;
                decimal dTotLines = 0;
                decimal dTotConvLines = 0;
                int dTotMins = 0;
                int dTotConvMins = 0;

                foreach (DataRow _drRow in _dtTransEdit.Select())
                {
                    lsvTransEdit.Items.Add(new ListItem_TransEdit(_drRow, iRowCount++));
                    dTotLines += Convert.ToDecimal(_drRow["file_lines"].ToString());
                    dTotConvLines += Convert.ToDecimal(_drRow["converted_lines"].ToString());
                    dTotMins += Convert.ToInt32(_drRow["CalSec"].ToString());
                    dTotConvMins += Convert.ToInt32(_drRow["Converted_Seconds"].ToString());
                }

                string oMins = sGetDuration(dTotMins);
                string oConvMins = sGetDuration(dTotConvMins);

                oListItem = new ListItem_TransEdit("Linecount summary for the date : " + Convert.ToDateTime(FromDate).ToString("dd-MM-yyyy"), oMins.ToString(), oConvMins.ToString(), dTotLines.ToString(), dTotConvLines.ToString());
                lsvTransEdit.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvTransEdit);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD TRANS EDIT FILES
        /// </summary>
        private void LoadTransEditFiles_Clinics()
        {
            try
            {
                lvTransEditOff.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_TransEdit_Clinics oListItem;

                string FromDate = Convert.ToDateTime(dtpTransEditFrom.Text).ToString("yyyy-MM-dd");
                string Todate = Convert.ToDateTime(dtpTransEditTo.Text).ToString("yyyy-MM-dd");

                DataTable _dtTransEditClinics = BusinessLogic.WS_Allocation.Get_TransEditDetails_CLINICS(Convert.ToDateTime(FromDate), Convert.ToDateTime(Todate));

                int iRowCount = 0;
                decimal dTotLines = 0;
                decimal dTotConvLines = 0;
                int dTotMins = 0;
                int dTotConvMins = 0;

                foreach (DataRow _drRow in _dtTransEditClinics.Select())
                {
                    lvTransEditOff.Items.Add(new ListItem_TransEdit_Clinics(_drRow, iRowCount++));
                    dTotLines += Convert.ToDecimal(_drRow["file_lines"].ToString());
                    dTotConvLines += Convert.ToDecimal(_drRow["converted_lines"].ToString());
                    dTotMins += Convert.ToInt32(_drRow["CalSec"].ToString());
                    dTotConvMins += Convert.ToInt32(_drRow["Converted_Seconds"].ToString());
                }

                string oMins = sGetDuration(dTotMins);
                string oConvMins = sGetDuration(dTotConvMins);

                oListItem = new ListItem_TransEdit_Clinics("Linecount summary for the date : " + Convert.ToDateTime(FromDate).ToString("dd-MM-yyyy"), oMins.ToString(), oConvMins.ToString(), dTotLines.ToString(), dTotConvLines.ToString());
                lvTransEditOff.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lvTransEditOff);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD FILE STATUS
        /// </summary>
        private void Load_File_Status()
        {
            try
            {
                //DataSet _dsFileStatus = new DataSet();
                //_dsFileStatus = BusinessLogic.WS_Allocation.Get_file_Status();

                //cbxFileStatus.ValueMember = Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT;
                //cbxFileStatus.DisplayMember = Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR;
                //cbxFileStatus.DataSource = _dsFileStatus.Tables[0];
                //cbxFileStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD ACCOUNT WISE INFO
        /// </summary>
        private void Load_Account_Wise_Minutes_New()
        {
            try
            {
                string iFromHour, iToHour;
                object oFromDate = null;
                object oTodate = null;

                lvAccountWiseInfo.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_AccountWiseInfo oListItem;

                string AccFromDate = Convert.ToDateTime(dtpAccountFrom.Text).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtpAccountTo.Text).ToString("yyyy-MM-dd");
                DataTable _dsAccountWiseInfo = null;

                iFromHour = cmb_Acc_Fromhour.SelectedItem.ToString();
                iToHour = cmb_Acc_Tohour.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = AccFromDate + " 23:59:59";
                else
                    oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = AccTodate + " 23:59:59";
                else
                    oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                if (rdb_downloaded.Checked == true)
                    //_dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWiseInfo(AccFromDate, AccTodate);
                    _dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWiseInfo(oFromDate, oTodate);
                else
                    //_dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWise_Processed_Info(AccFromDate, AccTodate);
                    _dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWise_Processed_Info_V2(oFromDate, oTodate, Convert.ToInt32(cmb_Acc_Status.SelectedValue));

                int iRowCount = 0;
                //decimal dTotLines = 0;
                int dTotMins = 0;

                foreach (DataRow _drRow in _dsAccountWiseInfo.Select())
                {
                    lvAccountWiseInfo.Items.Add(new ListItem_AccountWiseInfo(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                }

                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_AccountWiseInfo("Account Wise Minutes on: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins.ToString());
                lvAccountWiseInfo.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lvAccountWiseInfo);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD ACCOUNT WISE INFO
        /// </summary>
        private void Load_Account_Wise_Minutes()
        {
            try
            {
                lvAccountWiseInfo.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_AccountWiseInfo oListItem;

                string AccFromDate = Convert.ToDateTime(dtpAccountFrom.Text).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtpAccountTo.Text).ToString("yyyy-MM-dd");
                DataTable _dsAccountWiseInfo = null;

                if (rdb_downloaded.Checked == true)
                    _dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWiseInfo(AccFromDate, AccTodate);
                else
                    _dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWise_Processed_Info(AccFromDate, AccTodate);

                int iRowCount = 0;
                //decimal dTotLines = 0;
                int dTotMins = 0;

                foreach (DataRow _drRow in _dsAccountWiseInfo.Select())
                {
                    lvAccountWiseInfo.Items.Add(new ListItem_AccountWiseInfo(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                }

                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_AccountWiseInfo("Account Wise Minutes on: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins.ToString());
                lvAccountWiseInfo.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lvAccountWiseInfo);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD ACCOUNT WISE INFO
        /// </summary>
        private void Load_Offline_Account_Wise_Minutes()
        {
            try
            {
                lvOfflineAccountWise.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_Offline_AccountWiseInfo oListItem;

                string AccFromDate = Convert.ToDateTime(dtpOffAccountFrom.Text).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtpOffAccountTo.Text).ToString("yyyy-MM-dd");
                DataTable _dsAccountWiseInfo_Offline = null;

                if (rdb_offline_Downloaded.Checked == true)
                    _dsAccountWiseInfo_Offline = BusinessLogic.WS_Allocation.Get_accountWise_Downloaded_Info_Offline(AccFromDate, AccTodate);
                else
                    _dsAccountWiseInfo_Offline = BusinessLogic.WS_Allocation.Get_accountWiseInfo_Offline(AccFromDate, AccTodate);

                int iRowCount = 0;
                //decimal dTotLines = 0;
                int dTotMins = 0;

                foreach (DataRow _drRow in _dsAccountWiseInfo_Offline.Select())
                {
                    lvOfflineAccountWise.Items.Add(new ListItem_Offline_AccountWiseInfo(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                }

                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_Offline_AccountWiseInfo("Account Wise Minutes on: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins.ToString());
                lvOfflineAccountWise.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lvOfflineAccountWise);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD ACCOUNT WISE INFO BACKLOCK
        /// </summary>
        private void Load_Offline_Account_Wise_BackLock()
        {
            try
            {
                lsvBackLockFiles.Items.Clear();
                lvBackLocKME.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_Offline_AccountWiseInfo_BackLock oListItem_BackLock;

                string AccFromDate = Convert.ToDateTime(dtpBackFrom.Value).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtpBackTo.Value).ToString("yyyy-MM-dd");

                DataSet _dsBackLockInfo = BusinessLogic.WS_Allocation.Get_accountWiseInfo_Offline_BackLock(AccFromDate, AccTodate);

                int iRowCount = 0;
                //decimal dTotLines = 0;
                int dTotMins = 0;

                // MT BACKLOG

                foreach (DataRow _drRow in _dsBackLockInfo.Tables[0].Select())
                {
                    lsvBackLockFiles.Items.Add(new ListItem_Offline_AccountWiseInfo_BackLock(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                }

                string oMins = sGetDuration(dTotMins);

                oListItem_BackLock = new ListItem_Offline_AccountWiseInfo_BackLock("Back Lock Minutes On: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins.ToString());
                lsvBackLockFiles.Items.Add(oListItem_BackLock);
                BusinessLogic.Reset_ListViewColumn(lsvBackLockFiles);

                // ME BACKLOG

                int iRowCount1 = 0;
                //decimal dTotLines = 0;
                int dTotMins1 = 0;

                foreach (DataRow _drRowNew in _dsBackLockInfo.Tables[1].Select())
                {
                    lvBackLocKME.Items.Add(new ListItem_Offline_AccountWiseInfo_BackLock(_drRowNew, iRowCount1++));
                    if (Convert.ToInt32(_drRowNew["Tot_minutes"]) != 0)
                    {
                        if (_drRowNew["Tot_minutes"].ToString().Contains('.'))
                            dTotMins1 += Convert.ToInt32(_drRowNew["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins1 += Convert.ToInt32(_drRowNew["Tot_minutes"].ToString());
                    }
                }

                string oMins1 = sGetDuration(dTotMins1);

                oListItem_BackLock = new ListItem_Offline_AccountWiseInfo_BackLock("Back Lock Minutes On: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins1.ToString());
                lvBackLocKME.Items.Add(oListItem_BackLock);
                BusinessLogic.Reset_ListViewColumn(lvBackLocKME);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD MONTH NAMES INTO DROP DOWN
        /// </summary>
        private void Load_Month_Name()
        {
            try
            {
                //if ((tabOffline.SelectedTab.Name == "tabOfflineReport") && (tbConOfReport.SelectedTab.Name == "tbPageVolume"))
                //{
                //    //for (int i = 0; i < 12; i++)
                //    //{
                //    //    cmbVolMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);                        
                //    //}
                //    //cmbVolMonth.SelectedIndex = DateTime.Now.Month - 1;
                //    //cmbVolMonth.DropDownStyle = ComboBoxStyle.DropDownList;
                //}
                if ((tabAllocation.SelectedTab.Name == "tabPReports"))
                {
                    for (int i = 0; i < 12; i++)
                    {
                        cmbHoldPercentageMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                        comboBox18.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                    }
                    cmbHoldPercentageMonth.SelectedIndex = DateTime.Now.Month - 1;
                    cmbHoldPercentageMonth.DropDownStyle = ComboBoxStyle.DropDownList;

                    comboBox18.SelectedIndex = DateTime.Now.Month - 1;
                    comboBox18.DropDownStyle = ComboBoxStyle.DropDownList;                    
                }
                else if ((tabOffline.SelectedTab.Name == "tabOfflineReport") && (tbConOfReport.SelectedTab.Name == "tbTatPercentage"))
                {
                    cmbTatMonth.Items.Clear();
                    for (int i = 0; i < 12; i++)
                    {
                        cmbTatMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                    }
                    cmbTatMonth.SelectedIndex = DateTime.Now.Month - 1;
                    cmbTatMonth.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else
                {
                    cmbTargetMonth.Items.Clear();
                    comboBox4.Items.Clear();
                    while (cmbTargetMonth.Items.Count > 0)
                    {
                        cmbTargetMonth.Items.RemoveAt(0);
                    }

                    for (int i = 0; i < 12; i++)
                    {
                        cmbTargetMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                        comboBox4.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                    }
                    cmbTargetMonth.SelectedIndex = DateTime.Now.Month - 1;
                    cmbTargetMonth.DropDownStyle = ComboBoxStyle.DropDownList;

                    comboBox4.SelectedIndex = DateTime.Now.Month - 1;
                    comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void Load_Month_Name_ListView()
        {
            try
            {
                for (int i = 0; i < 12; i++)
                {
                    string sMonthName = CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i].ToString();
                    ListViewItem oItem = new ListViewItem();
                    oItem.Tag = sMonthName;
                    oItem.Text = sMonthName;
                    lvMonthName.Items.Add(oItem);

                    if (i % 2 == 1)
                        this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                    else
                        this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void Load_Account_ListView()
        {
            try
            {
                DataSet _dsAccount = new DataSet();
                _dsAccount = BusinessLogic.WS_Allocation.Get_ClientName();

                int iRowCount = 0;
                lsvAccountName.Items.Clear();
                foreach (DataRow _drAccount in _dsAccount.Tables[1].Select())
                    lsvAccountName.Items.Add(new ListItem_Account(_drAccount, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvAccountName);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD LOCATION BASED ON FILES
        /// </summary>
        private void Load_Location_ListView()
        {
            try
            {
                int iFileCount = 0;
                string dTotMins = string.Empty;
                int dTotSecs = 0;

                lvJobAccount.Items.Clear();
                int iOption = 1;
                if (cbxOffline_Trans.Checked == true)
                {
                    iOption = 1;
                }
                else if (cbxOffline_Editing.Checked == true)
                {
                    iOption = 2;
                }
                else if (cbxOffline_Review.Checked == true)
                {
                    iOption = 3;
                }
                DataSet _dsLocation = new DataSet();
                _dsLocation = BusinessLogic.WS_Allocation.Get_ClientName_Download(iOption);

                if ((tabControlMain.SelectedTab.Name == "tabPageOffline") && (tabOffline.SelectedTab.Name == "tabOffFileDeatils"))
                {
                    int iRowCount = 0;
                    lvJobAccount.Items.Clear();
                    foreach (DataRow _drAccount in _dsLocation.Tables[0].Select())
                    {
                        lvJobAccount.Items.Add(new Mylsvdownloaddetails(_drAccount, iRowCount++));
                        if (_drAccount["file_minutes_total"].ToString() != "")
                        {
                            if (Convert.ToInt32(_drAccount["file_minutes_total"]) != 0)
                            {
                                if (_drAccount["file_minutes_total"].ToString().Contains('.'))
                                    dTotSecs += Convert.ToInt32(_drAccount["file_minutes_total"].ToString().Split('.').GetValue(0));
                                else
                                    dTotSecs += Convert.ToInt32(_drAccount["file_minutes_total"].ToString());

                                iFileCount += Convert.ToInt32(_drAccount["File_Count"].ToString());
                            }
                        }
                    }

                    dTotMins = sGetDuration(dTotSecs);
                    Mylsvdownloaddetails oListItem;
                    oListItem = new Mylsvdownloaddetails("Total files for the day: ", dTotMins.ToString(), iFileCount.ToString());
                    lvJobAccount.Items.Add(oListItem);
                    BusinessLogic.Reset_ListViewColumn(lvJobAccount);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD YEAR INTO DROP DOWN
        /// </summary>
        private void LoadYear()
        {
            try
            {
                //if ((tabOffline.SelectedTab.Name == "tabOfflineReport") && (tbConOfReport.SelectedTab.Name == "tbPageVolume"))
                if ((tabAllocation.SelectedTab.Name == "tabPReports"))
                {
                    cmbHoldPercentageYear.Items.Clear();
                    comboBox19.Items.Clear();

                    int iCurrentYear = DateTime.Now.Year;
                    for (int i = 2014; i <= iCurrentYear; i++)
                    {
                        cmbHoldPercentageYear.Items.Add(i.ToString());
                        comboBox19.Items.Add(i.ToString());
                    }
                    cmbHoldPercentageYear.SelectedIndex = 0;
                    //cmbHoldPercentageYear.SelectedIndex = (cmbHoldPercentageYear.Items.Count + 1);
                    cmbHoldPercentageYear.DropDownStyle = ComboBoxStyle.DropDownList;

                    comboBox19.SelectedIndex = 0;
                    //comboBox19.SelectedIndex = (comboBox19.Items.Count + 1);
                    comboBox19.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if ((tabOffline.SelectedTab.Name == "tabOfflineReport") && (tbConOfReport.SelectedTab.Name == "tbTatPercentage"))
                {
                    cmbTatYear.Items.Clear();
                    int iCurrentYear = DateTime.Now.Year;
                    for (int i = 2014; i <= iCurrentYear; i++)
                    {
                        cmbTatYear.Items.Add(i.ToString());
                    }
                    cmbTatYear.SelectedIndex = 0;
                    cmbTatYear.SelectedIndex = (cmbTatYear.Items.Count + 1);
                    cmbTatYear.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else
                {
                    cmbTargetYear.Items.Clear();
                    comboBox5.Items.Clear();
                    int iCurrentYear = DateTime.Now.Year;
                    for (int i = 2014; i <= iCurrentYear; i++)
                    {
                        cmbTargetYear.Items.Add(i.ToString());
                        comboBox5.Items.Add(i.ToString());
                    }
                    cmbTargetYear.SelectedIndex = (cmbTargetYear.Items.Count - 1);
                    cmbTargetYear.DropDownStyle = ComboBoxStyle.DropDownList;
                    comboBox5.SelectedIndex = (comboBox5.Items.Count - 1);
                    comboBox5.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD TARGET
        /// </summary>
        private void Load_Target()
        {
            try
            {
                if (BusinessLogic.USERNAME == "Admin-Trivandrum")
                {
                    lvTarget.Items.Clear();
                    int sMonth = cmbTargetMonth.SelectedIndex + 1;
                    string iYear = cmbTargetYear.SelectedItem.ToString();
                    string sProductionID = string.Empty;
                    if (cmbEmployee.SelectedIndex != 0)
                        sProductionID = cmbEmployee.SelectedValue.ToString();
                    else
                        sProductionID = "-1";

                    DataSet _dsTarget = new DataSet();
                    _dsTarget = BusinessLogic.WS_Allocation.Get_Traget(sMonth, Convert.ToInt32(iYear), Convert.ToInt32(sProductionID), Convert.ToInt32(cmbTargetDisig.SelectedValue));

                    DataTable _dtWithLinqTarget = (from DataRow dr in _dsTarget.Tables[0].Select()
                                                   where dr["branch_id"].ToString() == "3"
                                                   select dr).CopyToDataTable();

                    int iRowCount = 0;
                    foreach (DataRow _drTarget in _dtWithLinqTarget.Select())
                        lvTarget.Items.Add(new ListItem_Target(_drTarget, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lvTarget);
                }
                else if (BusinessLogic.USERNAME == "Admin-Cochin")
                {
                    lvTarget.Items.Clear();
                    int sMonth = cmbTargetMonth.SelectedIndex + 1;
                    string iYear = cmbTargetYear.SelectedItem.ToString();
                    string sProductionID = string.Empty;
                    if (cmbEmployee.SelectedIndex != 0)
                        sProductionID = cmbEmployee.SelectedValue.ToString();
                    else
                        sProductionID = "-1";

                    DataSet _dsTarget = new DataSet();
                    _dsTarget = BusinessLogic.WS_Allocation.Get_Traget(sMonth, Convert.ToInt32(iYear), Convert.ToInt32(sProductionID), Convert.ToInt32(cmbTargetDisig.SelectedValue));

                    DataTable _dtWithLinqTarget = (from DataRow dr in _dsTarget.Tables[0].Select()
                                                   where dr["branch_id"].ToString() == "2"
                                                   select dr).CopyToDataTable();

                    int iRowCount = 0;
                    foreach (DataRow _drTarget in _dtWithLinqTarget.Select())
                        lvTarget.Items.Add(new ListItem_Target(_drTarget, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lvTarget);
                }
                else if (BusinessLogic.USERNAME == "Admin-Pondichery")
                {
                    lvTarget.Items.Clear();
                    int sMonth = cmbTargetMonth.SelectedIndex + 1;
                    string iYear = cmbTargetYear.SelectedItem.ToString();
                    string sProductionID = string.Empty;
                    if (cmbEmployee.SelectedIndex != 0)
                        sProductionID = cmbEmployee.SelectedValue.ToString();
                    else
                        sProductionID = "-1";

                    DataSet _dsTarget = new DataSet();
                    _dsTarget = BusinessLogic.WS_Allocation.Get_Traget(sMonth, Convert.ToInt32(iYear), Convert.ToInt32(sProductionID), Convert.ToInt32(cmbTargetDisig.SelectedValue));

                    DataTable _dtWithLinqTarget = (from DataRow dr in _dsTarget.Tables[0].Select()
                                                   where dr["branch_id"].ToString() == "4"
                                                   select dr).CopyToDataTable();

                    int iRowCount = 0;
                    foreach (DataRow _drTarget in _dtWithLinqTarget.Select())
                        lvTarget.Items.Add(new ListItem_Target(_drTarget, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lvTarget);
                }
                else
                {
                    lvTarget.Items.Clear();
                    int sMonth = cmbTargetMonth.SelectedIndex + 1;
                    string iYear = cmbTargetYear.SelectedItem.ToString();
                    string sProductionID = string.Empty;
                    if (cmbEmployee.SelectedIndex != 0)
                        sProductionID = cmbEmployee.SelectedValue.ToString();
                    else
                        sProductionID = "-1";

                    DataSet _dsTarget = new DataSet();
                    _dsTarget = BusinessLogic.WS_Allocation.Get_Traget(sMonth, Convert.ToInt32(iYear), Convert.ToInt32(sProductionID), Convert.ToInt32(cmbTargetDisig.SelectedValue));

                    int iRowCount = 0;
                    foreach (DataRow _drTarget in _dsTarget.Tables[0].Select())
                        lvTarget.Items.Add(new ListItem_Target(_drTarget, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lvTarget);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD NO ENTRY
        /// </summary>
        private void Load_Discrepancy()
        {
            try
            {
                lblMissingFiles.Text = string.Empty;
                lblChangeMinutes.Text = string.Empty;
                lblChangeLines.Text = string.Empty;
                lblChangeShift.Text = string.Empty;
                lblTransEdit.Text = string.Empty;
                lblChangeFilesStatus.Text = string.Empty;

                iMissingEntryCount = 1;
                iChangeMinutesCount = 1;
                iChangeLines = 1;
                iChangeShift = 1;
                iConvertTransEdit = 1;
                iChangeFilesStatus = 1;

                lsvDiscrepancy.Items.Clear();
                DataSet _dsDiscrepancy = new DataSet();
                _dsDiscrepancy = BusinessLogic.WS_Allocation.Get_Discrepancy();

                int iRowCount = 0;
                foreach (DataRow _drDiscrepancy in _dsDiscrepancy.Tables[0].Select())
                {
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 1)
                    {
                        lblMissingFiles.Text = iMissingEntryCount++.ToString();
                    }
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 2)
                    {
                        lblChangeMinutes.Text = iChangeMinutesCount++.ToString();
                    }
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 3)
                    {
                        lblChangeLines.Text = iChangeLines++.ToString();
                    }
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 4)
                    {
                        lblChangeShift.Text = iChangeShift++.ToString();
                    }
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 5)
                    {
                        lblTransEdit.Text = iConvertTransEdit++.ToString();
                    }
                    if (Convert.ToInt32(_drDiscrepancy[Framework.MASTER_DISCREPANCY.DISCREPANCY_ID].ToString()) == 6)
                    {
                        lblChangeFilesStatus.Text = iChangeFilesStatus++.ToString();
                    }

                    lsvDiscrepancy.Items.Add(new ListItem_Discrepancy(_drDiscrepancy, iRowCount++));
                }
                BusinessLogic.Reset_ListViewColumn(lsvDiscrepancy);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void Load_Discrepancy_Report()
        {
            try
            {
                lsvDiscrepancyReport.Items.Clear();

                DataSet _dsDiscrepancyReport = new DataSet();
                _dsDiscrepancyReport = BusinessLogic.WS_Allocation.Get_Discrepancy_Report();

                int iRowCount = 0;
                foreach (DataRow _drReport in _dsDiscrepancyReport.Tables[0].Select())
                    lsvDiscrepancyReport.Items.Add(new ListItem_Discrepancy_Report(_drReport, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvDiscrepancyReport);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void Load_Multiple_Entries()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                lvMultipleEntries.Items.Clear();
                string sFromdate = Convert.ToDateTime(dtpMFromDate.Value).ToString("yyyy-MM-dd");
                string sTo = Convert.ToDateTime(dtpMToDate.Value).ToString("yyyy-MM-dd");

                int iRowCount = 0;
                DataSet _dsMultipleEntries = new DataSet();
                _dsMultipleEntries = BusinessLogic.WS_Allocation.Get_Multiple_entries(Convert.ToDateTime(sFromdate), Convert.ToDateTime(sTo));

                foreach (DataRow _drRow in _dsMultipleEntries.Tables[0].Select())
                    lvMultipleEntries.Items.Add(new ListItem_Multiple_Entries(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvMultipleEntries);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void LoadEnteredFilesDetails()
        {
            try
            {
                lsvOnlineEntryFiles.Items.Clear();
                string sVoiceFileID = string.Empty;
                sVoiceFileID = txtEntryVoiceFIle.Text;
                DataSet _dsFetchMissingFiles = BusinessLogic.WS_Allocation.Get_Missing_File(sVoiceFileID);

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsFetchMissingFiles.Tables[0].Select())
                    lsvOnlineEntryFiles.Items.Add(new Listitem_MissingFileDetails(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvOnlineEntryFiles);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void LoadListView_InclusiveLines()
        {
            try
            {
                lsvInclusiveLines.Items.Clear();
                lsvInclusiveMinutes.Items.Clear();
                string sDate = Convert.ToDateTime(dtpInDatePicker.Text).ToString("yyyy/MM/dd");
                DataSet _dsFetchInclusiveLines = BusinessLogic.WS_Allocation.Get_Inclusive_Lines(Convert.ToDateTime(sDate));

                if (rbtInclusiveLines.Checked == true)
                {
                    int iRowCount = 0;
                    foreach (DataRow _drRow in _dsFetchInclusiveLines.Tables[1].Select())
                        lsvInclusiveLines.Items.Add(new ListItem_InclusiveLines(_drRow, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lsvInclusiveLines);
                }
                else if (rbtInclusiveMinutes.Checked == true)
                {
                    int iRowCount = 0;
                    foreach (DataRow _drRow in _dsFetchInclusiveLines.Tables[0].Select())
                        lsvInclusiveLines.Items.Add(new ListItem_InclusiveMinutes(_drRow, iRowCount++));

                    BusinessLogic.Reset_ListViewColumn(lsvInclusiveMinutes);
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// LOAD BUFFERING EVENT
        /// </summary>
        private void Load_ListView_Buffering()
        {
            try
            {
                ListViewHelper.EnableDoubleBuffer(lsvFileDetails);
                ListViewHelper.EnableDoubleBuffer(lsvLoginEmoloyees);
                ListViewHelper.EnableDoubleBuffer(lsvAllotedFiles);
                ListViewHelper.EnableDoubleBuffer(lsvProfiles);
                ListViewHelper.EnableDoubleBuffer(lvDesignation);
                ListViewHelper.EnableDoubleBuffer(lsvEmployeeList);
                ListViewHelper.EnableDoubleBuffer(lsvLeaveList);
                ListViewHelper.EnableDoubleBuffer(lsvManual);
                ListViewHelper.EnableDoubleBuffer(lsvExtractInfo);
                ListViewHelper.EnableDoubleBuffer(lvAllocationStatus);
                ListViewHelper.EnableDoubleBuffer(lvEmployeeAchievement);
                ListViewHelper.EnableDoubleBuffer(lsvOnlineEntryFiles);
                ListViewHelper.EnableDoubleBuffer(lvTarget);
                ListViewHelper.EnableDoubleBuffer(lsvLineCount);
                ListViewHelper.EnableDoubleBuffer(lsvHourlyrReports);
                ListViewHelper.EnableDoubleBuffer(lsvTransEdit);
                ListViewHelper.EnableDoubleBuffer(lvAccountWiseInfo);
                ListViewHelper.EnableDoubleBuffer(lsvDiscrepancy);
                ListViewHelper.EnableDoubleBuffer(lsvDiscrepancyReport);
                ListViewHelper.EnableDoubleBuffer(lsv_OfflieFile_Details);
                ListViewHelper.EnableDoubleBuffer(lsvLargeMinutes);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void LoadListView_AccountIncentive()
        {
            try
            {
                lsvAccViewIncent.Items.Clear();
                DataSet _dsInc = new DataSet();
                _dsInc = BusinessLogic.WS_Allocation.Get_Account_Incentive();

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsInc.Tables[0].Select())
                    lsvAccViewIncent.Items.Add(new Listitem_AccountIncentive(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvAccViewIncent);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void LoadHourlyReport()
        {
            try
            {
                lsvLineCount.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_LineCountDeatils oListItem;

                string sProductionID = string.Empty;
                sProductionID = cbxUserName.SelectedValue.ToString();
                DataSet _dsLineCount = new DataSet();

                object oFromDate = null;
                object oTodate = null;
                string iFromHour = "0";
                string iToHour = "0";

                string sUHourlyFromDtae = Convert.ToDateTime(dateTimeUserFrom.Value).ToString("yyyy-MM-dd");
                string sUHourlyToDate = Convert.ToDateTime(dateTimeUserTo.Value).ToString("yyyy-MM-dd");

                if ((Convert.ToInt32(cbxUFromHours.SelectedValue) > 0) || (cbxUFromHours.SelectedItem != null))
                    iFromHour = cbxUFromHours.SelectedItem.ToString();

                if ((Convert.ToInt32(cbxUToHours.SelectedValue) > 0) || (cbxUToHours.SelectedItem != null))
                    iToHour = cbxUToHours.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = sUHourlyFromDtae + " 23:59:59";
                else
                    oFromDate = sUHourlyFromDtae + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = sUHourlyToDate + " 23:59:59";
                else
                    oTodate = sUHourlyToDate + " " + iToHour + ":" + "00:00";

                _dsLineCount = BusinessLogic.WS_Allocation.Get_Transcription_view_VNEW(Convert.ToInt32(sProductionID), -1, -1, oFromDate, oTodate);

                int iRowCount = 0;
                decimal dTotLines = 0;
                int dTotMins = 0;
                foreach (DataRow _drLines in _dsLineCount.Tables[0].Select())
                {
                    lsvLineCount.Items.Add((new ListItem_LineCountDeatils(_drLines, iRowCount++)));

                    if (_drLines["Converted_lines"].ToString() != "")
                    {
                        dTotLines += Convert.ToDecimal(_drLines["Converted_lines"].ToString());
                    }

                    if (_drLines["Converted_Seconds"] != DBNull.Value)
                    {
                        if (Convert.ToInt32(_drLines["Converted_Seconds"]) != 0)
                        {
                            if (_drLines["Converted_Seconds"].ToString().Contains('.'))
                                dTotMins += Convert.ToInt32(_drLines["Converted_Seconds"].ToString().Split('.').GetValue(0));
                            else
                                dTotMins += Convert.ToInt32(_drLines["Converted_Seconds"].ToString());
                        }
                    }
                }
                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_LineCountDeatils("Linecount summary for the date : " + Convert.ToDateTime(oFromDate).ToString("dd-MM-yyyy"), oMins.ToString(), Math.Round(dTotLines, 2).ToString());
                lsvLineCount.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvLineCount);
            }

            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void Load_AllBranch_Lines()
        {
            try
            {
                //COIMBATORE DETAILS
                pictureBox1.BackColor = System.Drawing.ColorTranslator.FromHtml("#DAE3E9");
                lsvCoimbatoreDeatils.Items.Clear();
                lsvCochinDeatils.Items.Clear();
                lsvTrivandrum.Items.Clear();
                lsvPondicherryDetails.Items.Clear();

                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_AccountWiseInfo_BranchWise oListItem;

                string sBranchFromDate = Convert.ToDateTime(dtpLinesFromDate.Value).ToString("yyyy-MM-dd");
                string sBranchTodate = Convert.ToDateTime(dtpLinesToDate.Value).ToString("yyyy-MM-dd");

                DataSet _dsAccountBranchWise = BusinessLogic.WS_Allocation.Get_accountWiseInfoForBranch(sBranchFromDate, sBranchTodate);

                int iRowCount = 0;
                int dTotMins = 0;
                decimal dTotLines = 0;

                foreach (DataRow _drRow in _dsAccountBranchWise.Tables[0].Select())
                {
                    lsvCoimbatoreDeatils.Items.Add(new ListItem_AccountWiseInfo_BranchWise(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }

                    if (_drRow["FileLines"].ToString() != "")
                    {
                        dTotLines += Convert.ToDecimal(_drRow["FileLines"].ToString());
                    }
                    lsvCoimbatoreDeatils.BackColor = System.Drawing.ColorTranslator.FromHtml("#F781F3");

                }

                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_AccountWiseInfo_BranchWise("COIMBATORE ON: " + Convert.ToDateTime(sBranchFromDate).ToString("dd-MM-yyyy"), oMins.ToString(), dTotLines.ToString());
                lsvCoimbatoreDeatils.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvCoimbatoreDeatils);


                //COCHIN DETAILS

                int dTotMinsCOC = 0;
                decimal dTotLinesCOC = 0;

                foreach (DataRow _drRow in _dsAccountBranchWise.Tables[1].Select())
                {
                    lsvCochinDeatils.Items.Add(new ListItem_AccountWiseInfo_BranchWise(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMinsCOC += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMinsCOC += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }

                    if (_drRow["FileLines"].ToString() != "")
                    {
                        dTotLinesCOC += Convert.ToDecimal(_drRow["FileLines"].ToString());
                    }
                    lsvCochinDeatils.BackColor = System.Drawing.ColorTranslator.FromHtml("#FE9A2E");
                }

                string oMinsCOC = sGetDuration(dTotMinsCOC);

                oListItem = new ListItem_AccountWiseInfo_BranchWise("COCHIN ON: " + Convert.ToDateTime(sBranchFromDate).ToString("dd-MM-yyyy"), oMinsCOC.ToString(), dTotLinesCOC.ToString());
                lsvCochinDeatils.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvCochinDeatils);

                //TRIVANDRUM DETAILS

                int dTotMinsTRIV = 0;
                decimal dTotLinesTRIV = 0;

                foreach (DataRow _drRow in _dsAccountBranchWise.Tables[2].Select())
                {
                    lsvTrivandrum.Items.Add(new ListItem_AccountWiseInfo_BranchWise(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMinsTRIV += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMinsTRIV += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                    if (_drRow["FileLines"].ToString() != "")
                    {
                        dTotLinesTRIV += Convert.ToDecimal(_drRow["FileLines"].ToString());
                    }
                    lsvTrivandrum.BackColor = System.Drawing.ColorTranslator.FromHtml("#FA8258");
                }
                string oMinsTRIV = sGetDuration(dTotMinsTRIV);

                oListItem = new ListItem_AccountWiseInfo_BranchWise("TRIVANDRUM ON: " + Convert.ToDateTime(sBranchFromDate).ToString("dd-MM-yyyy"), oMinsTRIV.ToString(), dTotLinesTRIV.ToString());
                lsvTrivandrum.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvTrivandrum);


                // PONDICHERRY DETAILS

                int dTotMinsPONDI = 0;
                decimal dTotLinesPONDI = 0;

                foreach (DataRow _drRow in _dsAccountBranchWise.Tables[3].Select())
                {
                    lsvPondicherryDetails.Items.Add(new ListItem_AccountWiseInfo_BranchWise(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMinsPONDI += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMinsPONDI += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }

                    if (_drRow["FileLines"].ToString() != "")
                    {
                        dTotLinesPONDI += Convert.ToDecimal(_drRow["FileLines"].ToString());
                    }
                    lsvPondicherryDetails.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFCC00");
                }

                string oMinsPONDI = sGetDuration(dTotMinsPONDI);

                oListItem = new ListItem_AccountWiseInfo_BranchWise("PONDICHERRY ON: " + Convert.ToDateTime(sBranchFromDate).ToString("dd-MM-yyyy"), oMinsPONDI.ToString(), dTotLinesPONDI.ToString());
                lsvPondicherryDetails.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsvPondicherryDetails);




                decimal TotalLinesCounts = dTotLines + dTotLinesCOC + dTotLinesTRIV + dTotLinesPONDI;
                int TotalMinutesDone = dTotMins + dTotMinsCOC + dTotMinsTRIV + dTotMinsPONDI;
                string oTotalMins = sGetDuration(TotalMinutesDone);

                lsvAllTotal.Items.Clear();
                ListViewItem lvTotalAll = new ListViewItem();
                lvTotalAll.Text = "Total Lines and Minutes achieved:";
                lvTotalAll.SubItems.Add(oTotalMins);
                lvTotalAll.SubItems.Add(TotalLinesCounts.ToString());
                lsvAllTotal.Items.Add(lvTotalAll);

                lsvAllTotal.BackColor = System.Drawing.Color.SeaGreen;
                lsvAllTotal.ForeColor = System.Drawing.Color.White;
                BusinessLogic.Reset_ListViewColumn(lsvAllTotal);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void Load_All_Employees()
        {
            try
            {

                DataSet _dsAllEmployee = new DataSet();
                _dsAllEmployee = BusinessLogic.WS_Allocation.Get_All_Employees();

                //MT'S
                int iRowCount = 0;
                lsvMT.Items.Clear();
                foreach (DataRow _drRow in _dsAllEmployee.Tables[0].Select())
                    lsvMT.Items.Add(new Listitem_AllEmployees(_drRow, iRowCount++));
                lsvMT.BackColor = System.Drawing.ColorTranslator.FromHtml("#ff4141");
                lsvMT.ForeColor = System.Drawing.Color.White;
                BusinessLogic.Reset_ListViewColumn(lsvMT);

                //TED'S
                lsvTED.Items.Clear();
                foreach (DataRow _drRow in _dsAllEmployee.Tables[1].Select())
                    lsvTED.Items.Add(new Listitem_AllEmployees(_drRow, iRowCount++));
                lsvTED.BackColor = System.Drawing.ColorTranslator.FromHtml("#009100");
                lsvTED.ForeColor = System.Drawing.Color.White;
                BusinessLogic.Reset_ListViewColumn(lsvTED);

                //ME'S
                lsvME.Items.Clear();
                foreach (DataRow _drRow in _dsAllEmployee.Tables[2].Select())
                    lsvME.Items.Add(new Listitem_AllEmployees(_drRow, iRowCount++));
                lsvME.BackColor = System.Drawing.ColorTranslator.FromHtml("#b900ff");
                lsvME.ForeColor = System.Drawing.Color.White;
                BusinessLogic.Reset_ListViewColumn(lsvME);

                //AM'S
                lsvAM.Items.Clear();
                foreach (DataRow _drRow in _dsAllEmployee.Tables[3].Select())
                    lsvAM.Items.Add(new Listitem_AllEmployees(_drRow, iRowCount++));
                lsvAM.BackColor = System.Drawing.ColorTranslator.FromHtml("#c4b46c");
                lsvAM.ForeColor = System.Drawing.Color.Black;
                BusinessLogic.Reset_ListViewColumn(lsvAM);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void Load_Night_Shift_Marking()
        {
            try
            {
                lvNightShiftMarked.Items.Clear();
                DataSet _dsNightShiftAllowance = new DataSet();
                _dsNightShiftAllowance = BusinessLogic.WS_Allocation.Get_NightShift_Incentive_Datas();

                int iRowCount = 0;

                foreach (DataRow _drRow in _dsNightShiftAllowance.Tables[0].Select())
                    lvNightShiftMarked.Items.Add(new Listitem_NightShift_Marked(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvNightShiftMarked);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void Load_Night_Shift_Category()
        {
            try
            {
                lvNightShiftAllowance.Items.Clear();
                DataSet _dsCategory = new DataSet();
                _dsCategory = BusinessLogic.WS_Allocation.Get_NightShift_Category();

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsCategory.Tables[0].Select())
                    lvNightShiftAllowance.Items.Add(new Listitem_Category(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvNightShiftAllowance);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        //private void Load_LineCountReport()
        //{
        //    try
        //    {
        //        BusinessLogic.oMessageEvent.Start("Transferring data.");                
        //        List<BusinessLogic.LINECOUNT_DETAILS> oLinecount = new List<BusinessLogic.LINECOUNT_DETAILS>();

        //        lsvLineCountDetails.Items.Clear();
        //        lsvUserSummary.Items.Clear();
        //        string sProductionID = cmbLogUser.SelectedValue.ToString();
        //        ListItem_LineCountFileItem oListItem;
        //        int iRowCount = 0;
        //        int iTot_Mins = 0;
        //        int iConv_Mins = 0;
        //        decimal dTotal_Lines = 0;
        //        decimal dTotal_Conv_Lines = 0;

        //        DataSet _dsLineCountReport = new DataSet();
        //        if (!chkIsNightShiftUser.Checked)
        //        {
        //            _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2(Convert.ToInt32(sProductionID), Convert.ToDateTime(datDateFrom.Value), Convert.ToDateTime(datDateTo.Value));
        //        }
        //        else
        //        {
        //            string sFromDate = Convert.ToDateTime(datDateFrom.Text).ToString("yyyy/MM/dd 12:00:00");
        //            string sToDate = Convert.ToDateTime(datDateTo.Text).ToString("yyyy/MM/dd 12:00:00");
        //            if (sFromDate == sToDate)
        //            {
        //                _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2_NightShift(Convert.ToInt32(sProductionID), Convert.ToDateTime(sFromDate), Convert.ToDateTime(sFromDate).AddDays(1));
        //            }
        //            else
        //            {
        //                _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2(Convert.ToInt32(sProductionID), Convert.ToDateTime(datDateFrom.Value), Convert.ToDateTime(datDateTo.Value));
        //            }
        //        }
        //        if (_dsLineCountReport == null)
        //        {
        //            BusinessLogic.oMessageEvent.Start("No record found.");
        //            return;
        //        }

        //        if (_dsLineCountReport.Tables.Count <= 0)
        //        {
        //            BusinessLogic.oMessageEvent.Start("No record found.");
        //            return;
        //        }

        //        if (_dsLineCountReport.Tables[0].Rows.Count <= 0)
        //        {
        //            BusinessLogic.oMessageEvent.Start("No record found.");
        //            return;
        //        }

        //        foreach (DataRow _drTobegraded in _dsLineCountReport.Tables[0].Select())
        //        {
        //            oLinecount.Add(new BusinessLogic.LINECOUNT_DETAILS(_drTobegraded["client_name"].ToString(), _drTobegraded["location_name"].ToString(), 
        //                _drTobegraded["doctor_full_name"].ToString(), _drTobegraded["report_name"].ToString(), Convert.ToDateTime(_drTobegraded["file_date"].ToString()),
        //                _drTobegraded["file_minutes"].ToString(), _drTobegraded["Converted_minutes"].ToString(), Convert.ToDecimal(_drTobegraded["file_lines"].ToString()),
        //                Convert.ToDecimal(_drTobegraded["converted_lines"].ToString()), Convert.ToDateTime(_drTobegraded["submitted_time"].ToString()), 
        //                Convert.ToDateTime(_drTobegraded["Submit_Time"].ToString()),
        //                _drTobegraded["evaluated_date"].ToString(), _drTobegraded["transcription_status_description_1"].ToString(), _drTobegraded["transcription_status_description_1"].ToString(),
        //                _drTobegraded["template_description"].ToString(), Convert.ToDecimal(_drTobegraded["accuracy"].ToString()), string.Empty,
        //                Convert.ToInt32(_drTobegraded["CalSec"].ToString()), Convert.ToDecimal(_drTobegraded["Converted_Seconds"].ToString())));
        //        }

        //        for (var day = Convert.ToDateTime(datDateFrom.Value).Date; day.Date <= Convert.ToDateTime(datDateTo.Value).Date; day = day.AddDays(1))
        //        {
        //            DataSet dsNight_shift = BusinessLogic.WS_Allocation.Get_Check_Nightshift(Convert.ToInt32(sProductionID), Convert.ToDateTime(datDateFrom.Value), Convert.ToDateTime(datDateTo.Value));

        //            if (dsNight_shift != null)
        //            {
        //                if (dsNight_shift.Tables[0].Rows.Count > 0)
        //                {
        //                    string sFilter = " Login='" + Convert.ToDateTime(day).ToString("yyyy-MM-dd") + "'";
        //                    DataRow[] sFound = dsNight_shift.Tables[0].Select(sFilter);
        //                    int iCount = 0;
        //                    foreach (DataRow dr in sFound)
        //                        iCount = 1;

        //                    if (iCount == 0)        // For Day Shift
        //                    {
        //                        string sDay_Filter = " Submit_Time='" + Convert.ToDateTime(day).ToString("yyyy-MM-dd") + "' ";
        //                        DataRow[] drView = _dsLineCountReport.Tables[0].Select(sDay_Filter);
        //                        foreach (DataRow dr in drView)
        //                        {
        //                            oListItem = new ListItem_LineCountFileItem(dr, iRowCount++);
        //                            lsvLineCountDetails.Items.Add(oListItem);
        //                        }
        //                        var Tot_files = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) select c).Count();
        //                        var Tot_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.FILE_SEC into CP select new { TOTAL_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.FILE_SEC)) });
        //                        var Tot_Conv_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.CONV_SEC into CP select new { TOTAL_CONV_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.CONV_SEC)) });

        //                        var Total_File_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.FILELINES into CP select new { TOTAL_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.FILELINES)) });
        //                        var Total_File_Conv_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.CONVLINES into CP select new { TOTAL_CONV_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.CONVLINES)) });

        //                        foreach (var File_Mins in Tot_Mins)
        //                            iTot_Mins += Convert.ToInt32(File_Mins.TOTAL_FILE_SEC);

        //                        foreach (var File_Conv_Mins in Tot_Conv_Mins)
        //                            iConv_Mins += Convert.ToInt32(File_Conv_Mins.TOTAL_CONV_FILE_SEC);

        //                        foreach (var File_Lines in Total_File_Lines)
        //                            dTotal_Lines += Convert.ToInt32(File_Lines.TOTAL_FILE_LINES);

        //                        foreach (var File_Conv_Lines in Total_File_Conv_Lines)
        //                            dTotal_Conv_Lines += Convert.ToInt32(File_Conv_Lines.TOTAL_CONV_FILE_LINES);

        //                        //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), oMins.ToString(), oConMins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());
        //                        if (Tot_files > 0)
        //                        {
        //                            oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Lines, 2).ToString(), string.Empty, string.Empty);
        //                            lsvLineCountDetails.Items.Add(oListItem);
        //                        }   
        //                        //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Lines, 2).ToString(), string.Empty, string.Empty);
        //                        //lsvLineCountDetails.Items.Add(oListItem);
        //                        iCount = 0;
        //                        iTot_Mins = 0;
        //                        iConv_Mins = 0;
        //                        dTotal_Lines = 0;
        //                        dTotal_Conv_Lines = 0;

        //                    }
        //                    else                   // For Night Shift                            
        //                    {
        //                        DateTime end = DateTime.Now;
        //                        DateTime start = DateTime.Now;

        //                        string sFilter_Night = " Login='" + Convert.ToDateTime(day).ToString("yyyy-MM-dd") + "'";
        //                        DataRow[] sNight = dsNight_shift.Tables[0].Select(sFilter_Night);

        //                        foreach (DataRow dr_Night in sNight)
        //                        {
        //                            start = Convert.ToDateTime(dr_Night["login_time"].ToString());
        //                            end = Convert.ToDateTime(dr_Night["logoff_time"].ToString());
        //                        }

        //                        string sDay_Filter = " submitted_time>='" + Convert.ToDateTime(start) + "' and submitted_time<=  '" + Convert.ToDateTime(end) + "' ";
        //                        DataRow[] drView = _dsLineCountReport.Tables[0].Select(sDay_Filter);
        //                        foreach (DataRow dr in drView)
        //                        {
        //                            oListItem = new ListItem_LineCountFileItem(dr, iRowCount++);
        //                            lsvLineCountDetails.Items.Add(oListItem);
        //                        }
        //                        var Tot_files = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(start) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(end)) select c).Count();
        //                        var Tot_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(start) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(end)) group c by c.FILE_SEC into CP select new { TOTAL_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.FILE_SEC)) });
        //                        var Tot_Conv_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(start) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(end)) group c by c.CONV_SEC into CP select new { TOTAL_CONV_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.CONV_SEC)) });

        //                        var Total_File_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(start) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(end)) group c by c.FILELINES into CP select new { TOTAL_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.FILELINES)) });
        //                        var Total_File_Conv_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(start) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(end)) group c by c.CONVLINES into CP select new { TOTAL_CONV_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.CONVLINES)) });

        //                        foreach (var File_Mins in Tot_Mins)
        //                            iTot_Mins += Convert.ToInt32(File_Mins.TOTAL_FILE_SEC);

        //                        foreach (var File_Conv_Mins in Tot_Conv_Mins)
        //                            iConv_Mins += Convert.ToInt32(File_Conv_Mins.TOTAL_CONV_FILE_SEC);

        //                        foreach (var File_Lines in Total_File_Lines)
        //                            dTotal_Lines += Convert.ToInt32(File_Lines.TOTAL_FILE_LINES);

        //                        foreach (var File_Conv_Lines in Total_File_Conv_Lines)
        //                            dTotal_Conv_Lines += Convert.ToInt32(File_Conv_Lines.TOTAL_CONV_FILE_LINES);

        //                        //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), oMins.ToString(), oConMins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());

        //                        if (Tot_files > 0)
        //                        {
        //                            oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(start).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Lines, 2).ToString(), string.Empty, string.Empty);
        //                            lsvLineCountDetails.Items.Add(oListItem);
        //                        }                                
        //                        iCount = 0;
        //                        iTot_Mins = 0;
        //                        iConv_Mins = 0;
        //                        dTotal_Lines = 0;
        //                        dTotal_Conv_Lines = 0;
        //                    }
        //                }
        //                else
        //                {
        //                    string sDay_Filter = " Submit_Time='" + Convert.ToDateTime(day).ToString("yyyy-MM-dd") + "' ";
        //                    DataRow[] drView = _dsLineCountReport.Tables[0].Select(sDay_Filter);
        //                    foreach (DataRow dr in drView)
        //                    {
        //                        oListItem = new ListItem_LineCountFileItem(dr, iRowCount++);
        //                        lsvLineCountDetails.Items.Add(oListItem);
        //                    }
        //                    var Tot_files = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) select c).Count();
        //                    var Tot_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.FILE_SEC into CP select new { TOTAL_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.FILE_SEC)) });
        //                    var Tot_Conv_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.CONV_SEC into CP select new { TOTAL_CONV_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.CONV_SEC)) });

        //                    var Total_File_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.FILELINES into CP select new { TOTAL_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.FILELINES)) });
        //                    var Total_File_Conv_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBTIME) == Convert.ToDateTime(day)) group c by c.CONVLINES into CP select new { TOTAL_CONV_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.CONVLINES)) });

        //                    foreach (var File_Mins in Tot_Mins)
        //                        iTot_Mins += Convert.ToInt32(File_Mins.TOTAL_FILE_SEC);

        //                    foreach (var File_Conv_Mins in Tot_Conv_Mins)
        //                        iConv_Mins += Convert.ToInt32(File_Conv_Mins.TOTAL_CONV_FILE_SEC);

        //                    foreach (var File_Lines in Total_File_Lines)
        //                        dTotal_Lines += Convert.ToInt32(File_Lines.TOTAL_FILE_LINES);

        //                    foreach (var File_Conv_Lines in Total_File_Conv_Lines)
        //                        dTotal_Conv_Lines += Convert.ToInt32(File_Conv_Lines.TOTAL_CONV_FILE_LINES);

        //                    //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), oMins.ToString(), oConMins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());
        //                    if (Tot_files > 0)
        //                    {
        //                        oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Lines, 2).ToString(), string.Empty, string.Empty);
        //                        lsvLineCountDetails.Items.Add(oListItem);                            
        //                    }                            
        //                    iTot_Mins = 0;
        //                    iConv_Mins = 0;
        //                    dTotal_Lines = 0;
        //                    dTotal_Conv_Lines = 0;
        //                }

        //                BusinessLogic.Reset_ListViewColumn(lsvLineCountDetails);
        //            }
        //        }

        //        //FOR CUMULATIVE 
        //        DataTable dtCumulative = new DataTable();
        //        var Firstday = System.DateTime.Now;
        //        var Lastday = System.DateTime.Now;
        //        if (!chkIsNightShiftUser.Checked)
        //        {
        //            Firstday = datDateFrom.Value;
        //            Lastday = datDateTo.Value;
        //            dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday));
        //        }
        //        else
        //        {
        //            string sFromDate = string.Empty;
        //            string sToDate = string.Empty;
        //            sFromDate = Convert.ToDateTime(datDateFrom.Text).ToString("yyyy/MM/dd 12:00:00");
        //            Firstday = Convert.ToDateTime(sFromDate);
        //            sToDate = Convert.ToDateTime(datDateTo.Text).ToString("yyyy/MM/dd 12:00:00");
        //            Lastday = Convert.ToDateTime(sToDate);
        //            if (Firstday == Lastday)
        //            {
        //                dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New_Night_Shift(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday).AddDays(1));
        //            }
        //            else
        //            {
        //                Firstday = datDateFrom.Value;
        //                Lastday = datDateTo.Value;
        //                dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday));
        //            }
        //        }

        //        decimal oCumNightLines;
        //        decimal dNighAllowance = 0;
        //        decimal dAccount = 0;

        //        ListSummery oListSum;
        //        if (dtCumulative.Rows.Count > 0)
        //        {
        //            string CumulatveMins;

        //            CumulatveMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(dtCumulative.Rows[0]["Minutes"].ToString())));
        //            decimal dAcc = 0;
        //            decimal dHoldPer = Convert.ToDecimal(dtCumulative.Rows[0]["HoldPercentage"].ToString());
        //            oCumNightLines = Convert.ToDecimal(dtCumulative.Rows[0]["NightShift_Linecountss"].ToString());
        //            dNighAllowance = Convert.ToDecimal(dtCumulative.Rows[0]["Nightshift_Allowance"].ToString());
        //            dAcc = Convert.ToDecimal(dtCumulative.Rows[0]["Accuracy"].ToString());
        //            dAccount = Convert.ToDecimal(dtCumulative.Rows[0]["Account_Incentive"].ToString());

        //            decimal dApproxSal = 0;
        //            decimal dPunctuality = 0;
        //            decimal dLinecountSal = 0;


        //            DataRow drRow = ((DataRowView)cmbLogUser.SelectedItem).Row;
        //            string sEmployeeID = drRow["employee_id"].ToString();

        //            DataSet dsApproxSal = BusinessLogic.WS_Allocation.Get_ApproxSal_All(sEmployeeID, Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday), 0, dAcc, Convert.ToDecimal(dtCumulative.Rows[0]["Total_ConvertedLines"].ToString()), oCumNightLines);

        //            if (dsApproxSal != null)
        //            {
        //                if (dsApproxSal.Tables[0].Rows.Count > 0)
        //                {
        //                    dApproxSal = Convert.ToDecimal(dsApproxSal.Tables[0].Rows[0]["Earnings"].ToString());
        //                }

        //                if (dsApproxSal.Tables[1].Rows.Count > 0)
        //                {
        //                    dPunctuality = Convert.ToDecimal(dsApproxSal.Tables[1].Rows[0]["PunctualityIncentive"].ToString());
        //                }

        //                if (dsApproxSal.Tables[2].Rows.Count > 0)
        //                {
        //                    dLinecountSal = Convert.ToDecimal(dsApproxSal.Tables[2].Rows[0]["_linecountsalary"].ToString());
        //                }
        //            }

        //            //oListSum = new ListSummery(string.Empty, oCumTotfiles.ToString(), CumulatveMins.ToString(), oCumLinecount.ToString(), oCumConLines.ToString(), oCumNightLines.ToString(), dHoldPer.ToString(), dAcc.ToString(), dIncentive.ToString(), Math.Round(dLinecountSal, 2).ToString(), dNighAllowance.ToString(), dPunctuality.ToString(),  Math.Round(dApproxSal, 2).ToString());
        //            oListSum = new ListSummery(string.Empty, dtCumulative.Rows[0]["Totfiles"].ToString(), CumulatveMins.ToString(), dtCumulative.Rows[0]["Linecount"].ToString(), dtCumulative.Rows[0]["Final_Linecount"].ToString(), dtCumulative.Rows[0]["NightShift_Linecountss"].ToString(), dtCumulative.Rows[0]["Sunday_Shift_Lines"].ToString(), dtCumulative.Rows[0]["Extra_Support_Lines"].ToString(), dtCumulative.Rows[0]["HoldPercentage"].ToString(), dtCumulative.Rows[0]["Accuracy"].ToString(), dtCumulative.Rows[0]["Incentive_Lines"].ToString(), dtCumulative.Rows[0]["Sunday_Shift_Allowance"].ToString(), dtCumulative.Rows[0]["Extra_Support_Allowance"].ToString(), Math.Round(dLinecountSal, 2).ToString(), dNighAllowance.ToString(), dPunctuality.ToString(), dtCumulative.Rows[0]["Total_ConvertedLines"].ToString(), dAccount.ToString(), Math.Round(dApproxSal, 2).ToString());
        //            lsvUserSummary.Items.Add(oListSum);
        //        }

        //        BusinessLogic.Reset_ListViewColumn(lsvLineCountDetails);
        //        BusinessLogic.Reset_ListViewColumn(lsvUserSummary);
        //        BusinessLogic.oMessageEvent.Start("Ready");   
        //    }
        //    catch (Exception ex)
        //    {
        //        BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
        //        BusinessLogic.oMessageEvent.Start("This is a night shift user. He has not properly logged off!...");
        //    }
        //    finally
        //    {
        //        BusinessLogic.oProgressEvent.Start(false);
        //        this.Cursor = Cursors.Default;
        //        lsvLineCountDetails.EndUpdate();
        //        Application.DoEvents();
        //    }
        //}

        private void Load_LineCountReport()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring data.");

                lsvLineCountDetails.Items.Clear();
                lsvUserSummary.Items.Clear();

                string sProductionID = cmbLogUser.SelectedValue.ToString();
                DataSet _dsLineCountReport = new DataSet();
                if (!chkIsNightShiftUser.Checked)
                {
                    BusinessLogic.WS_Allocation.Timeout = 500000;
                    _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2(Convert.ToInt32(sProductionID), Convert.ToDateTime(datDateFrom.Value), Convert.ToDateTime(datDateTo.Value));
                }
                else
                {
                    string sFromDate = Convert.ToDateTime(datDateFrom.Text).ToString("yyyy/MM/dd 12:00:00");
                    string sToDate = Convert.ToDateTime(datDateTo.Text).ToString("yyyy/MM/dd 12:00:00");
                    if (sFromDate == sToDate)
                    {
                        BusinessLogic.WS_Allocation.Timeout = 500000;
                        _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2_NightShift(Convert.ToInt32(sProductionID), Convert.ToDateTime(sFromDate), Convert.ToDateTime(sFromDate).AddDays(1));
                    }
                    else
                    {
                        BusinessLogic.WS_Allocation.Timeout = 500000;
                        _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2(Convert.ToInt32(sProductionID), Convert.ToDateTime(datDateFrom.Value), Convert.ToDateTime(datDateTo.Value));
                    }
                }

                if (_dsLineCountReport == null)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                if (_dsLineCountReport.Tables.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                if (_dsLineCountReport.Tables[0].Rows.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                lsvLineCountDetails.BeginUpdate();
                ListItem_LineCountFileItem oListItem;
                int iRowCount = 0;
                int iNumberFilesForDay = 0;
                object CurrentDate = null;
                object PreviousDate = null;
                object oMinutes = 0;
                object dblConvertedMinutes = 0;
                double dblLines = 0;
                double dblConvertedLines = 0;
                double dblConvertedGradedLines = 0;
                double dblErrorPoints = 0;
                double dblAccuracy = 0;

                string oMins;
                string oConMins;
                int isNightShift = 0;

                foreach (DataRow _drTobegraded in _dsLineCountReport.Tables[0].Select())
                {
                    oMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(oMinutes)));
                    oConMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(dblConvertedMinutes)));
                    isNightShift = Convert.ToInt32(_drTobegraded["isNightShift"].ToString());

                    CurrentDate = BusinessLogic.ConvertToDateTime(_drTobegraded["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_SUBMITTED_TIME + ""]);
                    //int iDaydifferece = (int)((TimeSpan)Convert.ToDateTime(CurrentDate).Subtract(Convert.ToDateTime(PreviousDate))).Days;
                    int iDaydifferece = Convert.ToDateTime(CurrentDate).Day - Convert.ToDateTime(PreviousDate).Day;
                    if (PreviousDate == null)
                    {
                        //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(_drTobegraded["" + Data.TBL_TRANSCRIPTION_TRANSACTION.SUBMITTED_TIME + ""]).ToString("dd-MM-yyyy"));
                        //lsvLineCountDetails.Items.Add(oListItem);
                        oListItem = new ListItem_LineCountFileItem(_drTobegraded, iRowCount++);
                        lsvLineCountDetails.Items.Add(oListItem);
                        PreviousDate = CurrentDate;
                    }

                    //else if (DateTime.Compare(Convert.ToDateTime(CurrentDate), Convert.ToDateTime(PreviousDate))==0)
                    else if (iDaydifferece == 0)
                    {
                        oListItem = new ListItem_LineCountFileItem(_drTobegraded, iRowCount++);
                        lsvLineCountDetails.Items.Add(oListItem);
                        PreviousDate = CurrentDate;
                    }
                    else
                    {
                        dblAccuracy = 0;
                        if (dblConvertedGradedLines > 0)
                            dblAccuracy = Math.Round(100 - ((dblErrorPoints / dblConvertedGradedLines) * 100), 2);
                        else if (dblConvertedGradedLines == 0)
                            dblAccuracy = 0;

                        oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(PreviousDate).ToString("dd-MM-yyyy") + " Total Files : " + iNumberFilesForDay.ToString(), oMins.ToString(), oConMins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());
                        lsvLineCountDetails.Items.Add(oListItem);
                        oListItem = new ListItem_LineCountFileItem(_drTobegraded, iRowCount++);
                        lsvLineCountDetails.Items.Add(oListItem);
                        PreviousDate = CurrentDate;
                        oMinutes = 0;
                        dblConvertedMinutes = 0;
                        dblLines = 0;
                        iNumberFilesForDay = 0;
                        dblConvertedLines = 0;
                        dblErrorPoints = 0;
                        dblConvertedGradedLines = 0;
                    }


                    oMinutes = BusinessLogic.AddMinutes(oMinutes.ToString(), _drTobegraded["CalSec"].ToString());
                    dblConvertedMinutes = BusinessLogic.AddMinutes(dblConvertedMinutes.ToString(), _drTobegraded["Converted_Seconds"].ToString());
                    dblLines = dblLines + Convert.ToDouble(_drTobegraded["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_FILE_LINES_DECIMAL + ""]);
                    dblConvertedLines = dblConvertedLines + Convert.ToDouble(_drTobegraded["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_CONVERTED_LINES_DECIMAL + ""]);
                    if (_drTobegraded["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_IS_GRADED_BINT + ""].ToString().ToLower().Equals("yes"))
                        dblConvertedGradedLines = dblConvertedGradedLines + Convert.ToDouble(_drTobegraded["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_CONVERTED_LINES_DECIMAL + ""]);

                    dblErrorPoints = 0;
                    iNumberFilesForDay++;
                    oListItem.EnsureVisible();
                }

                oMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(oMinutes)));
                oConMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(dblConvertedMinutes)));

                dblAccuracy = 0;
                if (dblConvertedGradedLines > 0)
                    dblAccuracy = Math.Round(100 - ((dblErrorPoints / dblConvertedGradedLines) * 100), 2);

                oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(CurrentDate).ToString("dd-MM-yyyy") + " Total Files : " + iNumberFilesForDay.ToString(), oMins.ToString(), oConMins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());
                lsvLineCountDetails.Items.Add(oListItem);
                DataTable dtCumulative = new DataTable();
                var Firstday = System.DateTime.Now;
                var Lastday = System.DateTime.Now;
                if (!chkIsNightShiftUser.Checked)
                {
                    Firstday = datDateFrom.Value;
                    Lastday = datDateTo.Value;
                    dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday));
                }
                else
                {
                    string sFromDate = string.Empty;
                    string sToDate = string.Empty;
                    sFromDate = Convert.ToDateTime(datDateFrom.Text).ToString("yyyy/MM/dd 12:00:00");
                    Firstday = Convert.ToDateTime(sFromDate);
                    sToDate = Convert.ToDateTime(datDateTo.Text).ToString("yyyy/MM/dd 12:00:00");
                    Lastday = Convert.ToDateTime(sToDate);
                    if (Firstday == Lastday)
                    {
                        dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New_Night_Shift(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday).AddDays(1));
                    }
                    else
                    {
                        Firstday = datDateFrom.Value;
                        Lastday = datDateTo.Value;
                        dtCumulative = BusinessLogic.WS_Allocation.Get_Cumulative_Linecount_New(Convert.ToInt32(sProductionID), Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday));
                    }
                }

                decimal oCumNightLines;
                decimal dNighAllowance = 0;
                decimal dAccount = 0;

                ListSummery oListSum;
                if (dtCumulative.Rows.Count > 0)
                {
                    string CumulatveMins;

                    CumulatveMins = sGetDuration(Convert.ToInt32(Convert.ToDouble(dtCumulative.Rows[0]["Minutes"].ToString())));
                    decimal dAcc = 0;
                    decimal dHoldPer = Convert.ToDecimal(dtCumulative.Rows[0]["HoldPercentage"].ToString());
                    oCumNightLines = Convert.ToDecimal(dtCumulative.Rows[0]["NightShift_Linecountss"].ToString());
                    dNighAllowance = Convert.ToDecimal(dtCumulative.Rows[0]["Nightshift_Allowance"].ToString());
                    dAcc = Convert.ToDecimal(dtCumulative.Rows[0]["Accuracy"].ToString());
                    dAccount = Convert.ToDecimal(dtCumulative.Rows[0]["Account_Incentive"].ToString());

                    decimal dApproxSal = 0;
                    decimal dPunctuality = 0;
                    decimal dLinecountSal = 0;


                    DataRow drRow = ((DataRowView)cmbLogUser.SelectedItem).Row;
                    string sEmployeeID = drRow["employee_id"].ToString();

                    DataSet dsApproxSal = BusinessLogic.WS_Allocation.Get_ApproxSal_All(sEmployeeID, Convert.ToDateTime(Firstday), Convert.ToDateTime(Lastday), 0, dAcc, Convert.ToDecimal(dtCumulative.Rows[0]["Total_ConvertedLines"].ToString()), oCumNightLines);

                    if (dsApproxSal != null)
                    {
                        if (dsApproxSal.Tables[0].Rows.Count > 0)
                        {
                            dApproxSal = Convert.ToDecimal(dsApproxSal.Tables[0].Rows[0]["Earnings"].ToString());
                        }

                        if (dsApproxSal.Tables[1].Rows.Count > 0)
                        {
                            dPunctuality = Convert.ToDecimal(dsApproxSal.Tables[1].Rows[0]["PunctualityIncentive"].ToString());
                        }

                        if (dsApproxSal.Tables[2].Rows.Count > 0)
                        {
                            dLinecountSal = Convert.ToDecimal(dsApproxSal.Tables[2].Rows[0]["_linecountsalary"].ToString());
                        }
                    }

                    //oListSum = new ListSummery(string.Empty, oCumTotfiles.ToString(), CumulatveMins.ToString(), oCumLinecount.ToString(), oCumConLines.ToString(), oCumNightLines.ToString(), dHoldPer.ToString(), dAcc.ToString(), dIncentive.ToString(), Math.Round(dLinecountSal, 2).ToString(), dNighAllowance.ToString(), dPunctuality.ToString(),  Math.Round(dApproxSal, 2).ToString());
                    oListSum = new ListSummery(string.Empty, dtCumulative.Rows[0]["Totfiles"].ToString(), CumulatveMins.ToString(), dtCumulative.Rows[0]["Linecount"].ToString(), dtCumulative.Rows[0]["Final_Linecount"].ToString(), dtCumulative.Rows[0]["NightShift_Linecountss"].ToString(), dtCumulative.Rows[0]["Sunday_Shift_Lines"].ToString(), dtCumulative.Rows[0]["Extra_Support_Lines"].ToString(), dtCumulative.Rows[0]["HoldPercentage"].ToString(), dtCumulative.Rows[0]["Accuracy"].ToString(), dtCumulative.Rows[0]["Incentive_Lines"].ToString(), dtCumulative.Rows[0]["Sunday_Shift_Allowance"].ToString(), dtCumulative.Rows[0]["Extra_Support_Allowance"].ToString(), Math.Round(dLinecountSal, 2).ToString(), dNighAllowance.ToString(), dPunctuality.ToString(), dtCumulative.Rows[0]["Total_ConvertedLines"].ToString(), dAccount.ToString(), Math.Round(dApproxSal, 2).ToString());
                    lsvUserSummary.Items.Add(oListSum);
                }

                BusinessLogic.Reset_ListViewColumn(lsvLineCountDetails);
                BusinessLogic.Reset_ListViewColumn(lsvUserSummary);
                BusinessLogic.oMessageEvent.Start("Ready");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                this.Cursor = Cursors.Default;
                lsvLineCountDetails.EndUpdate();
                Application.DoEvents();
            }
        }

        private void Load_Offline_Volume()
        {
            try
            {
                int iOption = 0;

                object oFromDate = null;
                object oTodate = null;
                string iFromHour = "0";
                string iToHour = "0";

                string sUHourlyFromDtae = Convert.ToDateTime(dtpFDate.Value).ToString("yyyy-MM-dd");
                string sUHourlyToDate = Convert.ToDateTime(dtpTDate.Value).ToString("yyyy-MM-dd");

                if ((Convert.ToInt32(cmb_OfflineVolume_Fromhr.SelectedValue) > 0) || (cmb_OfflineVolume_Fromhr.SelectedItem != null))
                    iFromHour = cmb_OfflineVolume_Fromhr.SelectedItem.ToString();

                if ((Convert.ToInt32(cmb_OfflineVolume_ToHr.SelectedValue) > 0) || (cmb_OfflineVolume_ToHr.SelectedItem != null))
                    iToHour = cmb_OfflineVolume_ToHr.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = sUHourlyFromDtae + " 23:59:59";
                else
                    oFromDate = sUHourlyFromDtae + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = sUHourlyToDate + " 23:59:59";
                else
                    oTodate = sUHourlyToDate + " " + iToHour + ":" + "00:00";

                //_dsVolume = BusinessLogic.WS_Allocation.Get_OfflineFile_Details(Convert.ToDateTime(dtpFDate.Value).ToString("yyyy/MM/dd"), Convert.ToDateTime(dtpTDate.Value).ToString("yyyy/MM/dd"));
                _dsVolume = BusinessLogic.WS_Allocation.Get_OfflineFile_Details(oFromDate, oTodate);

                int iRowcount = 0;
                int dTotSecs = 0;
                int dTotFiles = 0;
                lsvOffAccounts.Items.Clear();
                lsv_OfflieFile_Details.Items.Clear();
                lsvPendingfiles.Items.Clear();
                ListItem_OfflineAccountVolume oListItem = null;
                foreach (DataRow dr in _dsVolume.Tables[0].Rows)
                {
                    lsvOffAccounts.Items.Add(new ListItem_OfflineAccountVolume(dr, iRowcount++));
                    dTotFiles = dTotFiles + Convert.ToInt32(dr["Tot_downloaded_file"].ToString());
                    dTotSecs = dTotSecs + Convert.ToInt32(dr["Tot_sec"].ToString());
                }
                string oMins = sGetDuration(dTotSecs);

                oListItem = new ListItem_OfflineAccountVolume(dTotFiles.ToString(), oMins.ToString());
                lsvOffAccounts.Items.Add(oListItem);

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lvFileAllotedStatus);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
                BusinessLogic.Reset_ListViewColumn(lsvOffAccounts);
            }
        }

        private void Load_Offline_Volume_Location(int client_id)
        {
            try
            {
                int iRowcount = 0;

                lsvAccountVolume.Items.Clear();
                foreach (DataRow dr in _dsVolume.Tables[1].Select("client_id=" + client_id))
                {
                    lsvAccountVolume.Items.Add(new ListItem_OfflineLocationVolume(dr, iRowcount++));
                }
                BusinessLogic.Reset_ListViewColumn(lsvAccountVolume);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Offline_Volume_Doctor(string Location_id)
        {
            try
            {

                int iRowcount = 0;

                lsvDoctroWise.Items.Clear();
                foreach (DataRow dr in _dsVolume.Tables[2].Select("location_id='" + Location_id + "'"))
                {
                    lsvDoctroWise.Items.Add(new ListItem_OfflineDoctorVolume(dr, iRowcount++));
                }
                BusinessLogic.Reset_ListViewColumn(lsvDoctroWise);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Offline_Pending_File(int Doctor_id)
        {
            try
            {

                int iRowcount = 0;

                lsvPendingfiles.Items.Clear();
                foreach (DataRow dr in _dsVolume.Tables[3].Select("doctor_id='" + Doctor_id + "'"))
                {
                    lsvPendingfiles.Items.Add(new ListItem_OfflinePendingFiles(dr, iRowcount++));
                }
                BusinessLogic.Reset_ListViewColumn(lsvPendingfiles);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void LoadTopTwentuHold()
        {
            try
            {
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        #endregion "METHODS"

        #region " MENUS "

        private void mtTrans_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 1, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.P_FILEMINS.ToString());
                        //insert into database             
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxTrans_CheckedChanged(this, e);
            }
        }

        private void baTrans_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 5, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxTrans_CheckedChanged(this, e);
            }
        }

        private void meTrans_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxTrans_CheckedChanged(this, e);
            }
        }

        private void tedTrans_Click(object sender, EventArgs e)
        {
            try
            {
                int iResult = 0;
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 3, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        if (BusinessLogic.MTMET_BATCH_ID == 1)
                            iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_TedAssign_V2(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, BusinessLogic.IS_TED_ASSIGN, Environment.UserName, Environment.MachineName);
                        else
                            iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_TedAssign_V2(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, BusinessLogic.IS_TED_ASSIGN, Environment.UserName, Environment.MachineName);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxTrans_CheckedChanged(this, e);
            }
        }

        private void mtAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), -1, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;


                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void baAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 5, 1);
                ofe.ShowDialog();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void tedAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 4, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void meAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void amAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), -1, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        private void meEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxEditing_CheckedChanged(this, e);
            }
        }

        private void amEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 4, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxEditing_CheckedChanged(this, e);
            }
        }

        private void amReview_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 4, 3);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        //insert into database

                        string sVoiceFileID = oFile.P_VOICE_FILE_NAME.ToString();
                        int sTransID;
                        if (cbxReview.Checked == true)
                        {
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                            sTransID = Convert.ToInt32(oFile.TRANSCRIPTIONID);

                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(sTransID, oFile.P_USERID, DateTime.Now, sVoiceFileID, BusinessLogic.iTATREQUIRED);

                            oFile.P_STATUS = "Allotted";
                            oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        }
                        else
                        {
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);

                            oFile.P_STATUS = "Allotted";
                            oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;
                        }
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxReview_CheckedChanged(this, e);
            }
        }

        private void meTransEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void tedTransEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 3, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void meEditReview_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 3, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), "", Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void amEditReview_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 3, 3);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        oFile.P_STATUS = "Allotted";
                        oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        #endregion

        #region " EVENTS "

        private void btn_Transonly_Online_Click(object sender, EventArgs e)
        {
            try
            {
                Get_TransOnly_Online();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        /// <summary>
        /// SIGN OUT BUTTON CLICK EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnsignout_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to sign out?", "RNDSOFT", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                this.Close();
            }
            else
            {
                return;
            }
        }

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

        /// <summary>
        /// FORM LOAD
        /// </summary>
        public int iIndex = 0;
        private void frmAllocation_Load(object sender, EventArgs e)
        {
            try
            {
                tabOffline.TabPages["tabOffPriorityMap"].Enabled = false;
                tabOffline.TabPages["tabOffPriorityMap"].HidePage();                               

                panel28.BackColor = System.Drawing.ColorTranslator.FromHtml("#DAE3E9");
                panel42.BackColor = System.Drawing.ColorTranslator.FromHtml("#DAE3E9");
                panel61.BackColor = System.Drawing.ColorTranslator.FromHtml("#DAE3E9");
                Load_ListView_Buffering();
                Control.CheckForIllegalCrossThreadCalls = false;
                lblWelcome.Text = "Welcome " + BusinessLogic.USERNAME.ToUpper();
                if ((BusinessLogic.USERNAME == "Admin-Coimbatore") || (BusinessLogic.USERNAME == "Admin-Trivandrum") || (BusinessLogic.USERNAME == "Admin-Trivandrum"))
                    btnChangePassword.Enabled = false;

                Load_TabPage(iIndex);
                dtpFromDate.Enabled = false;
                dtpTO.Enabled = false;
                cbxTrans.Checked = true;
                chkInclusiveMinutes_CheckedChanged(this, e);
                Load_Location_List();
                Load_HundredPercent_Doctors_List();
                Load_Batchs();
                Load_Work_Platform();
                Load_Emp_Consolidation();
                Load_MTFiles_List();
                Load_MEFiles_List();
                Load_Batch_Employee();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {               
                if (BusinessLogic.USERNAME == "Admin-Coimbatore")
                {
                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = true;
                    TabExtraction.TabPages["TabEmployee"].Enabled = false;
                    TabExtraction.TabPages["TabEmployee"].HidePage();
                }
                else if (BusinessLogic.IDESIG_ID == 281)
                {
                    #region "TEAM LEADS"
                    
                    tabAllocation.TabPages["tabFileDetails"].Enabled = false;
                    tabAllocation.TabPages["tabFileDetails"].HidePage();

                    tabAllocation.TabPages["tabDeAllocation"].Enabled = false;
                    tabAllocation.TabPages["tabDeAllocation"].HidePage();

                    tabAllocation.TabPages["tabAttendance"].Enabled = false;
                    tabAllocation.TabPages["tabAttendance"].HidePage();                    

                    tabAllocation.TabPages["tbCapacity"].Enabled = false;
                    tabAllocation.TabPages["tbCapacity"].HidePage();

                    tabAllocation.TabPages["tbNDSPDetails"].Enabled = false;
                    tabAllocation.TabPages["tbNDSPDetails"].HidePage();

                    tabAllocation.TabPages["tbNightshiftusers"].Enabled = false;
                    tabAllocation.TabPages["tbNightshiftusers"].HidePage();

                    tbpHabgUpProcess.TabPages["tabPage15"].Enabled = false;
                    tbpHabgUpProcess.TabPages["tabPage15"].HidePage();

                    tabAllocation.TabPages["tabDiscrepancy"].Enabled = false;
                    tabAllocation.TabPages["tabDiscrepancy"].HidePage();

                    tabAllocation.TabPages["tbIdleTime"].Enabled = false;
                    tabAllocation.TabPages["tbIdleTime"].HidePage();

                    tabOffline.TabPages["tabOffFileDeatils"].Enabled = false;
                    tabOffline.TabPages["tabOffFileDeatils"].HidePage();

                    tabOffline.TabPages["tablOffDeallot"].Enabled = false;
                    tabOffline.TabPages["tablOffDeallot"].HidePage();

                    tabOffline.TabPages["tbMTReallocation"].Enabled = false;
                    tabOffline.TabPages["tbMTReallocation"].HidePage();                                        

                    tabOffline.TabPages["tbAllocationPriority"].Enabled = false;
                    tabOffline.TabPages["tbAllocationPriority"].HidePage();

                    TabExtraction.TabPages["tabAccountIncentive"].Enabled = false;
                    TabExtraction.TabPages["tabAccountIncentive"].HidePage();

                    TabExtraction.TabPages["TabEmployee"].Enabled = false;
                    TabExtraction.TabPages["TabEmployee"].HidePage();

                    TabExtraction.TabPages["tabPageManual"].Enabled = false;
                    TabExtraction.TabPages["tabPageManual"].HidePage();

                    TabExtraction.TabPages["tabPageExtraction"].Enabled = false;
                    TabExtraction.TabPages["tabPageExtraction"].HidePage();

                    TabExtraction.TabPages["tabOnlineEntry"].Enabled = false;
                    TabExtraction.TabPages["tabOnlineEntry"].HidePage();

                    TabExtraction.TabPages["tabEmplList"].Enabled = false;
                    TabExtraction.TabPages["tabEmplList"].HidePage();

                    TabExtraction.TabPages["tabTarget"].Enabled = false;
                    TabExtraction.TabPages["tabTarget"].HidePage();

                    TabExtraction.TabPages["tabNightShift"].Enabled = false;
                    TabExtraction.TabPages["tabNightShift"].HidePage();

                    tabControlMain.TabPages["tabPageEmployee"].Enabled = false;
                    tabControlMain.TabPages["tabPageEmployee"].HidePage();

                    tabControlMain.TabPages["tabPage12"].Enabled = false;
                    tabControlMain.TabPages["tabPage12"].HidePage();

                    tbpHabgUpProcess.TabPages["tabPage15"].Enabled = false;
                    tbpHabgUpProcess.TabPages["tabPage15"].HidePage();

                    //tabControlMain.TabPages["tabPageOffline"].Enabled = false;
                    //tabControlMain.TabPages["tabPageOffline"].HidePage();

                    #endregion "TEAM LEADS"
                }
                else if ((BusinessLogic.USERNAME == "Admin-Trivandrum") || (BusinessLogic.USERNAME == "Admin-Cochin") || (BusinessLogic.USERNAME == "Admin-Pondichery"))
                {
                    #region "HIDE PAGES"

                    #region "Online Tabs"

                    tbpHabgUpProcess.TabPages["tabPage15"].Enabled = false;
                    tbpHabgUpProcess.TabPages["tabPage15"].HidePage();

                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;

                    tabAllocation.TabPages["tabFileDetails"].Enabled = false;
                    tabAllocation.TabPages["tabFileDetails"].HidePage();

                    tabAllocation.TabPages["tabDeAllocation"].Enabled = false;
                    tabAllocation.TabPages["tabDeAllocation"].HidePage();

                    tabAllocation.TabPages["tabAttendance"].Enabled = false;
                    tabAllocation.TabPages["tabAttendance"].HidePage();

                    tabAllocation.TabPages["tabDiscrepancy"].Enabled = false;
                    tabAllocation.TabPages["tabDiscrepancy"].HidePage();

                    tabAllocation.TabPages["tbIdleTime"].Enabled = false;
                    tabAllocation.TabPages["tbIdleTime"].HidePage();

                    tabAllocation.TabPages["tbCustomizeEntry"].Enabled = false;
                    tabAllocation.TabPages["tbCustomizeEntry"].HidePage();

                    tabAllocation.TabPages["tbCapacity"].Enabled = false;
                    tabAllocation.TabPages["tbCapacity"].HidePage();

                    tabAllocation.TabPages["tbNDSPDetails"].Enabled = false;
                    tabAllocation.TabPages["tbNDSPDetails"].HidePage();

                    tabAllocation.TabPages["tbNightshiftusers"].Enabled = false;
                    tabAllocation.TabPages["tbNightshiftusers"].HidePage();

                    tabAllocation.TabPages["tbHundred"].Enabled = false;
                    tabAllocation.TabPages["tbHundred"].HidePage();

                    TabExtraction.TabPages["tabOnlineEntry"].Enabled = false;
                    TabExtraction.TabPages["tabOnlineEntry"].HidePage();

                    TabExtraction.TabPages["tabAccountIncentive"].Enabled = false;
                    TabExtraction.TabPages["tabAccountIncentive"].HidePage();

                    TabExtraction.TabPages["TabEmployee"].Enabled = false;
                    TabExtraction.TabPages["TabEmployee"].HidePage();

                    TabExtraction.TabPages["tabPageManual"].Enabled = false;
                    TabExtraction.TabPages["tabPageManual"].HidePage();

                    TabExtraction.TabPages["tabPageExtraction"].Enabled = false;
                    TabExtraction.TabPages["tabPageExtraction"].HidePage();

                    TabExtraction.TabPages["tabPageAllocationStatus"].Enabled = false;
                    TabExtraction.TabPages["tabPageAllocationStatus"].HidePage();

                    TabExtraction.TabPages["tabEmplList"].Enabled = false;
                    TabExtraction.TabPages["tabEmplList"].HidePage();

                    TabExtraction.TabPages["tabNightShift"].Enabled = false;
                    TabExtraction.TabPages["tabNightShift"].HidePage();

                    tabReports.TabPages["tabUserFileMinutes"].Enabled = false;
                    tabReports.TabPages["tabUserFileMinutes"].HidePage();

                    tabReports.TabPages["HourlyReport"].Enabled = false;
                    tabReports.TabPages["HourlyReport"].HidePage();

                    tabReports.TabPages["AccountWiseMinutes"].Enabled = false;
                    tabReports.TabPages["AccountWiseMinutes"].HidePage();

                    tabReports.TabPages["TransEdittabPage"].Enabled = false;
                    tabReports.TabPages["TransEdittabPage"].HidePage();

                    tabReports.TabPages["tbMinutes"].Enabled = false;
                    tabReports.TabPages["tbMinutes"].HidePage();

                    tabReports.TabPages["tabPage3"].Enabled = false;
                    tabReports.TabPages["tabPage3"].HidePage();

                    tabReports.TabPages["tbOnlytranscribed"].Enabled = false;
                    tabReports.TabPages["tbOnlytranscribed"].HidePage();

                    tabReports.TabPages["tbTargetDetails"].Enabled = false;
                    tabReports.TabPages["tbTargetDetails"].HidePage();

                    tabReports.TabPages["tbBlankReport"].Enabled = false;
                    tabReports.TabPages["tbBlankReport"].HidePage();

                    tabReports.TabPages["tbOnline_Incentive"].Enabled = false;
                    tabReports.TabPages["tbOnline_Incentive"].HidePage();

                    tabReports.TabPages["tbConsolidated"].Enabled = false;
                    tabReports.TabPages["tbConsolidated"].HidePage();

                    tabReports.TabPages["tabPage9"].Enabled = false;
                    tabReports.TabPages["tabPage9"].HidePage();

                    tabReports.TabPages["taNightshiftlines"].Enabled = false;
                    tabReports.TabPages["taNightshiftlines"].HidePage();

                    tabReports.TabPages["tab_Logsheethourly"].Enabled = false;
                    tabReports.TabPages["tab_Logsheethourly"].HidePage();

                    tabControlMain.TabPages["tabPageEmployee"].Enabled = false;
                    tabControlMain.TabPages["tabPageEmployee"].HidePage();

                    tabControlMain.TabPages["tabPage12"].Enabled = false;
                    tabControlMain.TabPages["tabPage12"].HidePage();

                    tabControlMain.TabPages["tbpHangUp"].Enabled = false;
                    tabControlMain.TabPages["tbpHangUp"].HidePage();

                    tabControlMain.TabPages["tabPageOffline"].Enabled = false;
                    tabControlMain.TabPages["tabPageOffline"].HidePage();

                    #endregion "Online Tabs"

                    #endregion "HIDE PAGES"
                }
                else if (BusinessLogic.IDESIG_ID == 125)
                {

                    tabPageOffline.HidePage();
                    tabPageOnline.HidePage();
                    Load_Branch();
                    GetEmployeeList();
                }
                else if ((BusinessLogic.IDESIG_ID == 244) || (BusinessLogic.IDESIG_ID == 98))
                {
                    #region "Managers"

                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = true;

                    tabAllocation.TabPages["tabFileDetails"].Enabled = false;
                    tabAllocation.TabPages["tabFileDetails"].HidePage();

                    tabAllocation.TabPages["tabDeAllocation"].Enabled = false;
                    tabAllocation.TabPages["tabDeAllocation"].HidePage();

                    tabAllocation.TabPages["tabAttendance"].Enabled = false;
                    tabAllocation.TabPages["tabAttendance"].HidePage();

                    TabExtraction.TabPages["tabPageManual"].Enabled = false;
                    TabExtraction.TabPages["tabPageManual"].HidePage();

                    TabExtraction.TabPages["tabPageExtraction"].Enabled = false;
                    TabExtraction.TabPages["tabPageExtraction"].HidePage();

                    TabExtraction.TabPages["tabNightShift"].Enabled = false;
                    TabExtraction.TabPages["tabNightShift"].HidePage();

                    tabReports.TabPages["tabPage3"].Enabled = false;
                    tabReports.TabPages["tabPage3"].HidePage();

                    tabAllocation.TabPages["tbNightshiftusers"].Enabled = false;
                    tabAllocation.TabPages["tbNightshiftusers"].HidePage();

                    tabReports.TabPages["tbOnline_Incentive"].Enabled = false;
                    tabReports.TabPages["tbOnline_Incentive"].HidePage();

                    tbpHabgUpProcess.TabPages["tabPage15"].Enabled = false;
                    tbpHabgUpProcess.TabPages["tabPage15"].HidePage();

                    #endregion "Managers"
                }
                else
                {
                    tabPageEmployee.HidePage();
                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                    TabExtraction.TabPages["tabOnlineEntry"].Enabled = false;
                    TabExtraction.TabPages["tabOnlineEntry"].HidePage();

                    TabExtraction.TabPages["tabAccountIncentive"].Enabled = false;
                    TabExtraction.TabPages["tabAccountIncentive"].HidePage();

                    TabExtraction.TabPages["TabEmployee"].Enabled = false;
                    TabExtraction.TabPages["TabEmployee"].HidePage();

                    tabAllocation.TabPages["tabDiscrepancy"].Enabled = false;
                    tabAllocation.TabPages["tabDiscrepancy"].HidePage();

                    tabAllocation.TabPages["tbIdleTime"].Enabled = false;
                    tabAllocation.TabPages["tbIdleTime"].HidePage();

                    tabReports.TabPages["tbBlankReport"].Enabled = false;
                    tabReports.TabPages["tbBlankReport"].HidePage();

                    tabReports.TabPages["tbConsolidated"].Enabled = false;
                    tabReports.TabPages["tbConsolidated"].HidePage();

                    tbpHabgUpProcess.TabPages["tabPage15"].Enabled = false;
                    tbpHabgUpProcess.TabPages["tabPage15"].HidePage();
                }
            }
        }

        /// <summary>
        /// ALLOCATION LIST VIEW MOUSE UP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lsvFileDetails_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (cbxTrans.Checked && cbxReview.Checked && cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = ALL;
                }
                else if (cbxTrans.Checked && !cbxReview.Checked && !cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = TransDesignation;
                }
                else if (!cbxTrans.Checked && !cbxReview.Checked && cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = EditDesination;
                }

                else if (!cbxTrans.Checked && cbxReview.Checked && !cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = ReviewDesignation;
                }

                else if (!cbxTrans.Checked && cbxReview.Checked && cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = EditAndReview;
                }

                else if (cbxTrans.Checked && !cbxReview.Checked && cbxEditing.Checked)
                {
                    lsvFileDetails.ContextMenuStrip = TransAndEdit;
                }

                else
                {
                    lsvFileDetails.ContextMenuStrip = null;
                }

                iTotalFiles = 0;
                iTotalMins = 0;
                if (iTotalFiles == 0 || iTotalFiles == 1)
                {
                    foreach (MyAllocatioFile oItem in lsvFileDetails.SelectedItems)
                    {
                        iTotalFiles++;
                        iTotalMins = iTotalMins + Convert.ToDouble(oItem.MINUTES);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void lsvEmployee_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            Load_Priority_Mapping();
        }

        private void lsvLoginEmoloyees_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            Load_AllotedFilesForEmployees();
        }

        /// <summary>
        /// DE ALLOT BUTTON CLICK EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeallot_Click(object sender, EventArgs e)
        {
            try
            {
                if (lsvAllotedFiles.CheckedItems.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No Item is marked for deallocation.");
                    return;
                }

                BusinessLogic.oMessageEvent.Start("Transferring data..");
                BusinessLogic.oProgressEvent.Start(true);
                this.Cursor = Cursors.WaitCursor;

                string _sTanscriptionCollection = string.Empty;
                string _sVoiceFile_ID = string.Empty;
                foreach (ListItem_AllotedFilesForUsers oDeAllocationItem in lsvAllotedFiles.CheckedItems)
                {
                    _sTanscriptionCollection = oDeAllocationItem.TRANSCRIPTION_ID.ToString();
                    _sVoiceFile_ID = oDeAllocationItem.sVoiceFile_ID.ToString();

                    int iDeAllttotFiles = BusinessLogic.WS_Allocation.Set_Deallot_Files(Convert.ToInt32(_sTanscriptionCollection), _sVoiceFile_ID, -1);
                    if (iDeAllttotFiles > 0)
                    {
                        Load_AllotedFilesForEmployees();
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Failed");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                this.Cursor = Cursors.Default;
            }
        }


        private void btnSubmitLeave_Click(object sender, EventArgs e)
        {
            try
            {
                //ListItem_EmployeeList oCurrentItem = (ListItem_EmployeeList)lsvEmployeeList.SelectedItems[0];                
                string sDayStatus = string.Empty;
                if (rbtHalfDay.Checked)
                    sDayStatus = "H";
                else if (rbtFullDay.Checked)
                    sDayStatus = "L";
                else if (rbtExtraSupport.Checked)
                    sDayStatus = "E";
                else if (rbtSundaySchedule.Checked)
                    sDayStatus = "S";
                else if (rbtCompoff.Checked)
                    sDayStatus = "C";
                else if (rbtSectionalOff.Checked)
                    sDayStatus = "SO";

                string sComment = txtComment.Text;
                DateTime sDate = dtpAttendanceDate.Value;

                string _MTID = string.Empty;
                DataSet _dsInsert = new DataSet();
                foreach (ListItem_EmployeeList oMarkedItem in lsvEmployeeList.CheckedItems)
                {
                    oMarkedItem.Selected = true;
                    _MTID = oMarkedItem.PRODUCTION_ID;

                    _dsInsert = BusinessLogic.WS_Allocation.Set_Leave_List(Convert.ToInt32(_MTID), sDate, sDayStatus, sComment, 1);
                }

                if (_dsInsert.Tables.Count > 0)
                {
                    Load_Leave_List();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// CHANGE PASSWORD BUTTON CLICK EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChangePassword_Click(object sender, EventArgs e)
        {
            try
            {
                frmChangePassword ChangePassword = new frmChangePassword();
                if (ChangePassword.ShowDialog() == DialogResult.OK)
                {
                    if (BusinessLogic.WS_Allocation.ChangePassword(BusinessLogic.SEMPLOYEEID, ChangePassword._sUserName, ChangePassword._sNewPassword))
                    {
                        BusinessLogic.oMessageEvent.Start("Password Successfully Changed..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }


        /// <summary>
        /// CHECK BOX TRANS SELECTED INDEX CHANGED
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxTrans_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                lsvFileDetails.Items.Clear();
                lvJobAccount.Items.Clear();
                if (txtEditingVoice.Text != "")
                    sEditVoice = txtEditingVoice.Text.ToString();
                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_AllocationList(1, Convert.ToDateTime(dtpAll_From.Text).ToString("yyyy/MM/dd"), Convert.ToDateTime(dtpAllTo.Text).ToString("yyyy/MM/dd"), sEditVoice).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsvFileDetails.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsvFileDetails.Items.Add(new MyAllocatioFile(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsvFileDetails);
                }
            }

            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                txtEditingVoice.Text = "";
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// CHECK BOX EDITING SELECTED INDEX CHANGED
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxEditing_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                lsvFileDetails.Items.Clear();
                if (txtEditingVoice.Text != "")
                    sEditVoice = txtEditingVoice.Text.ToString();
                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_AllocationList(2, Convert.ToDateTime(dtpAll_From.Text).ToString("yyyy/MM/dd"), Convert.ToDateTime(dtpAllTo.Text).ToString("yyyy/MM/dd"), sEditVoice).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsvFileDetails.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsvFileDetails.Items.Add(new MyAllocatioFile(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsvFileDetails);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                txtEditingVoice.Text = "";
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");

            }
        }

        /// <summary>
        /// CHECK BOX REVIEW SELECTED INDEX CHANGED
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxReview_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                lsvFileDetails.Items.Clear();
                if (txtEditingVoice.Text != "")
                    sEditVoice = txtEditingVoice.Text.ToString();
                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_AllocationList(3, Convert.ToDateTime(dtpAll_From.Text).ToString("yyyy/MM/dd"), Convert.ToDateTime(dtpAllTo.Text).ToString("yyyy/MM/dd"), sEditVoice).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsvFileDetails.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsvFileDetails.Items.Add(new MyAllocatioFile(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsvFileDetails);
                }
                else
                {
                    lsvFileDetails.Items.Clear();
                    BusinessLogic.oMessageEvent.Start("No Items Found");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                txtEditingVoice.Text = "";
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// FORM KEY DOWN EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmAllocation_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.R)
                {
                    Load_TabPage(0);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void lbLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sAccount = lbClient.SelectedValue.ToString();
            string sLocation = lbLocation.SelectedValue.ToString();
            Load_Doctor(Convert.ToInt32(sAccount), sLocation);
        }

        private void cboExtractAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sAccount = cboExtractAccount.SelectedValue.ToString();
            Load_Location(Convert.ToInt32(sAccount));
        }

        private void cboManualAccount_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try
            {
                if ((BusinessLogic.USERNAME != "Admin-Trivandrum") && (BusinessLogic.USERNAME != "Admin-Cochin") && (BusinessLogic.USERNAME != "Admin-Pondichery") && (BusinessLogic.IDESIG_ID != 281) && (BusinessLogic.IDESIG_ID != 244) && (BusinessLogic.IDESIG_ID != 98))
                {
                    string sAccount = cboManualAccount.SelectedValue.ToString();
                    Load_Location(Convert.ToInt32(sAccount));
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cboManualLocation_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string sAccount = cboManualAccount.SelectedValue.ToString();
            string sLocation = cboManualLocation.SelectedValue.ToString();
            Load_Doctor(Convert.ToInt32(sAccount), sLocation);
        }

        private void tabAllocation_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                iIndex = e.TabPageIndex;
                Load_TabPage(e.TabPageIndex);
                TabExtraction_Selecting(TabExtraction, e);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// VOICE FILE BROWSE BUTTON CLICK EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                lsvManual.Items.Clear();
                if (FBD.ShowDialog() == DialogResult.OK)
                {
                    string sLocation = FBD.SelectedPath;
                    txt_Location.Text = sLocation.Trim();
                    Create_DataTable();
                    btnExtract.Enabled = (lsvManual.Items.Count > 0);
                    int iDictationCount = 0;

                    int iProgCount = 1;

                    BusinessLogic.oMessageEvent.Start("Loading Dictations...");
                    BusinessLogic.oProgressEvent.Start(iProgCount, 0, iDictationCount + 1);

                    //Clearing the list before adding the new dictations
                    lsvManual.Items.Clear();
                    foreach (string sDictationName in Directory.GetFiles(sLocation))
                    {
                        TimeSpan oTime = TimeSpan.FromSeconds(Convert.ToDouble(Get_Minutes(sDictationName)));
                        int iMinute = oTime.Minutes;
                        int iSeconds = oTime.Seconds;

                        //Finding the length for the dictation
                        int iDictationSize = 0;
                        FileInfo oFileInfo = new FileInfo(sDictationName);
                        iDictationSize = ((int)oFileInfo.Length / 1024);

                        DataRow oNewRow = dtManual.NewRow();
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_DOCTOR_ID_BINT] = Convert.ToInt32(cboManualDoctor.SelectedValue);
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR] = Path.GetFileNameWithoutExtension(sDictationName);
                        oNewRow["Dictation_path"] = sDictationName;
                        oNewRow["Client_Id"] = Convert.ToInt32(cboManualAccount.SelectedValue);
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_EXTENSION_STR] = Path.GetExtension(sDictationName);
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE] = Convert.ToDecimal(iMinute.ToString().PadLeft(2, '0') + "." + iSeconds.ToString().PadRight(2, '0'));
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_SIZE_BINT] = iDictationSize;
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT] = iTatHours;
                        oNewRow[Framework.MAINTRANSCRIPTION.FIELD_COMMENT_STR] = "New Dictation";

                        dtManual.Rows.InsertAt(oNewRow, iDictationCount);
                        iDictationCount = iDictationCount + 1;

                    }
                    iDictationCount = 1;
                    foreach (DataRow drM in dtManual.Rows)
                    {
                        lsvManual.Items.Add(new MylsvManualDetails(drM, iDictationCount));
                        iDictationCount = iDictationCount + 1;
                    }
                    btnBrowse.Enabled = false;
                    iProgCount++;
                }
                else
                    btnBrowse.Enabled = true;

                foreach (ColumnHeader oColumn in lsvManual.Columns)
                {
                    oColumn.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                    int iWidth = oColumn.Width;
                    oColumn.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                    if (iWidth > oColumn.Width)
                        oColumn.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                }

                //btnBrowse.Enabled = false;
                btnExtract.Enabled = true;
                BusinessLogic.oMessageEvent.Start("Ready");
                BusinessLogic.oProgressEvent.Start(false);
                dtManual.Clear();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// VOICE FILE EXTRACT CLICK EVENT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExtract_Click(object sender, EventArgs e)
        {
            try
            {
                Thread oThread = new Thread(Extract);
                oThread.Start();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// VIEW EXTRACTION LOG
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnViewLog_Click(object sender, EventArgs e)
        {
            Thread tExtThread = new Thread(Get_Extractdetails);
            tExtThread.Start();
        }

        private void TabExtraction_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabAllocation.SelectedTab.Name == "tabOnlineAllocation")
            {
                if ((TabExtraction.SelectedTab.Name == "tabPageManual") || (TabExtraction.SelectedTab.Name == "tabPageExtraction"))
                {
                    Load_Account();
                }
                else if (TabExtraction.SelectedTab.Name == "tabPageAllocationStatus")
                {
                    Thread tThreadNew = new Thread(Load_Allocation_Status);
                    tThreadNew.Start();
                }
                else if (TabExtraction.SelectedTab.Name == "TabEmployee")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    LoadEmployeeAchievement();
                }
                else if (TabExtraction.SelectedTab.Name == "tabOnlineEntry")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    LoadEmployeeAchievement();
                    Load_File_Status();
                    Load_Account();
                }
                else if (TabExtraction.SelectedTab.Name == "tabTarget")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    Load_Month_Name();
                    LoadYear();
                    Load_Target();
                }
                else if (TabExtraction.SelectedTab.Name == "tabAccountIncentive")
                {
                    Load_Month_Name_ListView();
                    Load_Account_ListView();
                    LoadListView_AccountIncentive();
                }
                else if (TabExtraction.SelectedTab.Name == "tabEmplList")
                {
                    Load_All_Employees();
                }
                else if (TabExtraction.SelectedTab.Name == "tabNightShift")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    Load_Night_Shift_Marking();
                    Load_Night_Shift_Category();
                    Load_Month_Name();
                    LoadYear();
                }
                else if (TabExtraction.SelectedTab.Name == "tbTargetdetails")
                {
                    dtp_Target_Fromdate.Value = DateTime.Now;
                    dtp_Target_Todate.Value = DateTime.Now;
                    Load_Branch();
                    Load_Batch_Employee();
                }
                else if ((tabOffline.SelectedTab.Name == "tabTLFileTrack"))
                {
                    Load_Batch_Employee();
                }
            }

        }

        /// <summary>
        /// SHOW ALLOCATION STATUS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btlShowStatus_Click(object sender, EventArgs e)
        {
            Thread btnThread = new Thread(Load_Allocation_Status);
            btnThread.Start();
        }

        private void chkIncludeDate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkIncludeDate.Checked)
            {
                dtpFromDate.Enabled = true;
                dtpTO.Enabled = true;
            }
            else
            {
                dtpFromDate.Enabled = false;
                dtpTO.Enabled = false;
            }
        }

        private void mEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 3);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                    {
                        string sVoiceFileID = oFile.P_VOICE_FILE_NAME.ToString();
                        int sTransID;
                        if (cbxReview.Checked == true)
                        {
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;
                            //CHECK IF VOICE FILE ID EXISTS
                            //DataSet _dsVoiceFileID = new DataSet();
                            //_dsVoiceFileID = BusinessLogic.WS_Allocation.Get_Voice_FileID(sVoiceFileID);
                            //if (_dsVoiceFileID.Tables[0].Rows.Count > 0)
                            //{
                            sTransID = Convert.ToInt32(oItem.TRANSCRIPTIONID);
                            //sTransID = Convert.ToInt32(_dsVoiceFileID.Tables[0].Rows[0]["" + Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT + ""].ToString());
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(sTransID, oFile.P_USERID, DateTime.Now, sVoiceFileID, BusinessLogic.iTATREQUIRED);

                            oFile.P_STATUS = "Allotted";
                            oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;
                            //}
                            //else
                            //{
                            //BusinessLogic.oMessageEvent.Start("No First Review Entry..!");
                            //return;
                            //}
                        }
                        else
                        {
                            //insert into database
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.P_USERID, DateTime.Now, oFile.P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);

                            oFile.P_STATUS = "Allotted";
                            oFile.P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.P_USERID = BusinessLogic.ALLOTEDUSERID;
                        }
                    }
                    double dTotmins = 0;
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxReview_CheckedChanged(this, e);
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Maximized;
            notifyIcon1.Visible = false;
            notifyIcon1.Text = "FA V 1.4";
        }

        private void frmAllocation_Resize(object sender, EventArgs e)
        {
            //if (this.WindowState == FormWindowState.Minimized)
            //{
            //    Hide();
            //    notifyIcon1.Visible = true;
            //    notifyIcon1.Text = "FA V 1.4";
            //}
        }

        private void btnEntry_Click(object sender, EventArgs e)
        {
            frmVfileKeyIn FormEMON = new frmVfileKeyIn();
            if (FormEMON.ShowDialog() == DialogResult.OK)
            {
                int iSaveVoFile = BusinessLogic.WS_Allocation.Set_Pending_Files(Convert.ToInt32(FormEMON._site_ID), FormEMON._voice_file_ID, Convert.ToDateTime(FormEMON._file_date));
                if (iSaveVoFile > 0)
                {
                    cbxReview_CheckedChanged(cbxReview, e);
                    Application.DoEvents();
                }
            }
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            if (cbxTrans.Checked == true)
            {
                cbxTrans_CheckedChanged(this, e);
                Application.DoEvents();
            }
            else if (cbxEditing.Checked == true)
            {
                cbxEditing_CheckedChanged(this, e);
                Application.DoEvents();
            }
            else if (cbxReview.Checked == true)
            {
                cbxReview_CheckedChanged(this, e);
                Application.DoEvents();
            }
        }

        private void mARKDONEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocatioFile oItem = (MyAllocatioFile)lsvFileDetails.SelectedItems[0];

                foreach (MyAllocatioFile oFile in lsvFileDetails.SelectedItems)
                {
                    //insert into database                
                    int iResult = BusinessLogic.WS_Allocation.SET_CLIENT_ALLOTED(oFile.TRANSCRIPTIONID);
                    if (iResult == 1)
                    {
                        if (cbxTrans.Checked == true)
                        {
                            cbxTrans_CheckedChanged(this, e);
                            Application.DoEvents();
                        }
                        else if (cbxTrans.Checked == true)
                        {
                            cbxEditing_CheckedChanged(this, e);
                            Application.DoEvents();
                        }
                        else if (cbxTrans.Checked == true)
                        {
                            cbxReview_CheckedChanged(this, e);
                            Application.DoEvents();
                        }
                        BusinessLogic.oMessageEvent.Start("Marked..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxReview_CheckedChanged(this, e);
            }
        }

        private void mARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mARKDONEToolStripMenuItem_Click(this, e);
            Application.DoEvents();
        }

        private void btnEmployeeAchievement_Click(object sender, EventArgs e)
        {
            Thread btnThread = new Thread(LoadEmployeeAchievement);
            btnThread.Start();
        }

        private void changeMinutesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)lsvDiscrepancy.SelectedItems[0];
                string sChangedMinutes = string.Empty;
                frmChangeMinutes frmChange = new frmChangeMinutes();
                sChangedMinutes = lsvOnAll.sMinutes;

                string sLeftValue = sChangedMinutes.Split(':').GetValue(0).ToString();
                string sRightValue = sChangedMinutes.Split(':').GetValue(1).ToString();
                if (sLeftValue.Length == 1)
                {
                    sChangedMinutes = "0" + sLeftValue + ":" + sRightValue.ToString();
                }
                else
                {
                    sChangedMinutes = sLeftValue + ":" + sRightValue.ToString();
                }

                frmChange.sMinutes = sChangedMinutes;
                if (frmChange.ShowDialog() == DialogResult.OK)
                {
                    int iSetMinutes = BusinessLogic.WS_Allocation.set_minutes(lsvOnAll.sVoiceFileID, Convert.ToDouble(frmChange.iMinutes));
                    if (iSetMinutes > 0)
                    {
                        int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                        if (iApproveDeny > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Done!..!");
                            Load_Discrepancy();
                        }

                        BusinessLogic.oProgressEvent.Start(false);
                        BusinessLogic.oMessageEvent.Start("Minutes Changed Successfully");
                    }
                }
                frmChange.sMinutes = "";
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
            }
        }

        private void cbxUserDesig_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cbxUserDesig.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesigID, sBranch_ID);
        }

        private void btnUserMin_Click(object sender, EventArgs e)
        {
            LoadHourlyReport();
        }

        private void brnViewReport_Click(object sender, EventArgs e)
        {
            try
            {
                lsvHourlyrReports.Items.Clear();
                txtUserID.Text = string.Empty;
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");

                object oFromDate = null;
                object oTodate = null;
                string iFromHour, iToHour;
                string iBatchID = "1";
                string iBranchID = "-1";
                int iCustomized = 0;
                int iNightshift = 0;
                int iCreated_by = 0;

                string sHourlyFromDate = Convert.ToDateTime(dtpHourFrom.Value).ToString("yyyy-MM-dd");
                string sHourlyTodate = Convert.ToDateTime(dtpHourTo.Value).ToString("yyyy-MM-dd");

                iFromHour = cbxFromHours.SelectedItem.ToString();
                iToHour = cbxToHours.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = sHourlyFromDate + " 23:59:59";
                else
                    oFromDate = sHourlyFromDate + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = sHourlyTodate + " 23:59:59";
                else
                    oTodate = sHourlyTodate + " " + iToHour + ":" + "00:00";

                if (cbxHourDesig.SelectedValue != null)
                    iBatchID = cbxHourDesig.SelectedValue.ToString();

                if (cbxHourBranch.SelectedValue != null)
                    iBranchID = cbxHourBranch.SelectedValue.ToString();

                if (chb_Hourly_Customized.Checked == true)
                {
                    //if (chb_Hourly_P75.Checked == false && chb_Hourly_P99.Checked == false)
                    //{
                    //    BusinessLogic.oMessageEvent.Start("Select whose customized report want to view");
                    //    return;
                    //}
                    //if (Convert.ToInt32(cmb_Customize_Group.SelectedValue) == 0)
                    //{
                    //    BusinessLogic.oMessageEvent.Start("Select customized group");
                    //    return;
                    //}
                    iCustomized = 1;

                }
                else
                    iCustomized = 0;


                if (chb_Hourly_P75.Checked == true)
                    iCreated_by = 190;
                else if (chb_Hourly_P99.Checked == true)
                    iCreated_by = 202;
                else
                    iCreated_by = 0;


                if (chb_Nightshift.Checked == true)
                    iNightshift = 1;
                else
                    iNightshift = 0;

                if (iBranchID == "0")
                    iBranchID = "-1";
                BusinessLogic.WS_Allocation.Timeout = 300000;

                if (iNightshift == 0)
                    _dsHourlyWiseReport = BusinessLogic.WS_Allocation.Get_hourly_log_branch_V4(oFromDate, oTodate, Convert.ToInt32(iBatchID), Convert.ToInt32(iBranchID), iCustomized, Convert.ToInt32(cbxHourPlatform.SelectedValue), Convert.ToInt32(cmb_hourly_group.SelectedValue), cmb_Hourly_Location.SelectedValue.ToString());
                else
                    _dsHourlyWiseReport = BusinessLogic.WS_Allocation.Get_hourly_log_branch_Nightshift_V2(oFromDate, oTodate, Convert.ToInt32(iBatchID), Convert.ToInt32(iBranchID), Convert.ToInt32(cbxHourPlatform.SelectedValue), cmb_Hourly_Location.SelectedValue.ToString());


                int iRowCount = 0;
                foreach (DataRow _drRow in _dsHourlyWiseReport.Tables[0].Select())
                    lsvHourlyrReports.Items.Add(new ListItem_EmployeeHourly_Log(_drRow, iRowCount++));

                mTooltip.SetToolTip(lsvHourlyrReports, string.Empty);
                BusinessLogic.Reset_ListViewColumn(lsvHourlyrReports);
                BusinessLogic.oMessageEvent.Start("Done");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                //xxxxxBusinessLogic.oMessageEvent.Start("Done..!");
            }
        }


        Point mLastPos = new Point(-1, -1);

        private void lsvHourlyrReports_MouseMove(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = lsvHourlyrReports.HitTest(e.X, e.Y);
            try
            {

                if (mLastPos != e.Location)
                {
                    if (info.Item != null && info.SubItem != null)
                    {
                        if (info.Item.Name.ToString() == string.Empty)
                        {

                            mTooltip.ToolTipTitle = null;
                            mTooltip.Hide(info.Item.ListView);

                        }
                        else
                        {

                            mTooltip.ToolTipTitle = info.Item.Text;
                            mTooltip.Show(info.Item.Name.ToString().Replace("||", System.Environment.NewLine), info.Item.ListView);
                        }
                    }
                }

                mLastPos = e.Location;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Hourly Report between " + dtpHourFrom.Text + " " + " and " + dtpHourTo.Text + " " + cbxFromHours.SelectedItem + " to " + cbxToHours.SelectedItem + ".xls";
            ExportToExcel(lsvHourlyrReports, sFolderNAme, sFileName);
        }

        private void btnAchieveExport_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Employee Achievement Report for " + cboDesignation.Text + " between " + dtpEmployeeFromDate.Text + " and " + dtpEmployeeToDate.Text + ".xls";
            ExportToExcel(lvEmployeeAchievement, sFolderNAme, sFileName);
        }

        private void btnUserMinutesExcel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Line Count Details Of " + cbxUserName.Text + " date between " + dateTimeUserFrom.Text + " and " + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lsvLineCount, sFolderNAme, sFileName);
        }

        private void btnTransEditView_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedTab.Name == "tabPage10")
            {
                Thread tTransEdit = new Thread(LoadTransEditFiles);
                tTransEdit.Start();
            }
            else
            {
                Thread tTransEditOFF = new Thread(LoadTransEditFiles_Clinics);
                tTransEditOFF.Start();
            }
        }

        private void btnTransEditExcel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Trans Edit Details Of " + cbxUserName.Text + " date between " + dateTimeUserFrom.Text + " and " + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lsvTransEdit, sFolderNAme, sFileName);
        }

        private void cbxOnlineEntry_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cbxOnlineEntry.SelectedValue.ToString();
            Load_Employee_Full_name(sDesigID, "-1");
        }

        string sMinutes = string.Empty;
        string iMinutes = string.Empty;
        string sConvertedMins = string.Empty;
        string sConvertedLines = string.Empty;


        private void btnOnlineEntry_Click(object sender, EventArgs e)
        {
            try
            {
                if (vFileFetch.Checked == false)
                {
                    string sBatchID = cbxOnlineEntry.SelectedValue.ToString();
                    string sProductionID = cbxEntryEmployee.SelectedValue.ToString();
                    DateTime dProcessedDate = dtpProcessedDate.Value;
                    string sVoiceFileID = txtEntryVoiceFIle.Text.ToString();

                    sMinutes = txtMinutesEntry.Text.ToString();
                    iMinutes = (Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(0)) * 60 + Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(1))).ToString();
                    int dMinutes = Convert.ToInt32(iMinutes);
                    string sFileLines = txtFileLines.Text;
                    sConvertedLines = sFileLines;
                    int iFileStatusID = 0;
                    int iDocID = 0;

                    string sAccountID = cmbEntryAccount.SelectedValue.ToString();
                    iDocID = BusinessLogic.WS_Allocation.Get_doctorid(sAccountID);

                    if (cbxFileStatus.SelectedItem.ToString() == "Transcribed")
                    {
                        iFileStatusID = 1;
                    }
                    else if (cbxFileStatus.SelectedItem.ToString() == "Edited")
                    {
                        if (!rbtDspStatus.Checked)
                            iFileStatusID = 13;
                        else if (rbtHold.Checked)
                            iFileStatusID = 3;
                        else
                            iFileStatusID = 2;
                    }
                    else if (cbxFileStatus.SelectedItem.ToString() == "Trans Edit")
                    {
                        if (rbtHold.Checked)
                            iFileStatusID = 9;
                        else
                            iFileStatusID = 8;
                    }
                    else
                    {
                        if (rbtBlank.Checked)
                            iFileStatusID = 5;
                        else
                            iFileStatusID = 6;
                    }
                    int iSaveOnlineEntry = BusinessLogic.WS_Allocation.Set_OnlineEntry(sVoiceFileID, dMinutes, iFileStatusID, Convert.ToInt32(sFileLines) * 65, "NA", iDocID, "NA", Convert.ToInt32(sProductionID), sVoiceFileID + "_" + sProductionID + ".doc", Convert.ToInt32(sBatchID), Convert.ToInt32(sBatchID), dMinutes, rbtHold.Checked ? 1 : 0, 0);
                    if (iSaveOnlineEntry > 0)
                    {
                        BusinessLogic.oProgressEvent.Start(false);
                        BusinessLogic.oMessageEvent.Start("Done..!");
                        cbxOnlineEntry.SelectedIndex = 0;
                        cmbEntryAccount.SelectedIndex = 0;
                        cbxFileStatus.SelectedIndex = 0;
                        cbxEntryEmployee.SelectedIndex = 0;
                        txtEntryVoiceFIle.Text = "";
                        txtMinutesEntry.Text = "";
                        txtFileLines.Text = "";
                        BusinessLogic.oMessageEvent.Start("Entry Successfull..!");
                    }
                }
                else
                {
                    LoadEnteredFilesDetails();
                    txtEntryVoiceFIle.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                LoadEnteredFilesDetails();
            }
        }

        private void cbxFileStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(cbxFileStatus.SelectedIndex) == 0)
            {
                rbtHold.Enabled = false;
                rbtBlank.Enabled = false;
                rbtDspStatus.Enabled = false;
                rbtDspStatus.Checked = true;
                lblConvertedMinutes.Text = "";
                lblConvertedLines.Text = "";
            }
            if (Convert.ToInt32(cbxFileStatus.SelectedIndex) == 1)
            {
                rbtHold.Enabled = true;
                rbtBlank.Enabled = false;
                rbtDspStatus.Enabled = true;
                rbtDspStatus.Checked = true;
                lblConvertedMinutes.Text = "";
                lblConvertedLines.Text = "";
            }
            if (Convert.ToInt32(cbxFileStatus.SelectedIndex) == 2)
            {
                rbtHold.Enabled = true;
                rbtBlank.Enabled = false;
                rbtDspStatus.Enabled = true;

                if (txtMinutesEntry.Text == "  :")
                {
                    errorProvider1.SetError(txtMinutesEntry, "Please Enter Minutes!");
                    return;
                }
                else
                {
                    errorProvider1.Clear();
                    sMinutes = txtMinutesEntry.Text.ToString();
                    iMinutes = (Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(0)) * 60 + Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(1))).ToString();
                }

                if (txtFileLines.Text == "")
                {
                    errorProvider1.Clear();
                    errorProvider1.SetError(txtFileLines, "Please Enter Lines!");
                    return;
                }
                else
                {
                    sConvertedMins = lblConvertedMinutes.Text = "Converted Minute:" + Convert.ToDecimal(iMinutes) * 2;
                    sConvertedLines = lblConvertedLines.Text = "Converted Lines:" + Convert.ToDecimal(txtFileLines.Text) * 2;
                }
            }
            if (Convert.ToInt32(cbxFileStatus.SelectedIndex) == 3)
            {
                rbtHold.Enabled = false;
                rbtBlank.Enabled = true;
                rbtDspStatus.Enabled = true;
                rbtDspStatus.Checked = true;
                lblConvertedMinutes.Text = "";
                lblConvertedLines.Text = "";
            }
        }

        private void convertToTransEditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem oItem in lsvDiscrepancy.SelectedItems)
                {
                    ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)oItem;

                    int iSetTransEdit = BusinessLogic.WS_Allocation.Set_Trans_Edit(lsvOnAll.sVoiceFileID.Split('.').GetValue(0).ToString(), lsvOnAll.iPorductionID);
                    if (iSetTransEdit > 0)
                    {
                        int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                        if (iApproveDeny > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Done!..!");
                            Load_Discrepancy();
                        }

                        BusinessLogic.oProgressEvent.Start(false);
                        BusinessLogic.oMessageEvent.Start("Done..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnAccountView_Click(object sender, EventArgs e)
        {
            lvAccountWiseInfo.Items.Clear();
            //Thread tAccountWiseInfo = new Thread(Load_Account_Wise_Minutes);
            Thread tAccountWiseInfo = new Thread(Load_Account_Wise_Minutes_New);
            tAccountWiseInfo.Start();
        }

        private void btnAccountExcel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Account Wise File Details For Online Accounts From" + "_date between_" + dtpAccountFrom.Text + "_and_" + dtpAccountTo.Text + ".xls";
            ExportToExcel(lvAccountWiseInfo, sFolderNAme, sFileName);
        }

        private void chkInclusiveLines_CheckedChanged(object sender, EventArgs e)
        {
            txtInclusiveLines.Enabled = true;
            txtInclusiveMinutes.Enabled = false;
        }

        private void chkInclusiveMinutes_CheckedChanged(object sender, EventArgs e)
        {
            txtInclusiveLines.Enabled = false;
            txtInclusiveMinutes.Enabled = true;
        }

        private void btnInclusiveMinAdd_Click(object sender, EventArgs e)
        {
            try
            {
                string sDesignationID = string.Empty;
                string sProductionID = string.Empty;
                sDesignationID = cbxInDesignation.SelectedValue.ToString();
                sProductionID = cbxInEmployee.SelectedValue.ToString();
                string sDate = string.Empty;
                sDate = Convert.ToDateTime(dtpInDatePicker.Value).ToString("yyyy/MM/dd HH:mm:ss");

                if (rbtInclusiveLines.Checked)
                {
                    string sInclusiveLines = string.Empty;
                    sInclusiveLines = txtInclusiveLines.Text;
                    string sComment = txtInComment.Text.ToString();

                    int iInclusiveLines = BusinessLogic.WS_Allocation.Set_Inclusive_Lines(Convert.ToInt32(sProductionID), sDate, Convert.ToDecimal(sInclusiveLines), sComment);
                    if (iInclusiveLines > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Inclusive Lines Inserted Successfully..!");
                        txtInclusiveLines.Text = "";
                        cbxInDesignation.SelectedIndex = 0;
                        cbxInEmployee.SelectedIndex = 0;
                        txtInComment.Text = "";
                    }
                }
                else if (rbtInclusiveMinutes.Checked)
                {
                    string sVoiceFileID = txtInclusiveMinutes.Text.ToString();
                    sMinutes = txtInclusiveMinutes.Text.ToString();
                    iMinutes = (Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(0)) * 60 + Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(1))).ToString();
                    string sComment = txtInComment.Text.ToString();

                    int iInclusiveMinutes = BusinessLogic.WS_Allocation.Set_Inclusive_Minutes(Convert.ToInt32(sProductionID), sDate, Convert.ToDouble(iMinutes), sComment);
                    if (iInclusiveMinutes > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Inclusive Minutes Inserted Successfully..!");
                        txtInclusiveMinutes.Text = "";
                        cbxInDesignation.SelectedIndex = 0;
                        cbxInEmployee.SelectedIndex = 0;
                        txtInComment.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                LoadListView_InclusiveLines();
            }
        }


        private void lvDesignation_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Thread NewThreadEmp = new Thread(Load_Employee_List);
                NewThreadEmp.Start();
            }
        }

        private void cmbTargetDisig_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranch = string.Empty;
            if (BusinessLogic.LOGIN_NAME == "Admin@rndsoftech.com")
            {
                sBranch = "1";
            }
            else if (BusinessLogic.LOGIN_NAME == "Cochin@rndsoftech.com")
            {
                sBranch = "2";
            }
            else if (BusinessLogic.LOGIN_NAME == "Trivandrum@rndsoftech.com")
            {
                sBranch = "3";
            }
            else if (BusinessLogic.LOGIN_NAME == "Pondichery@rndsoftech.com")
            {
                sBranch = "4";
            }
            else
            {
                sBranch = "-1";
            }
            string sDesigID = cmbTargetDisig.SelectedValue.ToString();
            Load_Employee_Full_name(sDesigID, sBranch);
        }

        private void cmbEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Target();
        }

        private void changeTargetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ListItem_Target lsvOnTarget = (ListItem_Target)lvTarget.SelectedItems[0];
                string sProdID = lsvOnTarget.iTargetProdID.ToString();
                string sTarget = lsvOnTarget.iTarget.ToString();

                int iChangeTarget = 0;
                frmChangeTarget ChangeTarget = new frmChangeTarget();
                ChangeTarget.sCurrentTarget = sTarget;
                if (ChangeTarget.ShowDialog() == DialogResult.OK)
                {
                    iChangeTarget = BusinessLogic.WS_Allocation.Set_Target(lsvOnTarget.iTragetID, Convert.ToInt32(sProdID), ChangeTarget.sChangedTarget);
                }

                if (iChangeTarget > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Target Changed..!");
                    Load_Target();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void btnAttendanceRerun_Click(object sender, EventArgs e)
        {
            frmAttendance frmAttendance = new frmAttendance();
            frmAttendance.Show();
        }

        private void btnViewAttendance_Click(object sender, EventArgs e)
        {
            Load_Leave_List();
        }

        //private void btnAppDeny_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        int iOptionID = 0;
        //        if (rbtApprove.Checked == true)
        //            iOptionID = 1;
        //        else if (rbtDeny.Checked == true)
        //            iOptionID = 2;

        //        int iApproveDeny = 0;
        //        if (lvNoEntry.CheckedItems.Count > 0)
        //        {
        //            foreach (ListItem_Discrepancy oMarkedItem in lvNoEntry.CheckedItems)
        //            {
        //                oMarkedItem.Selected = true;
        //                iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(oMarkedItem.iDiscrepancyID, iOptionID);
        //                if (iApproveDeny > 0)
        //                {
        //                    BusinessLogic.oMessageEvent.Start("Done!..!");
        //                    Load_Discrepency_For_No_Entry();
        //                }
        //            }
        //        }
        //        else if (lvOtherDiscrepancy.CheckedItems.Count > 0)
        //        {
        //            foreach (ListItem_Discrepancy oMarkedItem in lvOtherDiscrepancy.CheckedItems)
        //            {
        //                oMarkedItem.Selected = true;
        //                iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(oMarkedItem.iDiscrepancyID, iOptionID);
        //                if (iApproveDeny > 0)
        //                {
        //                    BusinessLogic.oMessageEvent.Start("Done!..!");
        //                    Load_Discrepency_For_Others();
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
        //        BusinessLogic.oMessageEvent.Start("Failed..!");
        //    }
        //}   

        private void btnUpdateAttendance_Click(object sender, EventArgs e)
        {
            try
            {
                string sDayStatus = string.Empty;
                if (rbtHalfDay.Checked)
                    sDayStatus = "H";
                else if (rbtFullDay.Checked)
                    sDayStatus = "L";
                else if (rbtExtraSupport.Checked)
                    sDayStatus = "E";
                else if (rbtSundaySchedule.Checked)
                    sDayStatus = "S";
                else if (rbtCompoff.Checked)
                    sDayStatus = "C";
                else if (rbtSectionalOff.Checked)
                    sDayStatus = "SO";

                int iUpdateLeave = 0;
                foreach (ListItem_LeaveDetails oMarkedItem in lsvLeaveList.CheckedItems)
                {
                    iUpdateLeave = BusinessLogic.WS_Allocation.Set_Update_Leave(oMarkedItem.DAY_ID, sDayStatus);
                }

                if (iUpdateLeave > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Updated!..!");
                    Load_Leave_List();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally { }
        }

        private void vFileFetch_CheckedChanged(object sender, EventArgs e)
        {
            if (vFileFetch.Checked == true)
            {
                btnOnlineEntry.Text = "Fetch";
                btnOnlineEntry.BackColor = System.Drawing.Color.Blue;
                btnOnlineEntry.ForeColor = System.Drawing.Color.White;
            }
            else
            {
                btnOnlineEntry.Text = "Entry";
                btnOnlineEntry.BackColor = System.Drawing.Color.Red;
                btnOnlineEntry.ForeColor = System.Drawing.Color.White;
            }
        }

        private void dtpInDatePicker_ValueChanged(object sender, EventArgs e)
        {
            LoadListView_InclusiveLines();
        }

        private void TpEditMinutes_Opening(object sender, CancelEventArgs e)
        {
            try
            {
                ListItem_Discrepancy oItem = (ListItem_Discrepancy)lsvDiscrepancy.SelectedItems[0];
                int iSelevtedValue = oItem.iDiscrepancyMasterID;
                if (iSelevtedValue == 1)
                {
                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                    TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = true;
                }
                else if (iSelevtedValue == 2)
                {
                    if ((BusinessLogic.USERNAME == "Admin-Coimbatore") || (BusinessLogic.IDESIG_ID == 244) || (BusinessLogic.IDESIG_ID == 98))
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = true;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                    else
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                }
                else if (iSelevtedValue == 3)
                {
                    if ((BusinessLogic.USERNAME == "Admin-Coimbatore") || (BusinessLogic.IDESIG_ID == 244) || (BusinessLogic.IDESIG_ID == 98))
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = true;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                    else
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                }
                else if (iSelevtedValue == 4)
                {
                    TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                    TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = true;
                    TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                    TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                }
                else if (iSelevtedValue == 5)
                {
                    if ((BusinessLogic.USERNAME == "Admin-Coimbatore") || (BusinessLogic.IDESIG_ID == 244) || (BusinessLogic.IDESIG_ID == 98))
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = true;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                    else
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                }
                else if (iSelevtedValue == 6)
                {
                    if ((BusinessLogic.USERNAME == "Admin-Coimbatore") || (BusinessLogic.IDESIG_ID == 244) || (BusinessLogic.IDESIG_ID == 98))
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = true;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                    else
                    {
                        TpEditMinutes.Items["convertToTransEditToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeMinutesToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeFileStatusToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeShiftToolStripMenuItem"].Enabled = false;
                        TpEditMinutes.Items["changeLinesToolStripMenu"].Enabled = false;
                        TpEditMinutes.Items["AddEntryToolStripMenu"].Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally { }
        }

        private void changeLinesToolStripMenu_Click(object sender, EventArgs e)
        {
            try
            {
                ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)lsvDiscrepancy.SelectedItems[0];
                string sChangeLines = string.Empty;
                frmChangeLines frmChangeLines = new frmChangeLines();
                sChangeLines = lsvOnAll.sLines;

                frmChangeLines.sLines = sChangeLines;
                if (frmChangeLines.ShowDialog() == DialogResult.OK)
                {
                    int iSetLines = BusinessLogic.WS_Allocation.set_Lines(lsvOnAll.sVoiceFileID, Convert.ToDecimal(frmChangeLines.iLines), lsvOnAll.iPorductionID);
                    if (iSetLines > 0)
                    {
                        int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                        if (iApproveDeny > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Done!..!");
                            Load_Discrepancy();
                        }

                        BusinessLogic.oProgressEvent.Start(false);
                        BusinessLogic.oMessageEvent.Start("Minutes Changed Successfully");
                    }
                }
                frmChangeLines.sLines = "";
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {

            }
        }

        private void changeFileStatusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                foreach (ListViewItem oItem in lsvDiscrepancy.SelectedItems)
                {
                    ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)oItem;
                    int iSetFileStatus = BusinessLogic.WS_Allocation.Set_File_Status(lsvOnAll.sVoiceFileID.Split('.').GetValue(0).ToString(), lsvOnAll.iPorductionID, lsvOnAll.iFileStatusID);
                    if (iSetFileStatus > 0)
                    {
                        int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                        if (iApproveDeny > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Done!..!");
                            Load_Discrepancy();
                        }

                        BusinessLogic.oProgressEvent.Start(false);
                        BusinessLogic.oMessageEvent.Start("Done..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void changeShiftToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                foreach (ListViewItem oItem in lsvDiscrepancy.SelectedItems)
                {
                    ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)oItem;
                    int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                    if (iApproveDeny > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Done!..!");
                        Load_Discrepancy();
                    }

                    BusinessLogic.oProgressEvent.Start(false);
                    BusinessLogic.oMessageEvent.Start("Done..!");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void AddEntryToolStripMenu_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                foreach (ListViewItem oItem in lsvDiscrepancy.SelectedItems)
                {
                    //ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)lsvDiscrepancy.SelectedItems[0];
                    ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)oItem;
                    string sVoice = lsvOnAll.sVoiceFileID;
                    string sMinutes = lsvOnAll.sMinutes;
                    string iCMinutes = (Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(0)) * 60 + Convert.ToDecimal(sMinutes.ToString().Split(':').GetValue(1))).ToString();
                    string sLines = lsvOnAll.sLines;
                    int iProductionID = lsvOnAll.iPorductionID;
                    int iFilesStatusID = lsvOnAll.iFileStatusID;
                    //DateTime dProcessedDate = lsvOnAll.dDateEntered;
                    DateTime dProcessedDate = lsvOnAll.dSubmitted_time;

                    int iAccountID = lsvOnAll.iAccountID;

                    int iSavedSuccessfully = 0;

                    iSavedSuccessfully = BusinessLogic.WS_Allocation.Set_First_Level_entry(Convert.ToInt32(iAccountID), sVoice.ToString(), Convert.ToInt32(iCMinutes), dProcessedDate, Convert.ToDecimal(sLines), iFilesStatusID, iProductionID);
                    if (iSavedSuccessfully > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Done..!");
                        int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 1);
                        if (iApproveDeny > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Done!..!");
                            Load_Discrepancy();
                        }
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Failed..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("All Entries Done!..");
            }
        }

        private void btnInceSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (lvMonthName.SelectedItems.Count == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Please choose a month..!");
                    return;
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Ready..!");
                }

                if (lsvAccountName.SelectedItems.Count == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Please choose a account..!");
                    return;
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Ready..!");
                }
                if (txtIncentiveRate.Text == "")
                {
                    BusinessLogic.oMessageEvent.Start("Please enter percentage..!");
                    return;
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Ready..!");
                }

                ListItem_Account lsvIncAccount = (ListItem_Account)lsvAccountName.SelectedItems[0];
                ListViewItem lsvMonth = (ListViewItem)lvMonthName.SelectedItems[0];
                string sClientID = lsvIncAccount.iClientID.ToString();
                string sClientName = lsvIncAccount.sClietnName.ToString();
                string sMonthName = lsvMonth.Tag.ToString();
                string sRate = txtIncentiveRate.Text.ToString();

                int iSaveIncRate = BusinessLogic.WS_Allocation.Set_Incentive_Rate(Convert.ToInt32(sClientID), sClientName, sMonthName, Convert.ToInt32(sRate));
                if (iSaveIncRate > 0)
                {
                    LoadListView_AccountIncentive();
                    txtIncentiveRate.Text = "";
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnLinesView_Click(object sender, EventArgs e)
        {
            Thread NewAllBranchThread = new Thread(Load_AllBranch_Lines);
            NewAllBranchThread.Start();
        }

        private void tabOffline_Selecting(object sender, TabControlCancelEventArgs e)
        {
            cbxOffline_Trans.Checked = true;
            if (tabOffline.SelectedTab.Name == "tabOffPriorityMap")
            {
                DataTable dtPriority = new DataTable();
                dtPriority.Columns.Add("Value", typeof(System.String));
                dtPriority.Columns.Add("Key", typeof(System.Int32));

                DataRow drNew = dtPriority.NewRow();
                drNew["Value"] = "MT";
                drNew["Key"] = 1;
                dtPriority.Rows.Add(drNew);

                drNew = dtPriority.NewRow();
                drNew["Value"] = "ME";
                drNew["Key"] = 2;
                dtPriority.Rows.Add(drNew);

                drNew = dtPriority.NewRow();
                drNew["Value"] = "BA";
                drNew["Key"] = 5;
                dtPriority.Rows.Add(drNew);

                drNew = dtPriority.NewRow();
                drNew["Value"] = "TED";
                drNew["Key"] = 3;
                dtPriority.Rows.Add(drNew);

                drNew = dtPriority.NewRow();
                drNew["Value"] = "AM";
                drNew["Key"] = 4;
                dtPriority.Rows.Add(drNew);

                lbDesig.DisplayMember = "Value";
                lbDesig.ValueMember = "Key";
                lbDesig.DataSource = dtPriority;

                Load_Priority_Mapping();
                Load_Account();
            }
            else if (tabOffline.SelectedTab.Name == "tablOffDeallot")
            {
                Load_Batch_Employee();
            }
            else if (tabOffline.SelectedTab.Name == "tbAllocationPriority")
            {
                Load_Batch_Employee();
                Load_Account();
            }
            else if (tabOffline.SelectedTab.Name == "tabOfflineReport")
            {
                LoadYear();
                Load_Month_Name();
                Load_Offline_Volume();
            }
            else if (tabOffline.SelectedTab.Name == "tabTLFileTrack")
            {                
                Load_Batch_Employee();
            }
        }

        private void tabReports_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabAllocation.SelectedTab.Name == "tabPReports")
            {
                if (tabReports.SelectedTab.Name == "TransEdittabPage")
                    btnTransEditView_Click(this, e);
                else if (tabReports.SelectedTab.Name == "HourlyReport")
                {

                    brnViewReport_Click(this, e);
                    Load_Clienttype();
                    Load_Branch();
                    Load_Customized_group();
                    Load_Batch_Employee();
                    Load_Location_Type(CLIENT_TYPE_ID);
                }
                else if (tabReports.SelectedTab.Name == "tbOnline_Incentive")
                {
                    Load_Branch();
                }
                else if (tabReports.SelectedTab.Name == "AccountWiseMinutes")
                {
                    Load_Acc_File_status();
                    Load_Acc_TL_File_Status();
                    btnAccountView_Click(this, e);
                }
                else if (tabReports.SelectedTab.Name == "tab_Accountwise_TL")
                {                    
                    Load_Acc_TL_File_Status();
                    btn_AccTL_View_Click(this, e);
                }
                else if (tabReports.SelectedTab.Name == "BranchWiseMinutesDeatails")
                    btnLinesView_Click(this, e);
                else if (tabReports.SelectedTab.Name == "tabUserFileMinutes")
                {
                    LoadHourlyReport();
                    cbxUFromHours.SelectedIndex = 0;
                    cbxUToHours.SelectedIndex = 0;
                }
                else if (tabReports.SelectedTab.Name == "tabLogSheet")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    datDateFrom.Value = DateTime.Now;
                    datDateTo.Value = DateTime.Now;
                }
                else if (tabReports.SelectedTab.Name == "tab_Logsheethourly")
                {
                    Load_Branch();
                    Load_Batch_Employee();
                    dtp_Log_HourlyFrom.Value = DateTime.Now;
                    dtp_Log_HourlyTo.Value = DateTime.Now;
                }
                else if (tabReports.SelectedTab.Name == "tbMinutes")
                {
                    LoadYear();
                    Load_Month_Name();
                }
                else if (tabReports.SelectedTab.Name == "tbTargetDetails")
                {
                    dtp_Target_Fromdate.Value = DateTime.Now;
                    dtp_Target_Todate.Value = DateTime.Now;
                    Load_Branch();
                    Load_Batch_Employee();
                }
                else if (tabReports.SelectedTab.Name == "tabPage3")
                {
                    LoadYear();
                    Load_Month_Name();
                }
                else if (tabReports.SelectedTab.Name == "tbOnlytranscribed")
                {
                    dtp_online_transonly_Fromdate.Value = DateTime.Now;
                    dtp_online_transonly_Todate.Value = DateTime.Now;
                    Get_TransOnly_Online();
                }
                else if (tabReports.SelectedTab.Name == "tbBlankReport")
                {
                    Load_All_Details();
                }
                else if (tabReports.SelectedTab.Name == "tbConsolidated")
                {
                    Load_All_Con_Details();
                }
                else if (tabReports.SelectedTab.Name == "tab_user_consolidated")
                {
                    LoadYear();
                    Load_Month_Name();
                    Load_Branch();
                    Load_Designation();
                    Load_Emp_List();
                }
                else if (tabReports.SelectedTab.Name == "tabPage23")
                {
                    LoadTopTwentuHold();
                }
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            int iBatchID = 0;
            int iOpType = 0;
            int iProductionID = 0;
            int iUpdateBatch = 0;

            if (rbtDeactivate.Checked == false) // TO CHANGE DESIGNtTION
            {
                iOpType = 1;
                if (rbtMT.Checked == true)
                {
                    iBatchID = 1;
                    if (lsvMT.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemMT in lsvMT.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnMT = (Listitem_AllEmployees)oItemMT;
                            iProductionID = lsvOnMT.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                    if (lsvME.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemME in lsvME.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnME = (Listitem_AllEmployees)oItemME;
                            iProductionID = lsvOnME.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                    if (lsvTED.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemTED in lsvTED.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnTED = (Listitem_AllEmployees)oItemTED;
                            iProductionID = lsvOnTED.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                }
                else if (rbtME.Checked == true)
                {
                    iBatchID = 2;
                    if (lsvMT.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemME in lsvME.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnME = (Listitem_AllEmployees)oItemME;
                            iProductionID = lsvOnME.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                    if (lsvMT.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemMT in lsvMT.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnMT = (Listitem_AllEmployees)oItemMT;
                            iProductionID = lsvOnMT.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                    if (lsvTED.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemTED in lsvTED.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnTED = (Listitem_AllEmployees)oItemTED;
                            iProductionID = lsvOnTED.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                    if (lsvAM.SelectedItems.Count > 0)
                    {
                        foreach (ListViewItem oItemAM in lsvAM.SelectedItems)
                        {
                            Listitem_AllEmployees lsvOnAM = (Listitem_AllEmployees)oItemAM;
                            iProductionID = lsvOnAM.iProductionID;

                            iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                            Load_All_Employees();
                        }
                    }
                }
                else if (rbtTED.Checked == true)
                {
                    iBatchID = 3;
                    foreach (ListViewItem oItemTED in lsvTED.SelectedItems)
                    {
                        Listitem_AllEmployees lvTED = (Listitem_AllEmployees)oItemTED;
                        iProductionID = lvTED.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                        Load_All_Employees();
                    }
                }
                else if (rbtAM.Checked == true)
                {
                    iBatchID = 4;
                    foreach (ListViewItem oItemAM in lsvAM.SelectedItems)
                    {
                        Listitem_AllEmployees lvAM = (Listitem_AllEmployees)oItemAM;
                        iProductionID = lvAM.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(iBatchID, iProductionID, iOpType);
                        Load_All_Employees();
                    }
                }
            }
            else // TO DEACTIVATE
            {
                iOpType = 2;
                if (rbtMT.Checked == true)
                {
                    iBatchID = 1;
                    foreach (ListViewItem oItemMT in lsvMT.SelectedItems)
                    {
                        Listitem_AllEmployees lsvOnMT = (Listitem_AllEmployees)oItemMT;
                        iProductionID = lsvOnMT.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(-1, iProductionID, iOpType);
                        Load_All_Employees();
                    }
                }
                else if (rbtME.Checked == true)
                {
                    iBatchID = 2;
                    foreach (ListViewItem oItemME in lsvME.SelectedItems)
                    {
                        Listitem_AllEmployees lsvOnME = (Listitem_AllEmployees)oItemME;
                        iProductionID = lsvOnME.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(-1, iProductionID, iOpType);
                    }
                }
                else if (rbtTED.Checked == true)
                {
                    iBatchID = 3;
                    foreach (ListViewItem oItemTED in lsvTED.SelectedItems)
                    {
                        Listitem_AllEmployees lvTED = (Listitem_AllEmployees)oItemTED;
                        iProductionID = lvTED.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(-1, iProductionID, iOpType);
                    }
                }
                else if (rbtAM.Checked == true)
                {
                    iBatchID = 4;
                    foreach (ListViewItem oItemAM in lsvAM.SelectedItems)
                    {
                        Listitem_AllEmployees lvAM = (Listitem_AllEmployees)oItemAM;
                        iProductionID = lvAM.iProductionID;

                        iUpdateBatch = BusinessLogic.WS_Allocation.set_update_batch(-1, iProductionID, iOpType);
                    }
                }
            }
        }

        private void lvNDesignation_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Thread NewThreadEmp = new Thread(Load_Employee_List_NightShiftMark);
                NewThreadEmp.Start();
            }
        }

        private void btnAddNightShift_Click(object sender, EventArgs e)
        {
            try
            {
                string sMTIDAllow = string.Empty;
                Listitem_Category oCategory = (Listitem_Category)lvNightShiftAllowance.SelectedItems[0];
                string sCatID = oCategory.iCategoryID.ToString();
                int sMonth = comboBox4.SelectedIndex + 1;
                string iYear = comboBox5.SelectedItem.ToString();
                foreach (ListItem_EmployeeList oMarkedItem in lvNEmployee.CheckedItems)
                {
                    oMarkedItem.Selected = true;
                    sMTIDAllow = oMarkedItem.PRODUCTION_ID;

                    int iSaveShiftAllovance = BusinessLogic.WS_Allocation.set_shift_allowance(Convert.ToInt32(sCatID), Convert.ToInt32(sMTIDAllow), Convert.ToInt32(sMonth), Convert.ToInt32(iYear));

                    if (iSaveShiftAllovance > 0)
                    {
                        Load_Night_Shift_Marking();
                        BusinessLogic.oMessageEvent.Start("Done..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnViewLogSheet_Click(object sender, EventArgs e)
        {
            try
            {
                Load_LineCountReport();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cbmLogBatch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cbmLogBatch.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesgination_ID, sBranch_ID);
        }

        private void changeMinutesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string sChangedMinutes = string.Empty;
                frmChangeMinutes frmChange = new frmChangeMinutes();
                string sVoiceFileID = string.Empty;
                if ((tabControlMain.SelectedTab.Name == "tabPageOffline"))
                {
                    ListItem_OfflineAllocationStatus lsvStatus = (ListItem_OfflineAllocationStatus)lsvOfflineDetail.SelectedItems[0];
                    sChangedMinutes = lsvStatus.sMinutes;
                    if (lsvStatus.sVoiceFileID.Contains("."))
                        sVoiceFileID = lsvStatus.sVoiceFileID.Split('.').GetValue(0).ToString();
                    else
                        sVoiceFileID = lsvStatus.sVoiceFileID.ToString();
                }

                else
                {
                    //ListItem_OnlineAllocationStatus lsvStatus = (ListItem_OnlineAllocationStatus)lvAllocationStatus.SelectedItems[0];
                    ListItem_LargeMinutes lsvMinutes = (ListItem_LargeMinutes)lsvLargeMinutes.SelectedItems[0];
                    sChangedMinutes = lsvMinutes.MINUTES.ToString();
                    if (lsvMinutes.VOICE_FILE_ID.Contains("."))
                        sVoiceFileID = lsvMinutes.VOICE_FILE_ID.Split('.').GetValue(0).ToString();
                    else
                        sVoiceFileID = lsvMinutes.VOICE_FILE_ID.ToString();

                    int totalSeconds = Convert.ToInt32(sChangedMinutes);
                    int seconds = totalSeconds % 60;
                    int minutes = totalSeconds / 60;
                    string time = minutes + ":" + seconds;
                    sChangedMinutes = time.ToString();
                }
                if (sChangedMinutes != "0")
                {
                    if (sChangedMinutes.Contains(":"))
                    {
                        string sLeftValue = sChangedMinutes.Split(':').GetValue(0).ToString();
                        string sRightValue = sChangedMinutes.Split(':').GetValue(1).ToString();
                        if (sLeftValue.Length == 1)
                        {
                            sChangedMinutes = "0" + sLeftValue + ":" + sRightValue.ToString();
                        }
                        else
                        {
                            sChangedMinutes = sLeftValue + ":" + sRightValue.ToString();
                        }
                    }
                }

                frmChange.sMinutes = sChangedMinutes;
                if (frmChange.ShowDialog() == DialogResult.OK)
                {
                    int iUpdateMinutes = BusinessLogic.WS_Allocation.set_minutes(sVoiceFileID, Convert.ToDouble(frmChange.iMinutes));
                    if (iUpdateMinutes > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Change Done..!");
                        Thread tThreadNew = new Thread(Load_Allocation_Status);
                        tThreadNew.Start();
                        BusinessLogic.oMessageEvent.Start("Change Done..!");
                        Thread tThreadNew1 = new Thread(Load_Allocation_Status_Offline);
                        tThreadNew1.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsvDeAllotBatch_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Thread NewThreadEmp = new Thread(Load_Employee_List);
                NewThreadEmp.Start();
            }
        }

        private void btnApproveDeny_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem oItem in lsvDiscrepancy.SelectedItems)
                {
                    ListItem_Discrepancy lsvOnAll = (ListItem_Discrepancy)oItem;

                    int iApproveDeny = BusinessLogic.WS_Allocation.Set_Discrepancy_Approve(lsvOnAll.iDiscrepancyID, 2);
                    if (iApproveDeny > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Done!..!");
                        Load_Discrepancy();
                    }

                    BusinessLogic.oProgressEvent.Start(false);
                    BusinessLogic.oMessageEvent.Start("Done..!");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnLogExport_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Log_Sheet_Of_" + cmbLogUser.Text + ".xls";
            ExportToExcel(lsvLineCountDetails, sFolderNAme, sFileName);
        }

        private void tabDescripancy_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabAllocation.SelectedTab.Name == "tabDiscrepancy")
            {
                if (tabDescripancy.SelectedTab.Name == "tabPage1")
                {
                    Thread NewDiscrepancy = new Thread(Load_Discrepancy);
                    NewDiscrepancy.Start();
                }
                else if (tabDescripancy.SelectedTab.Name == "tabPage2")
                {
                    Thread NewDiscrepancyReport = new Thread(Load_Discrepancy_Report);
                    NewDiscrepancyReport.Start();
                }
                else if (tabDescripancy.SelectedTab.Name == "tabPage4")
                {
                    Thread NewDiscrepancyReport = new Thread(Load_Multiple_Entries);
                    NewDiscrepancyReport.Start();
                }
            }
        }

        private void cmbDesignation_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                lsvMappingDeatils.Items.Clear();

                DataSet _dsTypistMap = new DataSet();
                string iBatchID = cmbDesignation.SelectedValue.ToString();
                _dsTypistMap = BusinessLogic.WS_Allocation.Get_Typist_Mapping(Convert.ToInt32(iBatchID));

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsTypistMap.Tables[0].Select())
                    lsvMappingDeatils.Items.Add(new ListIte_MappingDeatils_Priority(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvMappingDeatils);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void assignToDoctorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string sProduction_ID = string.Empty;
                foreach (ListViewItem oItem in lsvMappingDeatils.SelectedItems)
                {
                    ListIte_MappingDeatils_Priority lsvMap = (ListIte_MappingDeatils_Priority)oItem;
                    if (sProduction_ID == string.Empty)
                    {
                        sProduction_ID = lsvMap.PRODUCTION_ID.ToString();
                    }
                    else
                    {
                        sProduction_ID = sProduction_ID + "," + lsvMap.PRODUCTION_ID.ToString();
                    }
                }
                frmAssign_priority fPriority = new frmAssign_priority();
                fPriority._Productionid = sProduction_ID;
                fPriority.ShowDialog();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void tbConOfReport_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                DataSet dsBatch = new DataSet();
                dsBatch = BusinessLogic.WS_Allocation.Get_BatchDetails();

                if (tbConOfReport.SelectedTab.Name == "tbOfflineAccountWise")
                {
                    lvOfflineAccountWise.Items.Clear();
                    Thread tOffAccountWiseInfo = new Thread(Load_Offline_Account_Wise_Minutes);
                    tOffAccountWiseInfo.Start();
                }
                else if (tbConOfReport.SelectedTab.Name == "tbpTAT")
                {
                    Load_Location(0);
                }
                else if (tbConOfReport.SelectedTab.Name == "tbPageBackLock")
                {
                    Load_Offline_Account_Wise_BackLock();
                }
                else if (tbConOfReport.SelectedTab.Name == "tbTatPercentage")
                {
                    LoadYear();
                    Load_Month_Name();
                }
                else if (tbConOfReport.SelectedTab.Name == "tbTransonly_Offline")
                {
                    lsv_transonly_offline.Items.Clear();
                    Thread tTransonly_offline = new Thread(Get_TransOnly_Offline);
                    tTransonly_offline.Start();
                }
                else if (tbConOfReport.SelectedTab.Name == "tbMarquee")
                {
                    lsvMarquee.Items.Clear();
                    Thread tMarquee = new Thread(Get_Marquee_Text);
                    tMarquee.Start();
                }


                else if (tbConOfReport.SelectedTab.Name == "tbDaywise")
                {
                    cbYearDaywise.Items.Clear();
                    int iCurrentYear = DateTime.Now.Year;
                    for (int i = 2014; i <= iCurrentYear; i++)
                    {
                        cbYearDaywise.Items.Add(i.ToString());
                    }
                    cbYearDaywise.SelectedIndex = 0;
                    cbYearDaywise.SelectedIndex = (cbYearDaywise.Items.Count + 1);
                    cbYearDaywise.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else if (tbConOfReport.SelectedTab.Name == "tbUserwise")
                {
                    cbUserWiseBatch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                    cbUserWiseBatch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                    cbUserWiseBatch.DataSource = dsBatch.Tables[0];
                    cbUserWiseBatch.DropDownStyle = ComboBoxStyle.DropDownList;
                    dtpUserwiseFrom.Value = DateTime.Now;
                    dtpUserwiseTo.Value = DateTime.Now;
                }
                else
                {
                    Load_Clienttype();
                    Load_Branch();
                    Load_Batch_Employee();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
            }
        }

        private void btnViewVolume_Click(object sender, EventArgs e)
        {
            Thread tOffThread = new Thread(Load_Offline_Volume);
            tOffThread.Start();
        }

        private void lvJobAccount_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Get_Offline_AllocationDetails();
            }
        }

        private void btn_transonly_offline_Click(object sender, EventArgs e)
        {
            try
            {
                Get_TransOnly_Offline();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_UserTarget_Report_Click(object sender, EventArgs e)
        {
            Load_Target_Details();
        }

        private void cbxUserBranch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cbxUserBranch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_Employee_Full_name(sDesgination_ID, sBranchId);
        }

        private void cbmLogBranch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cbmLogBranch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_Employee_Full_name(sDesgination_ID, sBranchId);
        }

        private void cmb_Target_Branch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cmb_Target_Branch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_EmpName_List(sDesgination_ID, sBranchId);
        }

        private void cmb_Customized_Branch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cmb_Customized_Branch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_Employee_Full_name(sDesgination_ID, sBranchId);
        }

        private void cmb_Customized_Desig_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cmb_Customized_Desig.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesigID, sBranch_ID);
        }

        private void btn_Customized_Add_Click(object sender, EventArgs e)
        {
            try
            {
                //validation
                if (cmb_Customized_Employee.Text.Trim() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the employee Name");
                    cmb_Customized_Employee.Focus();
                    return;
                }

                //int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee(Convert.ToInt32(cmb_Customized_Employee.SelectedValue));
                //int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee_V2(Convert.ToInt32(cmb_Customized_Employee.SelectedValue), Convert.ToInt32(BusinessLogic.SPRODUCTIONID));
                int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee_NEW(Convert.ToInt32(cmb_Customized_Employee.SelectedValue), Convert.ToInt32(cmb_Customize_Group.SelectedValue));

                if (iResult > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Added Sucessfully");
                    Load_Customized_Employee_List();
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Already Added");
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (ICUSTOM_REMOVAL == 0)
                {
                    foreach (ListItem_Customized_Employee oItem in lsv_Customized_Employee.SelectedItems)
                    {
                        int iUpdate = BusinessLogic.WS_Allocation.Set_Customized_EmployeeList_Remove(Convert.ToInt32(oItem.IPRODUCTION_ID));
                        if (iUpdate > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Removed Sucessfully");
                            Load_Customized_Employee_List();
                        }
                    }
                }
                else
                {
                    foreach (ListItem_Customized_Employee oItem in lsv_Customized_Remove_Employee.SelectedItems)
                    {
                        int iUpdate = BusinessLogic.WS_Allocation.Set_Customized_EmployeeList_Remove_All(Convert.ToInt32(oItem.IPRODUCTION_ID));
                        if (iUpdate > 0)
                        {
                            BusinessLogic.oMessageEvent.Start("Removed Sucessfully");
                            Load_Customized_Employee_Removal_List();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsv_Customized_Employee_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                ICUSTOM_REMOVAL = 0;
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (ListItem_Customized_Employee oitem in lsv_Customized_Employee.SelectedItems)
                    {
                        Customized_Emp_contextMenuStrip.Visible = true;
                        Customized_Emp_contextMenuStrip.Show(PointToScreen(Control.MousePosition));
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        #endregion "EVENTS"

        #region " OFFLINE "

        #region " CLASSES "

        /// <summary>
        /// GET THE EMPLOYEE LIST BASED ON THE SELECTION OF BATCH
        /// </summary>
        public class MyDeallocate_Offline_EmployeeList : ListViewItem
        {
            public string EMP_PRODUCTION_ID;


            public MyDeallocate_Offline_EmployeeList(DataRow _dr, int i)
            {
                int AllottedFile = Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(1));
                string AllottedMin = _dr["alloted"].ToString().Split('-').GetValue(0).ToString();
                int AchivedFile = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(1));
                string AchivedMins = _dr["achived"].ToString().Split('-').GetValue(0).ToString();
                int TotFile = AchivedFile + AllottedFile;
                int Totsecs = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(2)) + Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(2));
                string TotMins = sGetDuration(Totsecs);

                this.Text = _dr["emp_id"].ToString();
                this.SubItems.Add(_dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString());
                this.SubItems.Add(_dr["designation"].ToString());
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                this.SubItems.Add(AllottedFile.ToString());
                this.SubItems.Add(AllottedMin);
                this.SubItems.Add(AchivedFile.ToString());
                this.SubItems.Add(AchivedMins);
                this.SubItems.Add(TotFile.ToString());
                this.SubItems.Add(TotMins);

                EMP_PRODUCTION_ID = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
            private string sGetDuration(int Seconds)
            {
                string sDuration = string.Empty;

                object oTotalDuration = null;
                oTotalDuration = Seconds;
                if (oTotalDuration != null)
                {
                    int iTotalSeconds = Convert.ToInt32(oTotalDuration);
                    int iMinutes = 0;
                    int iSeconds = 0;
                    if (iTotalSeconds > 0)
                    {
                        iMinutes = iTotalSeconds / 60;
                        iSeconds = iTotalSeconds % 60;
                    }
                    sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
                }
                return sDuration;
            }
        }



        /// <summary>
        /// GET THE EMPLOYEE LIST BASED ON THE SELECTION OF BATCH
        /// </summary>
        public class Offline_MT_Tracking_EmployeeList : ListViewItem
        {
            public string EMP_PRODUCTION_ID;


            public Offline_MT_Tracking_EmployeeList(DataRow _dr, int i)
            {
                int AllottedFile = Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(1));
                string AllottedMin = _dr["alloted"].ToString().Split('-').GetValue(0).ToString();
                int AchivedFile = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(1));
                string AchivedMins = _dr["achived"].ToString().Split('-').GetValue(0).ToString();
                int TotFile = AchivedFile + AllottedFile;
                int Totsecs = Convert.ToInt32(_dr["achived"].ToString().Split('-').GetValue(2)) + Convert.ToInt32(_dr["alloted"].ToString().Split('-').GetValue(2));
                string TotMins = sGetDuration(Totsecs);

                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.TARGET.FIELD_TARGET_MINS_INT].ToString());
                this.SubItems.Add(TotFile.ToString());
                this.SubItems.Add(TotMins);
                this.SubItems.Add(AllottedFile.ToString());
                this.SubItems.Add(AllottedMin);
                this.SubItems.Add(AchivedFile.ToString());
                this.SubItems.Add(AchivedMins);

                EMP_PRODUCTION_ID = _dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
            private string sGetDuration(int Seconds)
            {
                string sDuration = string.Empty;

                object oTotalDuration = null;
                oTotalDuration = Seconds;
                if (oTotalDuration != null)
                {
                    int iTotalSeconds = Convert.ToInt32(oTotalDuration);
                    int iMinutes = 0;
                    int iSeconds = 0;
                    if (iTotalSeconds > 0)
                    {
                        iMinutes = iTotalSeconds / 60;
                        iSeconds = iTotalSeconds % 60;
                    }
                    sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
                }
                return sDuration;
            }
        }


        /// <summary>
        /// LOAD THE EMPLOYEE DESIGNATION LIST
        /// </summary>
        public class MyDeallocatte_Offline_DesignationList : ListViewItem
        {
            public int iBatchID;
            public MyDeallocatte_Offline_DesignationList(DataRow dr, int i)
                : base()
            {
                Text = dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT].ToString();
                SubItems.Add(dr[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                iBatchID = Convert.ToInt32(dr[Framework.BATCH.FIELD_BATCH_BATCHID_INT]);
            }
        }

        /// <summary>
        /// Get the offline allocation list view items
        /// </summary>
        public class MyAllocation_Offline_List : ListViewItem
        {
            public int TRANSCRIPTIONID;
            public int CLIENTID;
            public int DOCTORID;
            public string MINUTES;
            public string STATUS;
            public string USERID;
            public string EMP_NAME;
            public string REMAINING_TAT;
            public double FILE_DURATION;

            public MyAllocation_Offline_List(DataRow dr, int i)
            {
                Name = dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString();
                Text = dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                SubItems.Add(dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                SubItems.Add(dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TAT_INT].ToString());
                SubItems.Add(dr["Hours_Completed"].ToString());
                SubItems.Add(dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIBED_BY_STR].ToString());
                SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_EDIT_BY_STR].ToString());
                //SubItems.Add(dr[Framework.MAINTRANSCRIPTION.FIELD_HOLD_REVIEW_BY_STR].ToString());

                STATUS = dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString();
                EMP_NAME = "";

                TRANSCRIPTIONID = Convert.ToInt32(dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString());
                CLIENTID = Convert.ToInt32(dr[Framework.CLIENT.FIELD_CLIENT_ID_BINT].ToString());
                DOCTORID = Convert.ToInt32(dr[Framework.DOCTOR.FIELD_DOCTOR_ID_BINT].ToString());
                MINUTES = dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString();
                REMAINING_TAT = dr["Hours_Completed"].ToString();
                FILE_DURATION = Convert.ToDouble(dr["fMin"].ToString());

                if (REMAINING_TAT.Contains("-"))
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }
                else if (i % 2 == 1)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                }
                else
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
                }
            }

            public string OFFLINE_P_STATUS
            {
                set
                {
                    this.SubItems[8].Text = value;
                    STATUS = value;
                    if (value.ToString() == "Allotted") //ENUM
                        this.BackColor = Color.Orange;
                    else if (value.ToString() == "Ready")
                    {
                        this.BackColor = Color.White;
                    }
                }
                get
                {
                    this.SubItems[8].Text = STATUS;
                    return STATUS;
                }
            }

            public string OFFLINE_P_EMPNAME
            {
                set
                {
                    this.SubItems[9].Text = value;
                    EMP_NAME = value;
                }
                get
                {
                    this.SubItems[9].Text = EMP_NAME;
                    return EMP_NAME;
                }
            }

            public string OFFLINE_P_USERID
            {
                set
                {
                    USERID = value;
                }
                get
                {
                    return USERID;
                }
            }

            public double OFFLINE_P_FILEMINS
            {
                set
                {
                    this.SubItems[4].Text = value.ToString();
                }
                get
                {
                    return Convert.ToDouble(this.SubItems[4].Text);
                }
            }

            public double OFFLINE_P_FILEMINS_V2
            {
                set
                {
                    FILE_DURATION = value;
                }
                get
                {
                    return FILE_DURATION;
                }
            }

            public string OFFLINE_P_VOICE_FILE_NAME
            {
                set
                {
                    this.SubItems[3].Text = value.ToString();
                }
                get
                {
                    return this.SubItems[3].Text;
                }
            }
        }

        public class MyAllocation_Offline_AllotedFiles_List : ListViewItem
        {
            public int TRANSCRIPTION_ID;
            public string sVoiceFile_ID;
            public string ALLOTED_FOR;

            public MyAllocation_Offline_AllotedFiles_List(DataRow _dr, int i)
            {
                Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                SubItems.Add(_dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                SubItems.Add(_dr[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                SubItems.Add(_dr["alloted_to"].ToString());
                SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_ALLOTED_DATE_DTIME].ToString());
                SubItems.Add(_dr[Framework.FILESTATUS.FIELD_FILE_STATUS_DESCRIPTION_STR].ToString());

                TRANSCRIPTION_ID = Convert.ToInt32(_dr[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT]);
                sVoiceFile_ID = _dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();
                ALLOTED_FOR = _dr[Framework.MAINTRANSCRIPTION.FIELD_ALLOTED_FOR_STR].ToString();

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class Mylsvdownloaddetails : ListViewItem
        {
            public string sLocation_id;
            public string dFileDate;

            public Mylsvdownloaddetails(DataRow dr, int iRowCount)
            {
                this.Text = dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(dr["file_minutes"].ToString());
                this.SubItems.Add(dr["File_count"].ToString());
                this.SubItems.Add(dr[Framework.CLIENT.FIELD_CLIENT_TAT_INT].ToString());

                sLocation_id = dr[Framework.LOCATION.FIELD_LOCATION_ID_STR].ToString();

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }

            public Mylsvdownloaddetails(string sMessage, string sTotalMinutes, string sTotFiles)
                : base()
            {
                this.Text = sMessage;
                this.SubItems.Add(sTotalMinutes);
                this.SubItems.Add(sTotFiles);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White;
            }
        }

        public class DownLoaded_Files_Deatils : ListViewItem
        {
            public string sStatus;

            public DownLoaded_Files_Deatils(DataRow _dr, int iRows)
                : base()
            {
                this.Text = _dr[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString();
                this.SubItems.Add(_dr["TotMinutes"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_dr["status"].ToString());

                sStatus = _dr["status"].ToString();

                if (sStatus == "To be alloted")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#F2F5A9");
                else if (sStatus == "In Process")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#F79F81");
                else if (sStatus == "For Editing")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#D8F781");
                else if (sStatus == "For Delivery")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#BCF5A9");
                else if (sStatus == "Formatted")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#A9E2F3");
            }
        }

        public class Employee_File_Alloted_Details : ListViewItem
        {
            public string VOICE_FILE_ID;
            public string TRANSCRIPTION_ID;
            public int is_tat;
            public int IS_AUTO_ALLOCATION;
            public string IS_STATUS;

            public Employee_File_Alloted_Details(DataRow _drRow, int iRowCount)
                : base()
            {
                this.Text = _drRow[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString();
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_MINUTES_DOUBLE].ToString());
                this.SubItems.Add(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_drRow[Framework.CLIENT.FIELD_CLIENT_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR].ToString());
                this.SubItems.Add(_drRow[Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT].ToString());
                this.SubItems.Add(_drRow["Remaining_tat"].ToString());

                TRANSCRIPTION_ID = _drRow[Framework.MAINTRANSCRIPTION.FIELD_TRANSCRIPTION_ID_BINT].ToString();
                is_tat = Convert.ToInt32(_drRow[Framework.MAINTRANSCRIPTION.FIELD_FILE_IS_TAT]);
                IS_AUTO_ALLOCATION = Convert.ToInt32(_drRow["is_auto_allocation"]);
                IS_STATUS = _drRow[Framework.FILESTATUS.FIELD_FILE_STATUS_ID_BINT].ToString();

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");

                if (is_tat == 1)
                {
                    this.BackColor = System.Drawing.Color.Red;
                    this.ForeColor = System.Drawing.Color.White;
                }

                if (IS_STATUS == "Opened")
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#ff5733");                    
                }
                else if (IS_STATUS == "Temp Hold")
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#33fff0");
                }
                else if (IS_STATUS == "In Process")
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#c7ff33");
                }
                else
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#a533ff");
                    this.ForeColor = System.Drawing.Color.White;
                }

                if (IS_AUTO_ALLOCATION == 1)
                {
                    this.BackColor = System.Drawing.Color.Blue;
                    this.ForeColor = System.Drawing.Color.White;
                } 
            }

            public Employee_File_Alloted_Details(string sMessage, string sTotMins)
                : base()
            {
                this.Text = sMessage;
                this.SubItems.Add(sTotMins);

                this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.BackColor = Color.Teal;
                this.ForeColor = Color.White; 
            }
        }


        public class ListIte_MappingDeatils_Priority : ListViewItem
        {
            public int PRODUCTION_ID;

            public ListIte_MappingDeatils_Priority(DataRow _dr, int i)
                : base()
            {
                this.Text = _dr[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());
                this.SubItems.Add(_dr[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                this.SubItems.Add(_dr["MappedDocCount"].ToString());
                this.SubItems.Add(_dr["MappedACCCount"].ToString());
                this.SubItems.Add(_dr["ClientNames"].ToString());
                this.SubItems.Add(_dr["DoctorNames"].ToString());

                PRODUCTION_ID = Convert.ToInt32(_dr[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID].ToString());

                if (i % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class List_Item_TATStatus : ListViewItem
        {
            public string STATUS;

            public List_Item_TATStatus(DataRow _dr, int iCount)
            {
                this.Text = _dr[Framework.LOCATION.FIELD_LOCATION_NAME_STR].ToString();
                this.SubItems.Add(_dr["ClientTat"].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_VOICE_FILE_ID_STR].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_FILE_DATE_DTIME].ToString());
                this.SubItems.Add(_dr[Framework.MAINTRANSCRIPTION.FIELD_EDIT_DATE_DTIME].ToString());
                this.SubItems.Add(_dr["TAT"].ToString());
                this.SubItems.Add(_dr["Status"].ToString());
                this.SubItems.Add(_dr["OUT_of_TAT_Reason"].ToString());


                STATUS = _dr["Status"].ToString();
                if (STATUS == "Out of TAT")
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#ff7d66");
                else
                    this.BackColor = System.Drawing.Color.Green;
            }
        }

        public class List_Item_List_Hold_Employee : ListViewItem
        {
            public string iProductionID = string.Empty;

            public List_Item_List_Hold_Employee(DataRow _drRow, int iRow)
                : base()
            {
                this.Text = _drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_drRow[Framework.BATCH.FIELD_BATCH_BATCHNAME_STR].ToString());

                iProductionID = Convert.ToInt32(_drRow[Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID]).ToString();

                if (iRow % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class List_Item_List_Employee_Moved : ListViewItem
        {
            public string sUnique_ID;
            public string sHoldTransaction;

            public List_Item_List_Employee_Moved(DataRow _drRow, int iRow)
                : base()
            {
                this.Text = _drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME].ToString();
                this.SubItems.Add(_drRow[Framework.EMPLOYEE.FIELD_EMPLOYEE_ID].ToString());
                this.SubItems.Add(_drRow[Framework.PRODUCTION_EMPLOYEES.FIELD_PRODUCTION_EMPLOYEE_TAG].ToString());
                this.SubItems.Add(_drRow["current_month"].ToString());
                this.SubItems.Add(_drRow["for_the_year"].ToString());
                this.SubItems.Add(_drRow["hold_allowed"].ToString() + " %");

                sUnique_ID = Convert.ToInt32(_drRow["unique_id"]).ToString();
                sHoldTransaction = Convert.ToInt32(_drRow["hold_slab_transaction_id"]).ToString();

                if (iRow % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class List_Item_List_Hold_Slab : ListViewItem
        {
            public string sHoldSlabID;

            public List_Item_List_Hold_Slab(DataRow _drRow, int iRow)
                : base()
            {
                this.Text = _drRow["current_month"].ToString();
                this.SubItems.Add(_drRow["hold_allowed"].ToString() + " %");

                sHoldSlabID = Convert.ToInt32(_drRow["hold_slab_id"]).ToString();

                if (iRow % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class List_item_NTS_Tat_Percentage : ListViewItem
        {
            public List_item_NTS_Tat_Percentage(DataRow _drRow, int iRow)
                : base()
            {
                this.Text = _drRow["location_name"].ToString();
                this.SubItems.Add(_drRow["tat_date"].ToString());
                this.SubItems.Add(_drRow["tat_percentage"].ToString());

                if (iRow % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class List_View_Employee : ListViewItem
        {
            public int iProductionID;
            public string iEscriptionID;
            public string iDictaphoneID;

            public string sEmployeeName = string.Empty;

            public List_View_Employee(DataRow _drRow, int iRowCount)
                : base()
            {
                Text = iRowCount.ToString();
                SubItems.Add(_drRow["emp_full_name"].ToString());
                SubItems.Add(_drRow["dictaphone_id"].ToString());
                SubItems.Add(_drRow["escription_id"].ToString());
                SubItems.Add(_drRow["branch_name"].ToString());
                SubItems.Add(_drRow["batch_name"].ToString());
                SubItems.Add(_drRow["work_platform"].ToString());
                SubItems.Add(_drRow["production_id"].ToString());
                SubItems.Add(_drRow["ptag_id"].ToString());
                SubItems.Add(_drRow["HT"].ToString());

                iProductionID = Convert.ToInt32(_drRow["production_id"]);
                iEscriptionID = _drRow["escription_id"].ToString();
                iDictaphoneID = _drRow["dictaphone_id"].ToString();

                sEmployeeName = _drRow["emp_full_name"].ToString();

                if (Convert.ToInt32(_drRow["status"]) == 1)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#fe8b72");
                    //this.ForeColor = Color.White;
                }
                else if (Convert.ToInt32(_drRow["status"]) == 2)
                {
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#b7fab0");
                }
            }
        }


        #endregion " CLASSES "

        #region " METHODS "
        private void Get_Offline_AllocationDetails()
        {
            try
            {
                Mylsvdownloaddetails lsvOffAccount = (Mylsvdownloaddetails)lvJobAccount.SelectedItems[0];
                lsv_OfflieFile_Details.Items.Clear();
                BusinessLogic.oMessageEvent.Start("Requesting Database...");

                int iType = 0;
                if (cbxOffline_Trans.Checked == true)
                {
                    iType = 1;
                }
                else if (cbxOffline_Editing.Checked == true)
                {
                    iType = 2;
                }
                else if (cbxOffline_Review.Checked == true)
                {
                    iType = 3;
                }
                DataTable dtAllocation = BusinessLogic.WS_Allocation.Get_Offline_Allocation(iType, null, null, null, lsvOffAccount.sLocation_id).Tables[0];
                if (dtAllocation != null)
                {
                    if (dtAllocation.Rows.Count > 0)
                    {
                        int i = 1;
                        lsv_OfflieFile_Details.Items.Clear();
                        foreach (DataRow dr in dtAllocation.Rows)
                        {
                            lsv_OfflieFile_Details.Items.Add(new MyAllocation_Offline_List(dr, i));
                            i++;
                        }
                        BusinessLogic.Reset_ListViewColumn(lsv_OfflieFile_Details);
                    }
                    else
                    {
                        lvJobAccount.Items.Clear();
                    }
                    BusinessLogic.oMessageEvent.Start("Loaded Dictations...");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                txtEditingVoice.Text = "";
            }
        }

        /// <summary>
        /// LOAD THE DESIGNATION DETAILS
        /// </summary>
        private void Load_Offline_Designation()
        {
            try
            {
                DataSet dsBatch = new DataSet();
                dsBatch = BusinessLogic.WS_Allocation.Get_BatchDetails();
                int iRowcount = 0;
                if (dsBatch != null)
                {
                    lsv_Offline_Deall_Designation.Items.Clear();
                    foreach (DataRow dr in dsBatch.Tables[0].Rows)
                        lsv_Offline_Deall_Designation.Items.Add(new MyDeallocatte_Offline_DesignationList(dr, iRowcount++));
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Can't load the designation details..!");
            }
        }

        /// <summary>
        /// LOAD THE EMPLOYEE LIST
        /// </summary>
        private void Load_Offline_Deallocation_EmpList()
        {
            try
            {
                int Batch = ((MyDeallocatte_Offline_DesignationList)lsv_Offline_Deall_Designation.SelectedItems[0]).iBatchID;
                DataTable _dsEmpList = BusinessLogic.WS_Allocation.GET_ALLOTED_DETAILS_NEW(Batch).Tables[0];

                lsv_Offline_Deall_EmplDetails.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drLogingEmp in _dsEmpList.Select())
                    lsv_Offline_Deall_EmplDetails.Items.Add(new MyDeallocate_Offline_EmployeeList(_drLogingEmp, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsv_Offline_Deall_EmplDetails);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD THE ALLOTEDFILES OF SELECTED EMPLOYEE
        /// </summary>
        private void Load_AllotedFilesForEmployees_Offline()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                DataTable _dsAllotedFiles_Offline = new DataTable();
                MyDeallocate_Offline_EmployeeList oCurrentLogin_Offline = (MyDeallocate_Offline_EmployeeList)lsv_Offline_Deall_EmplDetails.SelectedItems[0];
                //MyDeallocate_Offline_EmployeeList
                _dsAllotedFiles_Offline = BusinessLogic.WS_Allocation.Get_AllocationDetails(oCurrentLogin_Offline.EMP_PRODUCTION_ID, 1);

                lsv_Offline_Deallocate.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drAllFiles in _dsAllotedFiles_Offline.Select())
                    lsv_Offline_Deallocate.Items.Add(new MyAllocation_Offline_AllotedFiles_List(_drAllFiles, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsv_Offline_Deallocate);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// METHOD TO LOAD TRACKING DETAILS
        /// </summary>
        private void Load_MT_Tracking()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                lvRepEmp.Items.Clear();

                DataSet _dsTypistMap = new DataSet();
                string iBatchID = cbxReportDesig.SelectedValue.ToString();
                string iWork_Platform = cmb_MTTrack_Workplatform.SelectedValue.ToString();

                _dsTypistMap = BusinessLogic.WS_Allocation.GET_ALLOTED_DETAILS_NEW_V2_WORKPLATFORM(Convert.ToInt32(iBatchID), Convert.ToInt32(iWork_Platform));

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsTypistMap.Tables[0].Select())
                    lvRepEmp.Items.Add(new Offline_MT_Tracking_EmployeeList(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvRepEmp);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// LOAD IDLE TIME DAY WISE
        /// </summary>
        private void LoadIdleTimeDayWise(object sFromDate, object sToDate)
        {
            try
            {
                object CurrentDate = null;
                object PreviousDate = null;

                DataSet _dsIdleDayWise = new DataSet();
                ListItem_IdleSummary IdleTiming = (ListItem_IdleSummary)lvIdleSummary.SelectedItems[0];
                ListItem_IdleDateWise oListItem;

                _dsIdleDayWise = BusinessLogic.WS_Allocation.Get_Idle_DayWise(sFromDate, sToDate, Convert.ToInt32(IdleTiming.iProductionID)); ;


                int iRowCount = 1;
                int TotalMinutesIdle = 0;
                lvIdleProcess.Items.Clear();
                foreach (DataRow _drIdleSummary in _dsIdleDayWise.Tables[0].Select())
                {
                    CurrentDate = BusinessLogic.ConvertToDateTime(_drIdleSummary["idle_start_time"]);
                    int iDaydifferece = Convert.ToDateTime(CurrentDate).Day - Convert.ToDateTime(PreviousDate).Day;

                    if (PreviousDate == null)
                    {
                        oListItem = new ListItem_IdleDateWise(_drIdleSummary, iRowCount++);
                        lvIdleProcess.Items.Add(oListItem);
                        PreviousDate = CurrentDate;
                    }
                    else if (iDaydifferece == 0)
                    {
                        oListItem = new ListItem_IdleDateWise(_drIdleSummary, iRowCount++);
                        lvIdleProcess.Items.Add(oListItem);
                        PreviousDate = CurrentDate;
                    }
                    else
                    {
                        TotalMinutesIdle += Convert.ToInt32(_drIdleSummary["TotalIdle"].ToString());
                        oListItem = new ListItem_IdleDateWise(TotalMinutesIdle.ToString());
                        lvIdleProcess.Items.Add(oListItem);
                        oListItem = new ListItem_IdleDateWise(_drIdleSummary, iRowCount++);
                        lvIdleProcess.Items.Add(oListItem);

                        PreviousDate = CurrentDate;
                        TotalMinutesIdle = 0;
                    }

                    TotalMinutesIdle += Convert.ToInt32(_drIdleSummary["TotalIdle"].ToString());
                }


                oListItem = new ListItem_IdleDateWise(TotalMinutesIdle.ToString());
                lvIdleProcess.Items.Add(oListItem);

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lvIdleProcess);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void Get_Employee_Alloted_Files()
        {
            try
            {
                Offline_MT_Tracking_EmployeeList lsvEmp = (Offline_MT_Tracking_EmployeeList)lvRepEmp.SelectedItems[0];
                DataSet _dsAllotDetails = new DataSet();
                _dsAllotDetails = BusinessLogic.WS_Allocation.Get_alloted_Details(Convert.ToInt32(lsvEmp.EMP_PRODUCTION_ID));
                Employee_File_Alloted_Details oListItem;

                lvFileAllotedStatus.Items.Clear();
                int iRowCount = 1;
                int dTotMins = 0;

                if (cbxReportDesig.Text == "MT")
                {
                    foreach (DataRow _drAllFiles in _dsAllotDetails.Tables[0].Select())
                    {
                        lvFileAllotedStatus.Items.Add(new Employee_File_Alloted_Details(_drAllFiles, iRowCount++));
                        dTotMins += Convert.ToInt32(_drAllFiles["FileTot"].ToString());
                    }
                    string oMins = sGetDuration(dTotMins);

                    oListItem = new Employee_File_Alloted_Details("Total Minutes" + "Alloted: ", oMins.ToString());
                    lvFileAllotedStatus.Items.Add(oListItem);

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lvFileAllotedStatus);
                }
                else
                {
                    foreach (DataRow _drAllFiles in _dsAllotDetails.Tables[1].Select())
                    {
                        lvFileAllotedStatus.Items.Add(new Employee_File_Alloted_Details(_drAllFiles, iRowCount++));
                        dTotMins += Convert.ToInt32(_drAllFiles["FileTot"].ToString());
                    }
                    string oMins = sGetDuration(dTotMins);

                    oListItem = new Employee_File_Alloted_Details("Total Minutes" + "Alloted: ", oMins.ToString());
                    lvFileAllotedStatus.Items.Add(oListItem);

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lvFileAllotedStatus);
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void Load_Tat_Report()
        {
            int iOnTat = 1;
            int iOffTat = 1;
            int Rows = 1;

            lblTotDownloaded.Text = string.Empty;
            lblOnTat.Text = string.Empty;
            lblOffTat.Text = string.Empty;

            BusinessLogic.oProgressEvent.Start(true);
            try
            {
                string sDate = dtpTAT.Value.ToString("yyyy-MM-dd");
                string sLocationID = cmbLocationTat.SelectedValue.ToString();
                DataSet _dsTatStatus = new DataSet();
                _dsTatStatus = BusinessLogic.WS_Allocation.Get_Tat_Status(sDate, sLocationID);

                lsvTatStatus.Items.Clear();
                int iRowCount = 0;
                foreach (DataRow _drTat in _dsTatStatus.Tables[0].Select())
                {
                    lblTotDownloaded.Text = Rows.ToString();
                    lblTotDownloaded.ForeColor = System.Drawing.Color.Red;
                    lblOnTat.ForeColor = System.Drawing.Color.Red;
                    lblOffTat.ForeColor = System.Drawing.Color.Red;
                    if (_drTat["Status"].ToString() == "On-TAT")
                    {
                        lblOnTat.Text = iOnTat++.ToString();
                    }
                    else if (_drTat["Status"].ToString() == "Out of TAT")
                    {
                        lblOffTat.Text = iOffTat++.ToString();
                    }
                    lsvTatStatus.Items.Add(new List_Item_TATStatus(_drTat, iRowCount++));
                    Rows++;
                }

                BusinessLogic.Reset_ListViewColumn(lsvTatStatus);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void Load_EmpName_List(string BatchID, string BranchID)
        {
            int iRowcount = 0;
            int iWork_Type = 0;
            string sFilter = string.Empty;
            BusinessLogic.oProgressEvent.Start(true);
            try
            {
                if (BusinessLogic.USERNAME == "Admin-Trivandrum")
                {
                    if (tabControlMain.SelectedTab.Name == "tabPageOffline")
                        iWork_Type = 2;
                    else
                        iWork_Type = 1;

                    WORK_PLATFORM = iWork_Type;

                    sFilter = " Work_platform=" + iWork_Type + "";
                    DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_Desigwise_employees(Convert.ToInt32(BatchID));

                    DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                             where dr["branch_id"].ToString() == "3"
                                             select dr).CopyToDataTable();

                    lsv_Offline_Users.Items.Clear();
                    lsv_Target_NamesList.Items.Clear();  // 
                    foreach (DataRow dr in _dtWithLinq.Select(sFilter))
                    {
                        if (iWork_Type == 2)
                            lsv_Offline_Users.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                        else
                            lsv_Target_NamesList.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Target_NamesList);
                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Users);
                }
                else if (BusinessLogic.USERNAME == "Admin-Cochin")
                {
                    if (tabControlMain.SelectedTab.Name == "tabPageOffline")
                        iWork_Type = 2;
                    else
                        iWork_Type = 1;

                    WORK_PLATFORM = iWork_Type;

                    sFilter = " Work_platform=" + iWork_Type + "";
                    DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_Desigwise_employees(Convert.ToInt32(BatchID));

                    DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                             where dr["branch_id"].ToString() == "2"
                                             select dr).CopyToDataTable();

                    lsv_Offline_Users.Items.Clear();
                    lsv_Target_NamesList.Items.Clear();
                    foreach (DataRow dr in _dtWithLinq.Select(sFilter))
                    {
                        if (iWork_Type == 2)
                            lsv_Offline_Users.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                        else
                            lsv_Target_NamesList.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Target_NamesList);
                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Users);
                }
                else if (BusinessLogic.USERNAME == "Admin-Pondichery")
                {
                    if (tabControlMain.SelectedTab.Name == "tabPageOffline")
                        iWork_Type = 2;
                    else
                        iWork_Type = 1;

                    WORK_PLATFORM = iWork_Type;

                    sFilter = " Work_platform=" + iWork_Type + "";
                    DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_Desigwise_employees(Convert.ToInt32(BatchID));

                    DataTable _dtWithLinq = (from DataRow dr in _dsEmployee.Tables[0].Select()
                                             where dr["branch_id"].ToString() == "4"
                                             select dr).CopyToDataTable();

                    lsv_Offline_Users.Items.Clear();
                    lsv_Target_NamesList.Items.Clear();
                    foreach (DataRow dr in _dtWithLinq.Select(sFilter))
                    {
                        if (iWork_Type == 2)
                            lsv_Offline_Users.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                        else
                            lsv_Target_NamesList.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Target_NamesList);
                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Users);
                }
                else
                {
                    if (tabControlMain.SelectedTab.Name == "tabPageOffline")
                        iWork_Type = 2;
                    else
                        iWork_Type = 1;

                    WORK_PLATFORM = iWork_Type;

                    sFilter = " Work_platform=" + iWork_Type + "";
                    //DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_Desigwise_employees(Convert.ToInt32(BatchID));  Get_branch_Desigwise_employees
                    DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_branch_Desigwise_employees(Convert.ToInt32(BatchID), Convert.ToInt32(BranchID));
                    lsv_Offline_Users.Items.Clear();
                    lsv_Target_NamesList.Items.Clear();
                    foreach (DataRow dr in _dsEmployee.Tables[0].Select(sFilter))
                    {
                        if (iWork_Type == 2)
                            lsv_Offline_Users.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                        else
                            lsv_Target_NamesList.Items.Add(new ListItem_MTMEList(dr, iRowcount++));
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_Target_NamesList);
                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Users);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void Load_Target_Details()
        {
            try
            {
                int iRow = 0;
                string iFromHour, iToHour;
                object oFromDate = null;
                object oTodate = null;
                string sHourlyFromDate = string.Empty;
                string sHourlyTodate = string.Empty;
                string sStart_Time = string.Empty;
                string sEnd_Time = string.Empty;
                int IS_HT = 0;
                int IS_NIGHT = 0;
                int USER_TARGET = 0;
                int USER_TARGET_SEC = 0;

                ListItem_TargetDetails oListItem;
                List<BusinessLogic.TARGET_DETAILS> oTarget = new List<BusinessLogic.TARGET_DETAILS>();
                DateTime Display_date = DateTime.Now;

                int Tot_Achieve = 0;
                int Tot_Bal = 0;
                string sBal_Mins = string.Empty;
                decimal iCompleted_Mins = 0;
                decimal iAchievedLines = 0;
                decimal iTotalAchievedLines = 0;
                string sDetails = string.Empty;

                if (WORK_PLATFORM == 1)
                {
                    sHourlyFromDate = Convert.ToDateTime(dtp_Target_Fromdate.Value).ToString("yyyy-MM-dd");
                    sHourlyTodate = Convert.ToDateTime(dtp_Target_Todate.Value).ToString("yyyy-MM-dd");

                    sStart_Time = txt_StartingTime.Text.ToString();
                    sEnd_Time = txt_Endingtime.Text.ToString();

                    string[] sSplit_From = sStart_Time.Split(':');
                    string[] sSplit_To = sEnd_Time.Split(':');

                    iFromHour = sSplit_From[0].ToString();
                    iToHour = sSplit_To[0].ToString();

                    if (Convert.ToInt32(iFromHour) == 24)
                        oFromDate = sHourlyFromDate + " 23:59:59";
                    else
                        oFromDate = sHourlyFromDate + " " + iFromHour + ":" + "00:00";

                    if (Convert.ToInt32(iToHour) == 24)
                        oTodate = sHourlyTodate + " 23:59:59";
                    else
                        oTodate = sHourlyTodate + " " + iToHour + ":" + "00:00";

                    if (Convert.ToInt32(iFromHour) > Convert.ToInt32(iToHour))
                        IS_NIGHT = 1;
                    else
                        IS_NIGHT = 0;

                    foreach (ListItem_MTMEList oItem in lsv_Target_NamesList.SelectedItems)
                    {
                        DataSet dsTarget_Report = BusinessLogic.WS_Allocation.Get_Target_Details(Convert.ToInt32(oItem.IPRODUCTIONID), Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate));
                        if (dsTarget_Report != null)
                        {
                            if (dsTarget_Report.Tables[0].Rows.Count > 0)
                            {
                                lsv_Target_Report.Items.Clear();
                                lsv_Offline_Target.Items.Clear();

                                foreach (DataRow dr in dsTarget_Report.Tables[0].Rows)
                                {
                                    IS_HT = Convert.ToInt32(dr["is_ht_user"].ToString());
                                    //IS_NIGHT = Convert.ToInt32(dr["isNightShift"].ToString());
                                    USER_TARGET = Convert.ToInt32(dr["target_mins"].ToString());
                                    USER_TARGET_SEC = Convert.ToInt32(dr["target_secs"].ToString());
                                    iAchievedLines = Convert.ToDecimal(dr["total_lines"].ToString());
                                    //iTotalAchievedLines += Convert.ToInt32(iAchievedLines);
                                    sDetails = dr["details"].ToString();
                                    oTarget.Add(new BusinessLogic.TARGET_DETAILS(Convert.ToDateTime(dr["submitted_time"].ToString()).ToString("yyyy-MM-dd HH:mm"), Convert.ToInt32(dr["transcription_id"].ToString()), dr["file_minutes"].ToString(), Convert.ToInt32(dr["Conv_File_Mins"].ToString()), string.Empty, string.Empty, Convert.ToDecimal(dr["total_lines"].ToString()), sDetails));
                                }

                                //if (IS_HT == 0)
                                //{
                                //Display the details based on shift
                                if (IS_NIGHT == 0)
                                {
                                    for (var day = Convert.ToDateTime(sHourlyFromDate).Date; day.Date <= Convert.ToDateTime(sHourlyTodate).Date; day = day.AddDays(1))
                                    {
                                        iTotalAchievedLines = 0;

                                        DateTime end = Convert.ToDateTime(day);
                                        DateTime start = Convert.ToDateTime(day);

                                        if (Convert.ToInt32(iToHour) == 24)
                                            oTodate = " 23:59:59";
                                        else
                                            oTodate = iToHour + ":" + "00:00";

                                        object oStart = start.ToString("yyyy-MM-dd") + " " + iFromHour + ":" + "00:00";
                                        object oEnd = end.ToString("yyyy-MM-dd") + " " + oTodate;

                                        var Tot_Files = (from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) select c).Count();
                                        var Tot_Achieved_Mins = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVED_MINS into CP select new { TOTAL_ACHIEVED_MINS = CP.Sum(c => Convert.ToInt32(c.SACHIEVED_MINS)) };
                                        var Tot_Achieved_Lines = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVEDLINES into CP select new { TOTAL_ACHIEVED_LINES = CP.Sum(c => Convert.ToInt32(c.SACHIEVEDLINES)) };

                                        foreach (var ACHIEVED_Mins in Tot_Achieved_Mins)
                                        {
                                            Tot_Achieve += Convert.ToInt32(ACHIEVED_Mins.TOTAL_ACHIEVED_MINS);
                                        }

                                        foreach (var ACHIEVED_Line in Tot_Achieved_Lines)
                                        {
                                            iTotalAchievedLines += Convert.ToInt32(ACHIEVED_Line.TOTAL_ACHIEVED_LINES);
                                        }

                                        Tot_Bal = Tot_Achieve - USER_TARGET_SEC;
                                        bool bPos = Tot_Bal > 0;
                                        if (bPos == false)
                                        {
                                            int iBal = Math.Abs(Convert.ToInt32(Tot_Bal));
                                            sBal_Mins = "-" + sGetDuration(Convert.ToInt32(iBal));
                                        }
                                        else
                                            sBal_Mins = sGetDuration(Convert.ToInt32(Tot_Bal));

                                        iCompleted_Mins = Math.Round(((Convert.ToDecimal(Tot_Achieve) / Convert.ToDecimal(USER_TARGET_SEC)) * 100), 2);
                                        oListItem = new ListItem_TargetDetails(iRow++, Convert.ToDateTime(end).ToString("yyyy-MM-dd"), Tot_Files.ToString(), USER_TARGET.ToString(), sGetDuration(Tot_Achieve).ToString(), sBal_Mins, iCompleted_Mins.ToString() + "%", iTotalAchievedLines.ToString(), sDetails);
                                        lsv_Target_Report.Items.Add(oListItem);
                                        Tot_Achieve = 0;
                                        Tot_Bal = 0;
                                        BusinessLogic.Reset_ListViewColumn(lsv_Target_Report);
                                    }
                                }
                                else
                                {
                                    // FOR NIGHT SHIFT ENTRIES
                                    for (var day = Convert.ToDateTime(sHourlyFromDate).Date; day.Date <= Convert.ToDateTime(sHourlyTodate).Date; day = day.AddDays(1))
                                    {
                                        iTotalAchievedLines = 0;

                                        DateTime end = Convert.ToDateTime(day).AddDays(1);
                                        DateTime start = Convert.ToDateTime(day);

                                        if (Convert.ToInt32(iToHour) == 24)
                                            oTodate = " 23:59:59";
                                        else
                                            oTodate = iToHour + ":" + "00:00";

                                        object oStart = start.ToString("yyyy-MM-dd") + " " + iFromHour + ":" + "00:00";
                                        object oEnd = end.ToString("yyyy-MM-dd") + " " + oTodate;

                                        var Tot_Files = (from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) select c).Count();
                                        var Tot_Achieved_Mins = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVED_MINS into CP select new { TOTAL_ACHIEVED_MINS = CP.Sum(c => Convert.ToInt32(c.SACHIEVED_MINS)) };
                                        var Tot_Achieved_Lines = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVEDLINES into CP select new { TOTAL_ACHIEVED_LINES = CP.Sum(c => Convert.ToInt32(c.SACHIEVEDLINES)) };

                                        foreach (var ACHIEVED_Mins in Tot_Achieved_Mins)
                                        {
                                            Tot_Achieve += Convert.ToInt32(ACHIEVED_Mins.TOTAL_ACHIEVED_MINS);
                                        }
                                        foreach (var ACHIEVED_Line in Tot_Achieved_Lines)
                                        {
                                            iTotalAchievedLines += Convert.ToInt32(ACHIEVED_Line.TOTAL_ACHIEVED_LINES);
                                        }
                                        iTotalAchievedLines += Convert.ToInt32(iAchievedLines);
                                        Tot_Bal = Tot_Achieve - USER_TARGET_SEC;
                                        bool bPos = Tot_Bal > 0;
                                        if (bPos == false)
                                        {
                                            int iBal = Math.Abs(Convert.ToInt32(Tot_Bal));
                                            sBal_Mins = "-" + sGetDuration(Convert.ToInt32(iBal));
                                        }
                                        else
                                            sBal_Mins = sGetDuration(Convert.ToInt32(Tot_Bal));

                                        iCompleted_Mins = Math.Round(((Convert.ToDecimal(Tot_Achieve) / Convert.ToDecimal(USER_TARGET_SEC)) * 100), 2);

                                        oListItem = new ListItem_TargetDetails(iRow++, Convert.ToDateTime(start).ToString("yyyy-MM-dd"), Tot_Files.ToString(), USER_TARGET.ToString(), sGetDuration(Tot_Achieve).ToString(), sBal_Mins, iCompleted_Mins.ToString() + "%", iTotalAchievedLines.ToString(), sDetails);
                                        lsv_Target_Report.Items.Add(oListItem);
                                        Tot_Achieve = 0;
                                        Tot_Bal = 0;
                                        BusinessLogic.Reset_ListViewColumn(lsv_Target_Report);
                                    }
                                }
                            }
                            else
                            {
                                lsv_Target_Report.Items.Clear();
                            }
                        }
                    }
                }
                else
                {
                    sHourlyFromDate = Convert.ToDateTime(dtp_Offline_Fromdate.Value).ToString("yyyy-MM-dd");
                    sHourlyTodate = Convert.ToDateTime(dtp_Offline_Todate.Value).ToString("yyyy-MM-dd");

                    iFromHour = "0";
                    iToHour = "24";

                    if (Convert.ToInt32(iFromHour) == 24)
                        oFromDate = sHourlyFromDate + " 23:59:59";
                    else
                        oFromDate = sHourlyFromDate + " " + iFromHour + ":" + "00:00";

                    if (Convert.ToInt32(iToHour) == 24)
                        oTodate = sHourlyTodate + " 23:59:59";
                    else
                        oTodate = sHourlyTodate + " " + iToHour + ":" + "00:00";

                    if (Convert.ToInt32(iFromHour) > Convert.ToInt32(iToHour))
                        IS_NIGHT = 1;
                    else
                        IS_NIGHT = 0;


                    foreach (ListItem_MTMEList oItem in lsv_Offline_Users.SelectedItems)
                    {
                        DataSet dsTarget_Report = BusinessLogic.WS_Allocation.Get_Target_Details(Convert.ToInt32(oItem.IPRODUCTIONID), Convert.ToDateTime(dtp_Offline_Fromdate.Value), Convert.ToDateTime(dtp_Offline_Todate.Value));

                        if (dsTarget_Report != null)
                        {
                            if (dsTarget_Report.Tables[0].Rows.Count > 0)
                            {
                                lsv_Offline_Target.Items.Clear();
                                //foreach (DataRow dr in dsTarget_Report.Tables[0].Rows)
                                //    lsv_Offline_Target.Items.Add(new ListItem_TargetDetails(dr, iRow++));
                                //--------------------------------------------------------------
                                foreach (DataRow dr in dsTarget_Report.Tables[0].Rows)
                                {
                                    IS_HT = Convert.ToInt32(dr["is_ht_user"].ToString());
                                    //IS_NIGHT = Convert.ToInt32(dr["isNightShift"].ToString());
                                    USER_TARGET = Convert.ToInt32(dr["target_mins"].ToString());
                                    USER_TARGET_SEC = Convert.ToInt32(dr["target_secs"].ToString());
                                    oTarget.Add(new BusinessLogic.TARGET_DETAILS(Convert.ToDateTime(dr["submitted_time"].ToString()).ToString("yyyy-MM-dd HH:mm"), Convert.ToInt32(dr["transcription_id"].ToString()), dr["file_minutes"].ToString(), Convert.ToInt32(dr["Conv_File_Mins"].ToString()), string.Empty, string.Empty, Convert.ToInt32(dr["total_lines"].ToString()), dr["details"].ToString()));
                                }

                                for (var day = Convert.ToDateTime(sHourlyFromDate).Date; day.Date <= Convert.ToDateTime(sHourlyTodate).Date; day = day.AddDays(1))
                                {
                                    iTotalAchievedLines = 0;

                                    DateTime end = Convert.ToDateTime(day);
                                    DateTime start = Convert.ToDateTime(day);

                                    if (Convert.ToInt32(iToHour) == 24)
                                        oTodate = " 23:59:59";
                                    else
                                        oTodate = iToHour + ":" + "00:00";

                                    object oStart = start.ToString("yyyy-MM-dd") + " " + iFromHour + ":" + "00:00";
                                    object oEnd = end.ToString("yyyy-MM-dd") + " " + oTodate;

                                    var Tot_Files = (from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) select c).Count();
                                    var Tot_Achieved_Mins = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVED_MINS into CP select new { TOTAL_ACHIEVED_MINS = CP.Sum(c => Convert.ToInt32(c.SACHIEVED_MINS)) };
                                    var Tot_Achieved_Lines = from c in oTarget where (Convert.ToDateTime(c.DTARGET_DATE) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.DTARGET_DATE) <= Convert.ToDateTime(oEnd)) group c by c.SACHIEVEDLINES into CP select new { TOTAL_ACHIEVED_LINES = CP.Sum(c => Convert.ToInt32(c.SACHIEVEDLINES)) };

                                    foreach (var ACHIEVED_Mins in Tot_Achieved_Mins)
                                    {
                                        Tot_Achieve += Convert.ToInt32(ACHIEVED_Mins.TOTAL_ACHIEVED_MINS);                                        
                                    }
                                    foreach (var ACHIEVED_Line in Tot_Achieved_Lines)
                                    {
                                        iTotalAchievedLines += Convert.ToInt32(ACHIEVED_Line.TOTAL_ACHIEVED_LINES);
                                    }
                                    iTotalAchievedLines += Convert.ToInt32(iAchievedLines);
                                    Tot_Bal = Tot_Achieve - USER_TARGET_SEC;
                                    bool bPos = Tot_Bal > 0;
                                    if (bPos == false)
                                    {
                                        int iBal = Math.Abs(Convert.ToInt32(Tot_Bal));
                                        sBal_Mins = "-" + sGetDuration(Convert.ToInt32(iBal));
                                    }
                                    else
                                        sBal_Mins = sGetDuration(Convert.ToInt32(Tot_Bal));

                                    iCompleted_Mins = Math.Round(((Convert.ToDecimal(Tot_Achieve) / Convert.ToDecimal(USER_TARGET_SEC)) * 100), 2);

                                    oListItem = new ListItem_TargetDetails(iRow++, Convert.ToDateTime(end).ToString("yyyy-MM-dd"), Tot_Files.ToString(), USER_TARGET.ToString(), sGetDuration(Tot_Achieve).ToString(), sBal_Mins, iCompleted_Mins.ToString() + "%", iTotalAchievedLines.ToString(), sDetails);
                                    lsv_Offline_Target.Items.Add(oListItem);
                                    Tot_Achieve = 0;
                                    Tot_Bal = 0;
                                    BusinessLogic.Reset_ListViewColumn(lsv_Offline_Target);
                                }


                                //----------------------------------------------------------------

                                BusinessLogic.Reset_ListViewColumn(lsv_Offline_Target);
                            }
                            else
                            {
                                lsv_Target_Report.Items.Clear();
                                lsv_Offline_Target.Items.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //_Exception.WriteLog(ex);
            }
        }

        private void Load_Old_Target_Details()
        {
            try
            {
                int iRow = 0;

                if (WORK_PLATFORM == 1)
                {
                    object oFromDate = null;
                    object oTodate = null;
                    string iFromHour, iToHour;
                    object CurrentDate = null;
                    object PreviousDate = null;
                    object StartingDate = null;
                    ListItem_TargetDetails oListItem;

                    decimal iAchieved_Mins = 0;
                    decimal iBal_Mins = 0;
                    decimal iCompleted_Mins = 0;
                    int iTot_Files = 0;
                    int iTarget_Mins = 0;
                    decimal iTarget_Secs = 0;
                    string sDetails = string.Empty;

                    string sHourlyFromDate = Convert.ToDateTime(dtp_Target_Fromdate.Value).ToString("yyyy-MM-dd");
                    string sHourlyTodate = Convert.ToDateTime(dtp_Target_Todate.Value).ToString("yyyy-MM-dd");

                    string sStart_Time = txt_StartingTime.Text.ToString();
                    string sEnd_Time = txt_Endingtime.Text.ToString();

                    string[] sSplit_From = sStart_Time.Split(':');
                    string[] sSplit_To = sEnd_Time.Split(':');

                    iFromHour = sSplit_From[0].ToString();
                    iToHour = sSplit_To[0].ToString();

                    if (Convert.ToInt32(iFromHour) == 24)
                        oFromDate = sHourlyFromDate + " 23:59:59";
                    else
                        oFromDate = sHourlyFromDate + " " + iFromHour + ":" + "00:00";

                    if (Convert.ToInt32(iToHour) == 24)
                        oTodate = sHourlyTodate + " 23:59:59";
                    else
                        oTodate = sHourlyTodate + " " + iToHour + ":" + "00:00";

                    foreach (ListItem_MTMEList oItem in lsv_Target_NamesList.SelectedItems)
                    {
                        //DataSet dsTarget_Report = BusinessLogic.WS_Allocation.Get_Target_Details(Convert.ToInt32(oItem.IPRODUCTIONID), Convert.ToDateTime(dtp_Target_Fromdate.Value), Convert.ToDateTime(dtp_Target_Todate.Value));
                        DataSet dsTarget_Report = BusinessLogic.WS_Allocation.Get_Target_Details(Convert.ToInt32(oItem.IPRODUCTIONID), Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate));
                        if (dsTarget_Report != null)
                        {
                            if (dsTarget_Report.Tables[0].Rows.Count > 0)
                            {
                                lsv_Target_Report.Items.Clear();
                                lsv_Offline_Target.Items.Clear();

                                TimeSpan oStart_Time = new TimeSpan(Convert.ToInt32(iFromHour), 0, 0);
                                TimeSpan oTo_Time = new TimeSpan(Convert.ToInt32(iToHour), 0, 0);


                                var dates = new List<DateTime>();
                                for (var dt = Convert.ToDateTime(sHourlyFromDate); dt <= Convert.ToDateTime(sHourlyTodate); dt = dt.AddDays(1))
                                {
                                    PreviousDate = null;
                                    dates.Add(dt);

                                    if (Convert.ToInt32(iFromHour) > Convert.ToInt32(iToHour))
                                    {
                                        dt.AddDays(1);
                                    }

                                    string sdTodate = dt.ToString("yyyy-MM-dd " + iToHour + ":00:00");

                                    //WORK_PLATFORM  -- Radhika                                    

                                    string dtFromDate = dt.ToString("yyyy-MM-dd " + iFromHour + ":00:00");
                                    string dtTodate = string.Empty;


                                    if (dt != Convert.ToDateTime(sHourlyTodate))
                                    {
                                        if (Convert.ToInt32(iToHour) == 24)
                                            dtTodate = dt.AddDays(1).ToString("yyyy-MM-dd " + " 23:59:59");
                                        else
                                            dtTodate = dt.AddDays(1).ToString("yyyy-MM-dd " + iToHour + ":00:00");
                                    }
                                    else
                                    {
                                        if (Convert.ToInt32(iToHour) == 24)
                                            dtTodate = dt.ToString("yyyy-MM-dd " + " 23:59:59");
                                        else
                                            dtTodate = dt.ToString("yyyy-MM-dd " + iToHour + ":00:00");
                                    }

                                    foreach (DataRow dr in dsTarget_Report.Tables[0].Rows)
                                    {
                                        iTarget_Mins = Convert.ToInt32(dr["target_mins"].ToString());
                                        iTarget_Secs = Convert.ToDecimal(dr["target_secs"].ToString());
                                        TimeSpan oSubmitTime = new TimeSpan(Convert.ToDateTime(dr["submitted_time"].ToString()).Hour, 0, 0);
                                        string dSubmit_Time = (Convert.ToDateTime(dr["submitted_time"].ToString()).ToString("yyyy-MM-dd H:" + "00:00"));

                                        //if ((oSubmitTime > oStart_Time))
                                        if ((Convert.ToDateTime(dSubmit_Time) > Convert.ToDateTime(dtFromDate)) && (Convert.ToDateTime(dSubmit_Time) < Convert.ToDateTime(dtTodate)))
                                        {
                                            CurrentDate = BusinessLogic.ConvertToDateTime(dr["" + Framework.TRANSCRIPTIONTRANSACTION.FIELD_SUBMITTED_TIME + ""]);
                                            int iDaydifferece = Convert.ToDateTime(CurrentDate).Day - Convert.ToDateTime(PreviousDate).Day;
                                            if (PreviousDate == null)
                                            {
                                                iTot_Files += 1;
                                                iAchieved_Mins += Convert.ToInt32(dr["Conv_File_Mins"].ToString());
                                                StartingDate = CurrentDate;
                                                PreviousDate = CurrentDate;
                                            }
                                            else
                                            {
                                                iTot_Files += 1;
                                                iAchieved_Mins += Convert.ToInt32(dr["Conv_File_Mins"].ToString());
                                                PreviousDate = CurrentDate;
                                            }
                                        }
                                    }
                                    if (PreviousDate == CurrentDate)
                                    {
                                        string sBal_Mins = string.Empty;
                                        iBal_Mins = (iAchieved_Mins - Convert.ToDecimal(iTarget_Secs));
                                        bool bPos = iBal_Mins > 0;
                                        if (bPos == false)
                                        {
                                            int iBal = Math.Abs(Convert.ToInt32(iBal_Mins));
                                            sBal_Mins = "-" + sGetDuration(Convert.ToInt32(iBal));
                                        }
                                        else
                                            sBal_Mins = sGetDuration(Convert.ToInt32(iBal_Mins));

                                        iCompleted_Mins = Math.Round(((Convert.ToDecimal(iAchieved_Mins) / iTarget_Secs) * 100), 2);

                                        oListItem = new ListItem_TargetDetails(iRow++, Convert.ToDateTime(StartingDate.ToString()).ToString("yyyy-MM-dd"), iTot_Files.ToString(), iTarget_Mins.ToString(), sGetDuration(Convert.ToInt32(iAchieved_Mins)), sBal_Mins, iCompleted_Mins.ToString(), iAchieved_Mins.ToString() + "%", sDetails);
                                        lsv_Target_Report.Items.Add(oListItem);

                                        PreviousDate = CurrentDate;
                                        iTot_Files = 0;
                                        iAchieved_Mins = 0;
                                        iCompleted_Mins = 0;
                                    }

                                    //End for loop
                                }




                                BusinessLogic.Reset_ListViewColumn(lsv_Target_Report);
                            }

                            else
                            {
                                lsv_Target_Report.Items.Clear();
                                lsv_Offline_Target.Items.Clear();
                            }
                        }
                        else
                        {
                            lsv_Target_Report.Items.Clear();
                            lsv_Offline_Target.Items.Clear();
                        }
                    }
                }
                else
                {
                    foreach (ListItem_MTMEList oItem in lsv_Offline_Users.SelectedItems)
                    {
                        DataSet dsTarget_Report = BusinessLogic.WS_Allocation.Get_Target_Details(Convert.ToInt32(oItem.IPRODUCTIONID), Convert.ToDateTime(dtp_Offline_Fromdate.Value), Convert.ToDateTime(dtp_Offline_Todate.Value));
                        if (dsTarget_Report != null)
                        {
                            if (dsTarget_Report.Tables[0].Rows.Count > 0)
                            {
                                lsv_Offline_Target.Items.Clear();
                                foreach (DataRow dr in dsTarget_Report.Tables[0].Rows)
                                    lsv_Offline_Target.Items.Add(new ListItem_TargetDetails(dr, iRow++));

                                BusinessLogic.Reset_ListViewColumn(lsv_Offline_Target);
                            }
                            else
                            {
                                lsv_Target_Report.Items.Clear();
                                lsv_Offline_Target.Items.Clear();
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //_Exception.WriteLog(ex);
            }
        }

        private void LoadEmployeeForHold()
        {
            try
            {
                DataSet _ds = new DataSet();
                _ds = BusinessLogic.WS_Allocation.Get_EmployeeList(-1);

                lvAddHEmployee.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drRow in _ds.Tables[0].Select())
                    lvAddHEmployee.Items.Add(new List_Item_List_Hold_Employee(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvAddHEmployee);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
            }
        }

        private void LoadEmployeesMoved()
        {
            try
            {
                DataSet _ds = new DataSet();
                _ds = BusinessLogic.WS_Allocation.Get_EmployeeList_TED_TO_ME();

                lvHoldEmployeees.Items.Clear();
                lvHoldEmployeees_ToMove.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drRow in _ds.Tables[0].Select())
                {
                    lvHoldEmployeees.Items.Add(new List_Item_List_Employee_Moved(_drRow, iRowCount++));
                }
                foreach (DataRow _drRow in _ds.Tables[0].Select())
                {
                    lvHoldEmployeees_ToMove.Items.Add(new List_Item_List_Employee_Moved(_drRow, iRowCount++));
                }

                BusinessLogic.Reset_ListViewColumn(lvHoldEmployeees);
                BusinessLogic.Reset_ListViewColumn(lvHoldEmployeees_ToMove);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
            finally
            {
            }
        }

        private void LoadHoldSlot()
        {
            try
            {
                DataSet _ds = new DataSet();
                _ds = BusinessLogic.WS_Allocation.Get_Hold_Slab();

                lvHoldSlab.Items.Clear();
                int iRowCount = 1;
                foreach (DataRow _drRow in _ds.Tables[0].Select())
                    lvHoldSlab.Items.Add(new List_Item_List_Hold_Slab(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvHoldSlab);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        #endregion " METHODS "

        #region " MENUS "

        /// <summary>
        /// Click the MT's list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_MTTrans_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 1, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.OFFLINE_P_FILEMINS_V2.ToString());
                        //insert into database             
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                        if (iResult > 1)
                        {

                        }
                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Trans_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// CHOOSE THE ME'S MENU IN THE TRANSCRIPTION TAB
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_METrans_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.ME), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.FILE_DURATION.ToString());
                        //insert into database             
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Trans_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// CHOOSE THE BA'S MENU IN THE TRANSCRIPTION RADIO BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_BATrans_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.BA), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 5, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.FILE_DURATION.ToString());
                        //insert into database             
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Trans_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// CHOOSE THE TED'S MENU IN THE TRANSCRIPTION RADIO BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_TEDTrans_Click(object sender, EventArgs e)
        {
            try
            {
                iTotalFiles = 0;
                double dTotmins = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.TED), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 3, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.FILE_DURATION.ToString());
                        //insert into database             
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Trans_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// CHOOSE THE TED'S MENU IN THE TRANSCRIPTION RADIO BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_MEEdit_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.ME), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 2, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Editing_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// CHOOSE THE AM'S MENU IN THE TRANSCRIPTION RADIO BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_AMEEdit_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.AM), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 4, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Editing_CheckedChanged(this, e);
            }
        }

        /// <summary>
        /// MARK DONE DURING EDIT 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_MarkDoneEdit_Click(object sender, EventArgs e)
        {
            //TAT REQ FOR EDITING

        }

        /// <summary>
        /// AM REVIEW DETAILS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Offline_AMReview_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 4, 3);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        //insert into database

                        string sVoiceFileID = oFile.OFFLINE_P_VOICE_FILE_NAME.ToString();
                        int sTransID;
                        //cbxOffline_Review
                        if (cbxOffline_Review.Checked == true)
                        {
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                            //sTransID = Convert.ToInt32(oItem.TRANSCRIPTIONID);
                            sTransID = Convert.ToInt32(oFile.TRANSCRIPTIONID);
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(sTransID, oFile.OFFLINE_P_USERID, DateTime.Now, sVoiceFileID, BusinessLogic.iTATREQUIRED);

                            oFile.OFFLINE_P_STATUS = "Allotted";
                            oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        }
                        else
                        {
                            // int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);


                            oFile.OFFLINE_P_STATUS = "Allotted";
                            oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        }
                    }
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Review_CheckedChanged(this, e);
            }
        }

        private void Offline_MEReview_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.ME), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 2, 3);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        //insert into database

                        string sVoiceFileID = oFile.OFFLINE_P_VOICE_FILE_NAME.ToString();
                        int sTransID;
                        if (cbxOffline_Review.Checked == true)
                        {
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                            sTransID = Convert.ToInt32(oFile.TRANSCRIPTIONID);
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(sTransID, oFile.OFFLINE_P_USERID, DateTime.Now, sVoiceFileID, BusinessLogic.iTATREQUIRED);

                            oFile.OFFLINE_P_STATUS = "Allotted";
                            oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        }
                        else
                        {
                            //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                            int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);


                            oFile.OFFLINE_P_STATUS = "Allotted";
                            oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                            oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        }
                    }
                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Review_CheckedChanged(this, e);
            }
        }


        private void Offline_MEEditReview_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 3, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }

                    int filecount = lsvFileDetails.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), "", Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_AMEditReview_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 3, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_METransEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);
                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);

                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_TEDTransEdit_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 3, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_MTAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), -1, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;


                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_TEDAll_Click(object sender, EventArgs e)
        {

            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 4, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_MEAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Offline_BAAll_Click(object sender, EventArgs e)
        {
            try
            {
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(iTotalMins), 2, 1);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;

                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    double dTotmins = 0;
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        #endregion " MENUS "

        #region " EVENTS "
        /// <summary>
        /// Load the offline datas
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //tabPageOffline
                if (tabControlMain.SelectedTab.Name == "tabPageOffline")
                {
                    //Get_Offline_AllocationDetails();
                    Load_Location_ListView();
                    Load_Offline_Designation();
                    cbxOffline_Trans.Checked = true;
                }
                else if (tabControlMain.SelectedTab.Name == "tbpHangUp")
                {
                    LoadYear_HangUP();
                    Load_Month_Name_HangUP();
                    Thread tHigherLines1 = new Thread(LoadHigherLines);
                    tHigherLines1.Start();
                }
                else if (tabControlMain.SelectedTab.Name == "tabPageEmployee")
                {
                    Load_Branch();
                    GetEmployeeList();
                }
                else if (tabControlMain.SelectedTab.Name == "tbpMapping")
                {                    
                    Load_Branch();
                    Load_Mapping();
                }
                else
                {
                    LoadYear_BPM();
                    Load_Month_Name_BPM();
                    LoadListView_Complaints();
                    Load_Location_BPM();
                    Load_Location_BPM_CLINICS();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        
        private void Load_Mapping()
        {
            try
            {
                string sBranch = cmbBranchMapp.SelectedValue.ToString();
                DataSet dsMapping = new DataSet();
                dsMapping = BusinessLogic.WS_Allocation.GET_NTS_MAPPING(Convert.ToInt32(sBranch));

                lvMapping.Items.Clear();
                int iRowCount = 1;
                
                foreach (DataRow _drRow in dsMapping.Tables[0].Rows)
                    lvMapping.Items.Add(new List_View_Employee(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lvMapping);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }           
        }

        /// <summary>
        /// View the designations as per the selection of radio buttons
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lsv_OfflieFile_Details_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (cbxOffline_Trans.Checked && cbxOffline_Review.Checked && cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_All;
                }
                else if (cbxOffline_Trans.Checked && !cbxOffline_Review.Checked && !cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_TransDesignation;
                }
                else if (!cbxOffline_Trans.Checked && !cbxOffline_Review.Checked && cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_EditDesignation;
                }
                else if (!cbxOffline_Trans.Checked && cbxOffline_Review.Checked && !cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_ReviewDesignation;
                }
                else if (!cbxOffline_Trans.Checked && cbxOffline_Review.Checked && cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_EditDesignation;
                }
                else if (cbxOffline_Trans.Checked && !cbxOffline_Review.Checked && cbxOffline_Editing.Checked)
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = Offline_TransAndEdit;
                }
                else
                {
                    lsv_OfflieFile_Details.ContextMenuStrip = null;
                }

                iTotalFiles = 0;
                iTotalMins = 0;
                if (iTotalFiles == 0 || iTotalFiles == 1)
                {
                    foreach (MyAllocation_Offline_List oItem in lsv_OfflieFile_Details.SelectedItems)
                    {
                        iTotalFiles++;
                        iTotalMins = iTotalMins + Convert.ToDouble(oItem.MINUTES);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// Offline selection changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxOffline_Trans_CheckedChanged(object sender, EventArgs e)
        {
            Load_Location_ListView();
            Mylsvdownloaddetails lsvIncAccount = null;
            try
            {
                if (lvJobAccount.SelectedItems.Count > 0)
                {
                    lsvIncAccount = (Mylsvdownloaddetails)lvJobAccount.SelectedItems[0];
                }
                else
                {
                    return;
                }

                lsv_OfflieFile_Details.Items.Clear();

                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_Offline_Allocation(1, null, null, null, lsvIncAccount.sLocation_id).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsv_OfflieFile_Details.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsv_OfflieFile_Details.Items.Add(new MyAllocation_Offline_List(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_OfflieFile_Details);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// CHANGE THE QUERY AS PER THE SELECTION OF EDITING
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxOffline_Editing_CheckedChanged(object sender, EventArgs e)
        {
            Load_Location_ListView();
            Mylsvdownloaddetails lsvIncAccount = null;
            try
            {
                if (lsvAccountName.SelectedItems.Count > 0)
                {
                    lsvIncAccount = (Mylsvdownloaddetails)lsvAccountName.SelectedItems[0];
                }

                lsv_OfflieFile_Details.Items.Clear();

                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_Offline_Allocation(2, null, null, null, lsvIncAccount.sLocation_id).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsv_OfflieFile_Details.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsv_OfflieFile_Details.Items.Add(new MyAllocation_Offline_List(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_OfflieFile_Details);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        /// <summary>
        /// GET THE REVIEW FILES
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxOffline_Review_CheckedChanged(object sender, EventArgs e)
        {
            Load_Location_ListView();
            Mylsvdownloaddetails lsvIncAccount = null;
            try
            {
                if (lsvAccountName.SelectedItems.Count > 0)
                {
                    lsvIncAccount = (Mylsvdownloaddetails)lsvAccountName.SelectedItems[0];
                }

                lsv_OfflieFile_Details.Items.Clear();

                DataTable dtTrans = BusinessLogic.WS_Allocation.Get_Offline_Allocation(3, null, null, null, lsvIncAccount.sLocation_id).Tables[0];
                if (dtTrans.Rows.Count > 0)
                {
                    int i = 1;
                    lsv_OfflieFile_Details.Items.Clear();
                    foreach (DataRow dr in dtTrans.Rows)
                    {
                        lsv_OfflieFile_Details.Items.Add(new MyAllocation_Offline_List(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsv_OfflieFile_Details);
                }
                else
                {
                    lsv_OfflieFile_Details.Items.Clear();
                    BusinessLogic.oMessageEvent.Start("No Items Found");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                sEditVoice = null;
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void btnAll_Offline_View_Click(object sender, EventArgs e)
        {
            if (cbxOffline_Trans.Checked == true)
            {
                cbxOffline_Trans_CheckedChanged(this, e);
                Application.DoEvents();
            }
            else if (cbxOffline_Editing.Checked == true)
            {
                cbxOffline_Editing_CheckedChanged(this, e);
                Application.DoEvents();
            }
            else if (cbxOffline_Review.Checked == true)
            {
                cbxOffline_Review_CheckedChanged(this, e);
                Application.DoEvents();
            }
        }

        private void lsv_Offline_Deall_Designation_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {
                if (e.Item.Selected)
                {
                    Thread NewThreadEmp = new Thread(Load_Employee_List);
                    NewThreadEmp.Start();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        /// <summary>
        /// GET THE REVIEW FILES
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lsv_Offline_Deall_EmplDetails_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            Load_AllotedFilesForEmployees_Offline();
        }

        private void btn_Deallocate_Click(object sender, EventArgs e)
        {
            try
            {
                if (lsv_Offline_Deallocate.CheckedItems.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No Item is selected for deallocation.");
                    return;
                }

                BusinessLogic.oMessageEvent.Start("Transferring data..");
                BusinessLogic.oProgressEvent.Start(true);
                this.Cursor = Cursors.WaitCursor;

                string _sTanscriptionCollection = string.Empty;
                string _sVoiceFile_ID = string.Empty;
                foreach (MyAllocation_Offline_AllotedFiles_List oDeAllocationItem in lsv_Offline_Deallocate.CheckedItems)
                {
                    _sTanscriptionCollection = oDeAllocationItem.TRANSCRIPTION_ID.ToString();
                    _sVoiceFile_ID = oDeAllocationItem.sVoiceFile_ID.ToString();

                    int iDeAllttotFiles = BusinessLogic.WS_Allocation.Set_Deallot_Files(Convert.ToInt32(_sTanscriptionCollection), _sVoiceFile_ID, Convert.ToInt32(oDeAllocationItem.ALLOTED_FOR));
                    if (iDeAllttotFiles > 0)
                    {
                        Load_AllotedFilesForEmployees_Offline();
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Failed");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                this.Cursor = Cursors.Default;
            }
        }

        private void tATREQToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TRANS TAT REQUIRED
        }

        private void Offline_MarkDoneReview_Click(object sender, EventArgs e)
        {
            //TRANS REQ FOR REVIEW
        }

        private void lsvOffAccounts_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            ListItem_OfflineAccountVolume oitem;
            oitem = (ListItem_OfflineAccountVolume)e.Item;
            Load_Offline_Volume_Location(oitem.client_id);
        }

        private void lsvAccountVolume_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            ListItem_OfflineLocationVolume oItem;
            oItem = (ListItem_OfflineLocationVolume)e.Item;
            Load_Offline_Volume_Doctor(oItem.location_id);

        }
        private void lsvDoctroWise_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            ListItem_OfflineDoctorVolume oItem;
            oItem = (ListItem_OfflineDoctorVolume)e.Item;
            Load_Offline_Pending_File(oItem.Doctor_id);
        }


        private void cbxReportDesig_SelectedIndexChanged(object sender, EventArgs e)
        {
            lvFileAllotedStatus.Items.Clear();
            Thread tTrack = new Thread(Load_MT_Tracking);
            tTrack.Start();
        }

        private void lvRepEmp_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Get_Employee_Alloted_Files();
            }
        }

        private void btnViewOffAcc_Click(object sender, EventArgs e)
        {
            lvOfflineAccountWise.Items.Clear();
            Thread tOffAccountWiseInfo = new Thread(Load_Offline_Account_Wise_Minutes);
            tOffAccountWiseInfo.Start();
        }

        private void tATRequiredToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Employee_File_Alloted_Details oItem = (Employee_File_Alloted_Details)lvFileAllotedStatus.SelectedItems[0];

                foreach (Employee_File_Alloted_Details oFile in lvFileAllotedStatus.SelectedItems)
                {
                    int iSetTat = BusinessLogic.WS_Allocation.set_tat(Convert.ToInt32(oFile.TRANSCRIPTION_ID));
                    if (iSetTat > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Tat Set..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void btnExportOffAcc_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Files Processed in Offline on date between_" + dateTimeUserFrom.Text + "_and_" + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lvOfflineAccountWise, sFolderNAme, sFileName);
        }

        private void btnViewBack_Click(object sender, EventArgs e)
        {
            lsvBackLockFiles.Items.Clear();
            Thread tOffAccountWiseInfo = new Thread(Load_Offline_Account_Wise_BackLock);
            tOffAccountWiseInfo.Start();
        }

        private void btnExportBack_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "BackLock FIles Between_" + dateTimeUserFrom.Text + "_and_" + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lsvBackLockFiles, sFolderNAme, sFileName);
        }

        private void tEDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                iTotalFiles = 0;
                int iResult = 0;
                MyAllocation_Offline_List oItem = (MyAllocation_Offline_List)lsv_OfflieFile_Details.SelectedItems[0];
                foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.FILE_DURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.TED), Convert.ToInt32(oItem.CLIENTID), Convert.ToInt32(oItem.DOCTORID), iTotalFiles, Convert.ToInt32(dTotmins), 3, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyAllocation_Offline_List oFile in lsv_OfflieFile_Details.SelectedItems)
                    {
                        oFile.OFFLINE_P_STATUS = "Allotted";
                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;
                        dTotmins += Convert.ToDouble(oFile.FILE_DURATION.ToString());

                        //insert into database             
                        //iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);

                        if (BusinessLogic.MTMET_BATCH_ID == 1)
                            iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_TedAssign_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, BusinessLogic.IS_TED_ASSIGN, Environment.UserName, Environment.MachineName);
                        else
                            iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_TedAssign_V2(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED, BusinessLogic.IS_TED_ASSIGN, Environment.UserName, Environment.MachineName);
                    }
                    int filecount = lsv_OfflieFile_Details.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                cbxOffline_Trans_CheckedChanged(this, e);
            }
        }

        private void btnOVExcel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Files Processed in Offline on date between_" + dateTimeUserFrom.Text + "_and_" + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lsvOffAccounts, sFolderNAme, sFileName);
        }

        private void btnMeBackLog_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Files Processed in Offline on date between_" + dateTimeUserFrom.Text + "_and_" + dateTimeUserTo.Text + ".xls";
            ExportToExcel(lvBackLocKME, sFolderNAme, sFileName);
        }

        private void dtpTAT_ValueChanged(object sender, EventArgs e)
        {
            Thread tTatStatus = new Thread(Load_Tat_Report);
            tTatStatus.Start();
        }

        private void cmbLocationTat_SelectedIndexChanged(object sender, EventArgs e)
        {
            Thread tTatStatusN = new Thread(Load_Tat_Report);
            tTatStatusN.Start();
        }



        private void btn_Report_Click(object sender, EventArgs e)
        {
            Get_Realloted_List();
            Load_Allocation_Status_Offline();
        }

        private void Reallot_toolStripMenuItem_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void lsv_Reallocation_MouseUp(object sender, MouseEventArgs e)
        {
            if (lsv_Reallocation.SelectedItems.Count == 1)
            {
                Reallocation_contextMenuStrip.Items["Reallot_toolStripMenuItem"].Enabled = true;
                lsv_Reallocation.ContextMenuStrip = Reallocation_contextMenuStrip;
            }
            else if (lsv_Reallocation.SelectedItems.Count == 0)
            {
                Reallocation_contextMenuStrip.Items["Reallot_toolStripMenuItem"].Enabled = false;
            }
        }

        private void Reallot_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (MyReallocation oItem in lsv_Reallocation.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Reallocationfiles(Convert.ToInt32(oItem.ITRANSCRIPTION_ID), Convert.ToInt32(oItem.IPRODUCTION_ID));
                    if (iUpdate > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Allocated Successfully to " + oItem.SPTAG_ID.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Job Details " + ".xls";
            ExportToExcel(lsv_OfflieFile_Details, sFolderNAme, sFileName);
        }

        private void btnLMView_Click(object sender, EventArgs e)
        {
            Thread tLargeMinutes = new Thread(Load_LargeMinutes);
            tLargeMinutes.Start();
        }



        private void cmb_Target_Desig_SelectedIndexChanged(object sender, EventArgs e)
        {
            sDesgination_ID = cmb_Target_Desig.SelectedValue.ToString();
            Load_EmpName_List(cmb_Target_Desig.SelectedValue.ToString(), sBranch_ID);
        }

        private void lsv_Target_NamesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Target_Details();
        }

        private void cmb_Offline_Desig_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_EmpName_List(cmb_Offline_Desig.SelectedValue.ToString(), sBranch_ID);
        }

        private void lsv_Offline_Users_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Target_Details();
        }

        private void cmbHoldPercentageMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sYear = cmbHoldPercentageYear.SelectedItem.ToString();
            int sMonth = cmbHoldPercentageMonth.SelectedIndex + 1;
            Load_Hold_Percentage(sMonth, sYear);
        }

        private void btnAutoAllocation_Click(object sender, EventArgs e)
        {
            try
            {
                string sLocationID = string.Empty;
                string sDoctorID = string.Empty;
                string sTotalFiles = string.Empty;
                string sBatchID = string.Empty;

                sLocationID = lbLocation.SelectedValue.ToString();
                sDoctorID = lbDoctor.SelectedValue.ToString();
                sTotalFiles = txtTotFiles.Text.ToString();
                sBatchID = lbDesig.SelectedValue.ToString();

                int iSetAccAll = BusinessLogic.WS_Allocation.Set_Acc_Auto_Allocation(Convert.ToInt32(sBatchID), sLocationID, Convert.ToInt32(sDoctorID), Convert.ToInt32(sTotalFiles));
                if (iSetAccAll > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Successfully Saved..!");
                    Load_Priority_Mapping();
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lbClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sClientID = string.Empty;
            sClientID = lbClient.SelectedValue.ToString();
            Load_Location(Convert.ToInt32(sClientID));
        }

        private void deActivateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (Mylistitem_AutoAllocationUsers oUser in lsvAutoAllocationUserDetails.SelectedItems)
                {

                    int iSetInActive = BusinessLogic.WS_Allocation.Set_Allocation_Priority_UnActive(oUser.iPriorityID, 1, oUser.iStatus);
                    if (iSetInActive > 0)
                    {
                        btnautoallocationView_Click(this, e);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Allocation_Status_Offline()
        {
            BusinessLogic.oMessageEvent.Start("Transferring data..!");
            try
            {
                lsvOfflineDetail.Items.Clear();
                DataTable _dtFiledetails = new DataTable();
                string sVoiceFile = string.Empty;
                if (txt_Reallot_Voice.Text == "")
                    BusinessLogic.oMessageEvent.Start("Enter Voice file id!");
                else
                {
                    sVoiceFile = txt_Reallot_Voice.Text.Trim();
                }

                _dtFiledetails = BusinessLogic.WS_Allocation.Get_Reallocatingfiles(sVoiceFile.Trim(), Convert.ToDateTime(dtpFrom_Date.Value)).Tables[1];

                int iRowCount = 1;
                foreach (DataRow _drAll in _dtFiledetails.Select())
                    lsvOfflineDetail.Items.Add(new ListItem_OfflineAllocationStatus(_drAll, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lsvOfflineDetail);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void btnViewIdle_Click(object sender, EventArgs e)
        {
            try
            {
                string sFromDate = dtpIdleFrom.Value.ToString("yyyy/MM/dd");
                string sToDate = dtpIdleTo.Value.ToString("yyyy/MM/dd");
                DataSet _dsIdleSummary = new DataSet();
                _dsIdleSummary = BusinessLogic.WS_Allocation.Get_Idle_Summary(sFromDate, sToDate);

                lvIdleSummary.Items.Clear();
                int iRowCount = 1;

                foreach (DataRow _drIdleSummary in _dsIdleSummary.Tables[0].Select())
                    lvIdleSummary.Items.Add(new ListItem_IdleSummary(_drIdleSummary, iRowCount++));

                BusinessLogic.oMessageEvent.Start("Ready.");
                BusinessLogic.Reset_ListViewColumn(lvIdleSummary);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Done....!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void lvIdleSummary_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                string sFromDate = dtpIdleFrom.Value.ToString("yyyy/MM/dd");
                string sToDate = dtpIdleTo.Value.ToString("yyyy/MM/dd");
                LoadIdleTimeDayWise(sFromDate, sToDate);
            }
        }

        private void tabEntriesAndLines_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabEntriesAndLines.SelectedTab.Name == "tbPageTedMeSalary")
            {
                LoadEmployeeForHold();
                LoadEmployeesMoved();
                LoadHoldSlot();
            }

            if (tabEntriesAndLines.SelectedTab.Name == "tbPageMTTedLines")
            {
                Load_MT_List();
                Load_MTMETList(1);
            }
        }

        private void btnAddEmp_Click(object sender, EventArgs e)
        {
            try
            {
                string sProductionID = string.Empty;
                foreach (List_Item_List_Hold_Employee oHselectedItem in lvAddHEmployee.SelectedItems)
                {
                    oHselectedItem.Selected = true;
                    sProductionID = oHselectedItem.iProductionID.ToString();

                    int iMove = BusinessLogic.WS_Allocation.set_move_employee(Convert.ToInt32(sProductionID));
                    if (iMove > 0)
                    {
                        LoadEmployeesMoved();
                        BusinessLogic.oMessageEvent.Start("Successfully Added.");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
            }
        }

        private void btnUpdateEmpMonth_Click(object sender, EventArgs e)
        {
            try
            {
                List_Item_List_Hold_Slab oHoldSlab = (List_Item_List_Hold_Slab)lvHoldSlab.SelectedItems[0];
                string sSlabID = oHoldSlab.sHoldSlabID.ToString();
                string sTransactionID = string.Empty;

                foreach (List_Item_List_Employee_Moved oHEmployeeMoved in lvHoldEmployeees_ToMove.SelectedItems)
                {
                    oHEmployeeMoved.Selected = true;
                    sTransactionID = oHEmployeeMoved.sHoldTransaction.ToString();

                    int iUpdate = BusinessLogic.WS_Allocation.Set_Update_Hold_month(Convert.ToInt32(sTransactionID), Convert.ToInt32(sSlabID));
                    if (iUpdate > 0)
                    {
                        LoadEmployeesMoved();
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void btnDeactivate_Click(object sender, EventArgs e)
        {
            try
            {
                string sUnique_id = string.Empty;

                foreach (List_Item_List_Employee_Moved oHEmployeeMovedNew in lvHoldEmployeees_ToMove.SelectedItems)
                {
                    oHEmployeeMovedNew.Selected = true;
                    sUnique_id = oHEmployeeMovedNew.sUnique_ID.ToString();

                    int iUpdate = BusinessLogic.WS_Allocation.Set_Deactivate_User(Convert.ToInt32(sUnique_id));
                    if (iUpdate > 0)
                    {
                        LoadEmployeesMoved();
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void tbMetAndTedSalary_Selecting(object sender, TabControlCancelEventArgs e)
        {
            LoadEmployeeForHold();
            LoadEmployeesMoved();
            LoadHoldSlot();
        }

        private void viewMappedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string sProduction_ID = string.Empty;
                ListIte_MappingDeatils_Priority lsvMap = (ListIte_MappingDeatils_Priority)lsvMappingDeatils.SelectedItems[0];
                sProduction_ID = lsvMap.PRODUCTION_ID.ToString();

                eAllocation.UI.frmAutoAllocationDetails fAutoAllocation = new eAllocation.UI.frmAutoAllocationDetails();
                fAutoAllocation._Productionid = sProduction_ID;
                fAutoAllocation.ShowDialog();

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
            }
        }


        private void trackBar_StartTime_ValueChanged(object sender, EventArgs e)
        {
            if (trackBar_StartTime.Value < 10)
                txt_StartingTime.Text = "0" + trackBar_StartTime.Value.ToString() + "00";
            else
                txt_StartingTime.Text = trackBar_StartTime.Value.ToString() + "00";
        }

        private void trackBar_Endtime_ValueChanged(object sender, EventArgs e)
        {
            if (trackBar_Endtime.Value < 10)
                txt_Endingtime.Text = "0" + trackBar_Endtime.Value.ToString() + "00";
            else
                txt_Endingtime.Text = trackBar_Endtime.Value.ToString() + "00";
        }

        /// <summary>
        /// VIEW CAPACITY DETAILS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ViewCapacity_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue) == 2)
                    pnl_Current_Online_All.Visible = true;
                else
                    pnl_Current_cap_offline.Visible = true;

                Load_Capacity_Leave(Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToInt32(cmb_Capacity_Branch.SelectedValue));
                Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                Load_Capacity_Percentage(Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value));
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void lsv_Capacity_MouseUp(object sender, MouseEventArgs e)
        {
            iCurrent = 0;
            iOverall = 1;

            if (lsv_Capacity.SelectedItems.Count == 1)
            {
                Capacity_Remove.Items["RemoveCapacity_ToolStripMenuItem"].Visible = true;
                lsv_Capacity.ContextMenuStrip = Capacity_Remove;
            }
            else if (lsv_Capacity.SelectedItems.Count == 0)
            {
                Capacity_Remove.Items["RemoveCapacity_ToolStripMenuItem"].Visible = false;
            }
        }

        private void RemoveCapacity_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (iOverall == 1)
                {
                    foreach (ListItem_MTMECapacity oItem in lsv_Capacity.SelectedItems)
                    {
                        int iUpdate = BusinessLogic.WS_Allocation.SetCapacityRemove(Convert.ToInt32(oItem.IPRODUCTION_ID), 0, Convert.ToDateTime(dtp_Capacityfrom.Value));
                        if (iUpdate > 0)
                        {
                            //Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                            Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                            Load_Capacity_Percentage(Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value));
                        }
                    }
                }
                else
                {
                    foreach (ListItem_MTME_CurrentdateCapacity oItem in lsv_Currentdate_Capacity.SelectedItems)
                    {
                        int iUpdate = BusinessLogic.WS_Allocation.SetCapacityRemove(Convert.ToInt32(oItem.IPRODUCTION_ID), 1, Convert.ToDateTime(dtp_Capacityfrom.Value));
                        if (iUpdate > 0)
                        {
                            Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                            //Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                            Load_Capacity_Percentage(Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void btn_Addcapacity_Click(object sender, EventArgs e)
        {
            grp_AddCapacity.Visible = true;
            Load_Capacity_Usernames(Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToInt32(cmb_AddCap_Worktype.SelectedValue));
            Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
        }

        private void label136_Click(object sender, EventArgs e)
        {
            grp_AddCapacity.Visible = false;
            Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
            Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
        }

        private void btn_Add_Cap_Click(object sender, EventArgs e)
        {
            try
            {
                int iUpdate = BusinessLogic.WS_Allocation.SetCapacityRemove(Convert.ToInt32(cmb_AddCap_Name.SelectedValue), 0, Convert.ToDateTime(dtp_Capacityfrom.Value));
                if (iUpdate > 0)
                {
                    Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                    Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
                    Load_Capacity_Percentage(Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value));
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void lsv_Currentdate_Capacity_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                iCurrent = 1;
                iOverall = 0;
                if (lsv_Currentdate_Capacity.SelectedItems.Count == 1)
                {
                    Capacity_Remove.Items["RemoveCapacity_ToolStripMenuItem"].Visible = true;
                    lsv_Currentdate_Capacity.ContextMenuStrip = Capacity_Remove;
                }
                else if (lsv_Currentdate_Capacity.SelectedItems.Count == 0)
                {
                    Capacity_Remove.Items["RemoveCapacity_ToolStripMenuItem"].Visible = false;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }


        private void btnTarget_Export_Click(object sender, EventArgs e)
        {
            //lsv_Target_NamesList ListItem_MTMEList
            string sName = string.Empty;
            foreach (ListItem_MTMEList oItem in lsv_Target_NamesList.SelectedItems)
            {
                sName = oItem.SEMP_NAME.Trim();
            }
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Target Details " + sName.Trim() + ".xls";
            ExportToExcel(lsv_Target_Report, sFolderNAme, sFileName);
        }

        private void btn_Cap_Close_Click(object sender, EventArgs e)
        {
            grp_AddCapacity.Visible = false;
            Load_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
            Load_Overall_Capacity_UserList(Convert.ToInt32(cmb_Capacity_Branch.SelectedValue), Convert.ToInt32(cmb_Capacity_Batch.SelectedValue), Convert.ToDateTime(dtp_Capacityfrom.Value), Convert.ToDateTime(dtp_Capacityto.Value), Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue));
        }
        #endregion "EVENTS "

        #endregion " OFFLINE "

        #region "DSP and NDSP Details"


        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                object FromDate;
                object ToDate;

                lsvdspndsp.Items.Clear();
                lsvfreshDsp.Items.Clear();

                if (chkDate.Checked)
                {
                    FromDate = dtpFromDate.Value.ToString("yyyy-MM-dd");
                    ToDate = dtpToDate.Value.ToString("yyyy-MM-dd");
                }
                else
                {
                    FromDate = null;
                    ToDate = null;
                }
                DataSet dsNdspData = BusinessLogic.WS_QualityService.GET_NDSP_DETAILS(Convert.ToInt32(cboEmployee.SelectedValue.ToString().Split('/')[0].ToString()), cboClient.SelectedValue.ToString(), FromDate, ToDate);
                if (dsNdspData.Tables[0].Rows.Count < 0)
                    return;
                if (dsNdspData.Tables[0] == null)
                    return;

                int i = 1;

                foreach (DataRow dr in dsNdspData.Tables[0].Select("batch_id=" + cboDesignationNDSP.SelectedValue))
                {
                    lsvdspndsp.Items.Add(new listviewNdspDetails(dr, i));
                    i++;
                }
                BusinessLogic.Reset_ListViewColumn(lsvdspndsp);
                i = 1;
                foreach (DataRow dr in dsNdspData.Tables[1].Select("batch_id=" + cboDesignationNDSP.SelectedValue))
                {
                    lsvfreshDsp.Items.Add(new listviewFreshdspDetails(dr, i));
                    i++;
                }
                BusinessLogic.Reset_ListViewColumn(lsvfreshDsp);
            }
            catch (Exception ex)
            {
                BusinessLogic.oMessageEvent.Start("Error: " + ex.Message.ToString());
            }
        }

        public class listviewNdspDetails : ListViewItem
        {
            public int Detail_id;
            public string Production_id;
            public string Location_id;
            public string Comment;
            public string NDSP_ON;
            public string DSP_ON;
            public int batch_id;

            public listviewNdspDetails(DataRow dr, int i)
            {
                this.Text = i.ToString();
                this.SubItems.Add(dr["ptag_id"].ToString());
                this.SubItems.Add(dr["emp_full_name"].ToString());
                this.SubItems.Add(dr["location_name"].ToString());
                this.SubItems.Add(dr["Ndsp_on"].ToString());
                this.SubItems.Add(dr["dsp_on"].ToString());
                this.SubItems.Add(dr["comment"].ToString());

                batch_id = Convert.ToInt32(dr["batch_id"]);
                Detail_id = Convert.ToInt32(dr["detail_id"]);
                Production_id = dr["production_id"].ToString();
                Location_id = dr["location_id"].ToString();
                NDSP_ON = dr["Ndsp_on"].ToString();
                DSP_ON = dr["dsp_on"].ToString();
                Comment = dr["comment"].ToString();

                if (dr["dsp_on"].ToString() == string.Empty)
                {

                    this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    this.BackColor = Color.Yellow;
                    this.ForeColor = Color.Red;
                }
                if (dr["Ndsp_on"].ToString() == string.Empty)
                {

                    this.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    this.BackColor = Color.Violet;
                    this.ForeColor = Color.White;
                }
            }


        }

        public class listviewFreshdspDetails : ListViewItem
        {
            public int Detail_id;
            public string Production_id;
            public string Location_id;
            public string Comment;
            public string DSP_ON;
            public string NDSP_ON;
            public int batch_id;
            public bool freshDSDP;

            public listviewFreshdspDetails(DataRow dr, int i)
            {
                this.Text = i.ToString();
                this.SubItems.Add(dr["ptag_id"].ToString());
                this.SubItems.Add(dr["emp_full_name"].ToString());
                this.SubItems.Add(dr["location_name"].ToString());
                this.SubItems.Add(dr["dsp_on"].ToString());
                this.SubItems.Add(dr["comment"].ToString());

                batch_id = Convert.ToInt32(dr["batch_id"]);
                Detail_id = Convert.ToInt32(dr["detail_id"]);
                Production_id = dr["production_id"].ToString();
                Location_id = dr["location_id"].ToString();
                DSP_ON = dr["dsp_on"].ToString();
                NDSP_ON = dr["Ndsp_on"].ToString();
                Comment = dr["comment"].ToString();
                freshDSDP = Convert.ToBoolean(dr["fresh_dsp"]);
            }

        }
        /// <summary>
        /// To write the log details 
        /// </summary>
        /// <param name="sErrorDiscription"></param>
        /// <param name="ex"></param>
        public void LogException(string sErrorDiscription, Exception ex)
        {
            try
            {
                if (!BusinessLogic.CheckConnection())
                {
                    BusinessLogic.oMessageEvent.Start("Network error, Unable to connect to server.");
                    MessageBox.Show("Network Error, unable to connect to server" + Environment.NewLine + "Please contact the technical person.", " ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                BusinessLogic.oMessageEvent.Start("Error: " + sErrorDiscription + "->" + ex.Message.ToString());

            }
            catch
            {
                BusinessLogic.oMessageEvent.Start("Error: " + ex.Message.ToString());
            }
        }

        private void Load_Designation()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                DataSet _dsDesignation = BusinessLogic.WS_QualityService.Get_Designation();
                cboDesignationNDSP.DisplayMember = "" + Framework.BATCH.FIELD_BATCH_BATCHNAME_STR + "";
                cboDesignationNDSP.ValueMember = "" + Framework.BATCH.FIELD_BATCH_BATCHID_INT + "";
                cboDesignationNDSP.Text = "" + Framework.BATCH.FIELD_BATCH_BATCHNAME_STR + "";
                cboDesignationNDSP.DataSource = _dsDesignation.Tables[0];
                BusinessLogic.oMessageEvent.Start("Done.");


            }
            catch (Exception ex)
            {
                LogException("cboDesignationNDSP", ex);
                BusinessLogic.WS_QualityService.WriteException("Exception in Loading Designation | " + ex.ToString(), Environment.UserName, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);

            }
        }



        private void cboDesignationNDSP_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {

                if (cboDesignationNDSP.SelectedValue == null)
                    return;
                //Loading Employee
                DataSet _dsEmployee = BusinessLogic.WS_QualityService.Get_Employee(Convert.ToInt32(cboDesignationNDSP.SelectedValue));
                DataRow dr = _dsEmployee.Tables[0].NewRow();
                dr["" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + ""] = "-- ALL --";
                dr["" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + ""] = 0;
                _dsEmployee.Tables[0].Rows.InsertAt(dr, 0);

                cboEmployee.DisplayMember = "" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + "";
                cboEmployee.ValueMember = "" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + "";
                cboEmployee.DataSource = _dsEmployee.Tables[0];
            }
            catch (Exception ex)
            {
                LogException("cboDesignation_SelectedValueChanged", ex);
            }
        }


        private void cboEmployee_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {

                if (cboEmployee.SelectedValue == null)
                    return;
                //Loading Employee
                DataSet _dsClient = BusinessLogic.WS_QualityService.Get_Client_userwise(Convert.ToInt32(cboDesignationNDSP.SelectedValue), Convert.ToInt32(cboEmployee.SelectedValue.ToString().Split('/')[0].ToString()));
                DataRow dr = _dsClient.Tables[0].NewRow();
                dr["" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + ""] = "-- ALL --";
                dr["" + Framework.LOCATION.FIELD_LOCATION_ID_STR + ""] = 0;
                _dsClient.Tables[0].Rows.InsertAt(dr, 0);

                cboClient.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cboClient.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cboClient.DataSource = _dsClient.Tables[0];
            }
            catch (Exception ex)
            {
                LogException("cboEmployee_SelectedValueChanged", ex);
            }

        }

        private void dSPONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listviewNdspDetails oRecords = (listviewNdspDetails)lsvdspndsp.SelectedItems[0];

            frmaddNdsp fNew = new frmaddNdsp(oRecords.batch_id, oRecords.Production_id, oRecords.Location_id, oRecords.NDSP_ON, oRecords.DSP_ON, oRecords.Comment, oRecords.Detail_id, false, 2);
            if (fNew.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Updated Sucessfully");
                btnSearch_Click(this, e);
            }
            else
            {
                MessageBox.Show("Not Updated");
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int result = -1;
            listviewNdspDetails oRecords = (listviewNdspDetails)lsvdspndsp.SelectedItems[0];
            result = BusinessLogic.WS_QualityService.SET_NDSP_DETAILS(oRecords.Detail_id, Convert.ToInt32(oRecords.Production_id.Split('/')[0].ToString()), oRecords.Location_id, oRecords.Comment, null, oRecords.DSP_ON, 3, false);
            if (result == 1)
            {
                MessageBox.Show("Deleted Sucessfully");
                btnSearch_Click(this, e);
            }
            else
                MessageBox.Show("Not Deleted");
        }



        private void btnExport_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "FreshDSP.xls";

            ExportToExcel(lsvfreshDsp, sFolderNAme, sFileName);
            sFileName = "DSP and NDSP.xls";
            ExportToExcel(lsvdspndsp, sFolderNAme, sFileName);
        }



        private void tsModify_Click(object sender, EventArgs e)
        {
            listviewFreshdspDetails oRecords = (listviewFreshdspDetails)lsvfreshDsp.SelectedItems[0];

            frmaddNdsp fNew = new frmaddNdsp(oRecords.batch_id, oRecords.Production_id, oRecords.Location_id, oRecords.DSP_ON, oRecords.NDSP_ON, oRecords.Comment, oRecords.Detail_id, oRecords.freshDSDP, 2);
            if (fNew.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Updated Sucessfully");
                btnSearch_Click(this, e);
            }
            else
            {
                MessageBox.Show("Not Updated");
            }
        }

        private void tsdelete_Click(object sender, EventArgs e)
        {
            int result = -1;
            listviewFreshdspDetails oRecords = (listviewFreshdspDetails)lsvfreshDsp.SelectedItems[0];
            result = BusinessLogic.WS_QualityService.SET_NDSP_DETAILS(oRecords.Detail_id, Convert.ToInt32(oRecords.Production_id.Split('/')[0].ToString()), oRecords.Location_id, oRecords.Comment, null, oRecords.DSP_ON, 3, oRecords.freshDSDP);
            if (result == 1)
            {
                MessageBox.Show("Deleted Sucessfully");
                btnSearch_Click(this, e);
            }
            else
                MessageBox.Show("Not Deleted");

        }
        #endregion

        #region "Class Hangup"

        public class ListItem_HigherLines : ListViewItem
        {
            public ListItem_HigherLines(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["report_name"].ToString());
                SubItems.Add(dr["file_status_description"].ToString());
                SubItems.Add(dr["File_lines"].ToString());
                SubItems.Add(dr["File_Lines_Check"].ToString());
                SubItems.Add(dr["submitted_time"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_SplitMinutes : ListViewItem
        {
            public ListItem_SplitMinutes(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["file_date"].ToString());
                SubItems.Add(dr["trans_by"].ToString());
                SubItems.Add(dr["ted_by"].ToString());
                SubItems.Add(dr["edit_by"].ToString());
                SubItems.Add(dr["hold_by"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_TransEdit_Updated : ListViewItem
        {
            public ListItem_TransEdit_Updated(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["report_name"].ToString());
                SubItems.Add(dr["file_status_description"].ToString());
                SubItems.Add(dr["File_lines"].ToString());
                SubItems.Add(dr["submitted_time"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_ClientLines_Updated : ListViewItem
        {
            public ListItem_ClientLines_Updated(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["report_name"].ToString());
                SubItems.Add(dr["file_status_description"].ToString());
                SubItems.Add(dr["File_lines"].ToString());
                SubItems.Add(dr["submitted_time"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_NoDic_Updated : ListViewItem
        {
            public ListItem_NoDic_Updated(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["file_date"].ToString());
                SubItems.Add(dr["Hangup_lines"].ToString());
                SubItems.Add(dr["trans_by"].ToString());
                SubItems.Add(dr["ted_by"].ToString());
                SubItems.Add(dr["edit_by"].ToString());
                SubItems.Add(dr["hold_by"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_ManualEntry : ListViewItem
        {
            public ListItem_ManualEntry(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["report_name"].ToString());
                SubItems.Add(dr["file_status_description"].ToString());
                SubItems.Add(dr["File_lines"].ToString());
                SubItems.Add(dr["submitted_time"].ToString());
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_Zero_Minutes : ListViewItem
        {
            public ListItem_Zero_Minutes(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["voice_file_id"].ToString());
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["file_minutes"].ToString());
                SubItems.Add(dr["trans_by"].ToString());
                SubItems.Add(dr["edit_by"].ToString());
                SubItems.Add(dr["ted_by"].ToString());
                SubItems.Add(dr["hold_by"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class ListItem_Zero_Lines : ListViewItem
        {
            public ListItem_Zero_Lines(DataRow dr, int iRowcount)
                : base()
            {
                Text = iRowcount.ToString();
                SubItems.Add(dr["location_name"].ToString());
                SubItems.Add(dr["doctor_full_name"].ToString());
                SubItems.Add(dr["report_name"].ToString());
                SubItems.Add(dr["file_status_description"].ToString());
                SubItems.Add(dr["File_lines"].ToString());
                SubItems.Add(dr["submitted_time"].ToString());
                SubItems.Add(dr["ptag_id"].ToString());
                SubItems.Add(dr["emp_full_name"].ToString());

                if (iRowcount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }


        #endregion "Class Hangup"

        #region "HangUp Process"

        private void LoadYear_HangUP()
        {
            try
            {
                cmbHangUpYear.Items.Clear();
                int iCurrentYear = DateTime.Now.Year;
                for (int i = 2014; i <= iCurrentYear; i++)
                {
                    cmbHangUpYear.Items.Add(i.ToString());
                }
                cmbHangUpYear.SelectedIndex = 0;
                cmbHangUpYear.Text = Convert.ToString(iCurrentYear);
                cmbHangUpYear.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void Load_Month_Name_HangUP()
        {
            try
            {
                cmbHangUpMonth.Items.Clear();
                for (int i = 0; i < 12; i++)
                {
                    cmbHangUpMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                }
                cmbHangUpMonth.SelectedIndex = DateTime.Now.Month - 1;
                cmbHangUpMonth.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void tabControl4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (tbpHabgUpProcess.SelectedTab.Name == "tabPage15")
                {
                    Thread tHigherLines = new Thread(LoadHigherLines);
                    tHigherLines.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage16")
                {
                    Thread tNoDictation = new Thread(LoadNoDictation);
                    tNoDictation.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage17")
                {
                    Thread tTransEditMinutes = new Thread(LoadTransEditMinutes);
                    tTransEditMinutes.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage18")
                {
                    Thread tSplitMinutes = new Thread(LoadSplitMinutes);
                    tSplitMinutes.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage19")
                {
                    Thread tManualEntry = new Thread(LoadManualEntry);
                    tManualEntry.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage20")
                {
                    Thread tManualEntry = new Thread(LoadZeroMinutes);
                    tManualEntry.Start();
                }
                else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage21")
                {
                    Thread tZeroLines = new Thread(LoadZeroLines);
                    tZeroLines.Start();
                }
                //else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage22")
                //{
                //    Thread tClientLinesUpdate = new Thread(LoadClientLinesUpdated);
                //    tClientLinesUpdate.Start();
                //}
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void LoadHigherLines()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvHigherLines.Items.Clear();
                DataSet _dsHigherLines = new DataSet();
                _dsHigherLines = BusinessLogic.WS_Allocation.Get_Higher_lines(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsHigherLines.Tables[0].Select())
                    lvHigherLines.Items.Add(new ListItem_HigherLines(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvHigherLines);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void LoadSplitMinutes()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvSplitMinutesUpdated.Items.Clear();
                DataSet _dsSplitMinutes = new DataSet();
                _dsSplitMinutes = BusinessLogic.WS_Allocation.Get_Split_Minutes_updated(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsSplitMinutes.Tables[0].Select())
                    lvSplitMinutesUpdated.Items.Add(new ListItem_SplitMinutes(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvSplitMinutesUpdated);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void LoadTransEditMinutes()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvTransEdit.Items.Clear();
                DataSet _dsTransEdit = new DataSet();
                _dsTransEdit = BusinessLogic.WS_Allocation.Get_Trans_Edit_updated(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsTransEdit.Tables[0].Select())
                    lvTransEdit.Items.Add(new ListItem_TransEdit_Updated(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvTransEdit);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        //private void LoadClientLinesUpdated()
        //{
        //    try
        //    {
        //        BusinessLogic.oMessageEvent.Start("Transferring Data..!");
        //        BusinessLogic.oProgressEvent.Start(true);

        //        string sYear = cmbHangUpYear.Text.ToString();
        //        int sMonth = cmbHangUpMonth.SelectedIndex + 1;

        //        lvClientLinesUpdated.Items.Clear();
        //        DataSet _dsClientLines = new DataSet();
        //        _dsClientLines = BusinessLogic.WS_Allocation.Get_Client_lines_updated(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

        //        int iRow = 1;
        //        foreach (DataRow _drRow in _dsClientLines.Tables[0].Select())
        //            lvClientLinesUpdated.Items.Add(new ListItem_ClientLines_Updated(_drRow, iRow++));

        //        BusinessLogic.Reset_ListViewColumn(lvClientLinesUpdated);

        //    }
        //    catch (Exception ex)
        //    {
        //        BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
        //    }
        //    finally
        //    {
        //        BusinessLogic.oMessageEvent.Start("Data transferred..!");
        //        BusinessLogic.oProgressEvent.Start(false);
        //    }
        //}

        private void LoadNoDictation()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvNoDicFiles.Items.Clear();
                DataSet _dsNoDic = new DataSet();
                _dsNoDic = BusinessLogic.WS_Allocation.Get_No_dictation(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsNoDic.Tables[0].Select())
                    lvNoDicFiles.Items.Add(new ListItem_NoDic_Updated(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvNoDicFiles);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void LoadManualEntry()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvManualEntry.Items.Clear();
                DataSet _dsManualEntry = new DataSet();
                _dsManualEntry = BusinessLogic.WS_Allocation.Get_Entries_manually_Given(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsManualEntry.Tables[0].Select())
                    lvManualEntry.Items.Add(new ListItem_ManualEntry(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvManualEntry);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void LoadZeroMinutes()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvZeroMinutesUpdated.Items.Clear();
                DataSet _dsZeroMinutes = new DataSet();
                _dsZeroMinutes = BusinessLogic.WS_Allocation.Get_Zero_Minutes_updated(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsZeroMinutes.Tables[0].Select())
                    lvZeroMinutesUpdated.Items.Add(new ListItem_Zero_Minutes(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvZeroMinutesUpdated);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void LoadZeroLines()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                BusinessLogic.oProgressEvent.Start(true);

                string sYear = cmbHangUpYear.Text.ToString();
                int sMonth = cmbHangUpMonth.SelectedIndex + 1;

                lvZeroLinesUpdated.Items.Clear();
                DataSet _dsZeroLines = new DataSet();
                _dsZeroLines = BusinessLogic.WS_Allocation.Get_zero_lines_updated(Convert.ToInt32(sYear), Convert.ToInt32(sMonth));

                int iRow = 1;
                foreach (DataRow _drRow in _dsZeroLines.Tables[0].Select())
                    lvZeroLinesUpdated.Items.Add(new ListItem_Zero_Lines(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lvZeroLinesUpdated);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Data transferred..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }
        #endregion "HangUp Process"

        #region "Events 1"

        private void txtUserID_TextChanged(object sender, EventArgs e)
        {
            try
            {
                PopulateSearch(txtUserID.Text);
            }
            catch (Exception ex)
            {
                LogException("PopulateSearch()", ex);
            }

        }
        private void PopulateSearch(string queryStr)
        {
            lsvHourlyrReports.Items.Clear();

            //int i = 1;
            int iRowCount = 0;
            foreach (DataRow _drRow in _dsHourlyWiseReport.Tables[0].Select("Ptag_id like '%" + txtUserID.Text + "%'"))
                lsvHourlyrReports.Items.Add(new ListItem_EmployeeHourly_Log(_drRow, iRowCount++));

            mTooltip.SetToolTip(lsvHourlyrReports, string.Empty);
            BusinessLogic.Reset_ListViewColumn(lsvHourlyrReports);
        }

        private void cmb_Capacity_Worktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(cmb_Capacity_Worktype.SelectedIndex) == 0)
                pnl_Current_cap_offline.Visible = true;
            else
                pnl_Current_Online_All.Visible = true;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            //Offline
            if (Convert.ToInt32(cmb_Capacity_Worktype.SelectedValue) == 1)
            {
                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "Capacity for Offline : " + dtp_Capacityfrom.Text + " " + " and " + dtp_Capacityto.Text + ".xls";
                ExportToExcel(lsv_Currentdate_Capacity, sFolderNAme, sFileName);
            }
            else
            {
                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "Capacity on Online Iasis : " + dtp_Capacityfrom.Text + " " + " and " + dtp_Capacityto.Text + ".xls";
                ExportToExcel(lsv_Iasis_capcity, sFolderNAme, sFileName);

                string sFolderNAme_MKMG = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName_MKMG = "Capacity on Online MKMG : " + dtp_Capacityfrom.Text + " " + " and " + dtp_Capacityto.Text + ".xls";
                ExportToExcel(lsv_Mkmg_capcity, sFolderNAme_MKMG, sFileName_MKMG);

                string sFolderNAme_Night = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName_Night = "Capacity on Night Shift : " + dtp_Capacityfrom.Text + " " + " and " + dtp_Capacityto.Text + ".xls";
                ExportToExcel(lsv_Iasis_capcity, sFolderNAme_Night, sFileName_Night);

            }

            string sFolderNAme_Leave = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName_Leave = "Capacity for Leave : " + dtp_Capacityfrom.Text + " " + " and " + dtp_Capacityto.Text + ".xls";
            ExportToExcel(lsv_Currentdate_Capacity, sFolderNAme_Leave, sFileName_Leave);

        }

        private void GetEmployeeList()
        {
            try
            {
                int iRowCount = 1;
                if (cmb_Emp_Workplatform.SelectedValue == null)
                    cmb_Emp_Workplatform.SelectedValue = 0;

                DataSet _dsEmployee = BusinessLogic.WS_Allocation.Get_All_Employees_V2(Convert.ToInt32(cmb_Emp_Workplatform.SelectedValue), Convert.ToInt32(cmb_Emp_Batch.SelectedValue), Convert.ToInt32(cmb_Emp_Branch.SelectedValue));

                lsvAllEmployee.Items.Clear();
                foreach (DataRow _drRow in _dsEmployee.Tables[4].Select("isActive=1"))
                    lsvAllEmployee.Items.Add(new Listitem_AllEmployees_List(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvAllEmployee);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void txtEmployee_TextChanged(object sender, EventArgs e)
        {

            DataSet _dsEmployee = new DataSet();
            _dsEmployee = BusinessLogic.WS_Allocation.Get_All_Employees();
            lsvAllEmployee.Items.Clear();

            //int i = 1;
            int iRowCount = 0;
            foreach (DataRow _drRow in _dsEmployee.Tables[4].Select("emp_full_name like '" + txtEmployee.Text + "%' and isActive=" + chkViewAll.Checked))
                lsvAllEmployee.Items.Add(new Listitem_AllEmployees_List(_drRow, iRowCount++));
            BusinessLogic.Reset_ListViewColumn(lsvAllEmployee);
        }

        private void chkViewAll_CheckedChanged(object sender, EventArgs e)
        {
            DataSet _dsEmployee = new DataSet();
            _dsEmployee = BusinessLogic.WS_Allocation.Get_All_Employees();
            lsvAllEmployee.Items.Clear();

            //int i = 1;
            int iRowCount = 1;
            foreach (DataRow _drRow in _dsEmployee.Tables[4].Select("isActive=" + chkViewAll.Checked))
                lsvAllEmployee.Items.Add(new Listitem_AllEmployees_List(_drRow, iRowCount++));
            BusinessLogic.Reset_ListViewColumn(lsvAllEmployee);
        }

        private void btnDeactive_Click(object sender, EventArgs e)
        {
            try
            {
                //Listitem_AllEmployees_List oRecords = (Listitem_AllEmployees_List)lsvAllEmployee.SelectedItems[0];
                foreach (Listitem_AllEmployees_List oRecords in lsvAllEmployee.SelectedItems)
                    BusinessLogic.WS_Allocation.Set_Deactivate_Employee(oRecords.iProductionID, oRecords.isActive);
                GetEmployeeList();
            }
            catch (Exception ex)
            {
                LogException("PopulateSearch()", ex);
            }



        }

        private void pnl_Customized_Top_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btn_Customized_Remove_Click(object sender, EventArgs e)
        {
            try
            {
                //validation
                if (cmb_Customized_Employee.Text.Trim() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the employee Name");
                    cmb_Customized_Employee.Focus();
                    return;
                }

                //int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee_Removal(Convert.ToInt32(cmb_Customized_Employee.SelectedValue));
                //int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee_Removal_V2(Convert.ToInt32(cmb_Customized_Employee.SelectedValue), Convert.ToInt32(BusinessLogic.SPRODUCTIONID));
                int iResult = BusinessLogic.WS_Allocation.Set_Customized_Employee_Removal_V3(Convert.ToInt32(cmb_Customized_Employee.SelectedValue), Convert.ToInt32(cmb_Customize_Group.SelectedValue));

                if (iResult > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Added Sucessfully");
                    Load_Customized_Employee_Removal_List();
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Already Added");
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsv_Customized_Remove_Employee_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                ICUSTOM_REMOVAL = 1;
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (ListItem_Customized_Employee oitem in lsv_Customized_Remove_Employee.SelectedItems)
                    {
                        Customized_Emp_contextMenuStrip.Visible = true;
                        Customized_Emp_contextMenuStrip.Show(PointToScreen(Control.MousePosition));
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void dtp_incentive_todate_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                Load_Incentive_Mins();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_incentive_submit_Click(object sender, EventArgs e)
        {
            try
            {
                int iInsert = 0;
                if (lsv_Incentive.SelectedItems.Count == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Select the No to save");
                    return;
                }

                foreach (ListItemIncentive_Mins oItem in lsv_Incentive.SelectedItems)
                {
                    iInsert = BusinessLogic.WS_Allocation.Set_Incentive_amount(Convert.ToInt32(oItem.IPRODUCTION_ID), dtp_incentive_fromdate.Value, Convert.ToInt32(oItem.IMINS_DONE), oItem.DINCENTIVE_AMOUNT);
                }

                if (iInsert > 0)
                    BusinessLogic.oMessageEvent.Start("Inserted Successfully");
                else
                    BusinessLogic.oMessageEvent.Start("Already Exists");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_Incentive_View_Click(object sender, EventArgs e)
        {
            try
            {
                // Load_Incentive_View();
                //Load_Incentive_View2();
                Load_Incentive_View_New();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsv_Incentive_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (lsv_Incentive.SelectedItems.Count == 1)
                {
                    Incentive_Remove_contextMenuStrip.Items["Incentive_RemoveToolStripMenuItem"].Visible = true;
                    lsv_Incentive.ContextMenuStrip = Incentive_Remove_contextMenuStrip;
                }
                else if (lsv_Incentive.SelectedItems.Count == 0)
                {
                    Incentive_Remove_contextMenuStrip.Items["Incentive_RemoveToolStripMenuItem"].Visible = false;
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Incentive_RemoveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (ListItemIncentive_Mins oItem in lsv_Incentive.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Incentive_Remove(Convert.ToInt32(oItem.IPRODUCTION_ID), Convert.ToDateTime(dtp_incentive_fromdate.Value));
                    if (iUpdate > 0)
                    {
                        Load_Incentive_View();
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_View_ExtraSupport_Click(object sender, EventArgs e)
        {
            try
            {
                frmExtraSupport FES = new frmExtraSupport();
                if (FES.ShowDialog() == DialogResult.OK)
                {
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_NightshiftView_Click(object sender, EventArgs e)
        {
            frmNightshiftView FNV = new frmNightshiftView();
            FNV.Show();
        }

        private void lvNightShiftMarked_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (Listitem_NightShift_Marked oitem in lvNightShiftMarked.SelectedItems)
                    {
                        ChangeNightshift_Category_contextMenuStrip.Visible = true;
                        ChangeNightshift_Category_contextMenuStrip.Show(PointToScreen(Control.MousePosition));
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void changeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grp_Nightshift_Change.Visible = true;
            Load_Category_Details();
        }

        private void btn_ShiftUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                Update_Nightshift_Allowance();
                grp_Nightshift_Change.Visible = false;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void removeAllowanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Remove_Nightshift_Allowance();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void label153_Click(object sender, EventArgs e)
        {
            grp_Nightshift_Change.Visible = false;
        }

        private void cmb_MTTrack_Workplatform_SelectedIndexChanged(object sender, EventArgs e)
        {
            lvFileAllotedStatus.Items.Clear();
            Thread tTrack = new Thread(Load_MT_Tracking);
            tTrack.Start();
        }

        private void cmb_Nightshift_Branch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cmb_Nightshift_Branch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_Employee_Full_name(sDesgination_ID, sBranchId);
        }

        private void cmb_Nightshift_Batch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cmb_Nightshift_Batch.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesigID, sBranch_ID);
        }

        private void btn_Nightshift_Add_Click(object sender, EventArgs e)
        {
            try
            {
                Save_Nightshift_Users_List();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsv_Nightshift_Employee_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (ListItemNightshift oitem in lsv_Nightshift_Employee.SelectedItems)
                    {
                        Nightshift_Users_contextMenuStrip.Visible = true;
                        Nightshift_Users_contextMenuStrip.Show(PointToScreen(Control.MousePosition));
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void removeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Remove_Nightshift_Users();
        }

        private void btn_MTMETAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmb_MTMET_Name.Text.Trim() == "")
                {
                    BusinessLogic.oMessageEvent.Start("Select the employee Name");
                    cmb_MTMET_Name.Focus();
                    return;
                }

                int iResult = BusinessLogic.WS_Allocation.Set_MTMET_Users(Convert.ToInt32(cmb_MTMET_Name.SelectedValue), Convert.ToDateTime(dtp_MTMETDate.Value));

                if (iResult > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Added successfully..!");
                    Load_MTMETList(0);
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Already Added..!");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_MTMETView_Click(object sender, EventArgs e)
        {
            Load_MTMETList(1);
        }

        private void remov_MTMET_eToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Remove_MTMET_Users();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }

        private void lsv_MTMETOverall_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (ListItemMTMETList oitem in lsv_MTMETOverall.SelectedItems)
                    {
                        MTMET_contextMenuStrip.Visible = true;
                        MTMET_contextMenuStrip.Show(PointToScreen(Control.MousePosition));
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        private void btnBlankView_Click(object sender, EventArgs e)
        {
            try
            {

                Load_BlankCount_View();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_BlankCount_View()
        {
            try
            {
                string Account = string.Empty;
                string Location = string.Empty;
                string Doctor = string.Empty;
                int Type = 0;
                int MEorQC = 0;

                BusinessLogic.oMessageEvent.Start("Transfering Data...!");
                BusinessLogic.oProgressEvent.Start(true);
                DataTable dtView = null;
                if (cbxType.Text == "OnLine")
                    Type = 2;
                else if (cbxType.Text == "Offline")
                    Type = 1;
                if (chkME.Checked)
                    MEorQC = 1;
                else if (chkQC.Checked)
                    MEorQC = 2;
                else
                    MEorQC = 0;
                lsvBlankCount.Items.Clear();
                lsvBlankConCount.Items.Clear();
                dtView = BusinessLogic.WS_Allocation.Get_BlankCount_V1(Type, cbxAccount.Text, cbxLocation.Text, cbxDoctor.Text, cbxEmployee.Text.Split('/').GetValue(0).ToString(), MEorQC, dtp_blankFrom.Value, dtp_blankTo.Value);
                if (dtView.Rows.Count > 0)
                {
                    int iRowcount = 1;


                    foreach (DataRow dr in dtView.Rows)
                    {
                        lsvBlankCount.Items.Add(new ListItemBlankCount(dr, iRowcount++));
                    }

                    object Trans_Lines = dtView.Compute("Sum(Trans_Lines)", "Trans_Lines not in (-1,0)").ToString();
                    object Trans_Blank = dtView.Compute("Sum(Trans_Blank)", "Trans_Blank not in (-1,0)").ToString();
                    object TED_Lines = dtView.Compute("Sum(TED_Lines)", "TED_Lines not in (-1,0)").ToString();
                    object TED_Blank = dtView.Compute("Sum(TED_Blank)", "TED_Blank not in (-1,0)").ToString();
                    object NDSP_Lines = dtView.Compute("Sum(NDSP_Lines)", "NDSP_Lines not in (-1,0)").ToString();
                    object NDSP_Blank = dtView.Compute("Sum(NDSP_Blank)", "NDSP_Blank not in (-1,0)").ToString();
                    object Edit_Lines = dtView.Compute("Sum(Edit_Lines)", "Edit_Lines not in (-1,0)").ToString();
                    object Edit_Blank = dtView.Compute("Sum(Edit_Blank)", "Edit_Blank not in (-1,0)").ToString();
                    object QC_Lines = dtView.Compute("Sum(QC_Lines)", "QC_Lines not in (-1,0)").ToString();
                    object QC_Blank = dtView.Compute("Sum(QC_Blank)", "QC_Blank not in (-1,0)").ToString();
                    object Trans_Count = dtView.Compute("Count(Transcibed_by)", "").ToString();
                    object Ted_Count = dtView.Compute("Count(TED_by)", "").ToString();
                    object NDSP_Count = dtView.Compute("Count(NDSP_by)", "").ToString();
                    object ME_Count = dtView.Compute("Count(Edit_by)", "").ToString();
                    object QC_Count = dtView.Compute("Count(QC_by)", "").ToString();
                    string TransPercentage = string.Empty;
                    string TEDPercentage = string.Empty;
                    string NDSpPercentage = string.Empty;
                    string MEPercentage = string.Empty;
                    string QCPercentage = string.Empty;
                    object Total_Files = string.Empty;
                    object Pended_files = string.Empty;
                    object Files_With_Blank = string.Empty;
                    object Files_Without_Blank = string.Empty;

                    if (Trans_Count.ToString() == "" || Trans_Blank.ToString() == "")
                        TransPercentage = "0";
                    else
                        TransPercentage = Math.Round((Convert.ToInt32(Trans_Blank) / Convert.ToDecimal(Trans_Count) * 100), 2).ToString();
                    if (Ted_Count.ToString() == "" || TED_Blank.ToString() == "")
                        TEDPercentage = "0";
                    else
                        TEDPercentage = Math.Round((Convert.ToInt32(TED_Blank) / Convert.ToDecimal(Ted_Count) * 100), 2).ToString();

                    if (NDSP_Count.ToString() == "" || NDSP_Blank.ToString() == "")
                        NDSpPercentage = "0";
                    else
                        NDSpPercentage = Math.Round((Convert.ToInt32(NDSP_Blank) / Convert.ToDecimal(NDSP_Count) * 100), 2).ToString();
                    if (ME_Count.ToString() == "" || Edit_Blank.ToString() == "")
                        MEPercentage = "0";
                    else
                        MEPercentage = Math.Round((Convert.ToInt32(Edit_Blank) / Convert.ToDecimal(ME_Count) * 100), 2).ToString();
                    if (QC_Count.ToString() == "" || QC_Blank.ToString() == "")
                        QCPercentage = "0";
                    else
                        QCPercentage = Math.Round((Convert.ToInt32(QC_Blank) / Convert.ToDecimal(QC_Count) * 100), 2).ToString();

                    Total_Files = ME_Count;
                    if (ME_Count.ToString() == "" || Edit_Blank.ToString() == "")
                        Pended_files = "0";
                    else
                        Pended_files = dtView.Compute("Count(Edit_by)", "QC_by is not null").ToString();
                    Files_With_Blank = dtView.Compute("Count(Edit_by)", "QC_by is null and Edit_Blank > 0").ToString();
                    Files_Without_Blank = dtView.Compute("Count(Edit_by)", "QC_by is null and Edit_Blank in(0,-1)").ToString();

                    lsvBlankConCount.Items.Add(new ListItemConBlankCount(
                        Trans_Count.ToString(), Trans_Lines.ToString(), Trans_Blank.ToString(), TransPercentage,
                        Ted_Count.ToString(), TED_Lines.ToString(), TED_Blank.ToString(), TEDPercentage,
                        NDSP_Count.ToString(), NDSP_Lines.ToString(), NDSP_Blank.ToString(), NDSpPercentage,
                        ME_Count.ToString(), Edit_Lines.ToString(), Edit_Blank.ToString(), MEPercentage,
                        Total_Files.ToString(), Files_With_Blank.ToString(), Files_Without_Blank.ToString(), Pended_files.ToString(),
                        QC_Count.ToString(), QC_Lines.ToString(), QC_Blank.ToString(), QCPercentage));


                    BusinessLogic.Reset_ListViewColumn(lsvBlankCount);
                    BusinessLogic.Reset_ListViewColumn(lsvBlankConCount);

                }
                BusinessLogic.oMessageEvent.Start("Done!");
                BusinessLogic.oProgressEvent.Start(false);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_Consolidated_BlankCount_View()
        {
            try
            {
                string Account = string.Empty;

                int Type = 0;


                BusinessLogic.oMessageEvent.Start("Transfering Data...!");
                BusinessLogic.oProgressEvent.Start(true);
                DataTable dtView = null;
                if (cbxConType.Text == "OnLine")
                    Type = 2;
                else if (cbxConType.Text == "Offline")
                    Type = 1;


                lsvConsolidatedBlankCount.Items.Clear();
                dtView = BusinessLogic.WS_Allocation.Get_BlankCount_Consolidated(Type, cbxConAccount.Text, dtp_ConblankFrom.Value, dtp_ConblankTo.Value);
                if (dtView.Rows.Count > 0)
                {
                    int iRowcount = 1;


                    foreach (DataRow dr in dtView.Rows)
                    {
                        lsvConsolidatedBlankCount.Items.Add(new ListItemConBlankCount(dr, iRowcount++));
                    }



                    object Trans_Lines = dtView.Compute("Sum(Trans_Lines)", "Trans_Lines not in (-1,0)").ToString();
                    object Trans_Blank = dtView.Compute("Sum(Trans_Blank)", "Trans_Blank not in (-1,0)").ToString();
                    object TED_Lines = dtView.Compute("Sum(TED_Lines)", "TED_Lines not in (-1,0)").ToString();
                    object TED_Blank = dtView.Compute("Sum(TED_Blank)", "TED_Blank not in (-1,0)").ToString();
                    object NDSP_Lines = dtView.Compute("Sum(NDSP_Lines)", "NDSP_Lines not in (-1,0)").ToString();
                    object NDSP_Blank = dtView.Compute("Sum(NDSP_Blank)", "NDSP_Blank not in (-1,0)").ToString();
                    object Edit_Lines = dtView.Compute("Sum(Edit_Lines)", "Edit_Lines not in (-1,0)").ToString();
                    object Edit_Blank = dtView.Compute("Sum(Edit_Blank)", "Edit_Blank not in (-1,0)").ToString();
                    object QC_Lines = dtView.Compute("Sum(QC_Lines)", "QC_Lines not in (-1,0)").ToString();
                    object QC_Blank = dtView.Compute("Sum(QC_Blank)", "QC_Blank not in (-1,0)").ToString();
                    object Trans_Count = dtView.Compute("Sum(Trans_Count)", "").ToString();
                    object Ted_Count = dtView.Compute("Sum(TED_Count)", "").ToString();
                    object NDSP_Count = dtView.Compute("Sum(NDSP_Count)", "").ToString();
                    object ME_Count = dtView.Compute("Sum(Edit_Count)", "").ToString();
                    object QC_Count = dtView.Compute("Sum(QC_Count)", "").ToString();
                    object Total_Files = dtView.Compute("Sum(Total_Files)", "").ToString();
                    object Pended_files = dtView.Compute("Sum(Pended_files)", "").ToString();
                    object Files_With_Blank = dtView.Compute("Sum(Files_With_Blank)", "").ToString();
                    object Files_Without_Blank = dtView.Compute("Sum(Files_Without_Blank)", "").ToString();



                    string TransPercentage = string.Empty;
                    string TEDPercentage = string.Empty;
                    string NDSpPercentage = string.Empty;
                    string MEPercentage = string.Empty;
                    string QCPercentage = string.Empty;
                    if (Trans_Lines.ToString() == "" || Trans_Blank.ToString() == "")
                        TransPercentage = "0";
                    else
                        TransPercentage = Math.Round((Convert.ToInt32(Trans_Blank) / Convert.ToDecimal(Trans_Count) * 100), 2).ToString();
                    if (TED_Lines.ToString() == "" || TED_Blank.ToString() == "")
                        TEDPercentage = "0";
                    else
                        TEDPercentage = Math.Round((Convert.ToInt32(TED_Blank) / Convert.ToDecimal(Ted_Count) * 100), 2).ToString();

                    if (NDSP_Lines.ToString() == "" || NDSP_Blank.ToString() == "")
                        NDSpPercentage = "0";
                    else
                        NDSpPercentage = Math.Round((Convert.ToInt32(NDSP_Blank) / Convert.ToDecimal(NDSP_Count) * 100), 2).ToString();
                    if (Edit_Lines.ToString() == "" || Edit_Blank.ToString() == "")
                        MEPercentage = "0";
                    else
                        MEPercentage = Math.Round((Convert.ToInt32(Edit_Blank) / Convert.ToDecimal(ME_Count) * 100), 2).ToString();
                    if (QC_Lines.ToString() == "" || QC_Blank.ToString() == "")
                        QCPercentage = "0";
                    else
                        QCPercentage = Math.Round((Convert.ToInt32(QC_Blank) / Convert.ToDecimal(QC_Count) * 100), 2).ToString();


                    lsvConsolidatedBlankCount.Items.Add(new ListItemConBlankCount(
                        Trans_Count.ToString(), Trans_Lines.ToString(), Trans_Blank.ToString(), TransPercentage,
                        Ted_Count.ToString(), TED_Lines.ToString(), TED_Blank.ToString(), TEDPercentage,
                        NDSP_Count.ToString(), NDSP_Lines.ToString(), NDSP_Blank.ToString(), NDSpPercentage,
                        ME_Count.ToString(), Edit_Lines.ToString(), Edit_Blank.ToString(), MEPercentage,
                        Total_Files.ToString(), Pended_files.ToString(), Files_With_Blank.ToString(), Files_Without_Blank.ToString(),
                        QC_Count.ToString(), QC_Lines.ToString(), QC_Blank.ToString(), QCPercentage));


                    BusinessLogic.Reset_ListViewColumn(lsvConsolidatedBlankCount);


                }
                BusinessLogic.oMessageEvent.Start("Done!");
                BusinessLogic.oProgressEvent.Start(false);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_All_Details()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Loading Details...!");
                BusinessLogic.oProgressEvent.Start(true);
                DataSet _dsType = BusinessLogic.WS_QualityService.Get_ClientType();
                DataRow dr = _dsType.Tables[0].NewRow();
                dr["" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR + ""] = "-- ALL --";
                dr["" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT + ""] = 0;
                _dsType.Tables[0].Rows.InsertAt(dr, 0);

                cbxType.DisplayMember = "" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR + "";
                cbxType.ValueMember = "" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT + "";
                cbxType.DataSource = _dsType.Tables[0];


                DataSet _dsClient = BusinessLogic.WS_Allocation.Get_ClientName_V1(0);
                DataRow drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "-- ALL --";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = 0;
                _dsClient.Tables[0].Rows.InsertAt(drc, 0);

                drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "IASIS";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -1;
                _dsClient.Tables[0].Rows.InsertAt(drc, 1);

                drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "UHS";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -2;
                _dsClient.Tables[0].Rows.InsertAt(drc, 2);



                cbxAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                cbxAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                cbxAccount.DataSource = _dsClient.Tables[0];

                DataSet _dsLocaion = BusinessLogic.WS_Allocation.Get_Location_V1(0);
                DataRow drl = _dsLocaion.Tables[0].NewRow();
                drl["" + Framework.LOCATION.FIELD_LOCATION_ID_STR + ""] = "0";
                drl["" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + ""] = "-- ALL --";
                _dsLocaion.Tables[0].Rows.InsertAt(drl, 0);

                cbxLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cbxLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cbxLocation.DataSource = _dsLocaion.Tables[0];

                DataSet _dsDoctor = BusinessLogic.WS_Allocation.Get_Doctor_V1(0, "0");
                DataRow drd = _dsDoctor.Tables[0].NewRow();
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + ""] = 0;
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + ""] = "-- ALL --";
                _dsDoctor.Tables[0].Rows.InsertAt(drd, 0);

                cbxDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                cbxDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                cbxDoctor.DataSource = _dsDoctor.Tables[0];


                DataSet _dsDesig = BusinessLogic.WS_QualityService.Get_Designation();
                DataRow drde = _dsDesig.Tables[0].NewRow();
                drde["" + Framework.BATCH.FIELD_BATCH_BATCHID_INT + ""] = 0;
                drde["" + Framework.BATCH.FIELD_BATCH_BATCHNAME_STR + ""] = "-- ALL --";
                _dsDesig.Tables[0].Rows.InsertAt(drde, 0);

                cbxDesignation.DisplayMember = "" + Framework.BATCH.FIELD_BATCH_BATCHNAME_STR + "";
                cbxDesignation.ValueMember = "" + Framework.BATCH.FIELD_BATCH_BATCHID_INT + "";
                cbxDesignation.DataSource = _dsDesig.Tables[0];


                DataSet _dsEmployee = BusinessLogic.WS_QualityService.Get_Employee(0);
                DataRow dre = _dsEmployee.Tables[0].NewRow();
                dre["" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + ""] = 0;
                dre["" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + ""] = "-- ALL --";
                _dsEmployee.Tables[0].Rows.InsertAt(dre, 0);

                cbxEmployee.DisplayMember = "" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + "";
                cbxEmployee.ValueMember = "" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + "";
                cbxEmployee.DataSource = _dsEmployee.Tables[0];
                BusinessLogic.oMessageEvent.Start("Ready!");
                BusinessLogic.oProgressEvent.Start(false);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void Load_All_Con_Details()
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Loading Details...!");
                BusinessLogic.oProgressEvent.Start(true);
                DataSet _dsType = BusinessLogic.WS_QualityService.Get_ClientType();
                DataRow dr = _dsType.Tables[0].NewRow();
                dr["" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR + ""] = "-- ALL --";
                dr["" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT + ""] = 0;
                _dsType.Tables[0].Rows.InsertAt(dr, 0);

                cbxConType.DisplayMember = "" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_STR + "";
                cbxConType.ValueMember = "" + Framework.CLIENTTYPE.FIELD_CLIENTTYPE_ID_INT + "";
                cbxConType.DataSource = _dsType.Tables[0];


                DataSet _dsClient = BusinessLogic.WS_Allocation.Get_ClientName_V1(0);
                DataRow drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "-- ALL --";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = 0;
                _dsClient.Tables[0].Rows.InsertAt(drc, 0);

                drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "IASIS";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -1;
                _dsClient.Tables[0].Rows.InsertAt(drc, 1);

                drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "UHS";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -2;
                _dsClient.Tables[0].Rows.InsertAt(drc, 2);

                cbxConAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                cbxConAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                cbxConAccount.DataSource = _dsClient.Tables[0];
                BusinessLogic.oMessageEvent.Start("Ready!");
                BusinessLogic.oProgressEvent.Start(false);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cbxType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int type = Convert.ToInt32(cbxType.SelectedValue.ToString());

                DataSet _dsClient = BusinessLogic.WS_Allocation.Get_ClientName_V1(type);
                DataRow drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "-- ALL --";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = 0;
                _dsClient.Tables[0].Rows.InsertAt(drc, 0);

                if (type == 2)
                {
                    drc = _dsClient.Tables[0].NewRow();
                    drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "IASIS";
                    drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -1;
                    _dsClient.Tables[0].Rows.InsertAt(drc, 1);

                    drc = _dsClient.Tables[0].NewRow();
                    drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "UHS";
                    drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -2;
                    _dsClient.Tables[0].Rows.InsertAt(drc, 2);
                }
                cbxAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                cbxAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                cbxAccount.DataSource = _dsClient.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cbxAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int Account = Convert.ToInt32(cbxAccount.SelectedValue.ToString());
                DataSet _dsLocaion = BusinessLogic.WS_Allocation.Get_Location_V1(Account);
                DataRow drl = _dsLocaion.Tables[0].NewRow();
                drl["" + Framework.LOCATION.FIELD_LOCATION_ID_STR + ""] = "0";
                drl["" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + ""] = "-- ALL --";
                _dsLocaion.Tables[0].Rows.InsertAt(drl, 0);

                cbxLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cbxLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cbxLocation.DataSource = _dsLocaion.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cbxLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int Account = Convert.ToInt32(cbxAccount.SelectedValue.ToString());
                string Location = cbxLocation.SelectedValue.ToString();
                DataSet _dsDoctor = BusinessLogic.WS_Allocation.Get_Doctor_V1(Account, Location);
                DataRow drd = _dsDoctor.Tables[0].NewRow();
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + ""] = 0;
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + ""] = "-- ALL --";
                _dsDoctor.Tables[0].Rows.InsertAt(drd, 0);

                cbxDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                cbxDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                cbxDoctor.DataSource = _dsDoctor.Tables[0]; ;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }


        private void cbxDesignation_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int batch = Convert.ToInt32(cbxDesignation.SelectedValue.ToString());
                DataSet _dsEmployee = BusinessLogic.WS_QualityService.Get_Employee(batch);
                DataRow dre = _dsEmployee.Tables[0].NewRow();
                dre["" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + ""] = 0;
                dre["" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + ""] = "-- ALL --";
                _dsEmployee.Tables[0].Rows.InsertAt(dre, 0);

                cbxEmployee.DisplayMember = "" + Framework.EMPLOYEE.FIELD_EMPLOYEE_FULL_NAME + "";
                cbxEmployee.ValueMember = "" + Framework.PRODUCTION_EMPLOYEES.PRODUCTION_ID + "";
                cbxEmployee.DataSource = _dsEmployee.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }

        }


        private void btnBlankExport_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Exporting.....");
                BusinessLogic.oProgressEvent.Start(true);

                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "BlankCount Report between " + dtp_blankFrom.Text + " " + " and " + dtp_blankTo.Text + " .xls";
                ExportToExcel(lsvBlankCount, sFolderNAme, sFileName);
                BusinessLogic.oMessageEvent.Start("Done!");
                BusinessLogic.oProgressEvent.Start(false);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsvBlankCount_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            this.lsvBlankCount.ListViewItemSorter = new ListViewItemComparer(e.Column);
        }

        // Implements the manual sorting of items by columns.
        class ListViewItemComparer : IComparer
        {
            private int col;
            public ListViewItemComparer()
            {
                col = 0;
            }
            public ListViewItemComparer(int column)
            {
                col = column;
            }
            public int Compare(object x, object y)
            {
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
        }

        private void btnConview_Click(object sender, EventArgs e)
        {
            try
            {

                Load_Consolidated_BlankCount_View();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnConExport_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Exporting.....");
                BusinessLogic.oProgressEvent.Start(true);

                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "Consolidated BlankCount Report between " + dtp_blankFrom.Text + " " + " and " + dtp_blankTo.Text + " .xls";
                ExportToExcel(lsvConsolidatedBlankCount, sFolderNAme, sFileName);
                BusinessLogic.oMessageEvent.Start("Done!");
                BusinessLogic.oProgressEvent.Start(false);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void cbxConType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int type = Convert.ToInt32(cbxConType.SelectedValue.ToString());

                DataSet _dsClient = BusinessLogic.WS_Allocation.Get_ClientName_V1(type);
                DataRow drc = _dsClient.Tables[0].NewRow();
                drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "-- ALL --";
                drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = 0;
                _dsClient.Tables[0].Rows.InsertAt(drc, 0);

                if (type == 2)
                {
                    drc = _dsClient.Tables[0].NewRow();
                    drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "IASIS";
                    drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -1;
                    _dsClient.Tables[0].Rows.InsertAt(drc, 1);

                    drc = _dsClient.Tables[0].NewRow();
                    drc["" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + ""] = "UHS";
                    drc["" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + ""] = -2;
                    _dsClient.Tables[0].Rows.InsertAt(drc, 2);
                }
                cbxConAccount.DisplayMember = "" + Framework.CLIENT.FIELD_CLIENT_NAME_STR + "";
                cbxConAccount.ValueMember = "" + Framework.CLIENT.FIELD_CLIENT_ID_BINT + "";
                cbxConAccount.DataSource = _dsClient.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Load_Multiple_Entries();
        }

        private void lvMultipleEntries_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem oItem in lvMultipleEntries.SelectedItems)
                {
                    ListItem_Multiple_Entries lsvMultiple = (ListItem_Multiple_Entries)oItem;
                    string sTransID = lsvMultiple.TranscriptionID.ToString();

                    eAllocation.UI.frmDuplicateDetails frmDUP = new eAllocation.UI.frmDuplicateDetails();
                    frmDUP._Transcription_ID = sTransID.ToString();
                    frmDUP.Show();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btnProdView_Click(object sender, EventArgs e)
        {
            try
            {
                lvDailyProductivity.Items.Clear();
                BusinessLogic.oMessageEvent.Start("Transferring data");
                BusinessLogic.oProgressEvent.Start(true);

                string sDate = dtpProdDate.Value.ToString("yyyy/MM/dd");

                DataSet _dsProd = new DataSet();
                _dsProd = BusinessLogic.WS_Allocation.Get_Daily_productivity(Convert.ToDateTime(sDate));

                int iRowCount = 1;
                int sNO = 1;
                foreach (DataRow _drRow in _dsProd.Tables[0].Select())
                    lvDailyProductivity.Items.Add(new ListItem_ProductivityDeatils(_drRow, iRowCount++, sNO++));

                BusinessLogic.Reset_ListViewColumn(lvDailyProductivity);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void btnExportProd_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Rnd Daily Productivity on_" + System.DateTime.Now.ToString("yyyy_mm_dd hh:mm:ss").Replace(':', '_').ToString() + ".xls";
            ExportToExcel(lvDailyProductivity, sFolderNAme, sFileName);
        }

        private void Get_Marquee_Text()
        {
            try
            {
                lsvMarquee.Items.Clear();
                int i = 1;
                DataTable dtScroll = BusinessLogic.WS_Allocation.Get_Scroll_Text();
                foreach (DataRow dr in dtScroll.Select("Is_Active=1"))
                {
                    lsvMarquee.Items.Add(new ListItem_Marquee(dr, i));
                    i++;
                }
                BusinessLogic.Reset_ListViewColumn(lsvMarquee);

            }
            catch (Exception ex)
            {
                BusinessLogic.oMessageEvent.Start("Done");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void disableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ListItem_Marquee oitem = ((ListItem_Marquee)lsvMarquee.SelectedItems[0]);
            BusinessLogic.WS_Allocation.Set_Scroll_Text(oitem.iSCROLLID, oitem.sScrollText, oitem.iAdded_by, oitem.slogin_name, oitem.sCP_Number, chkMarquee.Checked ? 1 : 0);
            chkMarquee_CheckedChanged(this, e);
        }
        private void contextMenuStripMarquee_Opening(object sender, CancelEventArgs e)
        {
            int Status = ((ListItem_Marquee)lsvMarquee.SelectedItems[0]).iSTATUS;
            if (Status == 1)
                contextMenuStripMarquee.Items[0].Text = "Disable";
            else

                contextMenuStripMarquee.Items[0].Text = "Enable";
        }

        private void chkMarquee_CheckedChanged(object sender, EventArgs e)
        {
            if (chkMarquee.Checked)
            {
                try
                {
                    lsvMarquee.Items.Clear();
                    int i = 1;
                    DataTable dtScroll = BusinessLogic.WS_Allocation.Get_Scroll_Text();
                    foreach (DataRow dr in dtScroll.Select("Is_Active=0"))
                    {
                        lsvMarquee.Items.Add(new ListItem_Marquee(dr, i));
                        i++;
                    }
                    BusinessLogic.Reset_ListViewColumn(lsvMarquee);

                }
                catch (Exception ex)
                {
                    BusinessLogic.oMessageEvent.Start("Done");
                    BusinessLogic.oProgressEvent.Start(false);

                }
            }
            else
                Get_Marquee_Text();
        }

        private void btnMarquee_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtMarquee.Text.Length <= 0)
                {
                    MessageBox.Show("Enter Text");
                    return;
                }
                BusinessLogic.WS_Allocation.Set_Scroll_Text(0, txtMarquee.Text, Convert.ToInt32(BusinessLogic.SPRODUCTIONID), BusinessLogic.USERNAME, Environment.MachineName, chkMarquee.Checked ? 1 : 0);
                chkMarquee_CheckedChanged(this, e);
            }
            catch (Exception ex)
            {
                BusinessLogic.oMessageEvent.Start("Done");
                BusinessLogic.oProgressEvent.Start(false);

            }
        }

        private void lsv_Hund_Location_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {
                if (lsv_Hund_Location.SelectedItems.Count > 0)
                {
                    ListItemHundred_Location oItem = (ListItemHundred_Location)lsv_Hund_Location.SelectedItems[0];
                    Load_Doctor_List(oItem.LOCATION_ID.ToString());
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_Hund_Add_Click(object sender, EventArgs e)
        {
            try
            {
                //VALIDATION
                if (lsv_Hund_Location.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Select the location");
                    return;
                }

                if (lsv_Hund_Doctor.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Select the doctor");
                    return;
                }

                //INSERT THE RECORD                 
                int iResult = 0;
                foreach (ListItemHundred_Doctor oItem in lsv_Hund_Doctor.SelectedItems)
                {
                    iResult = BusinessLogic.WS_Allocation.Set_Hundred_Percent(Convert.ToInt32(oItem.DOCTOR_ID.ToString()));
                }

                if (iResult > 0)
                {
                    Load_HundredPercent_Doctors_List();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_Hund_Remove_Click(object sender, EventArgs e)
        {
            try
            {
                //VALIDATION
                if (lsv_Hundred_Percent.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Select the Doctor_Name");
                    return;
                }

                int iUpdate = 0;
                DialogResult sMessage = MessageBox.Show("Do you want to remove this doctor", "Allocation", MessageBoxButtons.YesNo);
                if (sMessage == DialogResult.Yes)
                {
                    ListItemHundred_Review oItem = (ListItemHundred_Review)lsv_Hundred_Percent.SelectedItems[0];
                    iUpdate = BusinessLogic.WS_Allocation.Get_Remove_HundredPercent_Doctors_List(Convert.ToInt32(oItem.IDOCTOR_ID));

                    if (iUpdate > 0)
                    {
                        Load_HundredPercent_Doctors_List();
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }


        public static string GiveMinutes(string sTotalSeconds)
        {
            int iMinutes = 0;
            int iSeconds = 0;
            string sDuration = string.Empty;
            int iTotalSeconds = Convert.ToInt32(sTotalSeconds);
            if (iTotalSeconds > 0)
            {
                iMinutes = iTotalSeconds / 60;
                iSeconds = iTotalSeconds % 60;
            }
            sDuration = iMinutes.ToString().PadLeft(2, '0') + ":" + iSeconds.ToString().PadLeft(2, '0');
            return sDuration;
        }

        private void btnDownloadReport_Click(object sender, EventArgs e)
        {
            try
            {
                int month = Convert.ToInt32(cbMonthDaywise.SelectedIndex.ToString()) + 1;
                int Year = Convert.ToInt32(cbYearDaywise.Text);
                lsv_DayWiseDownLoadReport.Items.Clear();
                DataTable dtDaywise = BusinessLogic.WS_Allocation.Get_OfflineFile_Details_Daywise(month, Year);

                int iRowCount = 0;
                foreach (DataRow dr in dtDaywise.Rows)
                    lsv_DayWiseDownLoadReport.Items.Add(new MyListVoiceLogItem(dr, iRowCount++));

                lsv_DayWiseDownLoadReport.Items.Add(new MyListVoiceLogItem(
                    dtDaywise.Compute("Sum(Day1)", "").ToString(), dtDaywise.Compute("Sum(Day1Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day2)", "").ToString(), dtDaywise.Compute("Sum(Day2Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day3)", "").ToString(), dtDaywise.Compute("Sum(Day3Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day4)", "").ToString(), dtDaywise.Compute("Sum(Day4Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day5)", "").ToString(), dtDaywise.Compute("Sum(Day5Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day6)", "").ToString(), dtDaywise.Compute("Sum(Day6Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day7)", "").ToString(), dtDaywise.Compute("Sum(Day7Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day8)", "").ToString(), dtDaywise.Compute("Sum(Day8Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day9)", "").ToString(), dtDaywise.Compute("Sum(Day9Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day10)", "").ToString(), dtDaywise.Compute("Sum(Day10Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day11)", "").ToString(), dtDaywise.Compute("Sum(Day11Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day12)", "").ToString(), dtDaywise.Compute("Sum(Day12Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day13)", "").ToString(), dtDaywise.Compute("Sum(Day13Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day14)", "").ToString(), dtDaywise.Compute("Sum(Day14Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day15)", "").ToString(), dtDaywise.Compute("Sum(Day15Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day16)", "").ToString(), dtDaywise.Compute("Sum(Day16Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day17)", "").ToString(), dtDaywise.Compute("Sum(Day17Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day18)", "").ToString(), dtDaywise.Compute("Sum(Day18Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day19)", "").ToString(), dtDaywise.Compute("Sum(Day19Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day20)", "").ToString(), dtDaywise.Compute("Sum(Day20Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day21)", "").ToString(), dtDaywise.Compute("Sum(Day21Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day22)", "").ToString(), dtDaywise.Compute("Sum(Day22Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day23)", "").ToString(), dtDaywise.Compute("Sum(Day23Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day24)", "").ToString(), dtDaywise.Compute("Sum(Day24Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day25)", "").ToString(), dtDaywise.Compute("Sum(Day25Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day26)", "").ToString(), dtDaywise.Compute("Sum(Day26Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day27)", "").ToString(), dtDaywise.Compute("Sum(Day27Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day28)", "").ToString(), dtDaywise.Compute("Sum(Day28Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day29)", "").ToString(), dtDaywise.Compute("Sum(Day29Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day30)", "").ToString(), dtDaywise.Compute("Sum(Day30Seconds)", "").ToString(),
                    dtDaywise.Compute("Sum(Day31)", "").ToString(), dtDaywise.Compute("Sum(Day31Seconds)", "").ToString()));
                BusinessLogic.Reset_ListViewColumn(lsv_DayWiseDownLoadReport);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void btnDownloadExport_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "DayWise Report For the Month" + cbMonthDaywise + " " + cbYearDaywise + ".xls";
            ExportToExcel(lsv_DayWiseDownLoadReport, sFolderNAme, sFileName);
        }

        private void btnHalfDay_Click(object sender, EventArgs e)
        {
            try
            {
                UI.frmhalfDay frmHalf = new UI.frmhalfDay();
                frmHalf.ShowDialog();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void btnUserwiseView_Click(object sender, EventArgs e)
        {
            try
            {
                string dFrom = dtpUserwiseFrom.Text;
                string dTo = dtpUserwiseTo.Text;
                DataTable dtUserwise = BusinessLogic.WS_Allocation.Get_UserWise(Convert.ToDateTime(dFrom), Convert.ToDateTime(dTo), Convert.ToInt32(cbUserWiseBatch.SelectedValue));

                int iRowCount = 1;
                foreach (DataRow dr in dtUserwise.Rows)
                    lsvUserwise_Report.Items.Add(new MyListItemUserwise(dr, iRowCount++));
                BusinessLogic.Reset_ListViewColumn(lsvUserwise_Report);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void btnUserwiseExport_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "UserWise Report between" + dtpUserwiseFrom.Text + " " + " and " + dtpUserwiseTo.Text + " .xls";
            ExportToExcel(lsvUserwise_Report, sFolderNAme, sFileName);
        }

        #region " NIGHT SHIFT LINES REPORT "
        #region " CLASSES "
        public class Mylistitem_Users : ListViewItem
        {
            public int IPRODUCTION_ID;
            public Mylistitem_Users(DataRow dr, int iRowCount)
            {
                Text = dr["ptag_id"].ToString();
                SubItems.Add(dr["emp_full_name"].ToString());

                IPRODUCTION_ID = Convert.ToInt32(dr["production_id"].ToString());

                if (iRowCount % 2 == 1)
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
                else
                    this.BackColor = System.Drawing.ColorTranslator.FromHtml("#81F7BE");
            }
        }

        public class Mylistitem_Nightshift_Lines : ListViewItem
        {
            public Mylistitem_Nightshift_Lines(DataRow dr)
            {
                Text = dr["7to7"].ToString();
                SubItems.Add(dr["7to1"].ToString());
                SubItems.Add(dr["1to5"].ToString());
                SubItems.Add(dr["5to7"].ToString());

                this.BackColor = System.Drawing.ColorTranslator.FromHtml("#90AFB0");
            }
        }
        #endregion " CLASSES "

        #region " METHODS "
        private void Load_Batchs()
        {
            try
            {
                DataSet dsBatch = BusinessLogic.WS_Allocation.Get_AM_Batch();
                if (dsBatch != null)
                {
                    if (dsBatch.Tables[0].Rows.Count > 0)
                    {
                        cmb_NightLines_Batch.DataSource = dsBatch.Tables[0];
                        cmb_NightLines_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                        cmb_NightLines_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                        //cmb_NightLines_Batch.SelectedIndex = 0;

                        cmb_Emp_Batch.DataSource = dsBatch.Tables[0];
                        cmb_Emp_Batch.DisplayMember = Framework.BATCH.FIELD_BATCH_BATCHNAME_STR;
                        cmb_Emp_Batch.ValueMember = Framework.BATCH.FIELD_BATCH_BATCHID_INT;
                        //cmb_Emp_Batch.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }


        private void Load_Batchwise_Users(int iBatchid)
        {
            try
            {
                DataSet dsUsers = BusinessLogic.WS_Allocation.Get_Batchwise_Empname(iBatchid);
                if (dsUsers != null)
                {
                    if (dsUsers.Tables[0].Rows.Count > 0)
                    {
                        int iRowCount = 0;
                        lsv_Night_Users.Items.Clear();
                        foreach (DataRow dr in dsUsers.Tables[0].Rows)
                            lsv_Night_Users.Items.Add(new Mylistitem_Users(dr, iRowCount++));

                        BusinessLogic.Reset_ListViewColumn(lsv_Night_Users);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void Load_Nightshift_lines(int iProduction_id, DateTime dFromdate, DateTime dTodate)
        {
            try
            {
                DataSet dsLines = BusinessLogic.WS_Allocation.Get_Empwise_nightshift_Lines(iProduction_id, dFromdate, dTodate);
                if (dsLines != null)
                {
                    if (dsLines.Tables[0].Rows.Count > 0)
                    {
                        lsv_Night_Lines.Items.Clear();
                        foreach (DataRow dr in dsLines.Tables[0].Rows)
                            lsv_Night_Lines.Items.Add(new Mylistitem_Nightshift_Lines(dr));

                        BusinessLogic.Reset_ListViewColumn(lsv_Night_Lines);
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }
        #endregion " METHODS "

        #region " EVENTS "

        private void btn_Nightlines_View_Click(object sender, EventArgs e)
        {
            try
            {
                Load_Batchwise_Users(Convert.ToInt32(cmb_NightLines_Batch.SelectedValue));
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        #endregion " EVENTS "

        private void lsv_Night_Users_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {
                if (lsv_Night_Users.SelectedItems.Count == 0)
                    return;

                Mylistitem_Users oItem = (Mylistitem_Users)lsv_Night_Users.SelectedItems[0];
                Load_Nightshift_lines(Convert.ToInt32(oItem.IPRODUCTION_ID), dtp_from_date.Value, dtp_To_date.Value);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }
        #endregion " NIGHT SHIFT LINES REPORT "

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                string sDate = dtpComplaintDate.Text.ToString();
                string sComplaint = txtComplaints.Text.ToString();
                string sLocation = cmbLocationBPM.SelectedValue.ToString();

                int iSave = BusinessLogic.WS_Allocation.Set_Complaints(Convert.ToDateTime(sDate), sComplaint, sLocation);
                if (iSave > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Successfully Added..!");
                    LoadListView_Complaints();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void LoadListView_Complaints()
        {
            try
            {
                int sMonth = cmbBpmMonth.SelectedIndex + 1;
                string sYear = string.Empty;
                if (tabControl3.SelectedTab.Name == "tabPage13")
                    sYear = cmbBpmYear.Text.ToString();
                else
                    sYear = cmbBpmYear1.Text.ToString();

                DataSet _dsComplaints = new DataSet();
                _dsComplaints = BusinessLogic.WS_Allocation.Get_Complaints(Convert.ToInt32(sMonth), Convert.ToInt32(sYear));

                lvComplaints.Items.Clear();
                int iRowCount = 0; ;
                foreach (DataRow _drRow in _dsComplaints.Tables[0].Select())
                    lvComplaints.Items.Add(new ListView_Complaints(_drRow, iRowCount++));

                lvClComplaintsClinics.Items.Clear();
                int iRowCountOff = 0; ;
                foreach (DataRow _drRow in _dsComplaints.Tables[1].Select())
                    lvClComplaintsClinics.Items.Add(new ListView_Complaints(_drRow, iRowCountOff++));

                BusinessLogic.Reset_ListViewColumn(lvComplaints);
                BusinessLogic.Reset_ListViewColumn(lvClComplaintsClinics);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void Load_Month_Name_BPM()
        {
            try
            {
                cmbBpmMonth.Items.Clear();
                cmbBpmMonth1.Items.Clear();
                for (int i = 0; i < 12; i++)
                {
                    cmbBpmMonth.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                    cmbBpmMonth1.Items.Insert(i, CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]);
                }
                cmbBpmMonth.SelectedIndex = DateTime.Now.Month - 1;
                cmbBpmMonth.DropDownStyle = ComboBoxStyle.DropDownList;

                cmbBpmMonth1.SelectedIndex = DateTime.Now.Month - 1;
                cmbBpmMonth1.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void LoadYear_BPM()
        {
            try
            {
                cmbBpmYear.Items.Clear();
                cmbBpmYear1.Items.Clear();
                int iCurrentYear = DateTime.Now.Year;
                for (int i = 2014; i <= iCurrentYear; i++)
                {
                    cmbBpmYear.Items.Add(i.ToString());
                    cmbBpmYear1.Items.Add(i.ToString());
                }
                cmbBpmYear.SelectedIndex = 0;
                cmbBpmYear.Text = Convert.ToString(iCurrentYear);
                //cmbBpmYear.SelectedIndex = (cmbBpmYear.Items.Count + 1);
                cmbBpmYear.DropDownStyle = ComboBoxStyle.DropDownList;

                cmbBpmYear1.SelectedIndex = 0;
                cmbBpmYear1.Text = Convert.ToString(iCurrentYear);
                //cmbBpmYear1.SelectedIndex = (cmbBpmYear1.Items.Count + 1);
                cmbBpmYear1.DropDownStyle = ComboBoxStyle.DropDownList;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void Load_Location_BPM()
        {
            try
            {
                DataSet _dsLocation = new DataSet();
                _dsLocation = BusinessLogic.WS_Allocation.Get_Location_List();

                cmbLocationBPM.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cmbLocationBPM.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cmbLocationBPM.DataSource = _dsLocation.Tables[0];

                cmbNTSClients.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cmbNTSClients.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cmbNTSClients.DataSource = _dsLocation.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void Load_Location_BPM_CLINICS()
        {
            try
            {
                DataSet _dsLocation = new DataSet();
                _dsLocation = BusinessLogic.WS_Allocation.Get_Location_List_Offline(1);

                cmbOfflineLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cmbOfflineLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cmbOfflineLocation.DataSource = _dsLocation.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void cmbBpmMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadListView_Complaints();
        }

        private void chb_Hourly_Customized_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //if (chb_Hourly_Customized.Checked == true)
                //    grp_Hourly_P7599.Visible = true;
                //else
                //    grp_Hourly_P7599.Visible = false;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string sDate = dtpOfflineComp.Text.ToString();
                string sComplaint = txtOfflineComplaints.Text.ToString();
                string sLocation = cmbOfflineLocation.SelectedValue.ToString();

                int iSave = BusinessLogic.WS_Allocation.Set_Complaints(Convert.ToDateTime(sDate), sComplaint, sLocation);
                if (iSave > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Successfully Added..!");
                    LoadListView_Complaints();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void cmbBpmMonth1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadListView_Complaints();
        }

        private void lvAccountWiseInfo_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (lvAccountWiseInfo.SelectedItems.Count > 0)
                {
                    ListItem_AccountWiseInfo oItems = (ListItem_AccountWiseInfo)lvAccountWiseInfo.SelectedItems[0];
                    string iFromHour, iToHour;
                    object oFromDate = null;
                    object oTodate = null;

                    string AccFromDate = Convert.ToDateTime(dtpAccountFrom.Text).ToString("yyyy-MM-dd");
                    string AccTodate = Convert.ToDateTime(dtpAccountTo.Text).ToString("yyyy-MM-dd");

                    iFromHour = cmb_Acc_Fromhour.SelectedItem.ToString();
                    iToHour = cmb_Acc_Tohour.SelectedItem.ToString();

                    if (Convert.ToInt32(iFromHour) == 24)
                        oFromDate = AccFromDate + " 23:59:59";
                    else
                        oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                    if (Convert.ToInt32(iToHour) == 24)
                        oTodate = AccTodate + " 23:59:59";
                    else
                        oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                    frmOnlineProcessed_Details OPD = new frmOnlineProcessed_Details(Convert.ToInt32(oItems.ICLIENT_ID), oItems.SLOCATION_ID.ToString(), Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate), Convert.ToInt32(cmb_Acc_Status.SelectedValue));
                    if (OPD.ShowDialog() == DialogResult.OK)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.LOGIN_NAME, Environment.MachineName);
            }
        }

        private void btn_Hourlyreport_view_Click(object sender, EventArgs e)
        {
            try
            {
                BusinessLogic.oMessageEvent.Start("Transferring data.");
                List<BusinessLogic.LINECOUNT_DETAILS> oLinecount = new List<BusinessLogic.LINECOUNT_DETAILS>();

                string sProductionID = (cmb_Log_Employee.SelectedValue.ToString());
                ListItem_LineCountFileItem oListItem;
                int iRowCount = 0;
                int iTot_Mins = 0;
                int iConv_Mins = 0;
                decimal dTotal_Lines = 0;
                decimal dTotal_Conv_Lines = 0;

                string oMins = string.Empty;
                string oConMins = string.Empty;
                double dblLines = 0;
                double dblConvertedLines = 0;
                double dblErrorPoints = 0;
                double dblAccuracy = 0;

                object oFromDate, oTodate;
                lsv_Hourlyreport.Items.Clear();
                string sFromDate, sToDate;

                DataSet _dsLineCountReport = new DataSet();

                if (Convert.ToInt32(cmb_Log_fromhr.Text.Trim()) == 24)
                    sFromDate = Convert.ToDateTime(dtp_Log_HourlyFrom.Value).ToString("yyyy/MM/dd") + " 23:59:59";
                else
                    sFromDate = Convert.ToDateTime(dtp_Log_HourlyFrom.Value).ToString("yyyy/MM/dd") + " " + Convert.ToInt32(cmb_Log_fromhr.Text.Trim()) + ":" + "00:00";

                if (Convert.ToInt32(cmb_Log_tohr.Text.Trim()) == 24)
                    sToDate = Convert.ToDateTime(dtp_Log_HourlyTo.Value).ToString("yyyy/MM/dd") + " 23:59:59";
                else
                    sToDate = Convert.ToDateTime(dtp_Log_HourlyTo.Value).ToString("yyyy/MM/dd") + " " + Convert.ToInt32(cmb_Log_tohr.Text.Trim()) + ":" + "00:00";
                
                BusinessLogic.WS_Allocation.Timeout = 500000;
                _dsLineCountReport = BusinessLogic.WS_Allocation.Get_Transcription_view_V2(Convert.ToInt32(sProductionID), Convert.ToDateTime(sFromDate), Convert.ToDateTime(sToDate));

                if (_dsLineCountReport == null)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                if (_dsLineCountReport.Tables.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                if (_dsLineCountReport.Tables[0].Rows.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No record found.");
                    return;
                }

                foreach (DataRow _drTobegraded in _dsLineCountReport.Tables[0].Select())
                {
                    oLinecount.Add(new BusinessLogic.LINECOUNT_DETAILS(_drTobegraded["client_name"].ToString(), _drTobegraded["location_name"].ToString(),
                        _drTobegraded["doctor_full_name"].ToString(), _drTobegraded["report_name"].ToString(), Convert.ToDateTime(_drTobegraded["file_date"].ToString()),
                        _drTobegraded["file_minutes"].ToString(), _drTobegraded["Converted_minutes"].ToString(), Convert.ToDecimal(_drTobegraded["file_lines"].ToString()),
                        Convert.ToDecimal(_drTobegraded["converted_lines"].ToString()), Convert.ToDateTime(_drTobegraded["submitted_time"].ToString()),
                        Convert.ToDateTime(_drTobegraded["Submit_Time"].ToString()),
                        _drTobegraded["evaluated_date"].ToString(), _drTobegraded["transcription_status_description_1"].ToString(), _drTobegraded["transcription_status_description_1"].ToString(),
                        _drTobegraded["template_description"].ToString(), Convert.ToDecimal(_drTobegraded["accuracy"].ToString()), string.Empty,
                        Convert.ToInt32(_drTobegraded["CalSec"].ToString()), Convert.ToDecimal(_drTobegraded["Converted_Seconds"].ToString())));
                }

                for (var day = Convert.ToDateTime(dtp_Log_HourlyFrom.Value).Date; day.Date <= Convert.ToDateTime(dtp_Log_HourlyTo.Value).Date; day = day.AddDays(1))
                {

                    DateTime end = Convert.ToDateTime(day);
                    DateTime start = Convert.ToDateTime(day);

                    if (Convert.ToInt32(cmb_Log_tohr.Text.Trim()) < Convert.ToInt32(cmb_Log_fromhr.Text.Trim()))
                    {
                        start = Convert.ToDateTime(day);
                        end = Convert.ToDateTime(day).AddDays(1);
                    }

                    if (Convert.ToInt32(cmb_Log_fromhr.Text.Trim()) == 24)
                        oFromDate = Convert.ToDateTime(start).ToString("yyyy/MM/dd") + " 23:59:59";
                    else
                        oFromDate = Convert.ToDateTime(start).ToString("yyyy/MM/dd") + " " + Convert.ToInt32(cmb_Log_fromhr.Text.Trim()) + ":" + "00:00";

                    if (Convert.ToInt32(cmb_Log_tohr.Text.Trim()) == 24)
                        oTodate = Convert.ToDateTime(end).ToString("yyyy/MM/dd") + " 23:59:59";
                    else
                        oTodate = Convert.ToDateTime(end).ToString("yyyy/MM/dd") + " " + Convert.ToInt32(cmb_Log_tohr.Text.Trim()) + ":" + "00:00";

                    object oStart = oFromDate;
                    object oEnd = oTodate;


                    //string sDay_Filter = " submitted_time<='" + oStart + "' and submitted_time>='" + oEnd + "' ";
                    string sDay_Filter = " submitted_time>='" + Convert.ToDateTime(oStart) + "' and submitted_time<=  '" + Convert.ToDateTime(oEnd) + "' ";
                    DataRow[] drView = _dsLineCountReport.Tables[0].Select(sDay_Filter);

                    foreach (DataRow dr in drView)
                    {
                        oListItem = new ListItem_LineCountFileItem(dr, iRowCount++);
                        lsv_Hourlyreport.Items.Add(oListItem);
                    }

                    var Tot_files = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(oEnd)) select c).Count();
                    var Tot_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(oEnd)) group c by c.FILE_SEC into CP select new { TOTAL_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.FILE_SEC)) });
                    var Tot_Conv_Mins = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(oEnd)) group c by c.CONV_SEC into CP select new { TOTAL_CONV_FILE_SEC = CP.Sum(c => Convert.ToInt32(c.CONV_SEC)) });

                    var Total_File_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(oEnd)) group c by c.FILELINES into CP select new { TOTAL_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.FILELINES)) });
                    var Total_File_Conv_Lines = (from c in oLinecount where (Convert.ToDateTime(c.SUBMITTEDTIME) >= Convert.ToDateTime(oStart) && Convert.ToDateTime(c.SUBMITTEDTIME) <= Convert.ToDateTime(oEnd)) group c by c.CONVLINES into CP select new { TOTAL_CONV_FILE_LINES = CP.Sum(c => Convert.ToInt32(c.CONVLINES)) });

                    foreach (var File_Mins in Tot_Mins)
                        iTot_Mins += Convert.ToInt32(File_Mins.TOTAL_FILE_SEC);

                    foreach (var File_Conv_Mins in Tot_Conv_Mins)
                        iConv_Mins += Convert.ToInt32(File_Conv_Mins.TOTAL_CONV_FILE_SEC);

                    foreach (var File_Lines in Total_File_Lines)
                        dTotal_Lines += Convert.ToInt32(File_Lines.TOTAL_FILE_LINES);

                    foreach (var File_Conv_Lines in Total_File_Conv_Lines)
                        dTotal_Conv_Lines += Convert.ToInt32(File_Conv_Lines.TOTAL_CONV_FILE_LINES);

                    oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), Tot_Mins.ToString(), Tot_Conv_Mins.ToString(), Math.Round(dblLines, 2).ToString(), Math.Round(dblConvertedLines, 2).ToString(), dblErrorPoints.ToString(), dblAccuracy.ToString());
                    if (Tot_files > 0)
                    {
                        oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Conv_Lines, 2).ToString(), string.Empty, string.Empty);
                        lsv_Hourlyreport.Items.Add(oListItem);
                    }
                    //oListItem = new ListItem_LineCountFileItem("Linecount summary for the date : " + Convert.ToDateTime(day).ToString("dd-MM-yyyy") + " Total Files : " + Tot_files.ToString(), sGetDuration(iTot_Mins).ToString(), sGetDuration(iConv_Mins).ToString(), Math.Round(dTotal_Lines, 2).ToString(), Math.Round(dTotal_Lines, 2).ToString(), string.Empty, string.Empty);
                    //lsv_Hourlyreport.Items.Add(oListItem);

                    iTot_Mins = 0;
                    iConv_Mins = 0;
                    dTotal_Lines = 0;
                    dTotal_Conv_Lines = 0;
                    Tot_files = 0;
                }
                BusinessLogic.Reset_ListViewColumn(lsv_Hourlyreport);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmb_Log_branch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sBranchId = cmb_Log_branch.SelectedValue.ToString();
            sBranch_ID = sBranchId;
            Load_Employee_Full_name(sDesgination_ID, sBranchId);
        }

        private void cmb_Log_desig_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = cmb_Log_desig.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesgination_ID, sBranch_ID);

        }

        private void btn_AddIncentive_Click(object sender, EventArgs e)
        {
            try
            {
                frmAdd_Incentives FAI = new frmAdd_Incentives();
                FAI.ShowDialog();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_Cust_Group_Click(object sender, EventArgs e)
        {
            try
            {
                frm_customized_Group FCG = new frm_customized_Group();
                if (FCG.ShowDialog() == DialogResult.OK)
                {
                    Load_Customized_group();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_ExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");

                if (tabControlMain.SelectedTab.Name == "tbpHangUp")
                {
                    if (tbpHabgUpProcess.SelectedTab.Name == "tabPage15")
                    {
                        string sFileName = "HangUp Process - Higher Lines for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvHigherLines, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage16")
                    {
                        string sFileName = "HangUp Process - No Dictation for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvNoDicFiles, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage17")
                    {
                        string sFileName = "HangUp Process - Trans Edit Files for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvTransEdit, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage18")
                    {
                        string sFileName = "HangUp Process - Split Minutes for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvSplitMinutesUpdated, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage19")
                    {
                        string sFileName = "HangUp Process - Manual Entry for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvManualEntry, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage20")
                    {
                        string sFileName = "HangUp Process - Zero minutes for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvZeroMinutesUpdated, sFolderNAme, sFileName);
                    }
                    else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage21")
                    {
                        string sFileName = "HangUp Process - Zero Lines for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                        ExportToExcel(lvZeroLinesUpdated, sFolderNAme, sFileName);
                    }

                    //else if (tbpHabgUpProcess.SelectedTab.Name == "tabPage22")
                    //{
                    //    string sFileName = "HangUp Process - Client Lines for the month of  " + cmbHangUpMonth.SelectedText + " " + " and Year " + cmbHangUpYear.SelectedText + ".xls";
                    //    ExportToExcel(lvClientLinesUpdated, sFolderNAme, sFileName);
                    //}
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }

        }

        private void btn_Excel_Click(object sender, EventArgs e)
        {
            try
            {
                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "Incentive Details For Online Accounts From" + "_date between_" + Convert.ToDateTime(dtp_incentive_fromdate.Value).ToString("YYYY-MM-dd") + "_and_" + Convert.ToDateTime(dtp_incentive_todate.Value).ToString("YYYY-MM-dd") + ".xls";
                ExportToExcel(lsv_Incentive, sFolderNAme, sFileName);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmb_Customize_Group_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Customized_Employee_List();
        }

        private void cbxHourPlatform_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Load_Location_Type(Convert.ToInt32(cbxHourPlatform.SelectedValue));
                CLIENT_TYPE_ID = (Convert.ToInt32(cbxHourPlatform.SelectedValue));
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void activeDeactiveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string sProduction_ID = string.Empty;
                ListIte_MappingDeatils_Priority lsvMap = (ListIte_MappingDeatils_Priority)lsvMappingDeatils.SelectedItems[0];
                sProduction_ID = lsvMap.PRODUCTION_ID.ToString();

                int iSetActive = BusinessLogic.WS_Allocation.Set_allocation_active(Convert.ToInt32(sProduction_ID), 0);
                if (iSetActive > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Successfully activated");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void deActivateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string sProduction_ID = string.Empty;
                ListIte_MappingDeatils_Priority lsvMap = (ListIte_MappingDeatils_Priority)lsvMappingDeatils.SelectedItems[0];
                sProduction_ID = lsvMap.PRODUCTION_ID.ToString();

                int iSetActive = BusinessLogic.WS_Allocation.Set_allocation_active(Convert.ToInt32(sProduction_ID), 1);
                if (iSetActive > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Successfully removed");
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btnAddTat_Click(object sender, EventArgs e)
        {
            try
            {
                string sTat = txtTatPercentage.Text;
                string sLocation = cmbNTSClients.SelectedValue.ToString();
                string sDate = dtpNtsTat.Text.ToString();

                int iSetTat = BusinessLogic.WS_Allocation.Set_Tat_NTS(Convert.ToDecimal(sTat), sLocation, Convert.ToDateTime(sDate));
                if (iSetTat > 0)
                {
                    Load_Tat_percentage_NTS(sDate);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_Processed_Status_Click(object sender, EventArgs e)
        {
            Thread tProcessedThread = new Thread(Load_Processed_Thread);
            tProcessedThread.Start();
        }

        private void dtpNtsTat_ValueChanged(object sender, EventArgs e)
        {
            Load_Tat_percentage_NTS(dtpNtsTat.Text);
        }

        private void Load_Tat_percentage_NTS(string dTDate)
        {
            try
            {
                string sDate = Convert.ToDateTime(dTDate).ToString("yyyy/MM/dd");
                DataSet _dsTatPercentage = new DataSet();
                _dsTatPercentage = BusinessLogic.WS_Allocation.Get_Nts_TAT(sDate);

                int iRowCount = 0;
                lsvNTSTatPercentage.Items.Clear();
                foreach (DataRow _drRow in _dsTatPercentage.Tables[0].Select())
                    lsvNTSTatPercentage.Items.Add(new List_item_NTS_Tat_Percentage(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsvNTSTatPercentage);

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void Load_Processed_Thread()
        {
            try
            {
                string iFromHour, iToHour;
                object oFromDate = null;
                object oTodate = null;

                //BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");

                string AccFromDate = Convert.ToDateTime(dtpAccountFrom.Text).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtpAccountTo.Text).ToString("yyyy-MM-dd");

                iFromHour = cmb_Acc_Fromhour.SelectedItem.ToString();
                iToHour = cmb_Acc_Tohour.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = AccFromDate + " 23:59:59";
                else
                    oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = AccTodate + " 23:59:59";
                else
                    oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                frmOnlineProcessed_Status PS = new frmOnlineProcessed_Status(Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate));
                PS.ShowDialog();
                BusinessLogic.oMessageEvent.Start("Done....!");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void Load_Processed_TL_Thread()
        {
            try
            {
                string iFromHour, iToHour;
                object oFromDate = null;
                object oTodate = null;

                //BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");

                string AccFromDate = Convert.ToDateTime(dtp_AccTL_Fromdate.Value).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtp_AccTL_Todate.Value).ToString("yyyy-MM-dd");

                iFromHour = cmb_AccTL_FromHRs.SelectedItem.ToString();
                iToHour = cmb_AccTL_ToHrs.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = AccFromDate + " 23:59:59";
                else
                    oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = AccTodate + " 23:59:59";
                else
                    oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                frmProcessed_Minutes_TL PS = new frmProcessed_Minutes_TL(Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate));
                PS.ShowDialog();
                BusinessLogic.oMessageEvent.Start("Done....!");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmbClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataSet _dsLocation = new DataSet();
                _dsLocation = BusinessLogic.WS_Allocation.Get_Location(Convert.ToInt32(cmbClient.SelectedValue));

                cmbLocation.DisplayMember = "" + Framework.LOCATION.FIELD_LOCATION_NAME_STR + "";
                cmbLocation.ValueMember = "" + Framework.LOCATION.FIELD_LOCATION_ID_STR + "";
                cmbLocation.DataSource = _dsLocation.Tables[0];
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmbLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int Account = Convert.ToInt32(cmbClient.SelectedValue.ToString());
                string Location = cmbLocation.SelectedValue.ToString();
                DataSet _dsDoctor = BusinessLogic.WS_Allocation.Get_Doctor_V1(Account, Location);
                DataRow drd = _dsDoctor.Tables[0].NewRow();
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + ""] = 0;
                drd["" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + ""] = "-- ALL --";
                _dsDoctor.Tables[0].Rows.InsertAt(drd, 0);

                cmbDoctor.DisplayMember = "" + Framework.DOCTOR.FIELD_DOCTOR_FULL_NAME_STR + "";
                cmbDoctor.ValueMember = "" + Framework.DOCTOR.FIELD_DOCTOR_ID_BINT + "";
                cmbDoctor.DataSource = _dsDoctor.Tables[0]; ;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btnautoallocationView_Click(object sender, EventArgs e)
        {
            try
            {
                int iDoctorID = Convert.ToInt32(cmbDoctor.SelectedValue);
                string iLocationID = cmbLocation.SelectedValue.ToString();
                DateTime dFromdate = Convert.ToDateTime(dtp_Filedate.Value);
                DateTime dTodate = Convert.ToDateTime(dtp_FileTodate.Value);
                int iBatch = Convert.ToInt32(cmb_Auto_Batch.SelectedValue);
                int iOption = 0;

                if (rdb_Doc_All.Checked == true)
                    iOption = 0;
                else if (rdb_Doc_login.Checked == true)
                    iOption = 1;
                else
                    iOption = 2;


                DataTable dsAutoAllocationDoctorwise = BusinessLogic.WS_Allocation.Get_priorityFiles_Doctorwise_V2(iDoctorID, iLocationID, (ChkInactive.Checked) ? 1 : 0, iBatch, iOption, dFromdate, dTodate);

                int rowcount = 1;
                if (dsAutoAllocationDoctorwise == null)
                    return;
                lsvAutoAllocationUserDetails.Items.Clear();
                foreach (DataRow dr in dsAutoAllocationDoctorwise.Rows)
                {
                    lsvAutoAllocationUserDetails.Items.Add(new Mylistitem_AutoAllocationUsers(dr, rowcount));
                    rowcount++;
                }

                Load_Doctorwise_filecount();
                BusinessLogic.Reset_ListViewColumn(lsvAutoAllocationUserDetails);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btnautoallocationSave_Click(object sender, EventArgs e)
        {
            try
            {
                int iResult = 0;
                if (lsvAutoAllocationUserDetails.SelectedItems.Count == 0)
                    return;

                if (Convert.ToInt32(txtFileCount.Text) == 0)
                {
                    BusinessLogic.oMessageEvent.Start("Enter the file count");
                    txtFileCount.Focus();
                    return;
                }

                foreach (Mylistitem_AutoAllocationUsers oUser in lsvAutoAllocationUserDetails.SelectedItems)
                {
                    iResult = BusinessLogic.WS_Allocation.Set_Autoallocation_Filecount(Convert.ToInt32(oUser.ICLIENTID), oUser.SLOCATIONID.ToString(), Convert.ToInt32(oUser.IDOCTORID), Convert.ToInt32(txtFileCount.Text), Convert.ToInt32(oUser.IPRODUCTION_ID), Convert.ToDateTime(dtp_Filedate.Value));
                }

                btnautoallocationView_Click(this, e);
                BusinessLogic.oMessageEvent.Start("Saved Successfully");

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }

        }

        //private void btnautoallocationSave_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //         int iDoctorID = Convert.ToInt32(cmbDoctor.SelectedValue);
        //        string iLocationID = cmbLocation.SelectedValue.ToString();
        //        int ifileCount = Convert.ToInt32(txtFileCount.Text);

        //        int AutoAllocationDoctorwise = BusinessLogic.WS_Allocation.Set_priorityFiles_Doctorwise(iDoctorID, iLocationID, ifileCount);
        //        if (AutoAllocationDoctorwise > 0)
        //            btnautoallocationView_Click(this, e);

        //    }
        //    catch (Exception ex)
        //    {
        //        BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
        //    }

        //}

        private void lsvMappingDeatils_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            this.lsvMappingDeatils.ListViewItemSorter = new ListViewItemComparer(e.Column);
        }

        private void btnDeallotHistory_Click(object sender, EventArgs e)
        {
            try
            {
                frmdeallotcatioHistory objdeallot = new frmdeallotcatioHistory();
                if (objdeallot.ShowDialog() == DialogResult.OK)
                    BusinessLogic.oMessageEvent.Start("Done.");
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_QuickAlert_Click(object sender, EventArgs e)
        {
            try
            {
                frmQuickAllot FQA = new frmQuickAllot();
                if (FQA.ShowDialog() == DialogResult.OK)
                {

                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_View_Week_Click(object sender, EventArgs e)
        {
            try
            {
                Load_Weekly_Mins();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btn_Export_Weekly_Click(object sender, EventArgs e)
        {
            try
            {
                string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
                string sFileName = "Weekly Processed Minutes " + ".xls";
                ExportToExcel(lsv_WeeklyProcessedMins, sFolderNAme, sFileName);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void moveToAnotherToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                DataTable dtAnother = BusinessLogic.WS_QualityService.Get_Employee(0).Tables[3];

                cmbEmpAnother.DisplayMember = "emp_full_name";
                cmbEmpAnother.ValueMember = "production_id";
                cmbEmpAnother.DataSource = dtAnother;
                PNLMOVEaNOTHER.Visible = true;

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void BtnMoveAnother_Click(object sender, EventArgs e)
        {
            int nCount = 0;
            foreach (ListItem_LineCountFileItem oitem in lsvLineCountDetails.SelectedItems)
            {
                nCount = BusinessLogic.WS_Allocation.Set_MoveFiletoAnother(oitem.REPORT_TRANSACTION_ID, Convert.ToInt32(cmbLogUser.SelectedValue), Convert.ToInt32(cmbEmpAnother.SelectedValue.ToString().Split('/').GetValue(0)), Convert.ToInt32(BusinessLogic.SPRODUCTIONID), Environment.MachineName.ToString());
            }
            if (nCount > 0)
            {
                MessageBox.Show("Moved Suceesfully");
                PNLMOVEaNOTHER.Visible = false;
            }
        }

        private void btnmoveCansel_Click(object sender, EventArgs e)
        {
            PNLMOVEaNOTHER.Visible = false;
        }

        private void cmsMovefiletoAnotherID_Opening(object sender, CancelEventArgs e)
        {
            if (Convert.ToInt32(BusinessLogic.SPRODUCTIONID) == 185)
                lsvLineCountDetails.ContextMenuStrip = cmsMovefiletoAnotherID;
            else
                lsvLineCountDetails.ContextMenuStrip = null;
        }

        private void btn_TempAllot_Click(object sender, EventArgs e)
        {
            try
            {
                frm_TempeAllot FQA = new frm_TempeAllot();
                if (FQA.ShowDialog() == DialogResult.OK)
                {

                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void btnempExport_Click(object sender, EventArgs e)
        {

            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Current Employee Report.xls";
            ExportToExcel(lsvAllEmployee, sFolderNAme, sFileName);
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                GetEmployeeList();
                Load_Emp_Consolidation();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmb_Emp_Workplatform_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmb_Emp_Workplatform.SelectedIndex >= 0)
                {
                    GetEmployeeList();
                    Load_Emp_Consolidation();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void cmb_Emp_Branch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                GetEmployeeList();
                Load_Emp_Consolidation();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        /// <summary>
        /// View the 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lsv_Quick_MT_TAT_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (lsv_Quick_MT_TAT.SelectedItems.Count == 1)
                {
                    MTTat_contextMenuStrip.Items["mTToolStripMenuItem"].Visible = true;
                    MTTat_contextMenuStrip.Items["mTToolStripMenuItem"].Enabled = true;
                    lsv_Quick_MT_TAT.ContextMenuStrip = MTTat_contextMenuStrip;
                }
                else
                {
                    MTTat_contextMenuStrip.Items["mTToolStripMenuItem"].Enabled = false;
                    MTTat_contextMenuStrip.Items["mTToolStripMenuItem"].Visible = false;
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void mTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Load_Emp_List();
        }

        private void deallotToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (lsv_Quick_MT_TAT.SelectedItems.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No Item is selected for deallocation.");
                    return;
                }

                BusinessLogic.oMessageEvent.Start("Transferring data..");
                BusinessLogic.oProgressEvent.Start(true);
                this.Cursor = Cursors.WaitCursor;

                string _sTanscriptionCollection = string.Empty;
                string _sVoiceFile_ID = string.Empty;
                foreach (MyListItem_QuickAllot_MT_TAT oDeAllocationItem in lsv_Quick_MT_TAT.SelectedItems)
                {
                    _sTanscriptionCollection = oDeAllocationItem.ITRANSCRIPTION_ID.ToString();
                    _sVoiceFile_ID = oDeAllocationItem.SVOICE_FILE_ID.ToString();

                    int iDeAllttotFiles = BusinessLogic.WS_Allocation.Set_Deallot_Files(Convert.ToInt32(_sTanscriptionCollection), _sVoiceFile_ID, Convert.ToInt32(oDeAllocationItem.SUSER_ID));
                    if (iDeAllttotFiles > 0)
                    {
                        Load_MTFiles_List();
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Failed");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                this.Cursor = Cursors.Default;
            }
        }

        private void reAllotToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Load_Emp_List();
        }

        private void tATToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyListItem_QuickAllot_MT_TAT oItem = (MyListItem_QuickAllot_MT_TAT)lsv_Quick_MT_TAT.SelectedItems[0];

                foreach (MyListItem_QuickAllot_MT_TAT oFile in lsv_Quick_MT_TAT.SelectedItems)
                {
                    int iSetTat = BusinessLogic.WS_Allocation.set_tat(Convert.ToInt32(oFile.ITRANSCRIPTION_ID));
                    if (iSetTat > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Tat Set..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void btnViewPer_Click(object sender, EventArgs e)
        {
            string sYear = string.Empty;
            int sMonth = 0;
            int OptionId = 0;
            string sFromDate = string.Empty;
            string sToDate = string.Empty;

            if (chkByDate.Checked == false)
            {
                sYear = cmbTatYear.SelectedItem.ToString();
                sMonth = cmbTatMonth.SelectedIndex + 1;
                sFromDate = null;
                sToDate = null;
                OptionId = 1;
            }
            else
            {
                sYear = "-1";
                sMonth = -1;
                sFromDate = drpTatPercFromDate.Text.ToString();
                sToDate = dtpTatPercToDate.Text.ToString();
                OptionId = 2;
            }
            LoadTatPercentageFiles(sYear, sMonth, sFromDate, sToDate, OptionId);
        }

        private void btnExportTatPerc_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "TAT Percentage Report" + ".xls";
            ExportToExcel(lvTatPercentage, sFolderNAme, sFileName);
        }

        private void lsv_TATMonitor_ME_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (lsv_TATMonitor_ME.SelectedItems.Count == 1)
                {
                    METAT_contextMenuStrip.Items["METattoolStripMenuItem"].Visible = true;
                    METAT_contextMenuStrip.Items["METattoolStripMenuItem"].Enabled = true;
                    lsv_TATMonitor_ME.ContextMenuStrip = METAT_contextMenuStrip;
                }
                else
                {
                    METAT_contextMenuStrip.Items["METattoolStripMenuItem"].Enabled = false;
                    METAT_contextMenuStrip.Items["METattoolStripMenuItem"].Visible = false;
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
        }

        private void METattoolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                double dTotmins = 0;
                int iTotalFiles = 0;
                MyListItem_QuickAllot_ME_TAT oItem = (MyListItem_QuickAllot_ME_TAT)lsv_TATMonitor_ME.SelectedItems[0];
                foreach (MyListItem_QuickAllot_ME_TAT oFile in lsv_TATMonitor_ME.SelectedItems)
                {
                    dTotmins += Convert.ToDouble(oItem.SDURATION);
                    iTotalFiles++;
                }
                frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.ME), Convert.ToInt32(oItem.ICLIENT_ID), Convert.ToInt32(oItem.IDOCTOR_ID), iTotalFiles, Convert.ToInt32(dTotmins), 2, 2);
                if (ofe.ShowDialog() == DialogResult.OK)
                {
                    foreach (MyListItem_QuickAllot_ME_TAT oFile in lsv_TATMonitor_ME.SelectedItems)
                    {
                        oFile.SSTATUS = "Allotted";
                        oFile.SEMP_NAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.SUSER_ID = BusinessLogic.ALLOTEDUSERID;

                        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;


                        //insert into database
                        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.ITRANSCRIPTION_ID, oFile.SUSER_ID, DateTime.Now, oFile.SVOICE_FILE_ID, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                    }
                    int filecount = lsv_Quick_MT_TAT.SelectedItems.Count;
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
            finally
            {
                Load_MEFiles_List();
            }
        }

        private void ME_Deallot_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (lsv_TATMonitor_ME.SelectedItems.Count <= 0)
                {
                    BusinessLogic.oMessageEvent.Start("No Item is selected for deallocation.");
                    return;
                }

                BusinessLogic.oMessageEvent.Start("Transferring data..");
                BusinessLogic.oProgressEvent.Start(true);
                this.Cursor = Cursors.WaitCursor;

                string _sTanscriptionCollection = string.Empty;
                string _sVoiceFile_ID = string.Empty;
                foreach (MyListItem_QuickAllot_ME_TAT oDeAllocationItem in lsv_TATMonitor_ME.SelectedItems)
                {
                    _sTanscriptionCollection = oDeAllocationItem.ITRANSCRIPTION_ID.ToString();
                    _sVoiceFile_ID = oDeAllocationItem.SVOICE_FILE_ID.ToString();

                    int iDeAllttotFiles = BusinessLogic.WS_Allocation.Set_Deallot_Files(Convert.ToInt32(_sTanscriptionCollection), _sVoiceFile_ID, Convert.ToInt32(oDeAllocationItem.SUSER_ID));
                    if (iDeAllttotFiles > 0)
                    {
                        Load_MEFiles_List();
                    }
                    else
                    {
                        BusinessLogic.oMessageEvent.Start("Failed");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                this.Cursor = Cursors.Default;
            }
        }

        private void ME_ReAllot_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                #region "Commented"
                //double dTotmins = 0;
                //int iTotalFiles = 0;
                //MyListItem_QuickAllot_ME_TAT oItem = (MyListItem_QuickAllot_ME_TAT)lsv_TATMonitor_ME.SelectedItems[0];
                //foreach (MyListItem_QuickAllot_ME_TAT oFile in lsv_TATMonitor_ME.SelectedItems)
                //{
                //    dTotmins += Convert.ToDouble(oItem.SDURATION);
                //    iTotalFiles++;
                //}
                //frmEmployee ofe = new frmEmployee(Convert.ToInt32(Framework.Variables.Employee_Mode.MT), Convert.ToInt32(oItem.ICLIENT_ID), Convert.ToInt32(oItem.IDOCTOR_ID), iTotalFiles, Convert.ToInt32(dTotmins), 2, 2);
                //if (ofe.ShowDialog() == DialogResult.OK)
                //{
                //    foreach (MyListItem_QuickAllot_ME_TAT oFile in lsv_TATMonitor_ME.SelectedItems)
                //    {
                //        oFile.SSTATUS = "Allotted";
                //        oFile.SEMP_NAME = BusinessLogic.ALLOTEDUSERNAME;
                //        oFile.SUSER_ID = BusinessLogic.ALLOTEDUSERID;

                //        oFile.OFFLINE_P_EMPNAME = BusinessLogic.ALLOTEDUSERNAME;
                //        oFile.OFFLINE_P_USERID = BusinessLogic.ALLOTEDUSERID;


                //        //insert into database
                //        //int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles(oFile.TRANSCRIPTIONID, oFile.OFFLINE_P_USERID, DateTime.Now, oFile.OFFLINE_P_VOICE_FILE_NAME, BusinessLogic.iTATREQUIRED);
                //        int iResult = BusinessLogic.WS_Allocation.Set_AllocationFiles_V2(oFile.ITRANSCRIPTION_ID, oFile.SUSER_ID, DateTime.Now, oFile.SVOICE_FILE_ID, BusinessLogic.iTATREQUIRED, Environment.UserName, Environment.MachineName);

                //    }
                //    int filecount = lsv_Quick_MT_TAT.SelectedItems.Count;
                //    int iUpdate = BusinessLogic.WS_Allocation.Set_Allotedlines(Convert.ToInt32(oItem.OFFLINE_P_USERID.ToString()), filecount, dTotmins);
                //}
                #endregion

                foreach (MyListItem_QuickAllot_ME_TAT oItem in lsv_TATMonitor_ME.SelectedItems)
                {
                    int iUpdate = BusinessLogic.WS_Allocation.Set_Reallocationfiles(Convert.ToInt32(oItem.ITRANSCRIPTION_ID), Convert.ToInt32(oItem.IPRODUCTION_ID));
                    if (iUpdate > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Allocated Successfully to " + oItem.ALLOTED_PTAG_ID.ToString());
                    }
                }

            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), Environment.UserName, Environment.MachineName);
            }
            finally
            {
                Load_MEFiles_List();
            }
        }

        private void ME_TAT_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MyListItem_QuickAllot_ME_TAT oItem = (MyListItem_QuickAllot_ME_TAT)lsv_TATMonitor_ME.SelectedItems[0];

                foreach (MyListItem_QuickAllot_ME_TAT oFile in lsv_TATMonitor_ME.SelectedItems)
                {
                    int iSetTat = BusinessLogic.WS_Allocation.set_tat(Convert.ToInt32(oFile.ITRANSCRIPTION_ID));
                    if (iSetTat > 0)
                    {
                        BusinessLogic.oMessageEvent.Start("Tat Set..!");
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
            }
        }

        private void btn_View_Click(object sender, EventArgs e)
        {
            Load_MTFiles_List();
        }

        private void btn_ME_View_Click(object sender, EventArgs e)
        {
            Load_MEFiles_List();
        }

        private void btn_AccTL_View_Click(object sender, EventArgs e)
        {
            try
            {
                string iFromHour, iToHour;
                object oFromDate = null;
                object oTodate = null;

                lsv_AccTL.Items.Clear();
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Processing Request....!");
                ListItem_AccountWiseInfo oListItem;

                string AccFromDate = Convert.ToDateTime(dtp_AccTL_Fromdate.Text).ToString("yyyy-MM-dd");
                string AccTodate = Convert.ToDateTime(dtp_AccTL_Todate.Text).ToString("yyyy-MM-dd");
                DataTable _dsAccountWiseInfo = null;

                iFromHour = cmb_AccTL_FromHRs.SelectedItem.ToString();
                iToHour = cmb_AccTL_ToHrs.SelectedItem.ToString();

                if (Convert.ToInt32(iFromHour) == 24)
                    oFromDate = AccFromDate + " 23:59:59";
                else
                    oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                if (Convert.ToInt32(iToHour) == 24)
                    oTodate = AccTodate + " 23:59:59";
                else
                    oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                _dsAccountWiseInfo = BusinessLogic.WS_Allocation.Get_accountWise_Processed_Info_V2(oFromDate, oTodate, Convert.ToInt32(cmb_AccTL_Status.SelectedValue));

                int iRowCount = 0;
                int dTotMins = 0;

                foreach (DataRow _drRow in _dsAccountWiseInfo.Select())
                {
                    lsv_AccTL.Items.Add(new ListItem_AccountWiseInfo(_drRow, iRowCount++));
                    if (Convert.ToInt32(_drRow["Tot_minutes"]) != 0)
                    {
                        if (_drRow["Tot_minutes"].ToString().Contains('.'))
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString().Split('.').GetValue(0));
                        else
                            dTotMins += Convert.ToInt32(_drRow["Tot_minutes"].ToString());
                    }
                }

                string oMins = sGetDuration(dTotMins);

                oListItem = new ListItem_AccountWiseInfo("Account Wise Minutes on: " + Convert.ToDateTime(AccFromDate).ToString("dd-MM-yyyy"), oMins.ToString());
                lsv_AccTL.Items.Add(oListItem);
                BusinessLogic.Reset_ListViewColumn(lsv_AccTL);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
                BusinessLogic.oMessageEvent.Start("Done..!");
            }
        }

        private void btn_AccTL_MinsWithStatus_Click(object sender, EventArgs e)
        {
            try
            {
                Thread tProcessedThread_TL = new Thread(Load_Processed_TL_Thread);
                tProcessedThread_TL.Start();
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void lsv_AccTL_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (lsv_AccTL.SelectedItems.Count > 0)
                {
                    ListItem_AccountWiseInfo oItems = (ListItem_AccountWiseInfo)lsv_AccTL.SelectedItems[0];
                    string iFromHour, iToHour;
                    object oFromDate = null;
                    object oTodate = null;

                    string AccFromDate = Convert.ToDateTime(dtp_AccTL_Fromdate.Text).ToString("yyyy-MM-dd");
                    string AccTodate = Convert.ToDateTime(dtp_AccTL_Todate.Text).ToString("yyyy-MM-dd");

                    iFromHour = cmb_AccTL_FromHRs.SelectedItem.ToString();
                    iToHour = cmb_AccTL_ToHrs.SelectedItem.ToString();

                    if (Convert.ToInt32(iFromHour) == 24)
                        oFromDate = AccFromDate + " 23:59:59";
                    else
                        oFromDate = AccFromDate + " " + iFromHour + ":" + "00:00";

                    if (Convert.ToInt32(iToHour) == 24)
                        oTodate = AccTodate + " 23:59:59";
                    else
                        oTodate = AccTodate + " " + iToHour + ":" + "00:00";

                    frmOnlineProcessed_Details OPD = new frmOnlineProcessed_Details(Convert.ToInt32(oItems.ICLIENT_ID), oItems.SLOCATION_ID.ToString(), Convert.ToDateTime(oFromDate), Convert.ToDateTime(oTodate), Convert.ToInt32(cmb_Acc_Status.SelectedValue));
                    if (OPD.ShowDialog() == DialogResult.OK)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
        }

        private void btn_AccTL_Excel_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Account Wise File Details For Online Accounts From" + "_date between_" + dtpAccountFrom.Text + "_and_" + dtpAccountTo.Text + ".xls";
            ExportToExcel(lvAccountWiseInfo, sFolderNAme, sFileName);
        }


        /// <summary>
        /// METHOD TO LOAD TRACKING DETAILS
        /// </summary>
        private void Load_TL_File_Tracking()
        {
            try
            {
                BusinessLogic.oProgressEvent.Start(true);
                BusinessLogic.oMessageEvent.Start("Transferring Data..!");
                lsv_TL_Names.Items.Clear();

                DataSet _dsTypistMap = new DataSet();
                string iBatchID = cmb_TL_Batch.SelectedValue.ToString();
                string iWork_Platform = "1";

                _dsTypistMap = BusinessLogic.WS_Allocation.GET_ALLOTED_DETAILS_NEW_V2_WORKPLATFORM(Convert.ToInt32(iBatchID), Convert.ToInt32(iWork_Platform));

                int iRowCount = 0;
                foreach (DataRow _drRow in _dsTypistMap.Tables[0].Select())
                    lsv_TL_Names.Items.Add(new Offline_MT_Tracking_EmployeeList(_drRow, iRowCount++));

                BusinessLogic.Reset_ListViewColumn(lsv_TL_Names);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed..!");
            }
            finally
            {
                BusinessLogic.oMessageEvent.Start("Ready..!");
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void lsv_TL_FileDetails_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmb_TL_Batch_SelectedIndexChanged(object sender, EventArgs e)
        {
            lsv_TL_FileDetails.Items.Clear();
            Thread tTrack = new Thread(Load_TL_File_Tracking);
            tTrack.Start();
        }

        private void Get_Employee_TLTracking__Alloted_Files()
        {
            try
            {
                Offline_MT_Tracking_EmployeeList lsvEmp = (Offline_MT_Tracking_EmployeeList)lsv_TL_Names.SelectedItems[0];
                DataSet _dsAllotDetails = new DataSet();
                _dsAllotDetails = BusinessLogic.WS_Allocation.Get_alloted_Details(Convert.ToInt32(lsvEmp.EMP_PRODUCTION_ID));
                Employee_File_Alloted_Details oListItem;

                lsv_TL_FileDetails.Items.Clear();
                int iRowCount = 1;
                int dTotMins = 0;

                if (cmb_TL_Batch.Text == "MT")
                {
                    BusinessLogic.oMessageEvent.Start("Trasferring Data");
                    foreach (DataRow _drAllFiles in _dsAllotDetails.Tables[0].Select())
                    {
                        lsv_TL_FileDetails.Items.Add(new Employee_File_Alloted_Details(_drAllFiles, iRowCount++));
                        dTotMins += Convert.ToInt32(_drAllFiles["FileTot"].ToString());
                    }
                    string oMins = sGetDuration(dTotMins);

                    oListItem = new Employee_File_Alloted_Details("Total Minutes" + "Alloted: ", oMins.ToString());
                    lsv_TL_FileDetails.Items.Add(oListItem);

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lsv_TL_FileDetails);
                }
                else
                {
                    BusinessLogic.oMessageEvent.Start("Trasferring Data");
                    foreach (DataRow _drAllFiles in _dsAllotDetails.Tables[1].Select())
                    {
                        lsv_TL_FileDetails.Items.Add(new Employee_File_Alloted_Details(_drAllFiles, iRowCount++));
                        dTotMins += Convert.ToInt32(_drAllFiles["FileTot"].ToString());
                    }
                    string oMins = sGetDuration(dTotMins);

                    oListItem = new Employee_File_Alloted_Details("Total Minutes" + "Alloted: ", oMins.ToString());
                    lsv_TL_FileDetails.Items.Add(oListItem);

                    BusinessLogic.oMessageEvent.Start("Ready.");
                    BusinessLogic.Reset_ListViewColumn(lsv_TL_FileDetails);
                }


            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
            finally
            {
                BusinessLogic.oProgressEvent.Start(false);
            }
        }

        private void lsv_TL_Names_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.Selected)
            {
                Get_Employee_TLTracking__Alloted_Files();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                lsvMonthConsolidated.Items.Clear();
                lsvYearConsolidated.Items.Clear();

                BusinessLogic.oMessageEvent.Start("Transferring data..!");
                BusinessLogic.oProgressEvent.Start(true);

                int iMonth = comboBox18.SelectedIndex + 1;
                string sYear = comboBox19.SelectedItem.ToString();
                string sProductionID = comboBox16.SelectedValue.ToString();

                DataSet _dsConsolidated = BusinessLogic.WS_Allocation.Get_Consolidated_User_wise(Convert.ToInt32(iMonth), Convert.ToInt32(sYear), Convert.ToInt32(sProductionID));

                int iRow = 0;
                foreach (DataRow _drRow in _dsConsolidated.Tables[0].Select())
                    lsvMonthConsolidated.Items.Add(new listItem_ConsolidatedMonth(_drRow, iRow++));

                BusinessLogic.Reset_ListViewColumn(lsvMonthConsolidated);

                int iRow1 = 0;
                foreach (DataRow _drRow in _dsConsolidated.Tables[1].Select())
                    lsvYearConsolidated.Items.Add(new listItem_ConsolidatedYear(_drRow, iRow1++));

                BusinessLogic.Reset_ListViewColumn(lsvYearConsolidated);

                BusinessLogic.oMessageEvent.Start("Ready..!");
                BusinessLogic.oProgressEvent.Start(true);
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
                BusinessLogic.oMessageEvent.Start("Failed.");
            }
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Batch_Employee();
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDesigID = comboBox17.SelectedValue.ToString();
            sDesgination_ID = sDesigID;
            Load_Employee_Full_name(sDesgination_ID, sBranch_ID);
        }

        private void lsv_Target_Report_MouseMove(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = lsv_Target_Report.HitTest(e.X, e.Y);
            try
            {

                if (mLastPos != e.Location)
                {
                    if (info.Item != null && info.SubItem != null)
                    {
                        if (info.Item.Name.ToString() == string.Empty)
                        {

                            mTooltip.ToolTipTitle = null;
                            mTooltip.Hide(info.Item.ListView);

                        }
                        else
                        {

                            mTooltip.ToolTipTitle = info.Item.Text;
                            mTooltip.Show(info.Item.Name.ToString().Replace("||", System.Environment.NewLine), info.Item.ListView);
                        }
                    }
                }

                mLastPos = e.Location;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void lsvYearConsolidated_MouseMove(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = lsvYearConsolidated.HitTest(e.X, e.Y);
            try
            {
                if (mLastPos != e.Location)
                {
                    if (info.Item != null && info.SubItem != null)
                    {
                        if (info.Item.Name.ToString() == string.Empty)
                        {

                            mTooltip.ToolTipTitle = null;
                            mTooltip.Hide(info.Item.ListView);

                        }
                        else
                        {

                            mTooltip.ToolTipTitle = info.Item.Text;
                            mTooltip.Show(info.Item.Name.ToString().Replace("||", System.Environment.NewLine), info.Item.ListView);
                        }
                    }
                }

                mLastPos = e.Location;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void lsvMonthConsolidated_MouseMove(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = lsvMonthConsolidated.HitTest(e.X, e.Y);
            try
            {
                if (mLastPos != e.Location)
                {
                    if (info.Item != null && info.SubItem != null)
                    {
                        if (info.Item.Name.ToString() == string.Empty)
                        {

                            mTooltip.ToolTipTitle = null;
                            mTooltip.Hide(info.Item.ListView);

                        }
                        else
                        {

                            mTooltip.ToolTipTitle = info.Item.Text;
                            mTooltip.Show(info.Item.Name.ToString().Replace("||", System.Environment.NewLine), info.Item.ListView);
                        }
                    }
                }

                mLastPos = e.Location;
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void idNoggToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                List_View_Employee oItem = (List_View_Employee)lvMapping.SelectedItems[0];
                int iProdID = oItem.iProductionID;

                int iUpdate = BusinessLogic.WS_Allocation.Set_Not_Required(Convert.ToInt32(iProdID));
                if (iUpdate > 0)
                {
                    BusinessLogic.oMessageEvent.Start("Marked!");
                    Load_Mapping();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                List_View_Employee oItem = (List_View_Employee)lvMapping.SelectedItems[0];
                int iProdID = oItem.iProductionID;

                frmIdUpdate frmUpdate = new frmIdUpdate();
                frmUpdate.iProductionID = iProdID;
                frmUpdate.iDictaphoneID = oItem.iDictaphoneID.ToString();
                frmUpdate.iEscriptionID = oItem.iEscriptionID.ToString();
                frmUpdate.ShowDialog();

                if (frmUpdate.iUpdateID > 0)
                {
                    Load_Mapping();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void deActivateToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                List_View_Employee oItem = (List_View_Employee)lvMapping.SelectedItems[0];
                int iProdID = oItem.iProductionID;
                string sName = oItem.sEmployeeName.ToString();

                int iDeActivate = BusinessLogic.WS_Allocation.Set_Deactivate_Employee_V1(Convert.ToInt32(iProdID), Convert.ToInt32(0));
                if (iDeActivate > 0)
                {
                    Load_Mapping();
                    MailMessage mailMessage = new MailMessage();
                    MailAddress fromAddress = new MailAddress("mathew@rndsoftech.com", "Mathew Samuel");
                    mailMessage.From = fromAddress;
                    mailMessage.To.Add("nandagopal@rndsoftech.com");
                    mailMessage.CC.Add("vedhambal@rndsoftech.com");
                    mailMessage.CC.Add("pvshankar@rndsoftech.com");

                    mailMessage.Subject = "Reg : Employee deactivation PRODUCTION";

                    mailMessage.Body = "Dear All, This is to inform you, " + sName + " has been deactivated from production....";
                    mailMessage.IsBodyHtml = true;

                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.Host = "172.16.2.55";
                    smtpClient.Send(mailMessage);
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Employee Active List" + ".xls";
            ExportToExcel(lvMapping, sFolderNAme, sFileName);
        }

        private void cmbBranchMapp_SelectedIndexChanged(object sender, EventArgs e)
        {
            Load_Mapping();
        }

        private void markHTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                List_View_Employee oItem = (List_View_Employee)lvMapping.SelectedItems[0];
                int iProdID = oItem.iProductionID;

                int iMark = BusinessLogic.WS_Allocation.Set_HT(Convert.ToInt32(iProdID));
                if (iMark > 0)
                {
                    Load_Mapping();
                }
            }
            catch (Exception ex)
            {
                BusinessLogic.WS_Allocation.WriteException(ex.ToString(), BusinessLogic.USERNAME, Environment.MachineName);
            }
        }


        #endregion "Events 1"              

        private void btnExportHourlyLogsheet_Click(object sender, EventArgs e)
        {
            string sFolderNAme = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Environment.UserName + " Report\\" + DateTime.Now.ToString("ddMMMMyyyy");
            string sFileName = "Userwise Logsheet Hourly report" + ".xls";
            ExportToExcel(lsv_Hourlyreport, sFolderNAme, sFileName);
        }
    }
}