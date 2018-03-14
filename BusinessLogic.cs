using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using eAllocation.Allocation;
using System.IO;
using System.Configuration;
using System.Net.NetworkInformation;
using System.Drawing;

namespace eAllocation
{
    public class BusinessLogic
    {
        #region " Variables "

            public static MessageEventClass oMessageEvent = new MessageEventClass();
            public static ProgressEventClass oProgressEvent = new ProgressEventClass();
            public static eAllocation.Allocation.AllocationService WS_Allocation = new eAllocation.Allocation.AllocationService();
            public static eAllocation.QualityService.QualityAssuranceService WS_QualityService = new eAllocation.QualityService.QualityAssuranceService(); 
            
            public static string iUSERID;
            public static string iDESIGID;
            public static string USERNAME;
            public static string MACRO_PATH;
            public static string ALLOTEDUSERID;
            public static string SPRODUCTIONID;
            public static string SEMPLOYEEID;
            public static string S_OLD_PASSWORD;
            public static string AUDIO_PLAYER_PATH;
            public static int iACTIVITYID;
            public static string ALLOTEDUSERNAME;
            public static string LOGIN_NAME;
            public static int iTATREQUIRED = 0;
            public static int IDESIG_ID = 0;
            public static int IS_TED_ASSIGN = 0;
            public static int MTMET_BATCH_ID = 0;


            public static DateTime SERVER_DATE = DateTime.Today;            
            public Framework.Methods COMMONMETHOD = new Framework.Methods();                         
            string TRANSCRIPTION_URL1 = ConfigurationSettings.AppSettings["TRANSCRIPTION_URL1"].ToString();
            string TRANSCRIPTION_URL2 = ConfigurationSettings.AppSettings["TRANSCRIPTION_URL2"].ToString();
                

            public static int IFILETRANSFER;

            public enum FILE_TRANSFER_MODE
            {
                NETWORK = 1,
                HTTP = 2,
            }
            
        #endregion " Variables "

        #region " LIST "

            /// <summary>
            /// GET WEEKLY PROCESSED MINUTES
            /// </summary>
            public struct WEEKLY_PROCESSED_MINS
            {
                public int ROWCOUNT;
                public DateTime START_DATE;
                public DateTime END_DATE;
                public int TOT_FILES;
                public string FILE_MINUTES;

                public WEEKLY_PROCESSED_MINS(int iRowCount, DateTime dStart_date, DateTime dTo_date, int iTot_Files, string sFile_mins)
                {
                    ROWCOUNT = iRowCount;
                    START_DATE = dStart_date;
                    END_DATE = dTo_date;
                    TOT_FILES = iTot_Files;
                    FILE_MINUTES = sFile_mins;
                }
            }

            public struct INCENTIVE_DATE
            {
                public DateTime DINCENTIVE_DATE;
                public string LOCATION_ID;
                public string SHIFT_NAME;

                public INCENTIVE_DATE(DateTime dIncentive_date, string sLocation_id, string sShift_name)
                {
                    this.DINCENTIVE_DATE = dIncentive_date;
                    this.LOCATION_ID = sLocation_id;
                    this.SHIFT_NAME = sShift_name;
                }
            }

            public struct INCENTIVE_TARGET
            {
                public int PRODUCTION_ID;
                public string EMPLOYEE_NAME;
                public string LOCATION_ID;
                public int TARGET_MINS;
                public string FILE_MINS;
                public string MINS_DIFFERENCE;
                public DateTime SUBMITTED_TIME;

                public INCENTIVE_TARGET(int iProduction_id, string sName, string sLocation_id, int iTarget, string sFile_Mins, string sMins_Diff, DateTime dSubmit_time)
                {
                    this.PRODUCTION_ID = iProduction_id;
                    this.EMPLOYEE_NAME = sName;
                    this.LOCATION_ID = sLocation_id;
                    this.TARGET_MINS = iTarget;
                    this.FILE_MINS = sFile_Mins;
                    this.MINS_DIFFERENCE = sMins_Diff;
                    this.SUBMITTED_TIME = dSubmit_time;
                }
            }
        /// <summary>
        /// TO GET THE FILE WISE DETAILS
        /// </summary>
            public struct LINECOUNT_DETAILS
            {
                public string ACCOUNT;
                public string LOCATION;
                public string DOCTOR;
                public string FILENAME;                
                public DateTime FILEDATE;
                public string FILEMINUTES;
                public string CONVMINUTES;
                public decimal FILELINES;
                public decimal CONVLINES;
                public DateTime SUBMITTEDTIME;
                public DateTime SUBTIME;
                public string EVALUATEDDATE;
                public string TRANSSTATUS;
                public string CURRENTSTATUS;
                public string TEMPLATE;
                public decimal ACCURACY;
                public string ISGRADED;
                public int FILE_SEC;
                public decimal CONV_SEC;

                public LINECOUNT_DETAILS(string sAccount, string sLocation, string sDoctor, string sFilename, DateTime dFiledate, string sFilemins,
                    string sConv_Mins, decimal dFile_lines, decimal dConv_Lines, DateTime dSubmitted_time, DateTime dSub_Time, string sEval_Date,
                    string sTrans_Status, string sCurrent_Status, string sTemplate, decimal dAccuracy, string sIs_Graded, int ifile_sec, decimal iConv_Sec)
                {
                    this.ACCOUNT = sAccount;
                    this.LOCATION = sLocation;
                    this.DOCTOR = sDoctor;
                    this.FILENAME = sFilename;
                    this.FILEDATE = dFiledate;
                    this.FILEMINUTES = sFilemins;
                    this.CONVMINUTES = sConv_Mins;
                    this.FILELINES = dFile_lines;
                    this.CONVLINES = dConv_Lines;
                    this.SUBMITTEDTIME = dSubmitted_time;
                    this.SUBTIME = dSub_Time;
                    this.EVALUATEDDATE = sEval_Date;
                    this.TRANSSTATUS = sTrans_Status;
                    this.CURRENTSTATUS = sCurrent_Status;
                    this.TEMPLATE = sTemplate;
                    this.ACCURACY = dAccuracy;
                    this.ISGRADED = sIs_Graded;
                    this.FILE_SEC = ifile_sec;
                    this.CONV_SEC = iConv_Sec;
                }
            }

        /// <summary>
        /// TO GET THE TOTAL TARGET DETAILS
        /// </summary>
            public struct TARGET_DETAILS
            {
                public string DTARGET_DATE;
                public int STOTAL_FILES;
                public string STARGET_MINS;
                public int SACHIEVED_MINS;
                public string SBALANCE_MINS;
                public string SCOMPLETED_PERCENTAGE;
                public decimal SACHIEVEDLINES;
                public string sDetails;

                public TARGET_DETAILS(string TARGET_DATE, int TOTAL_FILES, string TARGET_MINS, int ACHIEVED_MINS, string BALANCE_MINS, string COMP_PERCEN, decimal ACHIEVED_LINES, string sDetails)
                {
                    this.DTARGET_DATE = TARGET_DATE;
                    this.STOTAL_FILES = TOTAL_FILES;
                    this.STARGET_MINS = TARGET_MINS;
                    this.SACHIEVED_MINS = ACHIEVED_MINS;
                    this.SBALANCE_MINS = BALANCE_MINS;
                    this.SCOMPLETED_PERCENTAGE = COMP_PERCEN;
                    this.SACHIEVEDLINES = ACHIEVED_LINES;
                    this.sDetails = sDetails;
                }
            }

        #endregion " LIST "

            #region "Get Methods"


            /// <summary>
            /// Get the employee details
            /// </summary>
            /// <returns></returns>
            public static DataTable Get_EmployeeDetails()
            {
                try
                {                   
                    //return WS_Allocation.Get_EmployeeDetails();
                    return null;
                }
                catch (Exception ex)
                {
                    oMessageEvent.Start("...");
                    return null;
                }
            }

            /// <summary>
            /// Get the Allocation Details
            /// </summary>
            /// <returns></returns>
            public static DataTable Get_AllocationDetails(int client_type)
            {
                try
                {                   
                    return WS_Allocation.Get_AllocationDetails("-1",client_type);
                }
                catch (Exception ex)
                {
                    oMessageEvent.Start("...");
                    return null;
                }
            }

            /// <summary>
            /// get user name and password
            /// </summary>
            /// <param name="sUsername"></param>
            /// <param name="sPassword"></param>
            /// <returns></returns>
            public static string Get_Username(string sUsername, string sPassword)
            {
                try
                {
                    DataTable dtUserId = WS_Allocation.Get_UserId(sUsername, sPassword);
                    // DataTable dtUserId = null;
                    if (dtUserId != null)
                    {
                        if (dtUserId.Rows.Count > 0)
                        {
                            iUSERID = dtUserId.Rows[0]["" + Framework.EMPLOYEE.FIELD_EMPLOYEE_ID + ""].ToString();
                            iDESIGID = dtUserId.Rows[0]["" + Framework.EMPLOYEE.FIELD_DESIGNATION_ID + ""].ToString(); 
                        }
                    }
                    else
                    {
                        iUSERID = null;
                    }
                    string sConcat = iUSERID + "," + iDESIGID;
                    return sConcat;
                }
                catch (Exception ex)
                {
                    oMessageEvent.Start("Application Error.");
                    return null;
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

            /// <summary>
            /// Check the network mode
            /// </summary>
            /// <param name="sServerPath"></param>
            /// <returns></returns>
            public static int GET_FILE_TRANSFER_MODE(string sServerPath)
            {
                try
                {
                    if (Directory.Exists(sServerPath))
                        //return Convert.ToInt32(Framework.Variables.NETWORK_MODE.NETWORK_MODE);
                        return Convert.ToInt32(Framework.Variables.NETWORK_MODE.HTTP_MODE);
                    else
                        throw new Exception("Network Error");
                }
                catch
                {
                    try
                    {
                        return Convert.ToInt32(Framework.Variables.NETWORK_MODE.HTTP_MODE);
                    }
                    catch
                    {
                        throw new Exception("No network access, Please check your network/internet connection settings");
                    }
                }
            }

            /// <summary>
            /// CHECK CONNECTIONS
            /// </summary>
            /// <returns></returns>
            public bool CheckConnection()
            {
                do
                {
                    try
                    {
                        //Check if the primary service is connecting
                        WS_Allocation.Credentials = System.Net.CredentialCache.DefaultCredentials;
                        WS_Allocation.Url = TRANSCRIPTION_URL1;
                        WS_Allocation.CheckConnection();
                        IFILETRANSFER = Convert.ToInt32(BusinessLogic.FILE_TRANSFER_MODE.NETWORK);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        //if primary connection fails,  check the secondary connection
                        WS_Allocation.Credentials = System.Net.CredentialCache.DefaultCredentials;
                        WS_Allocation.Url = TRANSCRIPTION_URL2;
                        WS_Allocation.CheckConnection();
                        IFILETRANSFER = Convert.ToInt32(BusinessLogic.FILE_TRANSFER_MODE.HTTP);
                        return true;
                    }
                } while (true);
            }

            public DateTime ConvertToDateTime(object oValue)
            {
                try
                {
                    return Convert.ToDateTime(oValue);
                }
                catch
                {
                    try
                    {
                        var s = oValue.ToString();
                        var d = s.Split('/')[1] + "/" + s.Split('/')[0] + "/" + s.Split('/')[2];
                        return Convert.ToDateTime(d);
                    }
                    catch
                    {
                        return DateTime.Now;
                    }
                }
            }

            /// <summary>
            /// Adding two minutes
            /// </summary>
            /// <param name="sMinutesOne"></param>
            /// <param name="sMinutesTwo"></param>
            /// <returns></returns>
            public string AddMinutes(string sMinutesOne, string sMinutesTwo)
            {
                try
                {
                    if (sMinutesOne.Trim().Length == 0)
                        sMinutesOne = "0";

                    if (sMinutesTwo.Trim().Length == 0)
                        sMinutesTwo = "0";

                    return ConvertToMinutes(ConvertToSeconds(sMinutesOne) + ConvertToSeconds(sMinutesTwo));
                }
                catch (Exception ex)
                {
                    throw new Exception("Error in Adding Minutes" + Environment.NewLine + ex.ToString());
                }
            }

            /// <summary>
            /// Converting Seconds into minutes
            /// </summary>
            /// <param name="iPerfectSeconds"></param>
            /// <returns></returns>
            public string ConvertToMinutes(int iPerfectSeconds)
            {
                try
                {
                    string sPerfectMinutes = Convert.ToString((iPerfectSeconds / 60));
                    string sPerfectSeconds = Convert.ToString(iPerfectSeconds % 60).PadLeft(2, '0');
                    return sPerfectMinutes + "." + sPerfectSeconds;
                }
                catch (Exception ex)
                {
                    throw new Exception("Error in Converting Seconds into Minutes" + Environment.NewLine + ex.ToString());
                }
            }

            /// <summary>
            /// This method converts the minutes into seconds
            /// </summary>
            /// <param name="sMinutes"></param>
            /// <returns></returns>
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

            #endregion "Get Methods"
    }
}

