using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WorkAllocation_ReleaseQAaccounts : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        { 
        }
    }

    public class gridcolumns
    {
        public string SlNo { get; set; }
        public string ProcId { get; set; }
        public string Locked_by { get; set; }
        public string Claim_No { get; set; }
        public string Status { get; set; }
        public string Patient_Acct { get; set; }
        public string Patient { get; set; }
        public string Status_Date { get; set; }
        public string Created_Date { get; set; }
        public string Service_Date { get; set; }
        public string Payer { get; set; }
        public string Provider_Name { get; set; }
        public string Charges { get; set; }

    }

   [WebMethod]
    public static gridcolumns[] GetLockedAccounts()
    {
        Business objbusiness = new Business();
        StoredProcedures objStoredProcedure = new StoredProcedures();
        gridcolumns objprop = new gridcolumns();
        List<gridcolumns> GvData = new List<gridcolumns>();
        objbusiness.HtTable.Clear();
        objbusiness.HtTable.Add("Client_Id", Convert.ToInt32(HttpContext.Current.Session["CLIENT_ID"]));
        objbusiness.HtTable.Add("Subsite_id", Convert.ToInt32(HttpContext.Current.Session["SUBSITE_ID"]));
        DataSet ds = objbusiness.getDataSet(objStoredProcedure.strGetLockedAccounts);
        foreach (DataRow dr in ds.Tables[0].Rows)
        {
            objprop = new gridcolumns();
            objprop.SlNo =dr["Sl No"].ToString();
            objprop.ProcId = dr["PROCId"].ToString();
            objprop.Locked_by = dr["Locked by"].ToString();
            objprop.Claim_No = dr["Claim No"].ToString();
            objprop.Status = dr["Status"].ToString();
            objprop.Patient_Acct = dr["Patient Acct#"].ToString();
            objprop.Patient = dr["Patient"].ToString();
            objprop.Status_Date = dr["Status Date"].ToString();
            objprop.Created_Date = dr["Created Date"].ToString();
            objprop.Service_Date = dr["Service Date"].ToString();
            objprop.Payer = dr["Payer"].ToString();
            objprop.Provider_Name = dr["Provider Name"].ToString();
            objprop.Charges = dr["Charges"].ToString();
            GvData.Add(objprop);
        }
        return GvData.ToArray();
        
     
    }
       
}