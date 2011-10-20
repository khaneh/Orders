using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class UC_ReportViewer : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {
        CrystalReportViewer1.PrintMode = CrystalDecisions.Web.PrintMode.ActiveX;
        CrystalReportViewer1.ReuseParameterValuesOnRefresh = true;
        CrystalReportViewer1.ReportSource = CrystalReportSource1;
        if (Request.QueryString.Get("Id") != null)
        {
            SetParameters(Int32.Parse(Request.QueryString.Get("Id")));
        }
    }
    protected void SetParameters(int sessionid)
    {
        try
        {

            System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection("Server=" + ConfigurationSettings.AppSettings["SqlServer"] + ";initial catalog=" + ConfigurationSettings.AppSettings["SqlDatabase"] + ";password=" + ConfigurationSettings.AppSettings["SqlPassword"] + ";user id=" + ConfigurationSettings.AppSettings["SqlUsername"] + ";");
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand("SELECT * From ReportLog Where Id=" + sessionid);
            conn.Open();
            cmd.Connection = conn;
            System.Data.SqlClient.SqlDataReader red = cmd.ExecuteReader();
            if (red.Read())
            {
                string filename = red.GetString(red.GetOrdinal("ReportFileName"));
                string[] paramnames = red.GetString(red.GetOrdinal("ReportParameterNames")).Split((char)1);
                string[] paramvalues = red.GetString(red.GetOrdinal("ReportParameterValues")).Split((char)1);
                CrystalReportSource1.ReportDocument.Load(System.Configuration.ConfigurationSettings.AppSettings["ReportsRoot"] + filename);
                CrystalReportSource1.ReportDocument.DataSourceConnections[0].SetLogon(System.Configuration.ConfigurationSettings.AppSettings["SqlUsername"], System.Configuration.ConfigurationSettings.AppSettings["SqlPassword"]);
                CrystalReportSource1.ReportDocument.Refresh();
                for (int j = 0; j < paramnames.Length; j++)
                {
                    if (CrystalReportSource1.ReportDocument.ParameterFields["@" + paramnames[j]] != null)
                    {
                        CrystalReportSource1.ReportDocument.ParameterFields["@" + paramnames[j]].CurrentValues.AddValue(paramvalues[j]);
                    }
                    else if (CrystalReportSource1.ReportDocument.ParameterFields[paramnames[j]] != null)
                    {
                        CrystalReportSource1.ReportDocument.ParameterFields[paramnames[j]].CurrentValues.AddValue(paramvalues[j]);
                    }
                }

            }
            conn.Close();
            conn.Dispose();
        }
        catch (Exception e)
        {
            Response.Write(e.ToString());
        }

    }
}
