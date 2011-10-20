using System;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Web;

namespace Reports
{
	/// <summary>
	/// Summary description for GeneralReportWebForm.
	/// </summary>
	public class GeneralReportWebForm : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button Button1;
		protected System.Web.UI.HtmlControls.HtmlForm Form1;
		protected CrystalDecisions.Web.CrystalReportViewer viewer;
		protected GeneralReport report = null;
		protected ParameterFields parameterFields = new ParameterFields ();

		protected int reportLogID=0; 
		protected int totalPageCount=0;
		protected string errors="";
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			char failed='0';
			try
			{
				reportLogID = Int32.Parse(Request.QueryString.ToString());
			}
			catch 
			{
				errors+="Invalid ReportNumber\n<br>";
			}

			string[] strParameterNames=new string[] {};
			string[] strParameterValues=new string[] {};
			
			//SqlConnection conn = new SqlConnection("Data Source=localhost;Integrated Security=SSPI;Initial Catalog=sefareshat");
			SqlConnection conn = new SqlConnection("data source=localhost;initial catalog=sefareshat;password=5tgb;persist security info=True;user id=sefadmin;workstation id=appserv;packet size=4096");

			conn.Open();

			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection=conn;

			sqlCommand.CommandText="SELECT ReportFileName,ReportParameterNames,ReportParameterValues FROM ReportLog WHERE ID='"+ reportLogID +"' AND Checked=0";

			SqlDataReader dr = sqlCommand.ExecuteReader();

			if (dr.Read())
			{
				report=new GeneralReport(dr.GetString(0));
				string tmpParameterNames=dr.GetString(1);
				string tmpParameterValues=dr.GetString(2);
				if(tmpParameterNames!="")
				{
					strParameterNames=tmpParameterNames.Split((char)1);
					strParameterValues=tmpParameterValues.Split((char)1);
				}
			}
			else
			{
				errors+="No such Report\n<br>";
			}
			dr.Close();
			for (int i=0; i<strParameterNames.Length; i++)
			{
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				ParameterField paramField = new ParameterField ();

				paramField.ParameterFieldName = strParameterNames[i];
				discreteVal.Value = strParameterValues[i];
				paramField.CurrentValues.Add (discreteVal);
				parameterFields.Add (paramField);
			}
			if(! initViewer())
				errors+="Cannot Initialize Viewer.\n<br>";
			
			if (errors!="")
			{
				failed='1';
				Response.Write(errors);
			}

			if(errors.Length>180)
				errors = errors.Substring(0,180);

			sqlCommand.CommandText="UPDATE ReportLog Set Checked='1',Viewed='1',Failed='"+failed+"',ErrorDescription=N'"+ Util.sqlSafe(errors) +"',PrinterName='Not Printed' WHERE (ID='"+ reportLogID +"')";
			sqlCommand.ExecuteNonQuery();

			conn.Close();
//			if (report!=null) report.Dispose();
			if (report!=null)
			{
				// Response.Write("###+(" + totalPageCount + ") ###");
				if (totalPageCount<2)
					report.Dispose();
			}

			if (errors!="")
				Response.End();
		}

		private bool initViewer()
		{
			bool success=true;
			try
			{
				viewer.ReportSource = report;
				for (int i=0; i<report.DataDefinition.ParameterFields.Count; i++)
				{
					report.DataDefinition.ParameterFields[i].ApplyCurrentValues(parameterFields[i].CurrentValues);
				}
				viewer.RefreshReport();

				CrystalDecisions.Shared.ReportPageRequestContext reqContext = (CrystalDecisions.Shared.ReportPageRequestContext) viewer.RequestContext;

				totalPageCount = report.FormatEngine.GetLastPageNumber(reqContext); 

				/*
				*	Another way wich also works:
				* 
					for (int i=0; i<report.DataDefinition.ParameterFields.Count; i++)
					{
						report.DataDefinition.ParameterFields[i].CurrentValues.Add (parameterDiscreteValues[i]); 
					}
					CrystalReportViewer1.ReportSource = (ReportClass) report;
					CrystalReportViewer1.ParameterFieldInfo = parameterFields;
					CrystalReportViewer1.RefreshReport();
				*/
			}
			catch (Exception e)
			{
//				errors += e.ToString()+"\n<br><br>";
				success=false;
			}
			return success;
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
			DataBind();
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

	}
}