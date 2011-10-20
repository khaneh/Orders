using System;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Web;

namespace Reports
{
	/// <summary>
	/// Summary description for PrintReport.
	/// </summary>
	public class PrintReport : System.Web.UI.Page
	{
		protected GeneralReport report = null;
		protected ParameterFields parameterFields = new ParameterFields ();

		protected int reportLogID=0; 

		private void Page_Load(object sender, System.EventArgs e)
		{
			string errors="";
			string returnURL="";
			string reportName="";
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
			SqlConnection conn = new SqlConnection("data source=appserv;initial catalog=sefareshat;password=5tgb;persist security info=True;user id=sefadmin;workstation id=appserv;packet size=4096");
			conn.Open();

			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.Connection=conn;

			sqlCommand.CommandText="SELECT ReportFileName,ReportParameterNames,ReportParameterValues,ReturnURL FROM ReportLog WHERE ID='"+ reportLogID +"' AND Checked=0";
		
			SqlDataReader dr = sqlCommand.ExecuteReader();

			if (dr.Read())
			{
				reportName=dr.GetString(0);
				report=new GeneralReport(reportName);
				string tmpParameterNames=dr.GetString(1);
				string tmpParameterValues=dr.GetString(2);
				returnURL=dr.GetString(3);
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
			try
			{
				for (int i=0; i<strParameterNames.Length; i++)
				{
					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					ParameterField paramField = new ParameterField ();

					paramField.ParameterFieldName = strParameterNames[i];
					discreteVal.Value = strParameterValues[i];
					paramField.CurrentValues.Add (discreteVal);
					parameterFields.Add (paramField);
				}
			}
			catch
			{
				errors+="Invalid Report Parameters...\n<br>";
			}
			
			string printerName=Util.choosePrinter(Request.UserHostAddress,reportName);
			if(! Print(printerName))
				errors+="Cannot Print Report...\n<br>";

			if (errors!="")
			{
				failed='1';
				Response.Write(errors);
			}
			
			sqlCommand.CommandText="UPDATE ReportLog Set Checked='1',Printed='1',Failed='"+failed+"',ErrorDescription=N'"+ Util.sqlSafe(errors) +"',PrinterName='"+ Util.sqlSafe(printerName)+"' WHERE (ID='"+ reportLogID +"') AND (Printed=0) ";
			sqlCommand.ExecuteNonQuery();

			conn.Close();
			if (report!=null) report.Dispose();
			if (returnURL!="" && errors=="")
				Response.Redirect(returnURL,true);
			Response.End();
		}

		private bool Print(string PrinterName)
		{
			bool success=true;

			try
			{
				for (int i=0; i<report.DataDefinition.ParameterFields.Count; i++)
				{
					report.DataDefinition.ParameterFields[i].ApplyCurrentValues(parameterFields[i].CurrentValues);
				}
				report.PrintOptions.PrinterName = PrinterName;
				report.PrintToPrinter(1,false,1,99);
			}
			catch
			{
				success=false;
			}
			return success;
		}


		#region Web Form Designer generated code
		override protected void OnInit(System.EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
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