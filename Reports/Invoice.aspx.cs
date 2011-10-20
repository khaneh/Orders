using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Web;

namespace Reports
{
	/// <summary>
	/// Summary description for WebForm1.
	/// </summary>
	public class InvoiceWebForm : System.Web.UI.Page
	{
		protected CrystalDecisions.Web.CrystalReportViewer CrystalReportViewer1;
		protected System.Web.UI.WebControls.Button Button1;
		protected System.Web.UI.HtmlControls.HtmlForm Form1;
		protected ReportManager rm = new ReportManager();
		Invoice crReport = new Invoice();
		string paramName= "Invoice_ID";
		string param;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			param = Request.QueryString.Get(paramName);
			
			Boolean x = rm.initReport(crReport, CrystalReportViewer1, param, paramName);
			if(!x) 
			{
				Response.Write ("error: Parameter Needed");
				Response.End();
			}
		}


		private void Button1_Click(object sender, System.EventArgs e)
		{
			rm.SetPrinter(Request.UserHostAddress);
			Boolean x = rm.printReport(crReport, CrystalReportViewer1, param, paramName);
			if(x) 
			{
				Response.Write("<font face=titr color=red size=5><center>");
				Response.Write("›«ﬂ Ê— œ— ’› ç«Åê—  <br>");
				Response.Write(" ﬁ—«— ê—› "+crReport.PrintOptions.PrinterName);
				Response.Write("</center></font>");
			}
			else
			{
				Response.Write ("error: Problem in printing report");
				Response.End();
			}

		
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
			this.Button1.Click += new System.EventHandler(this.Button1_Click);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

	}
}
