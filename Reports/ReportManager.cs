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
	/// One object from this class will initialize and print CR reports
	/// </summary>

	public class ReportManager
	{
		protected String PrinterName;
		protected ParameterFields paramFields = new ParameterFields ();
		protected ParameterField paramField = new ParameterField ();
		protected ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
		protected ParameterField paramField2 = new ParameterField ();
		protected ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();

		public ReportManager()
		{
			//PrinterName="\\\\Orders1\\LQ300_O1";
			//PrinterName="\\\\mohammad\\LQ300_O2";
			//PrinterName="\\\\AHOKI\\LQ300_O3";
			//PrinterName="\\\\GOLABI\\LQ300_C";
			//PrinterName="\\\\KOOCHAK\\LQ300_P1";
			PrinterName="none";
		}


		public ReportManager(String Printer)
		{
			PrinterName=Printer;
		}

		public void SetPrinter(String IP)
		{
			PrinterName=Util.choosePrinter(IP);
		}


		public Boolean initReport(ReportClass crReport, CrystalReportViewer CrystalReportViewer1, String param, String ParamName)
		{

			if(param=="" || param==null	) return false;

			paramField.ParameterFieldName = ParamName;
			discreteVal.Value = param;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			CrystalReportViewer1.ReportSource = crReport;
			CrystalReportViewer1.ParameterFieldInfo = paramFields;
			CrystalReportViewer1.RefreshReport();
			return true;
		}


		public Boolean initReport(ReportClass crReport, CrystalReportViewer CrystalReportViewer1, String param1, String ParamName1, String param2, String ParamName2)
		{

			if(param1=="" || param1==null	) return false;
			if(param2=="" || param2==null	) return false;
			
			paramField.ParameterFieldName = ParamName1;
			discreteVal.Value = param1;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			crReport.DataDefinition.ParameterFields[0].CurrentValues.Add (discreteVal); 
			
			paramField2.ParameterFieldName = ParamName2;
			discreteVal2.Value = param2;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);
			crReport.DataDefinition.ParameterFields[1].CurrentValues.Add (discreteVal2); 
			
			CrystalReportViewer1.ReportSource = crReport;
			CrystalReportViewer1.ParameterFieldInfo = paramFields;
			CrystalReportViewer1.RefreshReport();
			return true;
		}



		public Boolean printReport(CrystalDecisions.CrystalReports.Engine.ReportClass crReport, CrystalReportViewer CrystalReportViewer1, String param, String ParamName)
		{
			bool success=false;
			if (PrinterName!="none") 
			{
				try
				{
					paramField.ParameterFieldName = ParamName;
					discreteVal.Value = param;
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);
				
					crReport.DataDefinition.ParameterFields[0].CurrentValues.Add (discreteVal); 
					CrystalReportViewer1.ReportSource = crReport;
					CrystalReportViewer1.ParameterFieldInfo = paramFields;
					CrystalReportViewer1.RefreshReport();
			
					ParameterValues tmpParam = new ParameterValues();
					tmpParam.Add(discreteVal);
					crReport.DataDefinition.ParameterFields[0].ApplyCurrentValues(tmpParam);
					crReport.PrintOptions.PrinterName = PrinterName;
					//crReport.PrintToPrinter(1,false,1,1);
					crReport.PrintToPrinter(1,false,1,99);
					success=true;
				}
				catch{}
			}
			return success;
		}

		public Boolean printReport(CrystalDecisions.CrystalReports.Engine.ReportClass crReport, CrystalReportViewer CrystalReportViewer1, String param1, String ParamName1, String param2, String ParamName2)
		{
			if (PrinterName=="none") return false;

			paramField.ParameterFieldName = ParamName1;
			discreteVal.Value = param1;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			crReport.DataDefinition.ParameterFields[0].CurrentValues.Add (discreteVal); 

			paramField2.ParameterFieldName = ParamName2;
			discreteVal2.Value = param2;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);
			crReport.DataDefinition.ParameterFields[1].CurrentValues.Add (discreteVal2); 
				
			CrystalReportViewer1.ReportSource = crReport;
			CrystalReportViewer1.ParameterFieldInfo = paramFields;
			CrystalReportViewer1.RefreshReport();
			
			ParameterValues tmpParam = new ParameterValues();
			ParameterValues tmpParam2 = new ParameterValues();
			tmpParam.Add(discreteVal);
			tmpParam2.Add(discreteVal2);
			crReport.DataDefinition.ParameterFields[0].ApplyCurrentValues(tmpParam);
			crReport.DataDefinition.ParameterFields[1].ApplyCurrentValues(tmpParam2);
			crReport.PrintOptions.PrinterName = PrinterName;
			//			crReport.PrintToPrinter(1,false,1,1);
			crReport.PrintToPrinter(1,false,1,99);

			//Response.write (PrinterName);
			return true;
		}
	
	}
}
