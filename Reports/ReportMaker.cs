using System;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Web;

namespace Reports
{
/// <summary>
/// Summary description for ReportMaker.
/// </summary>


public class ReportMaker 
{
	protected string PrinterName;

	public ReportMaker()
	{
		//PrinterName="\\\\Orders1\\LQ300_O1";
		//PrinterName="\\\\mohammad\\LQ300_O2";
		//PrinterName="\\\\AHOKI\\LQ300_O3";
		//PrinterName="\\\\GOLABI\\LQ300_C";
		//PrinterName="\\\\KOOCHAK\\LQ300_P1";
		PrinterName="none";
	}

	public void SetPrinter(String IP)
	{
		PrinterName=Util.choosePrinter(IP);
	}

	public Boolean printReport(ReportClass report, ParameterFields parameterFields)
	{
		if (PrinterName=="none") 
			return false;

		for (int i=0; i<report.DataDefinition.ParameterFields.Count; i++)
		{
			report.DataDefinition.ParameterFields[i].ApplyCurrentValues(parameterFields[i].CurrentValues);
		}
		report.PrintOptions.PrinterName = PrinterName;
		report.PrintToPrinter(1,false,1,99);

		return true;
	}
}
}
