using System;

namespace Reports
{
	/// <summary>
	/// Summary description for Util.
	/// </summary>
	public class Util
	{
/*		public Util()
		{
			//
			// TODO: Add constructor logic here
			//
		}
*/
		public static string choosePrinter(string IP)
		{
			//  Don't forget to Change 
			//		1- ReportManager.cs
			//		2- ReportMaker.cs
			//
			string PrinterName;
			switch(IP)
			{
				case "192.168.0.5":		// AHOKI	(Orders - Ehsan Jalali)
					PrinterName="\\\\AHOKI\\LQ300_O3";
					break;
				
				case "REM 192.168.0.52" :   // Bidaabaadi ( SALE )
					//case "192.168.0.111":	// Naseri	(Design4)
					//PrinterName="\\\\hasan\\Epson LQ-300 ESC/P 2";
					//PrinterName="\\\\192.168.0.6\\LQ300_O1";
					PrinterName="\\\\192.168.0.52\\LQ300_O1";
					break;
				case "192.168.0.8":		// MOHAMMAD	(Orders - Saman)
					PrinterName="\\\\mohammad\\LQ300_O2";
					break;
				case "192.168.0.10":	// Zamanie (Zamanie)
					PrinterName="\\\\192.168.0.10\\EpsonLQ-Zamani";
					//PrinterName="Oce 3165 Network Copier";
					break;
				case "192.168.0.22":		// Monfared ( old Naseri )
				case "192.168.0.21":		// Farahani ( old Tayefe - lito)
				
					PrinterName=@"\\offset\LQ300_P1";						
					break;
				case "192.168.0.117":		// Salimi
				case "192.168.0.18":	// Behnaz Ashraf
				case "192.168.0.55":	// Hossein Shahrabir
				case "192.168.0.6":		// Hasan	
				case "192.168.0.118":	// KOOFI	
					PrinterName=@"\\192.168.0.118\XEROX_3130";
					break;
				case "192.168.0.39":	// GOLABI	(Cashier - Golabi)
					//PrinterName="\\\\GOLABI\\LQ300_C";
					PrinterName=@"\\192.168.0.39\LQ300_C";
					break;
				//case "192.168.0.52":	// FARVARDIN	(Orders - Zargar)
				//	PrinterName="\\\\KHORDAD\\LQ300_O4";
				//	break;
				case "192.168.0.54":	// KHORDAD	(Orders - Saman)
					PrinterName="\\\\KHORDAD\\LQ300_O4";
					break;
				case "192.168.0.61":	// Shahami (Hesabdari)
					PrinterName="\\\\192.168.0.61\\LQ_Shahami";
					//					PrinterName="Oce 3165 Network Copier";
					break;
				case "192.168.0.64":		// Dehghan	(Dehghan - Hesabdari)
					PrinterName="\\\\192.168.0.64\\LQ300_A1";
					break;
				case "192.168.0.15":	// CopyShop
				case "192.168.0.51":
				case "192.168.0.52" :   // Bidaabaadi ( SALE )
				case "192.168.0.53":
				case "192.168.0.56":
				case "192.168.0.112":
				case "192.168.0.63":	// masoud babaee (Old Monafred)
				case "192.168.0.133":
				case "192.168.0.99":	// Server (Just For Testing)
				case "127.0.0.1":		// Server (Just For Testing)
					//PrinterName="\\\\192.168.0.112\\HP_1320";
					PrinterName=@"\\khordad\HP_1320";
					break;
				case "192.168.0.71":	// Beheshti	(Abbas Abad)
				case "192.168.0.72":
				case "192.168.0.73":
				case "192.168.0.74":
				case "192.168.0.75":
					PrinterName="\\\\"+IP+"\\Beheshti";
					break;
				case "192.168.0.94":	// Kid (For Testing)
					PrinterName="Oce 3165 Network Copier";
					break;
				case "192.168.0.89":	// Alix (For kiding)
				case "192.168.0.97":	// r (Laptop)
//				case "192.168.0.98":	// Mohaghegh (Laptop)
				case "192.168.0.62":		// Esterabi
					PrinterName="Oce 3165 Network Copier";
					break;
				case "192.168.0.206":	// Design6	(Tarrahi - Chalajour)
					PrinterName="\\\\192.168.0.206\\Tarrahi";
					break;
				default:
					PrinterName="none";
					break;
			}
			return PrinterName;
		}

		public static string choosePrinter(string IP,string reportName)
		{
			string PrinterName;
			switch(reportName)
			{
				case "Receipt.rpt":
					switch (IP)
					{
						case "192.168.0.15":	// CopyShop
						case "192.168.0.52" :   // Bidaabaadi ( SALE )
						case "192.168.0.51":
						case "192.168.0.53":
						case "192.168.0.56":
						case "192.168.0.112":
						case "192.168.0.133":
							//PrinterName="\\\\192.168.0.112\\XRX_O1";
							PrinterName=@"\\khordad\HP_1320";
							break;
						case "192.168.0.63":		// Monfared	(Monfared - Cashier)
							PrinterName=@"\\192.168.0.63\LQ300_C";

						break;
						default:
							PrinterName=choosePrinter(IP);
							break;
					};
					break;
				case "InvoicePrintForm.rpt":
					PrinterName=@"\\192.168.0.63\LQ300_C";
				break;
				default:
					PrinterName=choosePrinter(IP);
					break;
			}
			return PrinterName;
		}

		public static string sqlSafe(string inpStr)
		{
			return inpStr.Replace("'","''");
		}
	}
}
