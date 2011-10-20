<%
	Option Explicit

	Dim mla_cfg_theme, mla_cfg_themes(1, 4)
	Dim mla_cfg_lng, mla_cfg_lngs(1, 4)
	Dim mla_cfg_showsysdatabases, mla_cfg_showsystables, mla_cfg_showsysviews, mla_cfg_showsysprocedures, mla_cfg_showsysfunctions
	Dim mla_cfg_pagesize, mla_cfg_maxdisplayedchar, mla_cfg_maxdisplayedbin, mla_cfg_rowdelimiter, mla_cfg_firstdayofweek
	Dim mla_cfg_pagesize_default, mla_cfg_maxdisplayedchar_default, mla_cfg_maxdisplayedbin_default, mla_cfg_rowdelimiter_default, mla_cfg_firstdayofweek_default
	Dim mla_version_number, mla_release_number

	' ### Version info
	mla_version_number = "2.0 lite"
	mla_release_number = "096"

	' ### Default value (display)
	mla_cfg_pagesize_default = 20					' number of rows per page
	mla_cfg_maxdisplayedchar_default = 40		' number of char displayed for a string field
	mla_cfg_maxdisplayedbin_default = 16		' number of char displayed for a binary field
	mla_cfg_rowdelimiter_default = ";"				' filed delimiter for csv export
	mla_cfg_firstdayofweek_default = 1			' first day of week (1 = Sunday)

	' ### Skins info
	mla_cfg_themes(0, 0) = "classic"
	mla_cfg_themes(1, 0) = "SQL Server Like"
	mla_cfg_themes(0, 1) = "oraclelike"
	mla_cfg_themes(1, 1) = "Oracle Like"
	'mla_cfg_themes(0, 2) = "yourskinfolder"
	'mla_cfg_themes(1, 2) = "YourSkin Name"

	If Request.Cookies("Option")("Theme") = "" Then mla_cfg_theme = mla_cfg_themes(0, 0) Else mla_cfg_theme = Request.Cookies("Option")("Theme") End If

	' ### Languages info
	mla_cfg_lngs(0, 0) = "US"
	mla_cfg_lngs(1, 0) = "English"
	mla_cfg_lngs(0, 1) = "FR"
	mla_cfg_lngs(1, 1) = "French"
	'mla_cfg_lngs(0, 2) = "yourlanguageabbreviation"
	'mla_cfg_lngs(1, 2) = "yourlanguagelitteralname"

	If Request.Cookies("Option")("Lng") = "" Then mla_cfg_lng = mla_cfg_lngs(0, 0) Else mla_cfg_lng = Request.Cookies("Option")("Lng") End If

	' ### preferences saved in cookies
	If Request.Cookies("Option")("ShowSysDatabase") = "" Then mla_cfg_showsysdatabases = false Else mla_cfg_showsysdatabases = CBool(Request.Cookies("Option")("ShowSysDatabase")) End If
	If Request.Cookies("Option")("ShowSysTable") = ""  Then mla_cfg_showsystables = false Else mla_cfg_showsystables = CBool(Request.Cookies("Option")("ShowSysTable")) End If
	If Request.Cookies("Option")("ShowSysView") = ""  Then mla_cfg_showsysviews = false Else mla_cfg_showsysviews = CBool(Request.Cookies("Option")("ShowSysView")) End If
	If Request.Cookies("Option")("ShowSysProcedure") = ""  Then mla_cfg_showsysprocedures = false Else mla_cfg_showsysprocedures = CBool(Request.Cookies("Option")("ShowSysProcedure")) End If
	If Request.Cookies("Option")("ShowSysFunction") = ""  Then mla_cfg_showsysfunctions = false Else mla_cfg_showsysfunctions = CBool(Request.Cookies("Option")("ShowSysFunction")) End If
	If Request.Cookies("Option")("PageSize") = ""  Then mla_cfg_pagesize = mla_cfg_pagesize_default Else mla_cfg_pagesize = CInt(Request.Cookies("Option")("PageSize")) End If
	If Request.Cookies("Option")("MaxDisplayedChar") = ""  Then mla_cfg_maxdisplayedchar = mla_cfg_maxdisplayedchar_default Else mla_cfg_maxdisplayedchar = CInt(Request.Cookies("Option")("MaxDisplayedChar")) End If
	If Request.Cookies("Option")("MaxDisplayedBin") = ""  Then mla_cfg_maxdisplayedbin = mla_cfg_maxdisplayedbin_default Else mla_cfg_maxdisplayedbin = CInt(Request.Cookies("Option")("MaxDisplayedBin")) End If
	If Request.Cookies("Option")("RowDelimiter") = ""  Then mla_cfg_rowdelimiter = mla_cfg_rowdelimiter_default Else mla_cfg_rowdelimiter = Request.Cookies("Option")("RowDelimiter") End If
	If Request.Cookies("Option")("FirstDayOfWeek") = ""  Then mla_cfg_firstdayofweek = mla_cfg_firstdayofweek_default Else mla_cfg_firstdayofweek = CInt(Request.Cookies("Option")("FirstDayOfWeek")) End If

%>
