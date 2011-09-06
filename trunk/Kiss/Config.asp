<%
Option Explicit
Response.Buffer = true

dim StartTime
StartTime = timer()

'Database constants
Const DTAccess	 = 1
Const DTMSSQL	 = 2
Const DTMySQL	 = 3
Const DTDB2		 = 4
Const DTOracle	 = 5
Const DTExcel	 = 6

%>
<!--#include file = "Ext.asp" -->
<!--#include file = "Require.asp" -->
<!--#include file = "Router.asp" -->
<!--#include file = "View.asp" -->
<!--#include file = "File.asp" -->

<%
sub Finish()
	dim RunTime
	dim dbQueryTimes : dbQueryTimes = 0
	RunTime = round((timer()-StartTime), 3)
	if isobject(Db) then
		dbQueryTimes = Db.QueryTimes
	else
		dbQueryTimes = 0
	end if
	echo "<p style=""text-align:center;"">Page rendered in " & RunTime & " seconds and database query " & dbQueryTimes & " time(s).</p>"
	

	set require	= nothing 'Kiss_require
	set Router	= nothing 'Kiss_Router
	set View	= nothing 'Kiss_View
	
	'dim endHtml : endHtml = vbNewLine & "</body>" & vbNewLine & "</html>"
	'die(endHtml)
	die("")
end sub
%>