<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Microsoft Debugging class
'	File Name	: MSDebug.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** < % 'Create a session variable to test the debugger.
'*** Session("MySessionVar") = "Test variable"
'*** 
'*** 'Declare variables.
'*** dim x
'*** 
'*** 'Declare your Debug object.
'*** dim Debug
'*** 
'*** 'Instantiate the Debug object.
'*** set Debug = new MSDebug
'*** 
'*** 'Enable the Debug object. In the future, you can set it to false to disable it. 
'*** 'The rest of your debug code can be left in the page.
'*** Debug.Enabled = true
'*** 
'*** 'The following code shows how to use the Debug.Print method
'*** 'and how to provide a label to check your variables.
'*** x = 10
'*** Debug.Print "x before math", x
'*** x = x + 50
'*** Debug.Print "x after math", x
'*** 
'*** 'Create a cookie to test the debugger.
'*** Response.Cookies("TestCookie") = "Hello World"
'*** % >
'*** <HTML>
'*** <HEAD>
'*** <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
'*** </HEAD>
'*** <BODY>
'*** 
'*** <FORM action="" method=POST id=form1 name=form1>
'*** <INPUT type="text" id=text1 name=text1>
'*** <INPUT type="submit" value="Post Info" id=submit1 name=submit1>
'*** </FORM>
'*** 
'*** <FORM action="" method=get id=form2 name=form2>
'*** <INPUT type="text" id=text2 name=text2>
'*** <INPUT type="submit" value="get Info" id=submit2 name=submit2>
'*** </FORM>
'*** </BODY>
'*** </HTML>
'*** 
'*** < %'Call the end method to display all of the debug values.
'*** Debug.end
'*** set Debug = nothing % >

'*************************************************************
'	Initialize the class
'*************************************************************

dim msdebug
set msdebug = New Kiss_MSDebug 
msdebug.Enabled = true 

class Kiss_MSDebug

	dim blnEnabled
	dim dteRequestTime
	dim dteFinishTime
	dim objStorage

	public property get Enabled()
	   Enabled = blnEnabled
	end property

	public property let Enabled(bNewValue)
	   blnEnabled = bNewValue
	end property

	private sub class_Initialize()
	   dteRequestTime = Now()
	   set objStorage = Server.CreateObject("Scripting.Dictionary")
	end sub

	public sub Print(label, output)
	   if Enabled then
		   objStorage.Add label, output
	   end if
	end sub

	public sub [end]()
	   dteFinishTime = Now()
	   if Enabled then
		 PrintSummaryInfo()
		 PrintCollection "VARIABLE STORAGE", objStorage
		 PrintCollection "QUERYSTRING COLLECTION", Request.QueryString()
		 PrintCollection "FORM COLLECTION", Request.form()
		 PrintCollection "COOKIES COLLECTION", Request.Cookies()
		 PrintCollection "SESSION CONTENTS COLLECTION", Session.Contents()
		 PrintCollection "SERVER VARIABLES COLLECTION", Request.ServerVariables()
		 PrintCollection "APPLICATION CONTENTS COLLECTION", Application.Contents()
		 PrintCollection "APPLICATION STATICOBJECTS COLLECTION", Application.StaticObjects()
		 PrintCollection "SESSION STATICOBJECTS COLLECTION", Session.StaticObjects()
	   end if
	end sub

	private sub PrintSummaryInfo()
	   With Response
		.Write("<hr>")
		.Write("<b>SUMMARY INFO</b></br>")
		.Write("Time of Request = " & dteRequestTime) & "<br>"
		.Write("Time Finished = " & dteFinishTime) & "<br>"
		.Write("Elapsed Time = " & DateDiff("s", dteRequestTime, dteFinishTime) & " seconds<br>")
		.Write("Request Type = " & Request.ServerVariables("REQUEST_METHOD") & "<br>")
		.Write("Status Code = " & Response.Status & "<br>")
	   end With
	end sub

	private sub PrintCollection(Byval Name, Byval Collection)
	   dim varItem
	   Response.Write("<br><b>" & Name & "</b><br>")
	   for Each varItem in Collection
		 Response.Write(varItem & "=" & Collection(varItem) & "<br>")
	   next
	end sub

	private sub class_Terminate()
	   set objStorage = Nothing
	end sub

end class

%>