<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: URL Routing class
'	File Name	: Router.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** code

'*************************************************************
'	Initialize the class
'*************************************************************
dim Router
set Router = New Kiss_Router

class Kiss_Router
	public className	'Class name
	private Controller	'Controller name
	private Action		'Action name
	
	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
	private sub class_Initialize()
		className = "Kiss_Router"
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

	'************************************************************* 
	' Name: Dispatch
	' Purpose: Dispatch the router
	'************************************************************* 
	public default function Dispatch()
		Controller = rGet("c", fLCase)
		Action = rGet("a", fLCase)
		Controller = ucfirst(Controller)
		Action = ucfirst(Action)
		if Controller = "" then
		  Controller = "Default"
		  require(APP_ & "/controller/" & Controller &".asp")
		  eval ("public Ctl")
		  eval ("set Ctl = new Default")
		else
		  if (fileExists(APP_ & "/controller/" & Controller & ".asp")) then
			require(APP_ & "/controller/" & Controller & ".asp")
		    eval ("public Ctl")
		    eval ("set Ctl = new " & Controller)
		  else
			show404("Doesn't find the controller {<span style=""color:#f00; font-weight:bold;""> " & Controller & " </span>}")
		  end if
		end if

		me.executeAction(Action)
	end function

	'************************************************************* 
	' Name: executeAction
	' Param: Action as action name
	' Purpose: Execute the action
	'************************************************************* 
	public sub executeAction(Action)
	  if Action="" then
		Action="Index"
		if function_exists(Controller, "Index") then
		   eval ("Ctl.actionIndex()")
		else
		   show404("Doesn't find the action {<span style=""color:#f00; font-weight:bold;""> " & Action & " </span>}")
		end if
	  else
		'die(Controller & Action)
		if function_exists(Controller, Action) then
		   eval ("Ctl.action" & Action & "()")
		else
		   show404("Doesn't find the action {<span style=""color:#f00; font-weight:bold;""> " & Action & " </span>}")
		end if
	  end if
	end sub

end class
%>