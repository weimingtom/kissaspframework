<!--#include file = "Kiss/Config.asp" -->
<%

'Setting of application path
CONST APP_					= "App"					'Application path
CONST APP_CONTROLLER		= "App/controller/"		'Application controller path
CONST APP_VIEW				= "App/View/"			'Application template path

dim DbPath
DbPath = "Data/News.asa"

'Dispatch the Url
Router.Dispatch

'Clear all objects
Finish()

%>