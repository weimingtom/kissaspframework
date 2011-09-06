<% 
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: XML class
'	File Name	: Xml.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

Class Kiss_Xml
	Public XmlPath
	Private errorcode
	Private XMLMorntekDocument

	Private Sub Class_Initialize()
	 errorcode=-1
	end sub

	Private Sub Class_Terminate()

	end sub

 	'************************************************************* 
	' Name: Open
	' Purpose: Open=0, XMLMorntekDocument load a xml
	'************************************************************* 
	Public function Open()
	on error resume next
	   dim strSourceFile,strError
	   Set XMLMorntekDocument = Server.CreateObject(getXMLDOM)
		 If Err Then
		errorcode=-18239123
		Err.clear
		exit function
	   end if
	   XMLMorntekDocument.async = false  
	   strSourceFile = Server.MapPath(XmlPath) 
	   XMLMorntekDocument.load(strSourceFile) 
	   errorcode=XMLMorntekDocument.parseerror.errorcode
	end function 

 	'************************************************************* 
	' Name: OpenXML
	' Purpose: Create a xml document
	'************************************************************* 
	Public function OpenXML(xmlStr)
	on error resume next
	   dim strSourceFile,strError
	   Set XMLMorntekDocument = Server.CreateObject(getXMLDOM)
		 If Err Then
		errorcode=-18239123
		Err.clear
		exit function
	   end if
	   XMLMorntekDocument.async = false
	   XMLMorntekDocument.load(xmlStr)
	   errorcode=XMLMorntekDocument.parseerror.errorcode
	end function 

 	'************************************************************* 
	' Name: getError
	' Purpose: Get error infomation
	'************************************************************* 
	Public function getError()
	 getError=errorcode
	end function 

 	'************************************************************* 
	' Name: CloseXml
	' Purpose: Close a xml document
	'************************************************************* 
	Public function CloseXml() 
	 if IsObject(XMLMorntekDocument) then 
		 set XMLMorntekDocument=nothing 
	 end if 
	end function 

 	'************************************************************* 
	' Name: SelectXmlNodeText
	' Purpose: Get a element text
	'************************************************************* 
	Public function SelectXmlNodeText(elementname)  
		on error resume next
		dim temp
		temp=XMLMorntekDocument.getElementsByTagName(elementname).item(0).text
		selectXmlNodeText= temp
		if err then selectXmlNodeText=0
	end function

 	'************************************************************* 
	' Name: SelectXmlNode
	' Purpose: Get a element
	'************************************************************* 
	Public function SelectXmlNode(elementname,itemID) 
		dim temp
		set temp=XMLMorntekDocument.getElementsByTagName(elementname).item(itemID)
		set SelectXmlNode= temp
	end function

 	'************************************************************* 
	' Name: GetXmlNodeLength
	' Purpose: Get the length of a element
	'************************************************************* 
	Public function GetXmlNodeLength(elementname)  
		on error resume next 
		dim XmlLength
		XmlLength=XMLMorntekDocument.getElementsByTagName(elementname).length
		GetXmlNodeLength= XmlLength
		if err then GetXmlNodeLength=0
	end function

 	'************************************************************* 
	' Name: GetAttributes
	' Purpose: Get the attribute of a element
	'************************************************************* 
	Public function GetAttributes(elementname,nodeName,itemID)  
		dim XmlAttributes,i
		set XmlAttributes=XMLMorntekDocument.getElementsByTagName(elementname).item(itemID).attributes
		for i=0 to XmlAttributes.length-1
		 if XmlAttributes(i).name=nodeName then 
		  GetAttributes=XmlAttributes(i).value
		  exit function
		 end if
		next
		GetAttributes = 0
	end function  

 	'************************************************************* 
	' Name: SelectXmlNodeItemText
	' Purpose: Select a element text
	'************************************************************* 
	Public function SelectXmlNodeItemText(elementname,ID) 
		on error resume next 
		dim temp
		temp=XMLMorntekDocument.getElementsByTagName(elementname).item(ID).text
		SelectXmlNodeItemText= temp
		if err then SelectXmlNodeItemText=""
	end function

 	'************************************************************* 
	' Name: WriteXmlNodeItemText
	' Purpose: write text of a element
	'************************************************************* 
	Public function WriteXmlNodeItemText(elementname,ID,str) 
		on error resume next 
		WriteXmlNodeItemText=0
		dim temp,temp1
		set temp=XMLMorntekDocument.getElementsByTagName(elementname).item(ID)
		temp.childNodes(0).text=str
	  XMLMorntekDocument.save Server.MapPath(XmlPath)
		if err then WriteXmlNodeItemText=err.Description
	end function

 	'************************************************************* 
	' Name: IsXmlNode
	' Purpose: Check if elementname exists
	'************************************************************* 
	Public function IsXmlNode(elementname)
	 dim Temp
	 IsXmlNode=true
	 on error resume next
	 Temp=XMLMorntekDocument.getElementsByTagName(elementname).item(0).text
	 if err>0 then
	  err.clear
	  IsXmlNode=false
	 end if
	end function
end Class
%>