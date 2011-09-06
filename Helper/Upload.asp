<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A upload file class
'	File Name	: Upload.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** Dim up
'*** Set up=New Kiss_Upload
'*** '	up.GetDate(-1)
'*** Dim F_File,F_Type
'*** Set F_File=up.File("File")
'*** F_Name=randomStr(1)&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&"."&F_File.FileExt
'*** F_Type=FixName(F_File.FileExt)
'*** IF F_File.FileSize > Int(ULOADFILESIZE) Then
'*** 	echo("Error", "Error")
'*** Else
'*** 	F_File.SaveAs Server.MapPath("/upload/"&D_Name&"/"&F_Name)
'*** End IF
'*** Set F_File=Nothing
'*** Set up=Nothing

dim upload, Data_5xsoft
set upload = new Kiss_Upload

Class Kiss_Upload
  
	dim objForm,objFile,Version

	Public function Form(strForm)
	   strForm=lcase(strForm)
	   if not objForm.exists(strForm) then
		 Form=""
	   else
		 Form=objForm(strForm)
	   end if
	 end function

	Public function File(strFile)
	   strFile=lcase(strFile)
	   if not objFile.exists(strFile) then
		 set File=new FileInfo
	   else
		 set File=objFile(strFile)
	   end if
	 end function


	Private Sub Class_Initialize 
	  dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoend,tStream,iStart,theFile
	  dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
	  dim iFindStart,iFindend
	  dim iFormStart,iFormend,sFormName
	  Version="Upload files 0.1.1"
	  set objForm=Server.CreateObject("Scripting.Dictionary")
	  set objFile=Server.CreateObject("Scripting.Dictionary")
	  if Request.TotalBytes<1 then Exit Sub
	  set tStream = Server.CreateObject("adodb.stream")
	  set Data_5xsoft = Server.CreateObject("adodb.stream")
	  Data_5xsoft.Type = 1
	  Data_5xsoft.Mode =3
	  Data_5xsoft.Open
	  Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes)
	  Data_5xsoft.Position=0
	  RequestData =Data_5xsoft.Read 

	  iFormStart = 1
	  iFormend = LenB(RequestData)
	  vbCrlf = chrB(13) & chrB(10)
	  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1)
	  iStart = LenB (sStart)
	  iFormStart=iFormStart+iStart+1
	  while (iFormStart + 10) < iFormend 
		iInfoend = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3
		tStream.Type = 1
		tStream.Mode =3
		tStream.Open
		Data_5xsoft.Position = iFormStart
		Data_5xsoft.CopyTo tStream,iInfoend-iFormStart
		tStream.Position = 0
		tStream.Type = 2
		tStream.Charset ="gbk"
		sInfo = tStream.ReadText
		tStream.Close
		'Get form item name
		iFormStart = InStrB(iInfoend,RequestData,sStart)
		iFindStart = InStr(22,sInfo,"name=""",1)+6
		iFindend = InStr(iFindStart,sInfo,"""",1)
		sFormName = lcase(Mid (sinfo,iFindStart,iFindend-iFindStart))
		' if it is a file
		if InStr (45,sInfo,"filename=""",1) > 0 then
			set theFile=new FileInfo
			'Get the file name
			iFindStart = InStr(iFindend,sInfo,"filename=""",1)+10
			iFindend = InStr(iFindStart,sInfo,"""",1)
			sFileName = Mid (sinfo,iFindStart,iFindend-iFindStart)
			theFile.FileName=getFileName(sFileName)
			theFile.FilePath=getFilePath(sFileName)
			theFile.FileExt=getFileExt(sFileName)
			'Get the file type
			iFindStart = InStr(iFindend,sInfo,"Content-Type: ",1)+14
			iFindend = InStr(iFindStart,sInfo,vbCr)
			theFile.FileType =Mid (sinfo,iFindStart,iFindend-iFindStart)
			theFile.FileStart =iInfoend
			theFile.FileSize = iFormStart -iInfoend -3
			theFile.FormName=sFormName
			if not objFile.Exists(sFormName) then
			  objFile.add sFormName,theFile
			end if
		else
		'If it is a form item
			tStream.Type =1
			tStream.Mode =3
			tStream.Open
			Data_5xsoft.Position = iInfoend 
			Data_5xsoft.CopyTo tStream,iFormStart-iInfoend-3
			tStream.Position = 0
			tStream.Type = 2
			tStream.Charset ="gbk"
				sFormValue = tStream.ReadText 
				tStream.Close
			if objForm.Exists(sFormName) then
			  objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
			else
			  objForm.Add sFormName,sFormValue
			end if
		end if
		iFormStart=iFormStart+iStart+1
		wend
	  RequestData=""
	  set tStream =nothing
	end Sub

	Private Sub Class_Terminate  
	 if Request.TotalBytes>0 then
		objForm.RemoveAll
		objFile.RemoveAll
		set objForm=nothing
		set objFile=nothing
		Data_5xsoft.Close
		set Data_5xsoft =nothing
	 end if
	end Sub
	   
	 
	 Private function getFilePath(FullPath)
	  If FullPath <> "" Then
	   GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
	  Else
	   GetFilePath = ""
	  end If
	 end  function
	 
	 Private function getFileName(FullPath)
	  If FullPath <> "" Then
	   GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
	  Else
	   GetFileName = ""
	  end If
	 end  function
	 
	 Private function getFileExt(FullPath)
	  If FullPath <> "" Then
			GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
		Else
			GetFileExt = ""
	  end If
	 end function

end Class

Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
  Private Sub Class_Initialize 
	FileName = ""
	FilePath = ""
	FileSize = 0
	FileStart= 0
	FormName = ""
	FileType = ""
	FileExt=""
  end Sub
  
 Public function SaveAs(FullPath)
	dim dr,ErrorChar,i
	SaveAs=true
	if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function
	set dr=CreateObject("Adodb.Stream")
	dr.Mode=3
	dr.Type=1
	dr.Open
	Data_5xsoft.position=FileStart
	Data_5xsoft.copyto dr,FileSize
	dr.SaveToFile FullPath,2
	dr.Close
	set dr=nothing 
	SaveAs=false
  end function
 end Class

%>