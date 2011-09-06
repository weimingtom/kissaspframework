<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A database to excel class
'	File Name	: DbtoExcel.asp
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

class Kiss_DbtoExcel
	' 注意:此版本的类暂时只支持一次转换一个数据表，即一个数据表只能对应一个Excel文件。如果转换一个数据表后不更换TargetFile参数，则将覆盖以前的表数据！！！！
	' 使用方法请仔细阅读下面的注解说明!!
	''/***************************************************************************
	''/*
	''/*声明：使用此类必需服务器上装有Office（Excel）程序，否则使用时可能不能转移数据
	''/*      此版本的类暂时只支持一次转换一个数据表，即一个数据表只能对应一个Excel文件。
	''/*      如果转换一个数据表后不更换TargetFile参数，则将覆盖以前的表数据！！！！
	''/*用法：
	''/*方法一：（Access数据库文件 TO Excel数据库文件）
	''/*1、先设置源数据库文件SourceFile（可选）和目标数据库文件TargetFile（必选）
	''/*2、再使用Transfer("源表名","字段列表","转移条件")方法转移数据
	''/*例子：
	''/*   dim sFile,tFile,Objclass,sResult
	''/*   sFile=Server.MapPath("data/data.mdb")
	''/*   tFile=Server.Mappath(".")&"\back.xls"
	''/*   set Objclass=New DataBaseToExcel
	''/*   Objclass.SourceFile=sFile
	''/*   Objclass.TargetFile=tFile
	''/*   sResult=Objclass.Transfer("table1","","")
	''/*   if sResult then
	''/*      Response.Write "转移数据成功！"
	''/*   else
	''/*      Response.Write "转移数据失败！"
	''/*   end if
	''/*   set Objclass=Nothing
	''/*
	''/*方法二：（其它数据库文件 To Excel数据库文件)
	''/*1、设置目标数据库文件TargetFile
	''/*2、设置Adodb.Connection对象
	''/*3、再使用Transfer("源表名","字段列表","转移条件")方法转移数据
	''/*例子：(在此使用Access的数据源做例子，你可以使用其它数据源）
	''/*   dim Conn,ConnStr,tFile,Objclass,sResult
	''/*   tFile=Server.Mappath(".")&"\back.xls"
	''/*   set Conn=Server.CreateObject("ADODB.Connection")
	''/*   ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data/data.mdb")
	''/*   Conn.Open ConnStr
	''/*   set Objclass=New DataBaseToExcel
	''/*   set Objclass.Conn=Conn        ''此处关键
	''/*   Objclass.TargetFile=tFile
	''/*   sResult=Objclass.Transfer("table1","","")
	''/*   if sResult then
	''/*      Response.Write "转移数据成功！"
	''/*   else
	''/*      Response.Write "转移数据失败！"
	''/*   end if
	''/*   set Objclass=Nothing
	''/*   Conn.Close
	''/*   set Conn=Nothing
	''/*  
	''/*说明：TargetFile属性一定要设置！（备份文件地址，绝对地址！）
	''/*      如果不设置SourceFile则一定要设置Conn，这两个属性必选之一，但优先权是Conn
	''/*      方法：Transfer("源数据表名","字段列表","转移条件")
	''/*           “字段列表；转移条件”格式与SQL的“字段列表”,“查询条件”格式相同
	''/*            "字段列表"为空则是所有字段，“查询条件”为空则获取所有数据
	''/***************************************************************************
	private s_Conn
	private objExcelApp,objExcelSheet,objExcelBook
	private sChar,endChar
	''/***************************************************************************
	''/*             全局变量
	''/*外部直接使用：[Obj].SourceFile=源文件名   [Obj].TargetFile=目标文件名
	''/***************************************************************************
	public SourceFile,TargetFile

	private sub class_Initialize
		   sChar="ABCDEFGHIJKLNMOPQRSTUVWXYZ"
			  objExcelApp=Null
		   s_Conn=Null
	end sub
	private sub class_Terminate
		   if IsObject(s_Conn) And Not IsNull(s_Conn) then
				 s_Conn.Close
				 set s_Conn=Nothing
			  end if
		   CloseExcel
	end sub

	''/***************************************************************************
	''/*             设置/返回Conn对象
	''/*说明：添加这个是为了其它数据库(如：MSSQL）到ACCESS数据库的数据转移而设置的
	''/***************************************************************************
	public property set Conn(sNewValue)
		  if Not IsObject(sNewValue) then
			  s_Conn=Null
		   else
			  set s_Conn=sNewValue
		   end if
	end property
	public property get Conn
		  if IsObject(s_Conn) then
			  set Conn=s_Conn
		   else
			  s_Conn=Null
		   end if
	end property

	''/***************************************************************************
	''/*             数据转移
	''/*函数功能：转移源数据到TargetFile数据库文件
	''/*函数说明：利用SQL语句的Select Into In方法转移
	''/*函数返回：返回一些状态代码true = 转移数据成功   false = 转移数据失败
	''/*函数参数：sTableName = 源数据库的表名
	''/*         sCol = 要转移数据的字段列表，格式则同Select 的字段列表格式相同
	''/*         sSql = 转移数据时的条件 同SQL语句的Where 后面语句格式一样
	''/***************************************************************************
	public function Transfer(sTableName,sCol,sSql)
	On Error Resume next
	dim SQL,Rs
	dim iFieldsCount,iMod,iiMod,iCount,i
		   if TargetFile="" then         ''没有目标保存文件，转移失败
			  Transfer=false
				exit function
		   end if
		  if Not InitConn then          ''如果不能初始化Conn对象则转移数据出错
			  Transfer=false
				exit function
		   end if
		  if Not InitExcel then         ''如果不能初始化Excel对象则转移数据出错
			  Transfer=false
				exit function
		   end if
		  if sSql<>"" then              ''条件查询
			  sSql=" Where "&sSql
		   end if
		   if sCol="" then               ''字段列表,以","分隔
			  sCol="*"
		   end if
		   set Rs=Server.CreateObject("ADODB.Recordset")
		  SQL="SELECT "&sCol&" From ["&sTableName&"]"&sSql
		   Rs.Open SQL,s_Conn,1,1
		   if Err.Number<>0 then         ''出错则转移数据错误，否则转移数据成功
			  Err.Clear
			 Transfer=false
				set Rs=Nothing
				CloseExcel
				exit function
		   end if
		  iFieldsCount=Rs.Fields.Count
		  ''没字段和没有记录则退出
		  if iFieldsCount<1 Or Rs.Eof then
			 Transfer=false
				set Rs=Nothing
				CloseExcel
				exit function
		  end if
		  ''获取单元格的结尾字母
		  iMod=iFieldsCount Mod 26
		  iCount=iFieldsCount \ 26
		  if iMod=0 then
			   iMod=26
			   iCount=iCount
		  end if
		  endChar=""
		  Do While iCount>0
			   iiMod=iCount Mod 26
			   iCount=iCount \ 26
			   if iiMod=0 then
					iiMod=26
				  iCount=iCount
			   end if
			   endChar=Mid(sChar,iiMod,1)&endChar
		  Loop
		  endChar=endChar&Mid(sChar,iMod,1)
		  dim sExe    ''运行字符串

		  ''字段名列表
		  i=1
		  sExe="objExcelSheet.Range(""A"&i&":"&endChar&i&""").Value = Array("
		  for iMod=0 To iFieldsCount-1
			  sExe=sExe&""""&Rs.Fields(iMod).Name
				if iMod=iFieldsCount-1 then
					 sExe=sExe&""")"
				else
				   sExe=sExe&""","
				end if
		  next
		  Execute sExe      ''写字段名
		  if Err.Number<>0 then         ''出错则转移数据错误，否则转移数据成功
			 Err.Clear
			Transfer=false
			   Rs.Close
			   set Rs=Nothing
			   CloseExcel
			   exit function
		  end if
		  i=2
		  Do Until Rs.Eof
			   sExe="objExcelSheet.Range(""A"&i&":"&endChar&i&""").Value = Array("
			   for iMod=0 to iFieldsCount-1
				 sExe=sExe&""""&Rs.Fields(iMod).Value
				   if iMod=iFieldsCount-1 then
						sExe=sExe&""")"
					 else
					sExe=sExe&""","
					 end if
			   next
			   Execute sExe               ''写第i个记录
			   i=i+1
			   Rs.Movenext
		  Loop
		  if Err.Number<>0 then         ''出错则转移数据错误，否则转移数据成功
			 Err.Clear
			Transfer=false
			   Rs.Close
			   set Rs=Nothing
			   CloseExcel
			   exit function
		  end if
		  ''保存文件
		  objExcelBook.SaveAs  TargetFile
		  if Err.Number<>0 then         ''出错则转移数据错误，否则转移数据成功
			 Err.Clear
			Transfer=false
			   Rs.Close
			   set Rs=Nothing
			   CloseExcel
			   exit function
		  end if
		  Rs.Close
		  set Rs=Nothing
		  CloseExcel
		 Transfer=true
	end function

	''/***************************************************************************
	''/*             初始化Excel组件对象
	''/*
	''/***************************************************************************
	private function InitExcel()
	On Error Resume next
		   if Not IsObject(objExcelApp) Or IsNull(objExcelApp) then
				 set objExcelApp=Server.CreateObject("Excel.Application")
				 objExcelApp.DisplayAlerts = false
				 objExcelApp.Application.Visible = false
				 objExcelApp.WorkBooks.add
				 set objExcelBook=objExcelApp.ActiveWorkBook
			  set objExcelSheet = objExcelBook.Sheets(1)
				 if Err.Number<>0 then
				 CloseExcel
					  InitExcel=false
					  Err.Clear
					  exit function
				 end if
			  end if
			  InitExcel=true
	end function
	private sub CloseExcel
	On Error Resume next
		   if IsObject(objExcelApp) then
				 objExcelApp.Quit
				 set objExcelSheet=Nothing
				 set objExcelBook=Nothing
				 set objExcelApp=Nothing
			  end if
			objExcelApp=Null
	end sub

	''/***************************************************************************
	''/*             初始化Adodb.Connection组件对象
	''/*
	''/***************************************************************************
	private function InitConn()
	On Error Resume next
	dim ConnStr
		   if Not IsObject(s_Conn) Or IsNull(s_Conn) then
				 if SourceFile="" then
					InitConn=false
					  exit function
				 else
					set s_Conn=Server.CreateObject("ADODB.Connection")
					ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SourceFile
					s_Conn.Open ConnStr
					  if Err.Number<>0 then
						 InitConn=false
						   Err.Clear
						   s_Conn=Null
						   exit function
					  end if
				 end if
			  end if
			  InitConn=true
	end function
end class
%>