<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: TableDataGateway class
'	File Name	: TableDataGateway.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** dim test,i
'*** dim t
'*** set test = new Kiss_TableDataGateway

'*** With test
'*** 	.table "Blog"
'*** 	.fields "id,title"
'*** 	' .where "id = 28"
'*** 	' query
'*** 	.query
'*** 
'*** 	' another query method
'*** 	' .setSql "Select id,title from Blog"
'*** 	' .exec_sql t
'*** 
'*** 	'modify
'*** 	.fields "title"
'*** 	.fieldsValue "Kiss Asp Framework"
'*** 	.update
'*** 
'*** 	'add
'*** 	.fields "title"
'*** 	.fieldsValue "Kiss Asp Framework"
'*** 	.insert
'*** 
'*** 	'delete
'*** 	.table "Blog"
'*** 	.where "id=28"
'*** 	.delete
'*** 
'*** end With 

'*************************************************************
'	Initialize the class
'*************************************************************
dim tdg
set tdg = new Kiss_TableDataGateway

class Kiss_TableDataGateway
	public className	'Class name
	private strConn

	private objConn
	private objRs
	private objCmd
	
	private p_dbName
	private p_table
	private p_field
	private p_fieldVal
	private p_orderby
	private p_where
	private p_sql

	public rs

	private p_SqlLocalName,p_SqlDatabaseName,p_SqlUsername,p_SqlPassword

	private spliter
	private Status 'Command status 0 for success and 1 for failure

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
	private sub class_Initialize()
		classname = "Kiss_TableDataGateway"
		spliter = ","
	end sub

 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
		Call ClearRecord()
		Call ClearCommand()
		Call ClearConnection()
	end sub

 	'************************************************************* 
	' Name: open
	' Purpose: Open a database
	'************************************************************* 
	public sub open(pDbName, pDbType)
		dim tt

		p_dbName = pDbName
		if pDbType>6 or pDbType<1 then pDbType = DTAccess
		Select Case pDbType
			Case DTAccess
			'Access
				p_dbName = Server.MapPath(p_dbName)
				strConn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & p_dbName
			Case DTExcel
			'Excel
				p_dbName = Server.MapPath(p_dbName)
				strConn = "Provider = Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;Hdr=Yes"";Data Source = " & p_dbName
			Case DTMSSQL
			'SQL Server
				tt = split(p_dbName,"$$$")
				p_SqlLocalName = tt(1)
				p_SqlDatabaseName = tt(2)
				p_SqlUsername = tt(3)
				p_SqlPassword = tt(4)
				strConn = "Provider = Sqloledb; User ID = " & p_SqlUsername & "; Password = " & p_SqlPassword & "; Initial Catalog = " & p_SqlDatabaseName & "; Data Source = " & p_SqlLocalName & ";"
			Case DTMySQL
			'MySQL Server
				tt = split(p_dbName,"$$$")
				p_SqlLocalName = tt(1)
				p_SqlDatabaseName = tt(2)
				p_SqlUsername = tt(3)
				p_SqlPassword = tt(4)
				strConn = "driver = {MySQL ODBC 3.51 Driver}; UID = " & p_SqlUsername & "; Pwd = " & p_SqlPassword & "; Database = " & p_SqlDatabaseName & "; Server = " & p_SqlLocalName & ";option = 16386"
			Case DTDB2
			'DB2
				tt = split(p_dbName,"$$$")
				p_SqlLocalName = tt(1)	'table patternk Note: must add dot such as "NULLID." as prefix. Above three set it to 0
				'Cache("p_SqlLocalName") = p_SqlLocalName
				p_SqlDatabaseName = tt(2)	'ODBC data source name
				p_SqlUsername = tt(3)		'DB2 user
				p_SqlPassword = tt(4)		'DB2 password
				strConn = "DSN="& p_SqlDatabaseName &";Uid="& p_SqlUsername &";Pwd="& p_SqlPassword &";"
			Case DTOracle
			'Oracle
				tt = split(p_dbName,"$$$")
				p_SqlLocalName = tt(1)	'table space Note: must add dot such as "SYSTEM." as prefix.  Above three set it to 0
				'Cache("p_SqlLocalName") = p_SqlLocalName
				p_SqlDatabaseName = tt(2)	'Oracle database name
				p_SqlUsername = tt(3)		'Oracle user
				p_SqlPassword = tt(4)		'Oracle password
'				strConn = "Provider=MSDAORA.1;Persist Security Info=true;User ID=" & p_SqlUsername & ";Password=" & p_SqlPassword & ";Data Source="& p_SqlDatabaseName &";"
				strConn = "Provider=OraOLEDB.Oracle.1;Persist Security Info=true;User ID=" & p_SqlUsername & ";Password=" & p_SqlPassword & ";Data Source="& p_SqlDatabaseName &";"
		end select
		'Create database connection
		Connection()
	end sub

	'************************************************************* 
	' Name: ClearObject
	' Param: obj as a object
	' Purpose: Close a object
	'************************************************************* 
	private sub ClearObject(ByRef pObj)
		on error resume next
		if IsObject(pObj) = true then
			pObj.Close()
			set pObj = Nothing
		end if
		if err.Number <> 0 then
			throwException("Clearing object fails.")
		end if
	end sub

	'************************************************************* 
	' Name: Connection
	' Purpose: Create a database connection
	'************************************************************* 
	private sub Connection()
		set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open strConn
	end sub

	'************************************************************* 
	' Name: ClearConnection
	' Purpose: Close a database connection
	'************************************************************* 
	private sub ClearConnection()
		ClearObject(objConn)
	end sub

	'************************************************************* 
	' Name: Recordset
	' Purpose: Create a record set
	'************************************************************* 
	private sub Recordset()
		ClearObject(objRs)
		set objRs = Server.CreateObject("ADODB.Recordset")
	end sub

	'************************************************************* 
	' Name: ClearRecordset
	' Purpose: Close a record set
	'************************************************************* 
	private sub ClearRecordset()
		Call ClearObject(objRs)
	end sub

	'************************************************************* 
	' Name: Command
	' Purpose: Create a command object
	'************************************************************* 
	private sub Command()
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = objConn
	end sub

	'************************************************************* 
	' Name: ClearCommand
	' Purpose: Close a command object
	'************************************************************* 
	private sub ClearCommand()
		Call ClearObject(objcmd)
	end sub

	'************************************************************* 
	' Name: table
	' Param: pTable as a table name
	' Purpose: Get the data table
	'************************************************************* 
	public sub table(pTable)
		p_table = pTable
	end sub

	'************************************************************* 
	' Name: fields
	' Param: fields as data field name
	' Purpose: Get the data field name
	'************************************************************* 
	public sub fields(pFields)
		if isarray(pFields) then
			p_field = toString(pFields)
		else
			p_field = pFields
		end if
	end sub

	'************************************************************* 
	' Name: fields
	' Param: fields as record set
	' Purpose: Get the record set
	'************************************************************* 
	public sub fieldsVal(pFieldsVal)
		if isarray(pFieldsVal) then
			p_fieldVal = toString(pFieldsVal)
		else
			p_fieldVal = pFieldsVal
		end if
	end sub

	'************************************************************* 
	' Name: where
	' Param: pWhere as condition of sql
	' Purpose: Get the condition of sql
	'************************************************************* 
	public sub where(pWhere)
		p_where = pWhere
	end sub

	public sub orderby(pOrderby, ptype)
		if ptype = 1 then
			p_orderby = pOrderby
		else
			p_orderby = pOrderby & " desc"
		end if
	end sub

	public sub setSql(ByRef pSql)
		p_sql = ""
		p_sql = pSql
	end sub

	'************************************************************* 
	' Name: isSelecTable
	' Purpose: Check if the table is selected
	'************************************************************* 
	private sub isSelecTable()
		if sqlInject(p_table) then
			die "There is a not secure element of table."
		end if
		
		if p_table = "" then
			die("Please select the data table.")
		end if
	end sub

	'************************************************************* 
	' Name: checkField
	' Purpose: Prevent the inject sql
	'************************************************************* 
	private sub checkField()
		if sqlInject(p_field) then
			die "There is a not secure element of field."
		end if
	end sub

	'************************************************************* 
	' Name: clearSql
	' Purpose: Check if the table is selected
	'************************************************************* 
	private sub clearSql()
		p_field = ""
		p_fieldVal = "" 
		p_orderby = ""
		p_where = ""
	end sub

	'************************************************************* 
	' Name: query
	' Purpose: Query a record set
	'************************************************************* 
	public function query()
		set rs = nothing

		isSelecTable() ' Check if data table was selected
		p_sql = ""
		p_sql = "select "
		dim i,flag,ArrTemp
		flag = 0
		if p_field = "" then
			p_field = " * "
		else
			p_field = Replace(p_field,spliter,",")
		end if

		checkField()
		p_sql = p_sql & p_field
		p_sql = p_sql & " from " & p_table
		if p_where <> "" then
			p_sql = p_sql & " where " & p_where
		end if
		if p_orderby <> "" then
			p_sql = p_sql & " order by " & p_orderby
		end if
		
		'die p_sql
		Recordset()
		Command()
		objCmd.CommandText = p_sql
		set objRs = objCmd.Execute

		set rs = objRs

		clearSql()
	end function

	'************************************************************* 
	' Name: update
	' Purpose: Update the record(s) to database
	'************************************************************* 
	public sub update()
		isSelecTable()
		checkField()
		
		dim ArrField,ArrFieldValue
		p_sql = ""
		ArrField = ""
		ArrFieldValue = ""
		ArrField = Split(p_field,spliter)
		ArrFieldValue = Split(p_fieldVal,spliter)
		
		'Check if the number of field equal number of field value
		call Compare(Ubound(ArrField),UBound(ArrFieldValue))

		p_sql = "Update " & p_table & " set "
		for i = 0 To Ubound(ArrField)
			if IsNumeric(ArrFieldValue(i)) = false then
				p_sql = p_sql & ArrField(i) & " = '" & ArrFieldValue(i) & "'"
				if i <> UBound(ArrField) then p_sql = p_sql & " , "
			else
				p_sql = p_sql & ArrField(i) & " = " & ArrFieldValue(i)
				if i <> UBound(ArrField) then p_sql = p_sql & " , "
			end if
		next
		if p_where = "" then
			die("Please check the conditon of sql.")
		else
			p_sql = p_sql & " where " & p_where
		end if

		'die(p_sql)
		exec_sql()

		clearSql()
	end sub

	'************************************************************* 
	' Name: insert
	' Purpose: insert the record(s) to database
	'************************************************************* 
	public sub insert()
		isSelecTable()
		checkField()
		
		dim ArrFieldValue,ArrField,i
		ArrField = ""
		ArrFieldValue = ""
		p_sql = ""
		ArrField = Split(p_field,spliter)
		ArrFieldValue = Split(p_fieldVal,spliter)

		'Check if the number of field equal number of field value
		call Compare(Ubound(ArrField),UBound(ArrFieldValue))

		p_field = Replace(p_field,spliter,",")
		p_sql = "Insert into " & p_table & "(" & p_field & ") Values("
		for i = 0 To Ubound(ArrFieldValue)
			if IsNumeric(ArrFieldValue(i)) = false then
				p_sql = p_sql & "'" & ArrFieldValue(i) & "'"
				if i <> Ubound(ArrFieldValue) then p_sql = p_sql & ","
			else
				p_sql = p_sql & ArrFieldValue(i)
				if i <> Ubound(ArrFieldValue) then p_sql = p_sql & ","
			end if
		next
		p_sql = p_sql & ")"

		'die p_sql
		exec_sql()
		clearSql()
	end sub

	'************************************************************* 
	' Name: delete
	' Purpose: Delete record(s) from database
	'************************************************************* 
	public sub delete()
		isSelecTable()
		
		p_sql = ""
		if p_where = "" then
			die("Please check the conditon of sql.")
		end if
		p_sql = "Delete * from " & p_table & " where " & p_where

		'die p_sql
		exec_sql()
		clearSql()
	end sub

	'************************************************************* 
	' Name: exec_sql
	' Purpose: Execute a sql
	'************************************************************* 
	public function exec_sql()
		'die p_sql
		if p_sql <> "" then
			Command()
			With objCmd
				.CommandText = p_sql
				.execute
			end With
		end if
	end function

	private function Compare(pCompA, pCompB)
		if pCompA = pCompB then
			Compare = true
		else
			die("The field number is matched.")
			Compare = false
		end if
	end function

	function sqlInject(SqlStr)
		sqlInject = true
		dim TmpStr, ArrStr, OriginalLen
		TmpStr = "'',',or,not,and,--, ,chr,asc"
		OriginalLen = Len(SqlStr)
		ArrStr = Split(TmpStr, ",")
		TmpStr = UCase(TmpStr)
		for i = 0 To UBound(ArrStr)
			SqlStr = Replace(SqlStr, UCase(ArrStr(i)), "")
		next
		if Len(SqlStr) = OriginalLen then
			sqlInject = false
		end if
	end function

end class
%>