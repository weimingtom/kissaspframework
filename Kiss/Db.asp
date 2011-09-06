<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: DataBase Interface class
'	File Name	: Db.asp
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
dim Db
set Db = New Kiss_Db
dim C
set C = Db

class Kiss_Db
	public className
	public SqlNow
	public Dbs
	public Opened
	public DbsType
	private SqlLocalName,SqlDatabaseName,SqlUsername,SqlPassword
	private ConnStr
	public Conn
	public ADOX
	public QueryTimes	'search number
	
	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
	private sub class_Initialize()
		className = "Kiss_Db"
		SqlNow = "Now()"
		Opened = false
		QueryTimes = 0
	end sub

 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
		if isObject(Conn) then set Conn = Nothing
	end sub

 	'************************************************************* 
	' Name: Open
	' Purpose: Open the database connection, only construct the connecting string
	'************************************************************* 
	public sub Open(DbsName, nDbsType)

		dim tt
		' Support Access£¬MSSQL£¬MYSQL, DB2, ORACLE
		Dbs = DbsName
		DbsType = Cint("0" & nDbsType)

		if DbsType>6 Or DbsType<1 then DbsType = DTAccess
		Select Case DbsType
			Case DTAccess
			'Access
				Dbs = Server.MapPath(Dbs)
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbs
			Case DTExcel
			'Excel
				Dbs = Server.MapPath(Dbs)
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;Hdr=Yes"";Data Source = " & Dbs
			Case DTMSSQL
			'SQL Server
				tt = split(Dbs,"$$$")
				SqlLocalName = tt(1)
				SqlDatabaseName = tt(2)
				SqlUsername = tt(3)
				SqlPassword = tt(4)
				ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
			Case DTMySQL
			'MySQL Server
				tt = split(Dbs,"$$$")
				SqlLocalName = tt(1)
				SqlDatabaseName = tt(2)
				SqlUsername = tt(3)
				SqlPassword = tt(4)
				ConnStr = "driver = {MySQL ODBC 3.51 Driver}; UID = " & SqlUsername & "; Pwd = " & SqlPassword & "; Database = " & SqlDatabaseName & "; Server = " & SqlLocalName & ";option = 16386"
			Case DTDB2
			'DB2
				tt = split(Dbs,"$$$")
				SqlLocalName = tt(1)	'table patternk Note: must add dot such as "NULLID." as prefix. Above three set it to 0
				'Cache("SqlLocalName") = SqlLocalName
				SqlDatabaseName = tt(2)	'ODBC data source name
				SqlUsername = tt(3)		'DB2 user
				SqlPassword = tt(4)		'DB2 password
				ConnStr = "DSN="& SqlDatabaseName &";Uid="& SqlUsername &";Pwd="& SqlPassword &";"
			Case DTOracle
			'Oracle
				tt = split(Dbs,"$$$")
				SqlLocalName = tt(1)	'table space Note: must add dot such as "SYSTEM." as prefix.  Above three set it to 0
				'Cache("SqlLocalName") = SqlLocalName
				SqlDatabaseName = tt(2)	'Oracle database name
				SqlUsername = tt(3)		'Oracle user
				SqlPassword = tt(4)		'Oracle password
'				ConnStr = "Provider=MSDAORA.1;Persist Security Info=true;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
				ConnStr = "Provider=OraOLEDB.Oracle.1;Persist Security Info=true;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
		end select
		'Cache connection string
		'Cache("ConnStr") = ConnStr
	end sub
	
 	'************************************************************* 
	' Name: OpenDB
	' Purpose: Formally open database
	'************************************************************* 
	public sub OpenDB()
		if Opened = true then
			show404("The database have already opened.")
			exit sub
		end if

		'dim TemData
		'TemData = Cache("ConnStr")

		if ConnStr = "" then
			show404("Connection string is empty.")
			exit sub
		end if
		
		On error resume next
		err.clear
		
		set Conn = Server.CreateObject("ADODB.Connection")
		if err.Number>0 then throwException("Don't support the ADO.")

		Conn.open ConnStr
		
		if Conn.errors.Count>0 then
			'Cache("ConnStr") = ""
			show404("Connect the database" & Dbs & "fails, total" & Conn.errors.Count & " " & Conn.errors(Conn.errors.Count - 1))
			Conn.errors.Clear
		end if

		if err.Number>0 then
			set Conn = Nothing
			'Cache("ConnStr") = ""
			throwException("The database" & Dbs & "connection fails." & ConnStr)
			err.Clear
		end if
		set ADOX = Server.CreateObject("ADOX.Catalog")
		set ADOX.ActiveConnection = Conn
		if err.Number>0 then
			throwException("Don't support ADOX.")
		end if
		
		Opened = true
	end sub
	
 	'************************************************************* 
	' Name: Data
	' Purpose: Get the table name
	'************************************************************* 
	public Default property get Data(getType,TableName)
		
		if Opened = false then OpenDB
		
		on error resume next
		Select Case LCase(getType)
			Case "tables"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE
'				set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,"table"))
				set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,TableName))
			Case "table"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE
'				set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,"table"))
				set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,TableName,empty))
			Case "columns"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,COLUMN_NAME
				set Data = Conn.OpenSchema(adSchemaColumns,Array(empty,empty,TableName,empty))
			Case "primarykeys"
				'PK_TABLE_CATALOG,PK_TABLE_SCHEMA,PK_TABLE_NAME
				set Data = Conn.OpenSchema(adSchemaPrimaryKeys,Array(empty,empty,TableName))
			Case "sql"
				set Data = Conn.execute(TableName)
			Case else
				set Data = Conn.OpenSchema(getType, TableName)
		end select
		'show404("EOF = " & Data.Eof & "  BOF = " & Data.Bof & "  RecordCount = " & Data.RecordCount)
	end property
	
 	'************************************************************* 
	' Name: find
	' Purpose: Get the data record set
	'************************************************************* 
	public function find(sql,w)	'w: 0 can't alter the data set, 1 can
		'Debug.AddDb = "Sql " & sql

		if Opened = false then OpenDB

		QueryTimes = QueryTimes + 1

		if w<>0 And w<>1 then w = 0

		on error resume next
		err.clear
		set find = Server.CreateObject("ADODB.Recordset")
		find.open sql, Conn, 1, w * 2 + 1

		if Conn.errors.Count>0 then
			show404("The database object fails " & Conn.errors(Conn.errors.Count-1))
			Conn.errors.Clear
		end if

		if err.number<>0 then
			throwException("Error Sql " & sql)
		end if
		
		Const adStateClosed = 0
		if find.State = adStateClosed then show404("No opening the data set.")
		
		'show404("EOF = " & findAll.Eof & "  BOF = " & findAll.Bof & "  RecordCount = " & findAll.RecordCount)
	end function

 	'************************************************************* 
	' Name: findAll
	' Purpose: Get all the data record set
	'************************************************************* 
	public function findAll(name, w)
		dim sql
		sql = "select * from " & name

		if Opened = false then OpenDB

		QueryTimes = QueryTimes + 1

		if w<>0 And w<>1 then w = 0

		on error resume next
		err.clear
		set findAll = Server.CreateObject("ADODB.Recordset")
		findAll.open sql, Conn, 1, w * 2 + 1

		if Conn.errors.Count>0 then
			show404("The database object fails " & Conn.errors(Conn.errors.Count-1))
			Conn.errors.Clear
		end if

		if err.number <> 0 then
			throwException("Error Sql " & sql)
		end if
		
		Const adStateClosed = 0
		if findAll.State = adStateClosed then show404("No opening the data set.")
		
		'show404("EOF = " & findAll.Eof & "  BOF = " & findAll.Bof & "  RecordCount = " & findAll.RecordCount)
	end function
	
 	'************************************************************* 
	' Name: execute
	' Purpose: Execute SQL
	'************************************************************* 
	public function execute(sql)
		if Opened = false then OpenDB
		
		QueryTimes = QueryTimes + 1

		on error resume next
		err.clear
		set execute = Conn.execute(sql)
		
		if Conn.errors.Count>0 then
			show404("Error: " & Conn.errors(Conn.errors.Count-1))
			Conn.errors.Clear
		elseif err.number<>0 then
			throwException("Error of the SQL string " & sql)
		end if
	end function

 	'************************************************************* 
	' Name: createDatabase
	' Purpose: Create a database
	'************************************************************* 
	public function createDatabase(name)
		execute "CREATE DATABASE " & name
	end function
	
 	'************************************************************* 
	' Name: dropDatabase
	' Purpose: Delete a database
	'************************************************************* 
	public function dropDatabase(name)
		execute "DROP DATABASE " & name
	end function
	
 	'************************************************************* 
	' Name: renameTable
	' Purpose: Rename table
	'************************************************************* 
	public function renameTable(vTableName, vNewTableName)
		On error Resume next
		ADOX.Tables(vTableName).Name = vNewTableName
		if err.Number>0 then
			throwException(vTableName & "Renaming fails.")
			err.Clear
		end if
	end function
	
 	'************************************************************* 
	' Name: dropTable
	' Purpose: Delete a data talbe
	'************************************************************* 
	public function dropTable(vTableName)
		execute "DROP TABLE IF EXISTS " & vTableName
	end function
		
 	'************************************************************* 
	' Name: resetTable
	' Purpose: Reset a data table
	'************************************************************* 
	public function resetTable(vTableName, vIDName)
		execute "Alter Table " & vTableName & " ALTER COLUMN " & vIDName & " COUNTER (1, 1)"
	end function

 	'************************************************************* 
	' Name: showRs
	' Purpose: Show a data record set
	'************************************************************* 
	public sub showRs(ByRef iRs)
		dim i
		
		if Not IsObject(iRs) then
			show404("Record set type is wrong.")
			exit sub
		end if
		echo "<table><tr class=hth>" & vbNewLine
		for i = 0 To iRs.fields.Count-1
			echo "<td>"&iRs(i).name&"</td>"
		next
		echo "</tr>" & vbNewLine
		if iRs.eof And iRs.bof then
			show404("No data.")
'			exit sub
		else
			iRs.MoveFirst
			While Not iRs.eof
				echo "<tr>" & vbNewLine
				for i = 0 To iRs.fields.Count-1
					echo "<td class=dtd>" & Server.HTMLEncode(iRs(i) & "") & "</td>"
				next
				echo  vbNewLine & "</tr>" & vbNewLine
				iRs.Movenext
			Wend
			iRs.MoveFirst
		end if
		echo "</table>"
	end sub
	
 	'************************************************************* 
	' Name: toXML
	' Purpose: Print XML
	'************************************************************* 
	public function toXML(ByRef iRs)

		if Not IsObject(iRs) then
			show404("No data.")
			exit function
		end if

		toXML = ""
		if Not IsObject(iRs) Or iRs.RecordCount<1 then exit function
		dim i
		While Not iRs.eof
			toXML = toXML & vbNewLine & "<RsData>"
			for i = 0 To iRs.Fields.Count-1
				toXML = toXML & vbNewLine & "	<" & iRs(i).name & ">" & Server.HTMLEnCode(iRs(i).value & "") & "</" & iRs(i).name & ">"
			next
			toXML = toXML & vbNewLine & "</RsData>"
			iRs.Movenext
		wend
		echo toXML
	end function

 	'************************************************************* 
	' Name: sqlServer_To_Access
	' Purpose: SqlServer(97-2000) to Access(97-2000)
	'************************************************************* 
    public function sqlServer_To_Access(Sql)
        dim regEx, Matches, Match
        'Create regular expression
        set regEx = New RegExp
        regEx.IgnoreCase = true
        regEx.Global = true
        regEx.MultiLine = true

        'Convert: getDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"NOW()")

        'Convert: UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"UCASE($1)")

        'Convert:date forms
        '2004-23-23 11:11:10 standard form
        regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
        Sql = regEx.Replace(Sql,"#$1#")
        
        regEx.Pattern = "DATEDIFF\([\s]?(second|minute|hour|day|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
        set Matches = regEx.ExeCute(Sql)
        dim temStr
        for Each Match In Matches
            temStr = "DATEDIFF("
            Select Case lcase(Match.subMatches(0))
                Case "second" :
                    temStr = temStr & "'s'"
                Case "minute" :
                    temStr = temStr & "'n'"
                Case "hour" :
                    temStr = temStr & "'h'"
                Case "day" :
                    temStr = temStr & "'d'"
                Case "month" :
                    temStr = temStr & "'m'"
                Case "year" :
                    temStr = temStr & "'y'"
            end Select
            temStr = temStr & "," & Match.subMatches(1) & "," &  Match.subMatches(2) & Match.subMatches(3)
            Sql = Replace(Sql,Match.Value,temStr,1,1)
        next

        'Convert:Insert
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

        set regEx = Nothing
        sqlServer_To_Access = Sql
    end function
end class
%>