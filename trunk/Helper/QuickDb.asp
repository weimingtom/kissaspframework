<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A database class
'	File Name	: QuickDb.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*************************************************************
'	Initialize the class
'*************************************************************

Class Kiss_QuickDb
	private Conn, ConnStr
	private SqlDatabaseName, SqlPassword, SqlUsername, SqlLocalName, SqlNowString
	public rs

	private sub Class_Initialize()
		SqlDatabaseName = "db"
		SqlUsername = "sa"
		SqlPassword = "123456"
		SqlLocalName = "a01"
		SqlNowString = "GetDate()"
		OpenDb
	end sub

	private sub OpenDb()
		On error resume next
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & Replace(SqlPassword, Chr(0), "") & ";Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
		set Conn = CreateObject("ADODB.Connection")
		Conn.Open ConnStr
		if Err then
			Err.Clear
			set Conn = nothing
			On error goto 0
			Err.Raise 1, "MyClass", "Connection fails"
		end if
		set rs = server.createobject("ADODB.Recordset")
	end sub

	public sub setRs(strsql,CursorAndLockType) 
		dim c,l
		if CursorAndLockType="" then
			CursorAndLockType=13
		end if
		if CursorAndLockType<9 then
			CursorAndLockType=13
		end if
		c=left(CursorAndLockType,1)
		l=right(CursorAndLockType,1)
		rs.Open strsql, Conn, c,l
	end sub

	public sub Execute(sql,OutRs)
		if instr(Ucase(sql),Ucase("select"))>0 then
			set OutRs = Conn.Execute(sql)
			else
			Call Conn.Execute(sql)
			OutRs=1
		end if
	end sub

	public sub SelectDb(Table, Where,OutRs)
		dim sqlstr
		sqlstr = "Select * from " & Table & " Where " & Where
		Call Execute(sqlstr,OutRs)
	end sub

	public function Delete(Table, Where)
		dim Flag, sqlstr,NullTmp
		Flag = False
		On Error Resume next
		sqlstr = "delete from " & Table & " where " & Where
		Execute sqlstr,NullTmp
		if Err.Number = 0 then
			Flag = True
		end if
		Delete = Flag
	end function

	public function Insert(Table, MyFields, Values)
		dim sql,NullTmp
		Insert = False
		sql = "INSERT INTO Table1(fields) VALUES (values)"
		sql = Replace(sql, "Table1", Table)
		sql = Replace(sql, "fields", MyFields)
		sql = Replace(sql, "values", Values)
		On error Resume next
		Execute sql,NullTmp
		if Err.Number = 0 then
			Insert = True
		end if
		On error goto 0
	end function

	public function Update(Table,Field,Value,Where)
		Update=False
		dim SqlStr
		if SqlInject(Table) or SqlInject(Field) then
			die "There is a not secure element."
		end if
		SqlStr="Update [Table] set [Field]=Value Where Where1"
		SqlStr=Replace(SqlStr,"Table",Table)
		SqlStr=Replace(SqlStr,"Field",Field)
		SqlStr=Replace(SqlStr,"Value",Value)
		SqlStr=Replace(SqlStr,"Where1",Where)
		On error resume next
		dim QDb,TmpRs
		set QDb=new QuickDb
		Call QDb.Execute(SqlStr,TmpRs)
		if err.number=0 then
			if TmpRs=1 then
				Update=True
			end if
		end if
		set QDb=nothing
		On error goto 0
	end function

	function SqlInject(ByVal SqlStr)
		SqlInject = True
		dim TmpStr, ArrStr, OriginalLen
		TmpStr = "'',',or,not,and,--, ,chr,asc"
		OriginalLen = Len(SqlStr)
		ArrStr = Split(TmpStr, ",")
		TmpStr = UCase(TmpStr)
		for i = 0 To UBound(ArrStr)
			SqlStr = Replace(SqlStr, UCase(ArrStr(i)), "")
		next
		if Len(SqlStr) = OriginalLen then
			SqlInject = False
		end if
	end function

	private sub Class_Terminate()
		if IsObject(Conn) then
			if Conn.State <> 0 then
				Conn.Close
				set Conn = nothing
			end if
		end if

		if ifIsObject(rs) then
			if rs.State <> 0 then
				rs.Close
				set rs = nothing
			end if
		end if
	end sub
end class

%>