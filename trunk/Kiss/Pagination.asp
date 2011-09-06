<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A pagination class
'	File Name	: Pagination.asp
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

'*************************************************************
'具体用法
'dim strDbPath
'dim connstr
'dim mp
'set mp = New Page
'strDbPath = "fenye/db.mdb"
'connstr  = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
'connstr  = connstr & Server.MapPath(strDbPath)
'set conn  = Server.CreateObject("Adodb.Connection")
'conn.open connstr
'set rs = mp.Execute("select * from table1",conn,29)
'while not rs.eof
'    response.write rs("aaaa")&"<br>"
'    rs.Movenext
'wend
'mp.pageDispaly()
'*************************************************************

class Kiss_Pagination
	public className	'Class name
    private Page_Conn,Page_StrSql,Page_TotalStrSql,Page_RS,Page_TotalRS
    private Page_PageSize
    private Page_PageAbsolute,Page_PageTotal,Page_RecordTotal
    private Page_Url

    public property let conn(strConn)
		set Page_Conn = strConn
    end property

    public property let PageSize(intPageSize)
        Page_PageSize = Cint(intPageSize)
    end property

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
	private sub class_initialize()
		classname = "Kiss_Pagination"
	end sub

 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

    public function PageExecute(strSql)
        Page_PageAbsolute = Page_PageAbsoluteRequest()
        Page_TotalStrSql = formatPage_TotalStrSql(strSql) 
        set Page_TotalRS = Page_Conn.execute(Page_TotalStrSql)
        Page_RecordTotal = Page_TotalRS("total")
        Page_PageTotal = Cint(Page_RecordTotal/Page_PageSize)
        Page_StrSql = formatPage_StrSql(strSql)
        set Page_RS = Page_Conn.execute(Page_StrSql)
        dim i
        i = 0 
        while not Page_RS.eof and  i<(Page_PageAbsolute-1)*Page_PageSize
            i = i + 1
            Page_RS.Movenext
        wend
        set PageExecute = Page_RS 
    end function

    public function Execute(strSql,strConn,intPageSize)
        conn = strConn
        PageSize = intPageSize
        set Execute = PageExecute(strSql)
    end function

    public function pageDispaly()
        Page_Url = ReadPage_Url
        firstPageTag = "<font face=webdings>9</font>"  '|<<
        LastPageTag = "<font face=webdings>:</font>"  '>>|
        previewPageTag = "<font face=webdings>7</font>"  '<<
        nextPageTag = "<font face=webdings>8</font>"  '>>
        dim strAnd
        if instr(Page_Url,"?")=0 then
            strAnd = "?"
        else
            strAnd = "&"
        end if
        response.write "<table width=100%  border=0 cellspacing=0 cellpadding=0>"
        response.write "<tr>"
        response.write "<td align=left>"
        response.write  "页次:"&Page_PageAbsolute&"/"&Page_PageTotal&"页 "
        response.write  "主题数:"&Page_RecordTotal
        response.write "</td>"
        response.write "<td align=right>"
        response.write  "分页:"
        if Page_PageAbsolute>10 then
            response.write  "<a href='"&Page_Url&strAnd&"Page_PageNo=1'>"&firstPageTag&"</a>"
            response.write  "<a href='"&Page_Url&strAnd&"Page_PageNo="&(Page_PageAbsolute-10)&"'>"&previewPageTag&"</a>"
        else
            response.write  firstPageTag
            response.write  previewPageTag
        end if
        response.write " "
        dim CurrentStartPage,i
        i = 1
        CurrentStartPage=(Cint(Page_PageAbsolute)\10)*10+1
        if Cint(Page_PageAbsolute) mod 10=0 then
            CurrentStartPage = CurrentStartPage - 10
        end if
        while i<11 and CurrentStartPage<Page_PageTotal+1
            if CurrentStartPage < 10 then
                formatCurrentStartPage = "0" & CurrentStartPage
            else
                formatCurrentStartPage = CurrentStartPage
            end if
            response.write  "<a href='"&Page_Url&strAnd&"Page_PageNo="&CurrentStartPage&"'>"&formatCurrentStartPage&"</a> "
            i = i + 1
            CurrentStartPage = CurrentStartPage + 1
        wend
        if Page_PageAbsolute<(Page_PageTotal-10) then
            response.write  "<a href='"&Page_Url&strAnd&"Page_PageNo="&(Page_PageAbsolute+10)&"'>"&nextPageTag&"</a>"
            response.write  "<a href='"&Page_Url&strAnd&"Page_PageNo="&Page_PageTotal&"'>"&LastPageTag&"</a>"
        else
            response.write  nextPageTag
            response.write  LastPageTag
        end if
        response.write  ""
        response.write "</td>"
        response.write "</tr>" 
        response.write "</table>"
    end function

    public function getPageNo()
        getPageNo = cint(Page_PageAbsolute)
    end function

    public function getPageCount()
        getPageCount = cint(Page_PageTotal)
    end function

    public function getPageNoName()
        getPageNoName = "Page_PageNo"
    end function

    public function getPageSize()
        getPageSize = Page_PageSize
    end function

    public function getRecordTotal()
        getRecordTotal = Page_RecordTotal
    end function
    
    private function formatPage_TotalStrSql(strSql)
        formatPage_TotalStrSql = "select count(*) as total "
        formatPage_TotalStrSql = formatPage_TotalStrSql & Mid(strSql,instr(strSql,"from"))
        formatPage_TotalStrSql = Mid(formatPage_TotalStrSql,1,instr(formatPage_TotalStrSql&"order by","order by")-1)
    end function

    private function formatPage_StrSql(strSql)
        formatPage_StrSql = replace(strSql,"select","select top "&(Page_PageAbsolute*Cint(Page_PageSize)))
    end function

    private function Page_PageAbsoluteRequest()
        if request("Page_PageNo")="" then 
            Page_PageAbsoluteRequest = 1
        else
            if IsNumeric(request("Page_PageNo")) then
                Page_PageAbsoluteRequest = request("Page_PageNo")
            else
                Page_PageAbsoluteRequest = 1
            end if
        end if
    end function

    private function ReadPage_Url()
        ReadPage_Url = Request.ServerVariables("URL")
        if Request.QueryString<>"" then
            ReadPage_Url = ReadPage_Url & "?" & Request.QueryString 
        end if
        set re = new RegExp
        re.Pattern = "[&|?]Page_PageNo=\d+?"
        re.IgnoreCase = true
        re.multiLine = true
        re.global = true
        set Matches = re.Execute(ReadPage_Url) 
        for Each Match in Matches  
            tmpMatch = Match.Value
            ReadPage_Url = replace(ReadPage_Url,tmpMatch,"")
        next
    end function
end class
%> 