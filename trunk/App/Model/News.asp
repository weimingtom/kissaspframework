<%
public News
set News = new KNews

class KNews
	private rs, sql

	public function show()
		'Response.write "我是新闻"
		sql = "Select Top 30 * From NL_News"
		'set rs = Db.find(sql,0)
		set rs = Db.findAll("NL_News",0)
		'Db.toXML rs
		echo "<ul>" & vbNewLine
		do while not rs.eof
			echo "<li>" & rs("title") & "</li>" & vbNewLine
			rs.movenext
		loop
		echo "</ul>"

		rs.close
	end function

	public function getNews()
		dim temp
		temp = ""
		'Response.write "我是新闻"
		sql = "Select Top 30 * From NL_News"
		set rs = Db.find(sql,0)
		do while not rs.eof
			temp = temp & rs("title") & "<br />"
			rs.movenext
		loop

		getNews = temp

		rs.close
	end function

	public function list()
		View.UpdateBlock "news"
		sql = "Select Top 30 * From NL_News"
		set rs = Db.find(sql,0)
		do while not rs.eof
			View.assign "news", rs("title")&""
			View.assignBlock "news"
			rs.movenext
		loop
	end function

end class
%>