<%
'import("Cache")
'import("Session")
import("Db")
import("TableDataGateway")
import("Helper.MSDebug")

class Default

	private sub class_Initialize()
		Db.Open DbPath, DTAccess	' Open database connection
		tdg.open DbPath, DTAccess
	end sub

	private sub class_Terminate()
		set Db		= nothing 'Kiss_Db
		set Sn		= nothing 'Kiss_Session
		set Cache	= nothing 'Kiss_Cache
		set tdg		= nothing 'Kiss_TableDataGateway
	end sub

	public Default function actionIndex()
		dim re, temp, tmpField
		tmpField = array("id", "title")
		with tdg
			.table  "NL_News"
			.fields tmpField
			.where "id=28"
			.query
		end with
		
		'set re = tdg.query
		'echo typename(re)

		do while not tdg.rs.eof
			temp = temp & "<li>" & tdg.rs("title") & " === id: "&tdg.rs("id")&"</li>"
			tdg.rs.movenext
		loop
		echo "<ul>"&temp&"</ul>"

		
'		dim con, file : set file = new Kiss_File
'		con = file.readFile("readme.txt")
'		echo con 
'		set file =  nothing
	end function

	public function actionUpdate()
		with tdg
			.table "NL_News"
			.fields "title"
			.fieldsVal "Kiss Asp framework 0.2.2."
			.where "id=28"
			.update
		end with
		echo "Update"
	end function

	public function actionInsert()
		with tdg
			.table "NL_News"
			.fields "title"
			.fieldsVal "Kiss Asp framework 0.2.1."
			.insert
		end with
		echo "Insert"
	end function

	public function actionDel()
		with tdg
			.table "NL_News"
			.where "id=36"
			.delete
		end with
		echo "Delete"
	end function

	public function actionView()
		'start the debugging
		'dim output : output="Just output for debugging" 
		'debug.Print "Debugging", output

		import("App.Model.News")
		'Import("App.Model.News")

		'set the main template file
		View.setTemplateFile "layout.html"

		'Add some custom tags to the template
		View.assign "title", "ASP Template Example Script 1"

		'We add some more text to this tag
		View.Append "title", " - Main Page"

		'Load an external file into a tag.
		View.assignFile "content", "content.html"

		'Load another external file that will be used to draw some blocks
		View.assignFile "menu", "menu.html"

		'List news
		News.list()

		'setup the block
		View.UpdateBlock "menu_block"

		'set and parse the block
		View.assign "menu_text", "HOME"
		View.assignBlock "menu_block"

		'Do it multiple times
		View.assign "menu_text", "NEWS"
		View.assignBlock "menu_block"

		View.assign "menu_text", "CREDITS"
		View.assignBlock "menu_block"

		'Store the blocks in their directory service (order is important)

		View.UpdateBlock "c_block"
		View.UpdateBlock "b_block"
		View.UpdateBlock "a_block"

		View.assign "inner", "666666"
		View.assign "outer", "Outer Block (A)"

		View.assign "whatever", "Block C 1"
		View.assignBlock "c_block"
		View.assign "whatever", "Block C 2"
		View.assignBlock "c_block"
		View.assign "whatever", "Block C 3"
		View.assignBlock "c_block"

		View.assign "into_b", "Block B 1"
		View.assignBlock "b_block"

		View.assign "whatever", "Block C 1(b)"
		View.assignBlock "c_block"
		View.assign "whatever", "Block C 2(b)"
		View.assignBlock "c_block"

		View.assign "inner", "999999"

		View.assign "into_b", "Block B 2"
		View.assignBlock "b_block"

		View.assignBlock "a_block"


		'Generate the page
		View.display

		'end the debuggings

	end function

end class



%>