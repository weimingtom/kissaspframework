<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Template Engine class
'	File Name	: View.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** 'set the main template file
'*** View.setTemplateFile "layout.html"
'*** 
'*** 'Add some custom tags to the template
'*** View.assign "title", "ASP Template Example Script 1"
'*** 
'*** 'We add some more text to this tag
'*** View.Append "title", " - Main Page"
'*** 
'*** 'Load an external file into a tag.
'*** View.assignFile "content", "content.html"
'*** 
'*** 'Load another external file that will be used to draw some blocks
'*** View.assignFile "menu", "menu.html"
'*** 
'*** 'setup the block
'*** View.UpdateBlock "menu_block"
'*** 
'*** 'set and parse the block
'*** View.assign "menu_text", "HOME"
'*** View.assignBlock "menu_block"

'*************************************************************
'	Initialize the class
'*************************************************************
dim View
set View = new Kiss_View

class Kiss_View

	public className	'Class name
	' Contains the error objects
	private p_error
	
	' Print error messages?
	private p_print_errors
	
	' What to do with unknown tags (keep, remove or comment)?
	private p_unknowns
	
	' Opening delimiter (usually "{{")
	private p_var_tag_o
	
	' Closing delimiter (usually "}}")
	private p_var_tag_c

	'private p_start_block_delimiter_o
	'private p_start_block_delimiter_c
	'private p_end_block_delimiter_o
	'private p_end_block_delimiter_c
	
	'private p_int_block_delimiter
	
	private p_template
	private p_variables_list
	private p_blocks_list
	private p_blocks_name_list
	private	p_regexp
	private p_parsed_blocks_list

	private p_boolsubMatchesAllowed
	
	' Directory containing HTML templates
	private p_templates_dir
	
	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Constructor
	'************************************************************* 
	private sub class_Initialize
		classname = "Kiss_View"
		p_print_errors = FALSE
		p_unknowns = "keep"
		' Remember that opening and closing tags are being used in regular expressions
		' and must be explicitly escaped
		p_var_tag_o = "\{\{"
		p_var_tag_c = "\}\}"
		' Block delimiters are actually disabled and no longer available. Maybe they'll be again
		' in the future.
		'p_start_block_delimiter_o = "<!-- BEGIN "
		'p_start_block_delimiter_c = " -->"
		'p_end_block_delimiter_o = "<!-- END "
		'p_end_block_delimiter_c = " -->"
		'p_int_block_delimiter = "__"
		p_templates_dir = "App/View/"
		set p_variables_list = createobject("Scripting.Dictionary")
		set p_blocks_list = createobject("Scripting.Dictionary")
		set p_blocks_name_list = createobject("Scripting.Dictionary")
		set p_parsed_blocks_list = createobject("Scripting.Dictionary")
		p_template = ""
		p_boolsubMatchesAllowed = not (ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion < "5.5")
		set p_regexp = New RegExp   
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub
	
	'************************************************************* 
	' Name: setTemplatesDir
	' Param: dir as Variant Directory
	' Purpose: sets the directory containing html templates
	'************************************************************* 
	public sub setTemplatesDir(dir)
		p_templates_dir = dir
	end sub

	'************************************************************* 
	' Name: setTemplate
	' Param: template as Variant String containing the template
	' Purpose: sets a template passed through a string argument
	'************************************************************* 
	public sub setTemplate(template)
		p_template = template
	end sub
	
	'************************************************************* 
	' Name: getTemplate
	' Purpose: returns template as a string
	'************************************************************* 
	public function getTemplate
		getTemplate = p_template
	end function
	
	'************************************************************* 
	' Name: setUnknowns
	' Param: action as String containing the action to perform with unrecognized
	'    tags in the template
	' Purpose: sets a variable passed through a string argument
	' Remarks: The action can be one of the following:
	'  - 'keep': leave the tags untouched
	'  - 'remove': remove the tags from the output
	'  - 'comment': mark the tags as HTML comment
	'************************************************************* 
	public sub setUnknowns(action)
		if (action <> "keep") and (action <> "remove") and (action <> "comment") then
			p_unknowns = "keep"
		else
			p_unknowns = action
		end if
	end sub

	'************************************************************* 
	' Name: setTemplateFile
	' Param: inFileName as Variant Name of the file to read the template from
	' Purpose: sets a template given the filename to load the template from
	'************************************************************* 
	public sub setTemplateFile(inFileName)
		dim oFile : set oFile = new Kiss_File
		if len(inFileName) > 0 then
			' read the template file
			p_template = oFile.readFile(p_templates_dir & inFileName)
			set oFile = nothing
		else
			die("<p>Error: setTemplateFile missing filename.</p>")
		end if
		
	end sub
	
	'************************************************************* 
	' Name: assign
	' Param: s as Variant - Variable name
	'    v as Variant - Value
	' Purpose: sets a variable given it's name and value
	'************************************************************* 
	public function assign(s, v)
		if p_variables_list.Exists(s) then
			p_variables_list.Remove s
			p_variables_list.Add s, v
		else
			p_variables_list.Add s, v
		end if
	end function

	'************************************************************* 
	' Name: Append
	' Param: s as Variant - Variable name
	'    v as Variant - Value
	' Purpose: sets a variable appending the new value to the existing one
	'************************************************************* 
	public sub Append(s, v)
		dim tmp
		if p_variables_list.Exists(s) then
			tmp = p_variables_list.Item(s) & v
			p_variables_list.Remove s
			p_variables_list.Add s, tmp
		else
			p_variables_list.Add s, v
		end if
	end sub
		
	'************************************************************* 
	' Name: assignFile
	' Param: s as Variant Variable name
	'    inFileName as Variant Name of the file to read the value from
	' Purpose: Load a file into a variable's value
	'************************************************************* 
	public sub assignFile(s, inFileName)
		dim tmp, oFile : set oFile = new Kiss_File
		if len(inFileName) > 0 then
			' read the template file
			tmp = oFile.readFile(p_templates_dir & inFileName)
			ReplaceBlock s, tmp
			set oFile = nothing
		else
			die("<p>Error: Filename was never passed.</p>")
		end if
	end sub

	'************************************************************* 
	' Name: ReplaceBlock
	' Param: s as Variant Variable name
	'    inFile as Variant Content of the file to place in the template
	' Purpose: function used by assignFile to load a file and replace it
	'          into the template in place of a variable
	'************************************************************* 
	public sub ReplaceBlock(s, inFile)
		p_regexp.IgnoreCase = true
		p_regexp.Global = true
		assign s, inFile
		p_regexp.Pattern = p_var_tag_o & s & p_var_tag_c
		p_template = p_regexp.Replace(p_template, inFile)   
	end sub

	public property get getOutput
		dim Matches, match, MatchName
		
		'Replace the variables in the template
		p_regexp.IgnoreCase = true
		p_regexp.Global = true
		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			if p_boolsubMatchesAllowed then
				MatchName = match.subMatches(1)
			else
				MatchName = mid(match.Value,3,Len(match.Value) - 4)
			end if
			if p_variables_list.Exists(MatchName) then
				p_regexp.Pattern = match.Value
				p_template = p_regexp.Replace(p_template, p_variables_list.Item(MatchName))
			end if
			'response.write.write "getOutput (match): " & match.Value & "<br>"
		next
        
        'this removes any block placeholder tags that are left over
		p_regexp.Pattern = "__[_a-z0-9]*__"
		set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			'response.write.write "[[" & match.Value & "]]<br>"
			p_regexp.Pattern = match.Value
			p_template = p_regexp.Replace(p_template, "")
		next

		' deal with unknown tags
		select case p_unknowns
			case "keep"
				'do nothing, leave it
			case "remove"
				'all known matches have been replaced, remove every other match now
				p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
				set Matches = p_regexp.Execute(p_template)   
				for each match in Matches
					'response.write.Write "Found match: " & match & "<br>"
					p_regexp.Pattern = match.Value
					p_template = p_regexp.Replace(p_template, "")
				next
			case "comment"
				'all known matches have been replaced, HTML comment every other match
				p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
				set Matches = p_regexp.Execute(p_template)   
				for each match in Matches
					p_regexp.Pattern = match.Value
					if p_boolsubMatchesAllowed then
						p_template = p_regexp.Replace(p_template, "<!-- Template variable " & match.submatches(1) & " undefined -->")
					else
						p_template = p_regexp.Replace(p_template, "<!-- Template variable " & mid(match.Value,3,len(match) - 4) & " undefined -->")
					end if
				next
		end select
				
		getOutput = p_template
	end property

	public sub display
		dim parsed
		
		parsed = getOutput
		echo parsed
	end sub
			
	' TODO: if the block foud contains other blocks, it should recursively update all of them without the needing
	' of doing this by hand.
	public sub UpdateBlock(inBlockName)
		dim Matches, match, asubMatch
		dim braceStart, braceend
		
		p_regexp.IgnoreCase = true
		p_regexp.Global = true

		'p_regexp.Pattern = "<!--\s+BEGIN\s+(" & inBlockName & ")\s+-->([\s\S.]*)<!--\s+END\s+\1\s+-->"
		p_regexp.Pattern = "<!--\s*BEGIN\s+(" & inBlockName & ")\s*-->([\s\S.]*)<!--\s*END\s+\1\s*-->"
		set Matches = p_regexp.Execute(p_template)
		set match = Matches
		for each match in Matches
			if p_boolsubMatchesAllowed then
				asubMatch = match.subMatches(1)
			else
				braceStart = instr(match,"-->") + 3
				braceend = instrrev(match,"<!--")
				asubMatch = mid(match,braceStart,braceend - braceStart)
			end if
			'The following check let the user use the same template multiple times
			if p_blocks_list.Exists(inBlockName) then
				p_blocks_list.Remove(inBlockName)
				p_blocks_name_list.Remove(inBlockName)
			end if
			p_blocks_list.Add inBlockName, asubMatch
			p_blocks_name_list.Add inBlockName, inBlockName
			'printInternalTemplate "UpdateBlock: before replace"
			p_template = p_regexp.Replace(p_template, "__" & inBlockName & "__")
			'printInternalTemplate "UpdateBlock: after replace"
			'response.write.write "[[" & server.HTMLEncode(asubMatch) & "]]<br>"
		next
	end sub

	public sub assignBlock(inBlockName)
		dim Matches, match, tmp, w, asubMatch
		
		'debugPrint "Parsing: " + inBlockName
		
		w = getBlock(inBlockName)
		
        'See if there are any sub-blocks in this block
		p_regexp.IgnoreCase = true
		p_regexp.Global = true
		p_regexp.Pattern = "(__)([_a-z0-9]+)__"
		set Matches = p_regexp.Execute(w)
		set match = Matches
		for each match in Matches
			if p_boolsubMatchesAllowed then
				asubMatch = match.subMatches(1)
			else
				asubMatch = mid(match.Value,3,len(match) - 4)
			end if

            'if the sub-block has already been parsed, then replace the block
            'identifier with the already parsed text
            if p_parsed_blocks_list.Exists(asubMatch) then
			    p_regexp.Pattern = "__" & asubMatch & "__"
				w = p_regexp.Replace(w, p_parsed_blocks_list.Item(asubMatch))
				p_parsed_blocks_list.Remove(asubMatch)
			else
			    'if we are here, that means we are parsing a parent block
			    'that has a child block that has not yet been parsed.  We assume
			    'that means that the block should remain empty and we therefore
			    'need to remove the sub-block identifier from the parsed output
			    'of this block.  Otherwise, the sub-block identifiers will be
			    'replaced in the template on future requests to parse the
			    'sub-block.
			    'this removes any block placeholder tags that are left over
    			p_regexp.Pattern = "__" & asubMatch & "__"
    			w = p_regexp.Replace(w, "")
            end if
		next

		'if this block has already been parsed, append the output to the current
        'entry in the parsed_blocks_list.  Otherwise, create the entry.
		if p_parsed_blocks_list.Exists(inBlockName) then
			tmp = p_parsed_blocks_list.Item(inBlockName) & w
			p_parsed_blocks_list.Remove inBlockName
			p_parsed_blocks_list.Add inBlockName, tmp
		else
			p_parsed_blocks_list.Add inBlockName, w
		end if
        
        'Finally, replace the block identifier in the template with the text of
        'this block and then append the block identifier in case this block
        'is parsed again.
        'if the block is not found in the template, we assume that is because
        'the block is embedded in a parent block that will be parsed in the future.
        'When the parent block is parsed, the content of this block will be included
		p_regexp.IgnoreCase = true
		p_regexp.Global = true
		p_regexp.Pattern = "__" & inBlockName & "__"
		set Matches = p_regexp.Execute(p_template)
		set match = Matches
		for each match in Matches
			w = getParsedBlock(inBlockName)
			p_regexp.Pattern = "__" & inBlockName & "__"
			p_template = p_regexp.Replace(p_template, w & "__" & inBlockName & "__")
		next
        
        'printInternalVariables
        'printInternalTemplate "assignBlock: end of function ("+inBlockName+")"
	end sub
    
    'gets the text inside a block, parses and replaces variables, and returns
    'the block of text
	private property get getBlock(inToken)
		dim tmp, s
		
		'This routine checks the Dictionary for the text passed to it.
		'if it finds a key in the Dictionary it Display the value to the user.
		'if not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_blocks_list.Item(inToken)
			s = assignBlockVars(tmp)
			getBlock = s
			'response.write.write "s: " & s
		else
			getBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	private property get getParsedBlock(inToken)
		dim tmp, s
		
		'This routine checks the Dictionary for the text passed to it.
		'if it finds a key in the Dictionary it Display the value to the user.
		'if not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_parsed_blocks_list.Item(inToken)
			s = assignBlockVars(tmp)
			getParsedBlock = s
			'response.write.write "s: " & s
			p_parsed_blocks_list.Remove(inToken)
		else
			getParsedBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	public property get assignBlockVars(inText)
		dim Matches, match, asubMatch
		
		p_regexp.IgnoreCase = true
		p_regexp.Global = true

		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		set Matches = p_regexp.Execute(inText)

		for each match in Matches
			if p_boolsubMatchesAllowed then
				asubMatch = match.subMatches(1)
			else
				asubMatch = mid(match.Value,3,Len(match.Value) - 4)
			end if
			if p_variables_list.Exists(asubMatch) then
				p_regexp.Pattern = match.Value
				if IsNull(p_variables_list.Item(asubMatch)) then
					inText = p_regexp.Replace(inText, "")
				else
					inText = p_regexp.Replace(inText, p_variables_list.Item(asubMatch))
				end if
			end if
			'response.write.write "match.Value: " & match.Value & "<br>"
			'response.write.write "in text: " & inText & "<br>"
		next
		assignBlockVars = inText
	end property
    
    public sub printInternalVariables()
        'response.write "<b>p_variables_list:</b>"
    	'printr p_variables_list
    	'response.write "<b>p_blocks_list:</b>"
    	'printr p_blocks_list
    	'response.write "<b>p_blocks_name_list:</b>"
    	'printr p_blocks_name_list
    	echo "<b>p_parsed_blocks_list:</b>"
    	printr p_parsed_blocks_list
    end sub
    
    public sub printInternalTemplate( sPrefix )
        debugPrint sPrefix & "<br /><pre><blockquote>" & Server.HTMLEncode(p_template) & "</blockquote></pre><hr />"
    end sub
    
    private sub debugPrint( sText )
        echo sText + "<br />"
    end sub
end class
%>
