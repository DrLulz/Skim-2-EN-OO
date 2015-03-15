--------------------------------------------------------------
-- DrLulz Jan 2015   v-Alfred/OmniOutliner
-- 
-- CREDITS:
-- macscripter.net
-- macosxautomation.com/applescript/
-- OmniOutliner: Attach Image: forums.omnigroup.com/showpost.php?p=129355&postcount=5
-- Split Spaceless String into Words stackoverflow.com/questions/8870261/how-to-split-text-without-spaces-into-list-of-words
-------- Word Frequency Data: wordfrequency.info
-------- Word Frequency Data: invokeit.wordpress.com/frequency-word-lists/
--------------------------------------------------------------

--------------------------------------------------------------
-- Choose Export Options
--
--   -t = Titlecase Levels 1 - 3: this is cool --> This Is Cool
--   -c = Text Correction: Posi+on, Contraindica%ons, a`achments, sta-cally --> Position, Contraindications, attachments, statically
--   -s = Find Spaces (in Spaceless String): where,canifindthe beer --> where can i find the beer
--   -i = Extract Images: Finds boundries of Skims 'Box Note' and Creates PNG on Export
--        Define Resolution: Example: -t -i -300 (Default is 700)
--   -z = Top 20 Words in PDF
--  -oo = Use Own Template
-- -one = One Color: All Text of OO Export is Single Color (Only OmniOutliner)
--------------------------------------------------------------

on run argv
	
	set query to argv as text
	set wf_hfs to (POSIX file (POSIX path of ((path to me as text) & "::"))) as string

	set {extract_images, image_res, word_list} to {false, {}, {}}
	set export_options to my get_options(query)
	
	--------------------------------------------------------------
	-- Find Skim Notes, PDF Attributes, and Apply Options
	--------------------------------------------------------------
	tell application "Skim"
		
		if not (exists front document) then 
			activate
			display dialog ("No Document is Open") with icon path to resource "Skim.icns" in bundle (path to application "Skim")
			return
		end if
		
		set all_notes to every note of front document
		set doc_name to (name of front document)
		set pdf_name to text 1 thru ((offset of ".pdf" in doc_name) - 1) of doc_name
		set doc_text to text of front document
		set tot_wrds to count of words of text of front document
		set posix_path to (path of front document)
		set doc_alias to ((POSIX file (posix_path)) as Unicode text) as alias
		set file_url to my path2url(posix_path)
		set skimmer_url to "skimmer://" & file_url & "?page="
		set file_path to "file://localhost" & file_url
		--set {fc1, fc2, fc3, fc4, fc5, fc6} to favorite colors
		set fc1 to {23731, 23731, 23731, 9175}
		set fc2 to {7, 115, 65418, 5243}
		set fc3 to {64907, 32785, 2154, 14418}
		set fc4 to {64634, 63620, 29036, 39976}
		set fc5 to {21176, 65514, 15330, 25559}
		set fc6 to {64587, 609, 65481, 21627}
		set {highlight_note, box_note} to {highlight note, box note}
		
		
		--IMAGES
		if export_options contains "i" then
			repeat with n in all_notes
				if type of n is box note then
					set extract_images to true					
					set image_res to my get_res(query) --as text
					if image_res is "" then set image_res to "700"
					exit repeat
				end if
			end repeat
			if extract_images is false then
				activate
				display dialog ("There aren't any images selected") with title "Select Images First" with icon path to resource "Skim.icns" in bundle (path to application "Skim")
			end if
		end if
	
	
		-- CHOOSE WORD LIST FOR FIND SPACES
		--if export_options contains "s" then
		--	tell application "Alfred 2" to search "choose list"
		--	set word_list to my get_wrd_list()
		--end if
		
		
		-- TEXT CORRECTION & FIND SPACES
		if ("c" is in export_options) or ("s" is in export_options) then
			repeat with n in all_notes
				set note_text to text of n
				set text of n to my clean_txt(note_text, export_options)--, word_list)
			end repeat
		end if
		
		
		-- TOP 20 WORDS IN PDF
		if ("z" is in export_options) then
			my common_wrds(doc_name, pdf_name, doc_text, tot_wrds)
		end if
		
	end tell
	
	
	--------------------------------------------------------------
	-- OmniOutliner
	--------------------------------------------------------------
	tell application "System Events"
		if not (exists process "OmniOutliner") then
			do shell script "open -a \"OmniOutliner\""
		end if
	end tell
	
	tell application "Finder" to set dtb to bounds of window of desktop

	
	if ("oo" is in export_options) then
		
		--------------------------------------------------------------
		-- Use Users Own Custom Template
		--------------------------------------------------------------

		-- LOCATE USERS CUSTOM TEMPLATES & PROMPT FOR CHOICE
		tell application "Finder"
			set temp_folder to folder ((path to library folder from user domain as text) & ¬
				"Containers:com.omnigroup.OmniOutliner4:Data:Library:Application Support:The Omni Group:OmniOutliner:Templates:")
			set temp_files to name of every file in temp_folder whose name extension is "oo3template"
			
			set temp_names to {}
			repeat with n in temp_files
				set short_names to text 1 thru ((offset of "." in n) - 1) of n
				set end of temp_names to short_names
			end repeat
		end tell
		set template_name to (choose from list temp_names with prompt "Choose Template:" cancel button name "Cancel") as text
		
		
		-- OPEN USER TEMPLATE
		tell application id "OOut"
			set template_path to ((path to application support from user domain as string) & ¬
				"The Omni Group:OmniOutliner:Templates:" & template_name & ".oo3template:")
			open template_path
			make new row with properties {topic:pdf_name, note:file_path} at end of rows of front document
			tell application "System Events" to set frontmost of process "OmniOutliner" to true
			
			set doc to front document
			
			tell doc
				if "one" is in export_options then
					
					-- MAKE NEW NAMED STYLE TO MARK LEVELS 1-3, BUT USE TEMPLATE FONT STYLE
					set style_single to make new named style with properties {name:"Heading Mark"}
					set {style_one, style_two, style_three} to {style_single, style_single, style_single}
					set head_mark to name of style_single
				else
					
					-- USE TEMPLATE NAMED STYLES 1-3 FOR LEVELS 1-3
					set {style_one, style_two, style_three} to named styles of doc
					set head_mark to {name of style_one, name of style_two, name of style_three}
				end if
				
				-- MAKE HIGHLIGHT STYLES FOR SKIM FAVORITE COLORS 5 & 6
				if ("one" is in export_options) and ("only" is in export_options) then
					set style_five to make new named style with properties {name:"Green Highlight"}
					set style_six to make new named style with properties {name:"Red Highlight"}
				else
					set style_five to make new named style with properties {name:"Green Highlight"}
					set value of attribute "text-background-color" of style_five to {59367, 61166, 56540, 65535}
				
					set style_six to make new named style with properties {name:"Red Highlight"}
					set value of attribute "text-background-color" of style_six to {63479, 59110, 58339, 65535}
				end if
				
			end tell
		end tell
		
	else
		
		tell application id "OOut"
			
			--------------------------------------------------------------
			-- Make New Template From Script
			--------------------------------------------------------------
			
			-- MAKE DOCUMENT & DEFINE DEFAULT FONTS
			make new document with properties {name:pdf_name}
			set doc to document pdf_name
			make new row with properties {topic:pdf_name, note:file_path} at end of rows of doc
			set status visible of doc to false -- Hide OO Check-Box Column
			
			-- MAKE COLUMN FOR IMAGE PLACEMENT
			if extract_images is true then
				make new column at doc with properties {title:"Images", width:320}
			end if
			
			set index of window 1 where name contains pdf_name to 1
			activate window 1
			set bounds of window 1 to dtb
			
			-- DEFINE DEFAULT STYLE OF WHOLE DOCUMENT
			tell style of doc
				set value of attribute "font-size" to 16
				set value of attribute "font-family" to "Avenir Next"
				--set value of attribute "font-family" to "M+ 1p"
				set value of attribute "font-fill" to {21047, 21047, 21047, 65535}
				set value of attribute "item-to-note-space(com.omnigroup.OmniOutliner)" to "4.0"
				set value of attribute "item-child-indentation(com.omnigroup.OmniOutliner)" to "30.0"
				set value of attribute "shadow-color" to {59908, 59908, 59908, 65535}
				--set value of attribute "shadow-offset" to "-1.0" --Why doesn't this work?
				set value of attribute "shadow-radius" to "1.2"
			end tell
			
			tell column title style of doc
				set value of attribute "font-fill" to {57729, 57729, 57729, 65535}
				set value of attribute "underline-style" to "none"
				set value of attribute "font-weight" to 1.0
			end tell
			
			
			-- DEFINE HEADING & HIGHLIGHT STYLES
			-- HEADING STYLES USED TO DESIGNATE A PREVIOUS ROW AS HEADING ROW
			---- HEADING ROWS WILL AUGMENT SUBSEQUENT ROWS LEVEL OF INDENTATION
			tell doc
				
				if "one" is in export_options then
					
					-- MAKE NEW NAMED STYLE, IDENTICAL TO DEFAULT STYLE, TO MARK LEVELS 1-3
					set style_single to make new named style with properties {name:"Heading Mark"}
					set value of attribute "font-fill" of style_single to {21047, 21047, 21047, 65535}
					set {style_one, style_two, style_three} to {style_single, style_single, style_single}
					set head_mark to name of style_single
				else
					
					-- MAKE COLORED NAMED STYLES 1-3 FOR LEVELS 1-3
					set style_one to make new named style with properties {name:"Heading 1"}
					set value of attribute "font-fill" of style_one to {16587, 28329, 55163, 65535}
					
					set style_two to make new named style with properties {name:"Heading 2"}
					set value of attribute "font-fill" of style_two to {39981, 38846, 25155, 65535}
					
					set style_three to make new named style with properties {name:"Heading 3"}
					set value of attribute "font-fill" of style_three to {23514, 43610, 18771, 65535}
					
					set head_mark to {name of style_one, name of style_two, name of style_three}
				end if
				
				-- MAKE HIGHLIGHT STYLES FOR SKIM FAVORITE COLORS 5 & 6
				if ("one" is in export_options) and ("only" is in export_options) then
					set style_five to make new named style with properties {name:"Green Highlight"}
					set style_six to make new named style with properties {name:"Red Highlight"}
				else
					set style_five to make new named style with properties {name:"Green Highlight"}
					set value of attribute "text-background-color" of style_five to {59367, 61166, 56540, 65535}
				
					set style_six to make new named style with properties {name:"Red Highlight"}
					set value of attribute "text-background-color" of style_six to {63479, 59110, 58339, 65535}
				end if
				
			end tell
		end tell
		
	end if
	
	
	--------------------------------------------------------------
	-- Make Rows from Skim Annotations
	--------------------------------------------------------------
	tell application "Skim"
		repeat with n from 1 to count of all_notes
			
			set _note to item n of all_notes
			set type_note to type of _note
			set page_index to index of page of _note
			set note_url to skimmer_url & page_index
			set note_text to text of _note
			
			set _page to page of _note
			
			set _bounds to bounds of _note
			set _data to grab _page for _bounds
			
			set rgba to color of _note
			
			tell application id "OOut"
				set row_last to last row of doc
				set previous_rows_style to (name of named styles of style of row_last as string)
				
				if type_note = highlight_note then
					
					if rgba is fc1 then
						
						set head_one to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of rows of doc
						add style_one to named styles of style of head_one
						
					else if rgba is fc2 then
						
						try
							set head_two to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of last child of doc
						on error
							set head_two to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of rows of doc
						end try
						add style_two to named styles of style of head_two
						
					else if rgba is fc3 then
						
						set head_three to {}
						if previous_rows_style is in head_mark then
							set head_three to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of last row's children of doc
						else if level of parent of last row of doc ≤ 2 then
							set head_three to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of rows of parent of row_last
						else
							set head_three to make new row with properties {topic:my titlecap(note_text, export_options), note:note_url} at end of parent of last row's parent of doc
						end if
						add style_three to named styles of style of head_three
						
					else if rgba is fc4 then
						
						if previous_rows_style is in head_mark then
							make new row with properties {topic:note_text} at end of last row's children of doc
						else if level of row_last > 1 and (has subtopics of preceding sibling of row_last contains {} or has subtopics of preceding sibling of row_last contains {false}) then
							make new row with properties {topic:note_text} at end of parent of row_last
						else
							make new row with properties {topic:note_text} at end of last child of doc
						end if
						
					else if rgba is fc5 then
						
						set head_five to {}
						if previous_rows_style is in head_mark then
							set head_five to make new row with properties {topic:note_text} at end of last row's children of doc
						else if level of row_last > 1 and (has subtopics of preceding sibling of row_last contains {} or has subtopics of preceding sibling of row_last contains {false}) then
							set head_five to make new row with properties {topic:note_text} at end of parent of row_last
						else
							set head_five to make new row with properties {topic:note_text} at end of last child of doc
						end if
						add style_five to named styles of style of head_five
						
					else if rgba is fc6 then
						
						set head_six to {}
						if previous_rows_style is in head_mark then
							set head_six to make new row with properties {topic:note_text} at end of last row's children of doc
						else if level of row_last > 1 and (has subtopics of preceding sibling of row_last contains {} or has subtopics of preceding sibling of row_last contains {false}) then
							set head_six to make new row with properties {topic:note_text} at end of parent of row_last
						else
							set head_six to make new row with properties {topic:note_text} at end of last child of doc
						end if
						add style_six to named styles of style of head_six
						
					end if
					
				else if (type_note = box_note) and (extract_images = true) then
					
					my write_temp(_data, page_index, image_res)
					
				end if
				
			end tell
		end repeat
	end tell
	tell application id "OOut" to expandAll rows of front document
	try
		do shell script "rm -r " & quoted form of (POSIX path of (path to temporary items from user domain) & "pic_temp_folder")
	end try
end run

--------------------------------------------------------------
-- 
--                          Handlers
--
--------------------------------------------------------------

--------------------------------------------------------------
-- FIND OPTION SWITCHES IN ALFRED QUERY
--------------------------------------------------------------
on get_options(argv)
	set _cmd to "echo " & quoted form of argv & " | awk -F: '{gsub(\"-\",\"\",$1);print $1}'"
	do shell script _cmd
end get_options


--------------------------------------------------------------
-- ENCODE PDF PATH FOR SKIMMER
--------------------------------------------------------------
on path2url(the_path)
	return do shell script "python -c \"import urllib, sys; print (urllib.quote(sys.argv[1]))\" " & quoted form of the_path
end path2url


--------------------------------------------------------------
-- CONVERT LEVELS 1-3 TO TITLECASE
--------------------------------------------------------------
on titlecap(txt, option)
	if option does not contain "t" then
		return txt
	else
		return do shell script "python -c \"import sys; print unicode(sys.argv[1], 'utf8').title().encode('utf8')\" " & quoted form of txt
	end if
end titlecap


--------------------------------------------------------------
-- CLEAN UP WEIRD PDFs BY PASSING TO APPROPRIATE HANDLER
--------------------------------------------------------------
on clean_txt(txt, options)--, word_list)
	
	set corrected_txt to {}
	
	if (options contains "t") and (options contains "s") then
		set char_replace to my correct_char(txt)
		set corrected_txt to my find_spaces(char_replace)--, word_list)
	else if options contains "t" then
		set corrected_txt to my correct_char(txt)
	else if options contains "s" then
		set corrected_txt to my find_spaces(txt)--, word_list)
	end if
	
	return corrected_txt
end clean_txt


--------------------------------------------------------------
-- TEXT CORRECTION: FIND & REPLACE NON-ALPHA CHARACTERS
--------------------------------------------------------------
on correct_char(txt)
	set get_TI to "([a-zA-Z])([0-9+%-])([a-zA-Z])"
	set replace_TI to "\\1ti\\3"
	set get_space to "([a-zA-Z])([$)])([a-zA-Z])"
	set replace_space to "\\1 \\3"
	set get_TT to "([a-zA-Z])([.`])([a-zA-Z])"
	set replace_TT to "\\1tt\\3"

	set _cmd to "echo " & quoted form of txt & " \\
| sed -E 's!" & get_TI & "!" & replace_TI & "!g; s!" & get_space & "!" & replace_space & "!g; s!" & get_TT & "!" & replace_TT & "!g'"
	
	set corrected_txt to do shell script _cmd
	return corrected_txt
end correct_char


--------------------------------------------------------------
-- FIND SPACES IN SPACELESS STRING
--------------------------------------------------------------
on find_spaces(txt)--, word_list)
	
	set script_posix to quoted form of (POSIX path of ((path to me as text) & "::"))
	set _txt to quoted form of txt
	--set _list to quoted form of word_list
	set _list to quoted form of ((do shell script "pwd") & "/frequency_lists/wordlist_hz_333310.txt")
	
	set corrected_txt to do shell script "python " & script_posix & "/split_words.py" & " " & _txt & " " & _list
	
	return corrected_txt
end find_spaces
(*
on get_wrd_list()
	set proxy_path to quoted form of ((do shell script "pwd") & "/proxy.scpt")
	set time_start to do shell script "date +%s"
	set time_mod to do shell script "stat -f %m " & proxy_path

	repeat until time_mod > time_start
		set time_mod to do shell script "stat -f %m " & proxy_path
		set time_exit to do shell script "date +%s"
		if time_exit - time_start > 10 then exit repeat
	end repeat
	
	set proxy_value to load script ((POSIX file ((do shell script "pwd") & "/proxy.scpt")))
	return proxy_value's word_list	
end get_wrd_list
*)
--------------------------------------------------------------
-- GET PIC RESOLUTION FROM ALFRED QUERY IF USER DEFINED
--------------------------------------------------------------
on get_res(argv)
	set _cmd to "echo " & quoted form of argv & " | awk -F: '{gsub(\"-\",\"\",$1);print $1}' | awk '{gsub(\"[^[:digit:]]+\",\" \");print $1}'"
	do shell script _cmd
end get_res


--------------------------------------------------------------
-- CREATE PNG FROM SKIM BOX NOTES & SAVE TO TEMP FOLDER
--------------------------------------------------------------
on write_temp(pic_data, pic_page, pic_res)
	
	-- ~/Library/Caches/TemporaryItems
	set temp_folder to quoted form of (POSIX path of (path to temporary items from user domain) & "pic_temp_folder")
	
	try
		do shell script "mkdir " & temp_folder
	end try
	
	set target_folder to (path to temporary items from user domain as string) & "pic_temp_folder:box_page_"
	set target to target_folder & pic_page & ".pdf"
	set conversion_hfs to target_folder & pic_page & ".png"
	set conversion_posix to POSIX path of (target_folder & pic_page & ".png")
	
	set file_reference to (open for access target with write permission)
	write pic_data to file_reference starting at eof
	close access file_reference
	
	do shell script "sips -s format png " & (POSIX path of target) & " --out " & conversion_posix
	
	do shell script "sips -Z " & pic_res & " " & POSIX path of conversion_hfs
	
	set oo_pic to my oo_import(conversion_posix, temp_folder)
	
end write_temp


--------------------------------------------------------------
-- FIND THE 20 MOST FREQUENT WORDS OF THE PDF
--------------------------------------------------------------
on common_wrds(doc_name, pdf_name, doc_text, tot_wrds)
	
	set hz_file to (((path to desktop folder) as string) & pdf_name)
	
	set the_report to "PDF: " & doc_name & return & ¬
		"Total Words: " & tot_wrds & return & return & ¬
		"The 20 Most Frequent Words" & return & return & return
	set the_report to the_report & my wrd_hz(doc_text)
	
	try
		set the hz_file to hz_file as string
		set the open_hz_file to open for access file hz_file with write permission
		set eof of open_hz_file to 0
		write the_report to open_hz_file starting at eof
		close access the open_hz_file
	on error
		try
			close access file hz_file
		end try
	end try
	
	do shell script "open " & quoted form of ((POSIX path of (path to desktop)) & pdf_name) as string
	
end common_wrds

on wrd_hz(txt)
	set script_posix to quoted form of (POSIX path of ((path to me as text) & "::"))
	set _txt to quoted form of txt
	
	set common_txt to do shell script "python " & script_posix & "/common_words.py" & " " & _txt
end wrd_hz

--------------------------------------------------------------
-- OMNIOUTLINER HANDLERS
--------------------------------------------------------------
on oo_import(posix_path, temp_folder)
	tell front document of application "OmniOutliner"

		set target_cell to last cell of last row
		tell text of target_cell
			make new file attachment with properties {file name:POSIX file posix_path, embedded:true} at end of characters
		end tell

	end tell	
end oo_import
