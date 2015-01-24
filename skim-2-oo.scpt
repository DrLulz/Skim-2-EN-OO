--------------------------------------------------------------
-- DrLulz Jan 2015
-- 
-- CREDITS:
-- macscripter.net
-- macosxautomation.com/applescript/
-- OmniOutliner: Attach Image: forums.omnigroup.com/showpost.php?p=129355&postcount=5
-- Split Spaceless String into Words stackoverflow.com/questions/8870261/how-to-split-text-without-spaces-into-list-of-words
-------- Word Frequency Data: wordfrequency.info
-------- Word Frequency Data: invokeit.wordpress.com/frequency-word-lists/
-- Frequency of ALL WORDS IN PDF: macscripter.net/viewtopic.php?pid=136205
--------------------------------------------------------------

--------------------------------------------------------------
-- Choose Export Options
--
-- One Color: All Text of OO Export is Single Color (Only OmniOutliner)
-- Titlecase Levels 1 - 3: this is cool --> This Is Cool
-- Text Correction: Posi+on, Contraindica%ons, a`achments, sta-cally --> Position, Contraindications, attachments, statically
-- Find Spaces (in Spaceless String): where,canifindthe beer --> where can i find the beer
-- Extract Images: Finds boundries of Skims 'Box Note' and Creates PNG on Export (User Can Define Resolution)
--------------------------------------------------------------

property line_feed : (ASCII character 10)
property md_line_feed : (ASCII character 32) & (ASCII character 32) & (ASCII character 10)


set {template, template_name, extract_images, image_res, word_list} to {false, {}, false, {}, {}}
set export_options to {"One Color", "Titlecase Levels 1 - 3", "Text Correction", "Find Spaces", "Extract Images", "Most Used Words"}
set {template, template_name, extract_images, image_res, export_options, word_list} to ¬
	get_options(template, template_name, extract_images, image_res, export_options, word_list)

--------------------------------------------------------------
-- Ignore Words When Finding Most Frequent Words of PDF
--------------------------------------------------------------
set ignore_words to {"and", "the", "a", "for", "in", "on", "if", "which", "at", "this", "thus", "has", ¬
	"its", "but", "such", "these", "is", "to", "or", "of", "what", "it", "you", "", "with", "as", "from", "are", ¬
	"can", "that", "may", "be", "often", "most", "by", "an"}


--------------------------------------------------------------
-- Find Skim Notes & PDF Attributes
--------------------------------------------------------------

tell application "Skim"
	
	set all_notes to every note of front document
	set doc_name to (name of front document)
	set pdf_name to text 1 thru ((offset of "." in doc_name) - 1) of doc_name
	set posix_path to (path of front document)
	set doc_alias to ((POSIX file (posix_path)) as Unicode text) as alias
	set file_url to my path2url(posix_path)
	set skimmer_url to "skimmer://" & file_url & "?page="
	set file_path to "file://localhost" & file_url
	set {fc1, fc2, fc3, fc4, fc5, fc6} to favorite colors
	set {highlight_note, box_note} to {highlight note, box note}
	
	if ("Text Correction" is in export_options) or ("Find Spaces" is in export_options) then
		repeat with n in all_notes
			set note_text to text of n
			set text of n to my clean_txt(note_text, export_options, word_list)
		end repeat
	else if ("Most Used Words" is in export_options) then
		my main(doc_alias, ignore_words)
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

--------------------------------------------------------------
-- Make Template: Answered 'No' for 'Do you have a template?'
--------------------------------------------------------------

if template is false then
	
	tell application id "OOut"
		
		
		-- MAKE DOCUMENT & DEFINE DEFAULT FONTS
		
		make new document with properties {name:pdf_name}
		set doc to document pdf_name
		make new row with properties {topic:pdf_name, note:file_path} at end of rows of doc
		set status visible of doc to false -- Hide OO Check-Box Column
		
		if extract_images is true then
			make new column at doc with properties {title:"Images", width:320}
		end if
		
		set index of window 1 where name contains pdf_name to 1
		activate window 1
		set bounds of window 1 to dtb
		
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
			
			if "One Color" is in export_options then
				set style_single to make new named style with properties {name:"Heading Mark"}
				set value of attribute "font-fill" of style_single to {21047, 21047, 21047, 65535}
				set {style_one, style_two, style_three} to {style_single, style_single, style_single}
				set head_mark to name of style_single
			else
				set style_one to make new named style with properties {name:"Heading 1"}
				set value of attribute "font-fill" of style_one to {16587, 28329, 55163, 65535}
				
				set style_two to make new named style with properties {name:"Heading 2"}
				set value of attribute "font-fill" of style_two to {39981, 38846, 25155, 65535}
				
				set style_three to make new named style with properties {name:"Heading 3"}
				set value of attribute "font-fill" of style_three to {23514, 43610, 18771, 65535}
				
				set head_mark to {name of style_one, name of style_two, name of style_three}
			end if
			
			set style_five to make new named style with properties {name:"Green Highlight"}
			set value of attribute "text-background-color" of style_five to {59367, 61166, 56540, 65535}
			
			set style_six to make new named style with properties {name:"Red Highlight"}
			set value of attribute "text-background-color" of style_six to {63479, 59110, 58339, 65535}
			
		end tell
	end tell
	
else if template is true then
	
	--------------------------------------------------------------
	-- Users Template: Answered 'Yes' for 'Do you have a template?'
	--------------------------------------------------------------
	
	tell application id "OOut"
		set template_path to ((path to application support from user domain as string) & ¬
			"The Omni Group:OmniOutliner:Templates:" & template_name & ".oo3template:")
		open template_path
		make new row with properties {topic:pdf_name, note:file_path} at end of rows of front document
		tell application "System Events" to set frontmost of process "OmniOutliner" to true
		
		set doc to front document
		
		tell doc
			if "One Color" is in export_options then
				set style_single to make new named style with properties {name:"Heading Mark"}
				set {style_one, style_two, style_three} to {style_single, style_single, style_single}
				set head_mark to name of style_single
			else
				set {style_one, style_two, style_three} to named styles of doc
				set head_mark to {name of style_one, name of style_two, name of style_three}
			end if
			set style_five to make new named style with properties {name:"Green Highlight"}
			set value of attribute "text-background-color" of style_five to {59367, 61166, 56540, 65535}
			
			set style_six to make new named style with properties {name:"Red Highlight"}
			set value of attribute "text-background-color" of style_six to {63479, 59110, 58339, 65535}
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
					
					set head_one to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of rows of doc
					add style_one to named styles of style of head_one
					
					
				else if rgba is fc2 then
					
					try
						set head_two to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of last child of doc
					on error
						set head_two to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of rows of doc
					end try
					add style_two to named styles of style of head_two
					
					
				else if rgba is fc3 then
					
					set head_three to {}
					if previous_rows_style is in head_mark then
						set head_three to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of last row's children of doc
					else if level of parent of last row of doc ≤ 2 then
						set head_three to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of rows of parent of row_last
					else
						set head_three to make new row with properties {topic:my titlecase(note_text, export_options), note:note_url} at end of parent of last row's parent of doc
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





--------------------------------------------------------------
-- 
--                          Handlers
--
--------------------------------------------------------------


--------------------------------------------------------------
-- Get All Options
--------------------------------------------------------------
on get_options(template, template_name, extract_images, image_res, export_options, word_list)
	
	
	--OMNIOUTLINER
	
	--------------------------------------------------------------
	-- Use Custom OO Template or Make New
	--------------------------------------------------------------
	
	
	if button returned of (display dialog "Do you have a template?" buttons {"Yes", "No"} default button 2 with title "Use My Template or Use Script" with icon path to resource "OmniOutliner.icns" in bundle (path to application "OmniOutliner")) is "Yes" then -- Find Template
		set template to true
		
		--------------------------------------------------------------
		-- Locate User Template Folder & Promt User for Choice
		--------------------------------------------------------------
		
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
	end if
	
	set export_options to choose from list export_options default items {"Extract Images"} with prompt ("Options: Hold " & character id 8984 & " for Multiple") cancel button name "No Options" with multiple selections allowed and empty selection allowed
	
	
	-- FIND WORDS IN SPACELESS STRING
	
	if export_options contains "Find Spaces" then
		try
			display dialog ("This may take a few minutes." & return & "Results are not 100% perfect.") with title "Finding Spaces is Slow" with icon caution
		on error
			return
		end try
		set word_list to quoted form of (POSIX path of (choose file)) --as string
	end if
	
	
	--IMAGES
	
	if export_options contains "Extract Images" then
		
		tell front document of application "Skim"
			set all_notes to every note
			repeat with n in all_notes
				if type of n is box note then
					set extract_images to true
					set image_res to my get_res()
					exit repeat
				end if
			end repeat
		end tell
		if extract_images is false then
			display dialog ("There aren't any images selected") with title "Select Images First" with icon path to resource "Skim.icns" in bundle (path to application "Skim")
		end if
		
	end if
	
	return {template, template_name, extract_images, image_res, export_options, word_list}
	
end get_options


--------------------------------------------------------------
-- Encode PDF Filename for Use with Skimmer
--------------------------------------------------------------
on path2url(thepath)
	return do shell script "python -c \"import urllib, sys; print (urllib.quote(sys.argv[1]))\" " & quoted form of thepath
end path2url


--------------------------------------------------------------
-- Convert Levels 1-3 to Titlecase
--------------------------------------------------------------
on titlecase(txt, option)
	if option does not contain "Titlecase Levels 1 - 3" then
		return txt
	else
		return do shell script "python -c \"import sys; print unicode(sys.argv[1], 'utf8').title().encode('utf8')\" " & quoted form of txt
	end if
end titlecase


--------------------------------------------------------------
-- Clean Up Weird PDF's
--------------------------------------------------------------
on clean_txt(txt, options, word_list)
	
	set corrected_txt to {}
	
	if (options contains "Text Correction") and (options contains "Find Spaces") then
		set char_replace to my correct_char(txt)
		set corrected_txt to my find_spaces(char_replace, word_list)
	else if options contains "Text Correction" then
		set corrected_txt to my correct_char(txt)
	else if options contains "Find Spaces" then
		set corrected_txt to my find_spaces(txt, word_list)
	end if
	
	return corrected_txt
end clean_txt


--------------------------------------------------------------
-- Text Correction: Posi+on --> Position
--------------------------------------------------------------
on correct_char(txt)
	set get_TI to "([a-zA-Z])([0-9+%-])([a-zA-Z])"
	set replace_TI to "\\1ti\\3"
	set get_space to "([a-zA-Z])([$)])([a-zA-Z])"
	set replace_space to "\\1 \\3"
	set get_TT to "([a-zA-Z])([.`])([a-zA-Z])"
	set replace_TT to "\\1tt\\3"
	--set theWords to "Propriocep4ve, Posi+on, Contraindica%ons, a`achments, sta-cally, Lingual$nerve, ves6bule, supraglo.c, respiratory)center, restric-ve"
	set _cmd to "echo " & quoted form of txt & " \\
| sed -E 's!" & get_TI & "!" & replace_TI & "!g; s!" & get_space & "!" & replace_space & "!g; s!" & get_TT & "!" & replace_TT & "!g'"
	
	set corrected_txt to do shell script _cmd
	return corrected_txt
end correct_char


--------------------------------------------------------------
-- Find Spaces: where,canifindthe beer --> where can i find the beer
--------------------------------------------------------------
on find_spaces(txt, word_list)
	
	set note_text to quoted form of txt
	set _cmd to quoted form of "

from math import log 
import string

words = open(\"" & word_list & quoted form of "\").read().split()
wordcost = dict((k, log((i+1)*log(len(words)))) for i,k in enumerate(words))
maxword = max(len(x) for x in words)
table = string.maketrans(\"\",\"\")
l = \"\".join(\"" & note_text & quoted form of "\".split()).lower()

def infer_spaces(s):

    def best_match(i):
        candidates = enumerate(reversed(cost[max(0, i-maxword):i]))
        return min((c + wordcost.get(s[i-k-1:i], 9e999), k+1) for k,c in candidates)

    cost = [0]
    for i in range(1,len(s)+1):
        c,k = best_match(i)
        cost.append(c)

    out = []
    i = len(s)
    while i>0:
        c,k = best_match(i)
        assert c == cost[i]
        out.append(s[i-k:i])
        i -= k

    return \" \".join(reversed(out))

def test_trans(s):
    return s.translate(table, string.punctuation)
    
s = test_trans(l)
print(infer_spaces(s))"
	
	set corrected_txt to do shell script "python -c " & _cmd
	return corrected_txt
end find_spaces


--------------------------------------------------------------
-- Choose Resolution of Image Extraction
--------------------------------------------------------------
on get_res()
	set choice_res to choose from list {"Low", "Medium", "High", "Custom"} default items {"Medium"} with title "Scale Images" with prompt "Scale to what size?"
	if choice_res = false then return
	set choice_res to item 1 of choice_res
	
	if choice_res = "Low" then
		set pic_res to 300
	else if choice_res = "Medium" then
		set pic_res to 640
	else if choice_res = "High" then
		set pic_res to 1280
		
	else
		set choice_icon to note
		set custom_res to ""
		set choice_icon to path to resource "appicon.icns" in bundle (path to application "Grab")
		repeat
			set pic_res to text returned of (display dialog custom_res & "Please specify a maximum number of pixels for the longest side:" default answer "320" with icon choice_icon)
			try
				set pic_res to pic_res as integer
				exit repeat
			on error
				set custom_res to "You must enter a number. "
				set choice_icon to caution
			end try
		end repeat
	end if
end get_res


--------------------------------------------------------------
-- Create PNG from Skim Box Highlights
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


--------------------------------------------------------------
-- FIND THE 50 MOST FREQUENT WORDS OF THE PDF
--------------------------------------------------------------
on main(pdf_file, ignore_words)
	script o
		property wrds : missing value
		property scores : {}
		
		-- Custom comparison handler for the sort.
		-- This one compares the end items of passed lists in such a way as to produce a reversed sort.
		on isGreater(a, b)
			(end of a < end of b)
		end isGreater
	end script
	
	tell application "Skim"
		set doc_name to name of front document
		set o's wrds to words of text of front document
		set posix_path to path of front document
		set pdf_name to text 1 thru ((offset of "." in doc_name) - 1) of doc_name
		--set save_path to text 1 thru ((offset of "." in posix_path) - 1) of posix_path
		set total_words to count of words of text of front document
	end tell
	
	set this_file to (((path to desktop folder) as string) & pdf_name)
	
	-- Sort the list of words into groups of equal words.
	sort_words(o's wrds, 1, -1, {})
	
	-- Go through the sorted list, counting the instances of each word. 
	-- Store each word and its score in a list in the 'scores' list in the script object above.
	set current_word to item 1 of o's wrds
	set c to 1
	repeat with i from 2 to (count o's wrds)
		set this_word to item i of o's wrds
		if this_word is not in ignore_words then
			if (this_word is current_word) then
				set c to c + 1
			else
				set end of o's scores to {current_word, c}
				set current_word to this_word
				set c to 1
			end if
		end if
	end repeat
	set end of o's scores to {current_word, c}
	
	-- Reverse-sort the list of word/score lists by the scores themselves.
	sort_words(o's scores, 1, -1, {comparer:o})
	
	-- Report the 100 most frequently use words, if there are that many.
	set n to (count o's scores)
	if (n > 50) then set n to 50
	
	set the_report to "PDF: " & doc_name & ".pdf" & return & ¬
		"Total Words: " & total_words & return & return & ¬
		"The " & n & " Most Frequent Words" & return & return & return
	repeat with i from 1 to n
		set x to item i of o's scores
		set the_report to the_report & "x" & end of x & " - " & beginning of x & return
	end repeat
	
	
	try
		set the this_file to this_file as string
		set the open_this_file to open for access file this_file with write permission
		set eof of open_this_file to 0
		write the_report to open_this_file starting at eof
		close access the open_this_file
	on error
		try
			close access file this_file
		end try
	end try
	
	--do shell script "open " & quoted form of save_path
	do shell script "open " & quoted form of ((POSIX path of (path to desktop)) & pdf_name) as string
	
end main

on sort_words(wrd_list, l, r, customiser)
	script o
		property comparer : me
		property slave : me
		property lst : wrd_list
		
		on shsrt(l, r)
			set inc to (r - l + 1) div 2
			repeat while (inc > 0)
				slave's setInc(inc)
				repeat with j from (l + inc) to r
					set v to item j of o's lst
					repeat with i from (j - inc) to l by -inc
						tell item i of o's lst
							if (comparer's isGreater(it, v)) then
								set item (i + inc) of o's lst to it
							else
								set i to i + inc
								exit repeat
							end if
						end tell
					end repeat
					set item i of o's lst to v
					slave's shift(i, j)
				end repeat
				set inc to (inc / 2.2) as integer
			end repeat
		end shsrt
		
		on isGreater(a, b)
			(a > b)
		end isGreater
		
		on shift(a, b)
		end shift
		
		on setInc(a)
		end setInc
	end script
	
	set listLen to (count wrd_list)
	if (listLen > 1) then
		if (l < 0) then set l to listLen + l + 1
		if (r < 0) then set r to listLen + r + 1
		if (l > r) then set {l, r} to {r, l}
		
		if (customiser's class is record) then set {comparer:o's comparer, slave:o's slave} to (customiser & {comparer:o, slave:o})
		
		o's shsrt(l, r)
	end if
	
	return
end sort_words
