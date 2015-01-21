--------------------------------------------------------------
-- DrLulz Jan 2015
-- 
-- CREDITS:
-- macscripter.net
-- macosxautomation.com/applescript/
-- Evernote: Idea & Most of Script from packal.org/workflow/skimmer
-------- Scale Images: tuaw.com/2013/01/28/applescripting-evernote-embedded-image-scaler/
-- Split Spaceless String into Words stackoverflow.com/questions/8870261/how-to-split-text-without-spaces-into-list-of-words
-------- Word Frequency Data: wordfrequency.info
-------- Word Frequency Data: invokeit.wordpress.com/frequency-word-lists/
-- Frequency of ALL WORDS IN PDF: macscripter.net/viewtopic.php?pid=136205
--------------------------------------------------------------

--------------------------------------------------------------
-- Choose Export Options
--
-- Titlecase Levels 1 - 3: this is cool --> This Is Cool
-- Text Correction: Posi+on, Contraindica%ons, a`achments, sta-cally --> Position, Contraindications, attachments, statically
-- Find Spaces (in Spaceless String): where,canifindthe beer --> where can i find the beer
-- Extract Images: Finds boundries of Skims 'Box Note' and Creates PNG on Export (User Can Define Resolution)
--------------------------------------------------------------

property line_feed : (ASCII character 10)
property md_line_feed : (ASCII character 32) & (ASCII character 32) & (ASCII character 10)

set {extract_images, image_res, word_list} to {false, {}, {}}
set export_options to {"Titlecase Levels 1 - 3", "Text Correction", "Find Spaces", "Extract Images", "Most Used Words"}

set export_options to choose from list export_options default items {"Extract Images"} with prompt ("Options: Hold " & character id 8984 & " for Multiple") cancel button name "No Options" with multiple selections allowed and empty selection allowed

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
	set file_url to my encode_text(posix_path, false, false)
	set skimmer_url to "skimmer://" & file_url & "?page="
	set file_path to "file://localhost" & file_url
	set {fc1, fc2, fc3, fc4, fc5, fc6} to favorite colors
	set {highlight_note, box_note} to {highlight note, box note}
	
	
	--IMAGES
	if export_options contains "Extract Images" then
		repeat with n in all_notes
			if type of n is box note then
				set extract_images to true
				set image_res to my get_res()
				exit repeat
			end if
		end repeat
		if extract_images is false then
			display dialog ("There aren't any images selected") with title "Select Images First" with icon path to resource "Skim.icns" in bundle (path to application "Skim")
		end if
	end if
	
	
	-- FIND WORDS IN SPACELESS STRING
	if export_options contains "Find Spaces" then
		try
			display dialog ("This may take a few minutes." & return & "Results are not 100% perfect.") with title "Finding Spaces is Slow" with icon caution
		on error
			return
		end try
		set word_list to quoted form of (POSIX path of (choose file)) --as string
	end if
	
	
	--TEXT CORRECTION
	if ("Text Correction" is in export_options) or ("Find Spaces" is in export_options) then
		repeat with n in all_notes
			set note_text to text of n
			set text of n to my clean_txt(note_text, export_options, word_list)
		end repeat
	else if ("Most Used Words" is in export_options) then
		my main(doc_alias, ignore_words)
	end if
	
	set annotations to line_feed & line_feed & md_line_feed & line_feed & line_feed
	
	repeat with n from 1 to count of all_notes
		set _note to item n of all_notes
		set type_note to type of _note
		set page_index to index of page of _note
		set page_num to page_index as string
		set this_url to skimmer_url & page_index
		set rgba to color of _note
		set fav_colors to favorite colors
		set note_text to text of _note
		set _bounds to bounds of _note
		set _page to page of _note
		set _data to grab _page for _bounds
		
		set en_bullet to my get_bullet(rgba, fav_colors)
		
		if type_note = highlight_note then
			set annotations to my annotations & (my en_html(rgba, fav_colors, en_bullet, note_text, this_url, page_num, export_options))
		else if (type_note = box_note) and (extract_images = true) then
			my write_temp(_data, page_index, image_res)
			set annotations to my annotations & (my en_import(page_index))
		end if
		
	end repeat
	
	set final_text to ""
	set final_text to final_text & annotations
	
	
	--------------------------------------------------------------
	-- Evernote
	--------------------------------------------------------------
	
	tell application "System Events"
		if not (exists process "Evernote") then
			do shell script "open -a \"Evernote\""
		end if
	end tell
	
	tell application id "com.evernote.Evernote"
		activate
		
		try
			if miniaturized of front window is true then
				set miniaturized of front window to false
			end if
		on error
			repeat until front window exists
				try
					if miniaturized of front window is true then
						set miniaturized of front window to false
					end if
				end try
			end repeat
		end try
		
		
		set en_note to create note title pdf_name with html final_text notebook my en_inbox()
		try
			do shell script "rm -r " & quoted form of (POSIX path of (path to temporary items from user domain) & "pic_temp_folder")
		end try
		
		
		--Choose/Create Tags for Export
		set make_tags to "ENTER YOUR TAGS HERE"
		set dialog_title to "Assign Tags On Exported Skim Notes"
		display dialog "Enter Your Tags Separated By Commas" with title dialog_title default answer make_tags buttons {"No Tags", "Assign"} default button "Assign" with icon path to resource "Evernote.icns" in bundle (path to application "Evernote")
		set dialog_result to the result
		set user_input to text returned of dialog_result
		set button_sel to button returned of dialog_result
		set the_delims to {", ", ","}
		set en_tag to my tag_list(user_input, the_delims)
		set en_tag_list to my tag_check(en_tag)
		
		if button_sel is "Assign" then
			assign en_tag_list to en_note
		end if
		
		repeat until isSynchronizing is false
			synchronize
		end repeat
		repeat until isSynchronizing is false
		end repeat
		
	end tell
	
end tell

--------------------------------------------------------------
-- 
--                          Handlers
--
--------------------------------------------------------------


--------------------------------------------------------------
-- Encode PDF Filename for Use with Skimmer
-- github.com/smargh/alfred_skimmer
--------------------------------------------------------------
on encode_text(this_text, encode_URL_A, encode_URL_B)
	set the standard_characters to "abcdefghijklmnopqrstuvwxyz0123456789"
	set the URL_A_chars to "$+!'/?;&@=#%><{}[]\"~`^\\|*"
	set the URL_B_chars to ".-_:"
	set the acceptable_characters to the standard_characters
	if encode_URL_A is false then set the acceptable_characters to the acceptable_characters & the URL_A_chars
	if encode_URL_B is false then set the acceptable_characters to the acceptable_characters & the URL_B_chars
	set the encoded_text to ""
	repeat with this_char in this_text
		if this_char is in the acceptable_characters then
			set the encoded_text to (the encoded_text & this_char)
		else
			set the encoded_text to (the encoded_text & encode_char(this_char)) as string
		end if
	end repeat
	return the encoded_text
end encode_text

on encode_char(this_char)
	set the ASCII_num to (the ASCII number this_char)
	set the hex_list to {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"}
	set x to item ((ASCII_num div 16) + 1) of the hex_list
	set y to item ((ASCII_num mod 16) + 1) of the hex_list
	return ("%" & x & y) as string
end encode_char


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
	
end write_temp


--------------------------------------------------------------
-- EVERNOTE HANDLERS
--
-- Change font_ (family, color, size, style, weight)
-- Idea from gist.github.com/smargh/ea15a72df99debae411c
--------------------------------------------------------------

on en_inbox()
	tell application "Evernote" to return name of (every notebook whose default is true)
end en_inbox


on en_html(rgba, fav_colors, _title, note_text, hyperlink, page_num, options)
	
	set {fc1, fc2, fc3, fc4, fc5, fc6} to fav_colors
	
	set font_family to "Avenir, sans-serif"
	set font_color to "rgb(75, 75, 75)"
	set font_size to "18px"
	set font_style to "normal"
	set font_weight to "100"
	
	if rgba is equal to fc1 then
		
		set top to "<br></br><p style=\"font-size: 0px;\"><hr/><p style=\"font-size: 10px;\"></p>"
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & my titlecase(note_text, options) & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return top & space & pre & space & body & space & page & line_feed & line_feed
		
		
	else if rgba is equal to fc2 then
		
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & my titlecase(note_text, options) & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return pre & space & body & space & page & line_feed & line_feed
		
		
	else if rgba is equal to fc3 then
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & my titlecase(note_text, options) & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return pre & space & body & space & page & line_feed & line_feed
		
		
	else if rgba is equal to fc4 then
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & note_text & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return pre & space & body & space & line_feed & line_feed
		
		
	else if rgba is equal to fc5 then
		
		set font_color to "rgb(204, 121, 40)"
		
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & note_text & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return pre & space & body & space & line_feed & line_feed
		
		
	else if rgba is equal to fc6 then
		
		set font_color to "rgb(175, 0, 0)"
		
		set pre to "<p style=\"font-size: 14px;\">" & "<a href=\"" & hyperlink & "\" style=\"text-decoration: none\">" & _title & "</a>"
		set body to "<span style=\"color: " & font_color & "; font-family: " & font_family & "; font-size: " & font_size & "; font-style: " & font_style & "; font-variant: normal; font-weight: " & font_weight & "; letter-spacing: normal; orphans: auto; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(254, 254, 254); float: none;\">" & note_text & "</span>"
		set page to "" & "<font size=\"1\" color=\"#4C4C4C\">(" & page_num & ")</font></p>"
		return pre & space & body & space & page & line_feed & line_feed
		
	end if
end en_html

--------------------------------------------------------------
-- Create Icon for Hyperlinked Bullets
-- Idea from gist.github.com/smargh/ea15a72df99debae411c
--------------------------------------------------------------
on get_bullet(rgba, fav_colors)
	set colorText to ""
	set html_tab to "&nbsp;"
	
	if rgba is (item 1 of fav_colors) then
		set colorText to "◾️" & html_tab
	else if rgba is (item 2 of fav_colors) then
		set colorText to my str_multiply(4, html_tab) & "▪️"
	else if rgba is (item 3 of fav_colors) then
		set colorText to my str_multiply(7, html_tab) & "▫️"
	else if rgba is (item 4 of fav_colors) then
		set colorText to my str_multiply(16, html_tab)
	else if rgba is (item 5 of fav_colors) then
		set colorText to str_multiply(11, html_tab) & "⭐️"
	else if rgba is (item 6 of fav_colors) then
		set colorText to str_multiply(11, html_tab) & "❗️"
	end if
	
	return colorText
end get_bullet

on str_multiply(x, str)
	set new to str
	repeat x times
		set new to new & str
	end repeat
	return new
end str_multiply

--------------------------------------------------------------
-- Create Evernote Tag List
-- veritrope.com/code/apple-mail-to-evernote/
--------------------------------------------------------------
on tag_list(user_input, the_delims)
	set oldDelims to text item delimiters
	set theList to {user_input}
	repeat with aDelim in the_delims
		set text item delimiters to aDelim
		set newList to {}
		repeat with anItem in theList
			set newList to newList & text items of anItem
		end repeat
		set theList to newList
	end repeat
	return theList as list
	set text item delimiters to oldDelims
end tag_list

--------------------------------------------------------------
-- Create Evernote Tags if the User Types New Tags
--------------------------------------------------------------
on tag_check(theTags)
	tell application id "com.evernote.Evernote"
		set finalTags to {}
		repeat with theTag in theTags
			if (not (tag named theTag exists)) then
				set makeTag to make tag with properties {name:theTag}
				set end of finalTags to makeTag
			else
				set end of finalTags to tag theTag
			end if
		end repeat
	end tell
	return finalTags
end tag_check


--------------------------------------------------------------
-- Used to Make List of Images for Evernote Import
--------------------------------------------------------------
on en_import(page_index)
	
	set pic_posix to posix path of (path to temporary items from user domain as string) & "pic_temp_folder/box_page_" & page_index & ".png"
	set image_url to "file:///" & pic_posix
	set image_html to "<div><img src=\"" & image_url & "\" style=\"\"/></div>"
	
end en_import


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
