#!/usr/bin/env ruby
require 'roo'

Dir.glob("**/*.xlsx") do |file|
  xlsx = Roo::Spreadsheet.open(file)
  chapter = ""
  bookname = xlsx.column(1)
  book_name = bookname[1]
  verse = ""
  book = {}
  chapter = {}
  verses = {}

  check_chapter = ""
  xlsx.each_with_pagename do |name, sheet|
    sheet.each(book: "Book", chapter: "Chapter", verse: "Verse" , notes: "TRANSLATION" )  do |s|
      if s[:chapter] != "Chapter"
        verse = s[:verse].split("-")[0]
        if(check_chapter != s[:chapter])
          verses = {}
          check_chapter = s[:chapter]

          verses[verse] = s[:notes]
          chapter[s[:chapter]] = verses
          

        elsif(check_chapter == s[:chapter])
          verses[verse] = s[:notes]
        end
      end
    end
  end

  chapter.each do |chapter_v, v|
    Dir.mkdir("#{book_name}") unless File.exists?("#{book_name}")
    Dir.mkdir("#{book_name}/#{chapter_v}") unless File.exists?("#{book_name}/#{chapter_v}")
      verse_m = chapter[chapter_v]
      verse_m.each do |verse, v|
        output_name = "#{book_name}/#{chapter_v}/#{File.basename(verse.partition('-').first, '.*')}.md"
        output      = File.open("#{output_name}", 'w')
        p_tn_data = verse_m[verse].split(/\R+/)
        split_bullet = p_tn_data.join('').split(/(?=•)/) 

        iterate_split_bullet = split_bullet.each
        iterate_split_bullet.next

        loop do
          begin
            final_data = iterate_split_bullet.next
            before_hyphan = final_data.partition('-').first.partition('•').last # It will show only string which is before hyphan.
            after_hyphan  = final_data.partition('-').last # It will show only string which is after hyphan.
            # Need to add here logic so that obly require varse should go in file. 

            output << "#"+"#{before_hyphan}\n\n"
            output << "#{after_hyphan} \n"
          rescue StopIteration
            break
          end
        end
      output.close
    end
  end  
end



