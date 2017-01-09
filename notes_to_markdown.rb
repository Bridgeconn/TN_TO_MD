#!/usr/bin/env ruby
require 'roo'

Dir.glob("**/*.xlsx") do |file|
  xlsx = Roo::Spreadsheet.open(file)

  bookname             = xlsx.column(1)
  cahpter_number_array = xlsx.column(2).uniq
  
  cahpter_number_array.each do |chapter|
    if (chapter != "Chapter")
      book_name      = bookname[1] if bookname
      chapter_number = chapter if (chapter != "Chapter")

      Dir.mkdir(book_name) unless File.exists?(book_name)
      Dir.mkdir("#{book_name}/#{chapter_number}") unless File.exists?("#{book_name}/#{chapter_number}")

      xlsx.column(3).each do |md|
        if (md != "Verse")
          output_name = "#{book_name}/#{chapter_number}/#{File.basename(md.partition('-').first, '.*')}.md"
          output      = File.open("#{output_name}", 'w')

          xlsx.column(4).each_with_index do |item, index|
            p_tn_data = item.split(/\R+/)
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
          end
          output.close
        end
      end
    end
  end

  # xlsx.column(3).each_with_index do |cl, index|
  #   if (cl != "Verse")
  #     directory_name = bookname[1]
  #     puts cl.partition('-').first

  #     Dir.mkdir(directory_name) unless File.exists?(directory_name)
  #     if cl.partition('-').first
  #       output_name = "#{directory_name}/#{File.basename(cl.partition('-').first, '.*')}.md"
  #     end
  #     output = File.open(output_name, 'w')

  #     xlsx.column(4).each_with_index do |item, index|
  #       # tn_data = tn[index]
  #       p_tn_data = item.split(/\R+/)
  #       split_bullet = p_tn_data.join('').split(/(?=•)/) 

  #       iterate_split_bullet = split_bullet.each
  #       iterate_split_bullet.next

  #       loop do
  #         begin
  #           final_data = iterate_split_bullet.next
  #           before_hyphan = final_data.partition('-').first.partition('•').last # It will show only string which is before hyphan.
  #           after_hyphan  = final_data.partition('-').last # It will show only string which is after hyphan.
  #           # Need to add here logic so that obly require varse should go in file. 

  #           output << "#"+"#{before_hyphan}\n\n"
  #           output << "#{after_hyphan} \n"
  #         rescue StopIteration
  #           break
  #         end
  #       end

  #     end
  #     output.close
  #   end
  # end
end


