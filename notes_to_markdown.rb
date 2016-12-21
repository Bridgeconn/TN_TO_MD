#!/usr/bin/env ruby
require 'roo'

Dir.glob("**/*.xlsx") do |file|
  xlsx = Roo::Spreadsheet.open(file)
  bookname = xlsx.column(1)

  xlsx.column(3).each do |cl|
    if (cl != "Verse")
      directory_name = bookname[1]
      Dir.mkdir(directory_name) unless File.exists?(directory_name)
      if cl[0]
        output_name = "#{directory_name}/#{File.basename(cl[0], '.*')}.md"
      end
      output = File.open(output_name, 'w')

      tn = xlsx.column(5)
      tn_data = tn[1]
      p_tn_data = tn_data.split(/\R+/)
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
