#Ruby program using RubyXL to parse some spreadsheets for my wife.
#runs as soon as program is started
BEGIN { system "cls" #clears the screen for ease of use
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Ruby Spreadsheet Conversion v1.26"
  puts " "
  puts "  by lincoln131"
  puts "  2017"
  puts " "
  puts "-----------------------------------------------------------------------"}


END {  #runs at end of program
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Program Finished."
  puts " "
  puts "-----------------------------------------------------------------------"
  puts " "}

require 'rubyXL' #ensures spreadsheet gem is installed and used
workbook = RubyXL::Parser.parse("./speaking.xlsx") #parsing an existing workbook
worksheet = workbook[0] #set active worksheet to first sheet of the workbook
worksheet2 = workbook[1] #defines the output sheet

column_name = [] #make a var
row = 1 #starting row, skips headings
column = 0 #starting column
error_message = "Something broke!" # Custom Error message

#get input about number of students, warn make input integer
puts "Step 1"
puts "------------------------------------------------------------------------"
puts " "
puts "How many students? (1 - 50) "
puts "* Will not parse if not an integer * "
puts "* Will also max at 50 * "
puts " "
total_students = gets.chomp.to_i #prompts for number of students
system "cls" #clears the screen for ease of use
puts "Step 2"
puts "-------------------------------------------------------------------------"
puts " "
puts "Too many students. Limiting amount of students..." if total_students > 50 #warn about too many students
sleep 1 if total_students > 50 #sleeps so user can see output
total_students = 50 if total_students > 50 #set to max if too many
puts "Number of students set to #{total_students}" if total_students <= 50 #print number if appropriate
puts " "
puts "-------------------------------------------------------------------------"
puts " "
sleep 2 #sleeps so user can see output
system "cls" #clears the screen for ease of use

#verbosity settings
puts "Step 3"
puts "-------------------------------------------------------------------------"
puts " "
puts "Do you want verbose mode? (y/n)"
verbose_mode = gets.chomp.downcase #prompts for y/n
#check user input for verbosity then sets variable and other stuff
(verbose = 1) && (puts "Maximum verbosity enabled!") && (sleep 2) if verbose_mode == "y"
(verbose = 0) if verbose_mode == "n"
(verbose = 1) && (puts "Invalid input! Setting to verbose mode!") && (sleep 2) if verbose_mode != "y" && verbose_mode != "n"

#output before processing!
system "cls" #clears the screen for ease of use
puts "Step 4 "
puts "-------------------------------------------------------------------------"
puts " "
puts "Preparing to parse to ./speaking.xlsx for #{total_students} students" if verbose_mode == "y" #being verbose
puts " "
puts "-------------------------------------------------------------------------"
puts " "

sleep 2.5 if verbose_mode == "y" #sleep if verbose_mode so user can see output

#tells user something is happening
system "cls" #clears the screen for ease of use
puts "Processing..."
sleep 1 #sleeps so user can see output

while row <= total_students #main loop runs once per row. Each row expected to be seperate student
column = 0 #makes sure column goes back to zero
obj_group = 5 #starting point for objective column
#break if worksheet.sheet_data[row][column] == nil #can't figure out how to make the damn thing break when it hits an empty row on spreadsheet.

#blanks the objective variables for the next loop
objective1 = ""

student = worksheet.sheet_data[row][1].value #gets the student for current row
class_period = worksheet.sheet_data[row][3].value #gets class period student is in

if verbose == 1
      puts "-----------------------------------------------------------------------------------------------------------"
      puts " "
      puts "Student with email address of #{student} processed"
      puts "Row with number #{row} processed"
end

88.times do #loop to check each field and concatenate the submissions.
            #I know there is a better way to do this, and I'll tidy it up as i get more proficient.
            #It is mostly repated, so I'm just commenting the first instance.
      if column == obj_group  #checks the column. I know I can clean this up when I get more proficient
                     #columns are in three objective groups, due to spreadsheet having inputs for three seperate submissions
        submission =  worksheet.sheet_data[row][column].value # if the column is correct, assigns variable
      case submission #case for variable
        when "Demonstrated"
          objective1 << "A" #if correct column and marked 'Demonstrated', pushes an 'A' to variable. Look below for B and C
        when "Not Demonstrated"
          objective1 << "-"#if correct column and marked 'Not Demonstrated', pushes a '-' to variable
        else
          puts "#{error_message}" #if correct column and unexpected input, displays an error message defined above
        end
    elsif column == obj_group+27
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective1 << "B" #if correct column and marked 'Demonstrated', appends a 'B' to variable
        when "Not Demonstrated"
          objective1 << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
        else
          puts "#{error_message}"
        end
    elsif column == obj_group+54
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective1 << "C" #if correct column and marked 'Demonstrated', appends a 'C' to variable
        when "Not Demonstrated"
          objective1 << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
        else
          puts "#{error_message}"
      end

if verbose == 1
  puts "#{objective1} for Objective '#{worksheet.sheet_data[0][column].value}'"
end

worksheet2.add_cell(column+2, row, "#{objective1}")

column = column + 1 #next column
obj_group += 1 #next obj_group
end #end of bigass loop for each student's column

if verbose == 1
      puts " "
      sleep 0.33 #take a nap so user can see output
elsif verbose == 0
      puts "#{row} - #{student} - Done!" #simple output for no verbosity
else
      puts "Something Broke!" #shouldn't ever see this
end

=begin
#check for verbosity then output
if verbose == 1
      puts "-----------------------------------------------------------------------------------------------------------"
      puts " "
      puts "Student with email address of #{student} processed"
      puts "Row with number #{row} processed"
      puts "#{objective1} for Objective '#{worksheet.sheet_data[0][5].value}'"
      puts "#{objective2} for Objective '#{worksheet.sheet_data[0][6].value}'"
      puts "#{objective3} for Objective '#{worksheet.sheet_data[0][7].value}'"
      puts "#{objective4} for Objective '#{worksheet.sheet_data[0][8].value}'"
      puts "#{objective5} for Objective '#{worksheet.sheet_data[0][9].value}'"
      puts "#{objective6} for Objective '#{worksheet.sheet_data[0][10].value}'"
      puts "#{objective7} for Objective '#{worksheet.sheet_data[0][11].value}'"
      puts "#{objective8} for Objective '#{worksheet.sheet_data[0][12].value}'"
      puts "#{objective9} for Objective '#{worksheet.sheet_data[0][13].value}'"
      puts "#{objective10} for Objective '#{worksheet.sheet_data[0][14].value}'"
      puts "#{objective11} for Objective '#{worksheet.sheet_data[0][15].value}'"
      puts "#{objective12} for Objective '#{worksheet.sheet_data[0][16].value}'"
      puts "#{objective13} for Objective '#{worksheet.sheet_data[0][17].value}'"
      puts "#{objective14} for Objective '#{worksheet.sheet_data[0][18].value}'"
      puts "#{objective15} for Objective '#{worksheet.sheet_data[0][19].value}'"
      puts "#{objective16} for Objective '#{worksheet.sheet_data[0][20].value}'"
      puts "#{objective17} for Objective '#{worksheet.sheet_data[0][21].value}'"
      puts "#{objective18} for Objective '#{worksheet.sheet_data[0][22].value}'"
      puts "#{objective19} for Objective '#{worksheet.sheet_data[0][23].value}'"
      puts "#{objective20} for Objective '#{worksheet.sheet_data[0][24].value}'"
      puts "#{objective21} for Objective '#{worksheet.sheet_data[0][25].value}'"
      puts "#{objective22} for Objective '#{worksheet.sheet_data[0][26].value}'"
      puts "#{objective23} for Objective '#{worksheet.sheet_data[0][27].value}'"
      puts " "
      sleep 0.33 #take a nap so user can see output
elsif verbose == 0
      puts "#{row} - #{student} - Done!" #simple output for no verbosity
else
      puts "Something Broke!" #shouldn't ever see this
end



worksheet2 = workbook[1] #defines the output sheet

#outputs the crap. I should be able to clean this up too.
worksheet2.add_cell(0, row, "#{student}")
worksheet2.add_cell(1, row, "#{class_period}")
worksheet2.add_cell(2, row, "#{objective1}")
worksheet2.add_cell(3, row, "#{objective2}")
worksheet2.add_cell(4, row, "#{objective3}")
worksheet2.add_cell(5, row, "#{objective4}")
worksheet2.add_cell(6, row, "#{objective5}")
worksheet2.add_cell(7, row, "#{objective6}")
worksheet2.add_cell(8, row, "#{objective7}")
worksheet2.add_cell(9, row, "#{objective8}")
worksheet2.add_cell(10, row, "#{objective9}")
worksheet2.add_cell(11, row, "#{objective10}")
worksheet2.add_cell(12, row, "#{objective11}")
worksheet2.add_cell(13, row, "#{objective12}")
worksheet2.add_cell(14, row, "#{objective13}")
worksheet2.add_cell(15, row, "#{objective14}")
worksheet2.add_cell(16, row, "#{objective15}")
worksheet2.add_cell(17, row, "#{objective16}")
worksheet2.add_cell(18, row, "#{objective17}")
worksheet2.add_cell(19, row, "#{objective18}")
worksheet2.add_cell(20, row, "#{objective19}")
worksheet2.add_cell(21, row, "#{objective20}")
worksheet2.add_cell(22, row, "#{objective21}")
worksheet2.add_cell(23, row, "#{objective22}")
worksheet2.add_cell(24, row, "#{objective23}")

=end

#fills in first column with objective titles
for x in 2..24
  worksheet2.add_cell(x, 0, "#{worksheet.sheet_data[0][x+3].value}")
end

workbook.write("./speaking.xlsx") #writes the whole book
row = row + 1 #sets up the next row

end #end of the main loop for each student
end
