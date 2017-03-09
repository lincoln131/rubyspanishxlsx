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

#break if worksheet.sheet_data[row][column] == nil #can't figure out how to make the damn thing break when it hits an empty row on spreadsheet.

#blanks the objective variables for the next loop
objective1 = ""
objective2 = ""
objective3 = ""
objective4 = ""
objective5 = ""
objective6 = ""
objective7 = ""
objective8 = ""
objective9 = ""
objective10 = ""
objective11 = ""
objective12 = ""
objective13 = ""
objective14 = ""
objective15 = ""
objective16 = ""
objective17 = ""
objective18 = ""
objective19 = ""
objective20 = ""
objective21 = ""
objective22 = ""
objective23 = ""
objective24 = ""

88.times do #loop to check each field and concatenate the submissions.
            #I know there is a better way to do this, and I'll tidy it up as i get more proficient.
            #It is mostly repated, so I'm just commenting the first instance.
      if column == 5 #checks the column. I know I can clean this up when I get more proficient
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
    elsif column == 32
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective1 << "B" #if correct column and marked 'Demonstrated', appends a 'B' to variable
        when "Not Demonstrated"
          objective1 << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
        else
          puts "#{error_message}"
        end
    elsif column == 59
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective1 << "C" #if correct column and marked 'Demonstrated', appends a 'C' to variable
        when "Not Demonstrated"
          objective1 << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
        else
          puts "#{error_message}"
      end
    elsif column == 6
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective2 << "A"
        when "Not Demonstrated"
          objective2 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 33
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective2 << "B"
        when "Not Demonstrated"
          objective2 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 60
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective2 << "C"
        when "Not Demonstrated"
          objective2 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 7
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective3 << "A"
        when "Not Demonstrated"
          objective3 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 34
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective3 << "B"
        when "Not Demonstrated"
          objective3 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 61
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective3 << "C"
        when "Not Demonstrated"
          objective3 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 8
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective4 << "A"
        when "Not Demonstrated"
          objective4 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 35
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective4 << "B"
        when "Not Demonstrated"
          objective4 << "-"
        else
          objective4 << "*"
          puts "#{error_message}"
        end
    elsif column == 62
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective4 << "C"
        when "Not Demonstrated"
          objective4 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 9
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective5 << "A"
        when "Not Demonstrated"
          objective5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 36
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective5 << "B"
        when "Not Demonstrated"
          objective5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 63
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective5 << "C"
        when "Not Demonstrated"
          objective5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 10
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective6 << "A"
        when "Not Demonstrated"
          objective6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 37
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective6 << "B"
        when "Not Demonstrated"
          objective6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 64
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective6 << "C"
        when "Not Demonstrated"
          objective6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 11
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective7 << "A"
        when "Not Demonstrated"
          objective7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 38
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective7 << "B"
        when "Not Demonstrated"
          objective7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 65
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective7 << "C"
        when "Not Demonstrated"
          objective7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 12
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective8 << "A"
        when "Not Demonstrated"
          objective8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 39
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective8 << "B"
        when "Not Demonstrated"
          objective8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 66
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective8 << "C"
        when "Not Demonstrated"
          objective8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 13
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective9 << "A"
        when "Not Demonstrated"
          objective9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 40
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective9 << "B"
        when "Not Demonstrated"
          objective9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 67
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective9 << "C"
        when "Not Demonstrated"
          objective9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 14
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective10 << "A"
        when "Not Demonstrated"
          objective10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 41
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective10 << "B"
        when "Not Demonstrated"
          objective10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 68
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective10 << "C"
        when "Not Demonstrated"
          objective10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 15
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective11 << "A"
        when "Not Demonstrated"
          objective11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 42
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective11 << "B"
        when "Not Demonstrated"
          objective11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 69
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective11 << "C"
        when "Not Demonstrated"
          objective11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 16
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective12 << "A"
        when "Not Demonstrated"
          objective12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 43
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective12 << "B"
        when "Not Demonstrated"
          objective12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 70
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective12 << "C"
        when "Not Demonstrated"
          objective12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 17
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective13 << "A"
        when "Not Demonstrated"
          objective13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 44
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective13 << "B"
        when "Not Demonstrated"
          objective13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 71
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective13 << "C"
        when "Not Demonstrated"
          objective13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 18
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective14 << "A"
        when "Not Demonstrated"
          objective14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 45
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective14 << "B"
        when "Not Demonstrated"
          objective14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 72
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective14 << "C"
        when "Not Demonstrated"
          objective14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 19
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective15 << "A"
        when "Not Demonstrated"
          objective15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 46
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective15 << "B"
        when "Not Demonstrated"
          objective15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 73
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective15 << "C"
        when "Not Demonstrated"
          objective15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 20
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective16 << "A"
        when "Not Demonstrated"
          objective16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 47
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective16 << "B"
        when "Not Demonstrated"
          objective16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 74
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective16 << "C"
        when "Not Demonstrated"
          objective16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 21
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective17 << "A"
        when "Not Demonstrated"
          objective17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 48
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective17 << "B"
        when "Not Demonstrated"
          objective17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 75
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective17 << "C"
        when "Not Demonstrated"
          objective17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 22
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective18 << "A"
        when "Not Demonstrated"
          objective18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 49
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective18 << "B"
        when "Not Demonstrated"
          objective18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 76
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective18 << "C"
        when "Not Demonstrated"
          objective18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 23
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective19 << "A"
        when "Not Demonstrated"
          objective19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 50
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective19 << "B"
        when "Not Demonstrated"
          objective19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 77
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective19 << "C"
        when "Not Demonstrated"
          objective19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 24
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective20 << "A"
        when "Not Demonstrated"
          objective20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 51
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective20 << "B"
        when "Not Demonstrated"
          objective20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 78
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective20 << "C"
        when "Not Demonstrated"
          objective20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 25
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective21 << "A"
        when "Not Demonstrated"
          objective21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 52
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective21 << "B"
        when "Not Demonstrated"
          objective21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 79
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective21 << "C"
        when "Not Demonstrated"
          objective21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 26
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective22 << "A"
        when "Not Demonstrated"
          objective22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 53
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective22 << "B"
        when "Not Demonstrated"
          objective22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 80
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective22 << "C"
        when "Not Demonstrated"
          objective22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 27
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective23 << "A"
        when "Not Demonstrated"
          objective23 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 54
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective23 << "B"
        when "Not Demonstrated"
          objective23 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 81
      submission =  worksheet.sheet_data[row][column].value
      case submission
        when "Demonstrated"
          objective23 << "C"
        when "Not Demonstrated"
          objective23 << "-"
        else
          puts "#{error_message}"
      end
    end

    column = column + 1 #next column
end #end of bigass loop for each student's column

student = worksheet.sheet_data[row][1].value #gets the student for current row
class_period = worksheet.sheet_data[row][3].value #gets class period student is in

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

#fills in first column with objective titles
for x in 2..24
  worksheet2.add_cell(x, 0, "#{worksheet.sheet_data[0][x+3].value}")
end

workbook.write("./speaking.xlsx") #writes the whole book
row = row + 1 #sets up the next row

end #end of the main loop for each student
