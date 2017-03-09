#Ruby program using RubyXL to parse some spreadsheets for my wife.
#runs as soon as program is started
BEGIN { num_of_standards = 23
        path = "./speaking.xlsx"
  system "cls" #clears the screen for ease of use
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Speaking Responses Spreadsheet Conversion v2.0"
  puts " "
  puts "  by lincoln131"
  puts "  2017"
  puts " "
  puts "This program is currently defaulted to #{num_of_standards} standards."
  puts "The default path is currently #{path}."
  puts "-----------------------------------------------------------------------"}


END {  #runs at end of program
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Program Finished."
  puts " "
  puts "-----------------------------------------------------------------------"
  puts " "}

require 'rubyXL' #ensures spreadsheet gem is installed and used
workbook = RubyXL::Parser.parse(path) #parsing an existing workbook
worksheet = workbook[0] #set active worksheet to first sheet of the workbook
worksheet2 = workbook[1] #defines the output sheet

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
(verbose = 1) && (puts "Maximum verbosity enabled!") if verbose_mode == "y"
(verbose = 0) if verbose_mode == "n"
(verbose = 1) if (verbose_mode != "y") && (verbose_mode != "n")
sleep 2 if verbose_mode == "y"
system "cls" #clears the screen for ease of use

#output before processing
puts "Step 4 "
puts "-------------------------------------------------------------------------"
puts " "
puts "Preparing to parse to ./speaking.xlsx for #{total_students} students" if verbose_mode == "y" #being verbose
puts " "
puts "-------------------------------------------------------------------------"
puts " "
sleep 1.5 if verbose_mode == "y" #sleep if verbose_mode so user can see output

#tells user something is happening
puts "Processing..."
sleep 1 #sleeps so user can see output

#Main Loop for each student
while row <= total_students #main loop runs once per row. Each row expected to be seperate student
column = 1 #makes sure column goes back to default
obj_group = 5 #starting point for objective column
#break if worksheet.sheet_data[row][column] == nil #can't figure out how to make the damn thing break when it hits an empty row on spreadsheet.

#blanks the objective variables for the next loop
objective = " "

student = worksheet.sheet_data[row][1].value #gets the student for current row
url = worksheet.sheet_data[row][2].value #gets url for student
class_period = worksheet.sheet_data[row][3].value #gets class period student is in
worksheet2.add_cell(0, row, "#{student}")    #writes the student for current row
worksheet2.add_cell(1, row, "#{url}")    #writes the url for current student
worksheet2.add_cell(2, row, "#{class_period}")    #writes the classperiod for current student

if verbose == 1 && column = 5 #verbosity output
      puts "----------------------------------------------------------------------------"
      puts " "
      puts "Student with email address of #{student} processed"
      puts "Row with number #{row} processed"
end

#Loop for each student's objectives
num_of_standards.times do #loop to check each field and concatenate the submissions. 23 is the number of standards in this particular skill
      submission =  worksheet.sheet_data[row][column].value
      if column == obj_group
        submission1 =  worksheet.sheet_data[row][column].value
        submission2 =  worksheet.sheet_data[row][column+27].value
        submission3 =  worksheet.sheet_data[row][column+54].value
        case submission #case for variable
          when "Demonstrated"
            objective << "A" #if correct column and marked 'Demonstrated', pushes an 'A' to variable. Look below for B and C
          when "Not Demonstrated"
            objective << "-"#if correct column and marked 'Not Demonstrated', pushes a '-' to variable
          else
            objective = "Skipped"
          end
        case submission2
          when "Demonstrated"
            objective << "B" #if correct column and marked 'Demonstrated', appends a 'B' to variable
          when "Not Demonstrated"
            objective << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
          else
              objective = "Skipped"
        end
        case submission
          when "Demonstrated"
            objective << "C" #if correct column and marked 'Demonstrated', appends a 'C' to variable
          when "Not Demonstrated"
            objective << "-" #if correct column and marked 'Not Demonstrated', appends a '-' to variable
          else
            objective = "Skipped"
        end
        worksheet2.add_cell(column-2, row, "#{objective}") unless objective == "Skipped"
      end

      if verbose == 1 && row >= 1 && obj_group >= 5 #more verbosity controls
        puts "#{objective} for Objective '#{worksheet.sheet_data[0][column].value}'"
      end

      obj_group += 1 if column == obj_group #next obj_group if it was one that was used
      column = column + 1 #next column
      objective = " " #blanks objective for next pass

end #end of bigass loop for each student's column

if verbose == 1 #more verbosity outputs
      puts " "
      #sleep 0.25 #take a nap so user can see output
elsif verbose == 0
      puts "#{row} - #{student} - Done!" #simple output for no verbosity
end

#fills in first column with objective titles
for x in 3..25
  worksheet2.add_cell(x, 0, "#{worksheet.sheet_data[0][x+2].value}")
end

workbook.write("./speaking.xlsx") #writes the whole book
row = row + 1 #sets up the next row

end #end of the main loop for each student
