#runs as soon as program is started
BEGIN {
  system "cls"
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Ruby Spreadsheet Conversion v1.26"
  puts " "
  puts "  by lincoln131"
  puts "  2017"
  puts " "
  puts "-----------------------------------------------------------------------"
}

#runs at end of program
END {
  puts "-----------------------------------------------------------------------"
  puts " "
  puts "Program Finished."
  puts " "
  puts "-----------------------------------------------------------------------"
  puts " "
}

#ensures spreadsheet gem is installed and used
require 'rubyXL'

#parksing an existing workbook
workbook = RubyXL::Parser.parse("./speaking.xlsx")

#set active worksheet to first sheet of the workbook
worksheet = workbook[0]

#define some variables for later
column_name = []
#starting row
row = 1

#get input about number of students, warn make input integer
puts "Step 1"
puts "-------------------------------------------------------------------------"
puts " "
puts "How many students? (1 - 50) "
puts "* Will not parse if not an integer * "
puts "* Will also max at 50 * "
puts " "
total_students = gets.chomp.to_i
system "cls"
puts "Step 2"
puts "-------------------------------------------------------------------------"
puts " "
puts "Too many students. Limiting amount of students..." if total_students > 50 #warn about too many students
sleep 1 if total_students > 50
total_students = 50 if total_students > 50 #set to max if too many
puts "Number of students set to #{total_students}" if total_students <= 50 #print number if appropriate
puts " "
puts "-------------------------------------------------------------------------"
puts " "
sleep 2
system "cls"
#verbosity settings
puts "Step 3"
puts "-------------------------------------------------------------------------"
puts " "
puts "Do you want verbose mode? (y/n)"
verbose_mode = gets.chomp
verbose_mode = verbose_mode.downcase
if verbose_mode == "y"
    verbose = 1
    puts "Maximum verbosity enabled!"
elsif verbose_mode == "n"
    verbose = 0
else
    puts "Invalid input! Setting to verbose mode!"
    verbose = 1
end

puts "Step 4 "
puts "-------------------------------------------------------------------------"
puts " "
puts "Preparing to parse to ./speaking.xlsx for #{total_students} students" if verbose_mode == "y" #being verbose
puts " "
puts "-------------------------------------------------------------------------"
puts " "

sleep 5 if verbose_mode == "y" #sleep if verbose_mode


#set a custom error message
error_message = "!!!!!"

#tells user something is happening
puts "Processing..."
sleep 1

while row <= total_students
column = 0
heading1 = ""
heading2 = ""
heading3 = ""
heading4 = ""
heading5 = ""
heading6 = ""
heading7 = ""
heading8 = ""
heading9 = ""
heading10 = ""
heading11 = ""
heading12 = ""
heading13 = ""
heading14 = ""
heading15 = ""
heading16 = ""
heading17 = ""
heading18 = ""
heading19 = ""
heading20 = ""
heading21 = ""
heading22 = ""
heading23 = ""
heading24 = ""

88.times do
  column_name.push(worksheet.sheet_data[row][column].value)
    if column == 5
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading1 << "A"
        when "Not Demonstrated"
          heading1 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 32
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading1 << "B"
        when "Not Demonstrated"
          heading1 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 59
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading1 << "C"
        when "Not Demonstrated"
          heading1 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 6
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading2 << "A"
        when "Not Demonstrated"
          heading2 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 33
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading2 << "B"
        when "Not Demonstrated"
          heading2 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 60
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading2 << "C"
        when "Not Demonstrated"
          heading2 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 7
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading3 << "A"
        when "Not Demonstrated"
          heading3 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 34
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading3 << "B"
        when "Not Demonstrated"
          heading3 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 61
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading3 << "C"
        when "Not Demonstrated"
          heading3 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 8
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading4 << "A"
        when "Not Demonstrated"
          heading4 << "-"
        else
          puts "#{error_message}"
        end
    elsif column == 35
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading4 << "B"
        when "Not Demonstrated"
          heading4 << "-"
        else
          heading4 << "*"
          puts "#{error_message}"
        end
    elsif column == 62
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading4 << "C"
        when "Not Demonstrated"
          heading4 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 9
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading5 << "A"
        when "Not Demonstrated"
          heading5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 36
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading5 << "B"
        when "Not Demonstrated"
          heading5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 63
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading5 << "C"
        when "Not Demonstrated"
          heading5 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 10
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading6 << "A"
        when "Not Demonstrated"
          heading6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 37
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading6 << "B"
        when "Not Demonstrated"
          heading6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 64
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading6 << "C"
        when "Not Demonstrated"
          heading6 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 11
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading7 << "A"
        when "Not Demonstrated"
          heading7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 38
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading7 << "B"
        when "Not Demonstrated"
          heading7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 65
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading7 << "C"
        when "Not Demonstrated"
          heading7 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 12
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading8 << "A"
        when "Not Demonstrated"
          heading8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 39
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading8 << "B"
        when "Not Demonstrated"
          heading8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 66
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading8 << "C"
        when "Not Demonstrated"
          heading8 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 13
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading9 << "A"
        when "Not Demonstrated"
          heading9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 40
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading9 << "B"
        when "Not Demonstrated"
          heading9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 67
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading9 << "C"
        when "Not Demonstrated"
          heading9 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 14
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading10 << "A"
        when "Not Demonstrated"
          heading10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 41
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading10 << "B"
        when "Not Demonstrated"
          heading10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 68
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading10 << "C"
        when "Not Demonstrated"
          heading10 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 15
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading11 << "A"
        when "Not Demonstrated"
          heading11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 42
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading11 << "B"
        when "Not Demonstrated"
          heading11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 69
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading11 << "C"
        when "Not Demonstrated"
          heading11 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 16
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading12 << "A"
        when "Not Demonstrated"
          heading12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 43
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading12 << "B"
        when "Not Demonstrated"
          heading12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 70
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading12 << "C"
        when "Not Demonstrated"
          heading12 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 17
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading13 << "A"
        when "Not Demonstrated"
          heading13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 44
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading13 << "B"
        when "Not Demonstrated"
          heading13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 71
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading13 << "C"
        when "Not Demonstrated"
          heading13 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 18
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading14 << "A"
        when "Not Demonstrated"
          heading14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 45
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading14 << "B"
        when "Not Demonstrated"
          heading14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 72
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading14 << "C"
        when "Not Demonstrated"
          heading14 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 19
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading15 << "A"
        when "Not Demonstrated"
          heading15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 46
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading15 << "B"
        when "Not Demonstrated"
          heading15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 73
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading15 << "C"
        when "Not Demonstrated"
          heading15 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 20
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading16 << "A"
        when "Not Demonstrated"
          heading16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 47
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading16 << "B"
        when "Not Demonstrated"
          heading16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 74
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading16 << "C"
        when "Not Demonstrated"
          heading16 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 21
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading17 << "A"
        when "Not Demonstrated"
          heading17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 48
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading17 << "B"
        when "Not Demonstrated"
          heading17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 75
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading17 << "C"
        when "Not Demonstrated"
          heading17 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 22
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading18 << "A"
        when "Not Demonstrated"
          heading18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 49
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading18 << "B"
        when "Not Demonstrated"
          heading18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 76
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading18 << "C"
        when "Not Demonstrated"
          heading18 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 23
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading19 << "A"
        when "Not Demonstrated"
          heading19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 50
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading19 << "B"
        when "Not Demonstrated"
          heading19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 77
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading19 << "C"
        when "Not Demonstrated"
          heading19 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 24
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading20 << "A"
        when "Not Demonstrated"
          heading20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 51
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading20 << "B"
        when "Not Demonstrated"
          heading20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 78
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading20 << "C"
        when "Not Demonstrated"
          heading20 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 25
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading21 << "A"
        when "Not Demonstrated"
          heading21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 52
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading21 << "B"
        when "Not Demonstrated"
          heading21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 79
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading21 << "C"
        when "Not Demonstrated"
          heading21 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 26
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading22 << "A"
        when "Not Demonstrated"
          heading22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 53
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading22 << "B"
        when "Not Demonstrated"
          heading22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 80
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading22 << "C"
        when "Not Demonstrated"
          heading22 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 27
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading23 << "A"
        when "Not Demonstrated"
          heading23 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 54
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading23 << "B"
        when "Not Demonstrated"
          heading23 << "-"
        else
          puts "#{error_message}"
      end
    elsif column == 81
      first = worksheet.sheet_data[row][column].value
      case first
        when "Demonstrated"
          heading23 << "C"
        when "Not Demonstrated"
          heading23 << "-"
        else
          puts "#{error_message}"
      end
    end
    column = column + 1
end

student = worksheet.sheet_data[row][1].value

#check for verbosity then output
if verbose == 1
      puts "-----------------------------------------------------------------------------------------------------------"
      puts " "
      puts "Student with email address of #{student} processed"
      puts "Row with #{row} number processed"
      puts "#{heading1} for Objective 'NOVICE ask/respond to  predictable, formulaic questions'"
      puts "#{heading2} for Objective 'NOVICE list, name, identify'"
      puts "#{heading3} for Objective 'INTERMEDIATE ask, answer variety of questions'"
      puts "#{heading4} for Objective 'INTERMEDIATE initiate, maintain, end conversation for basic needs/transactions'"
      puts "#{heading5} for Objective 'INTERMEDIATE communicate beyond 'here & now'"
      puts "#{heading6} for Objective 'NOVICE practiced words, phrases, sentences'"
      puts "#{heading7} for Objective 'NOVICE formulaic/memorized questions'"
      puts "#{heading8} for Objective 'INTERMEDIATE strings of sentences'"
      puts "#{heading9} for Objective 'INTERMEDIATE connected sentences'"
      puts "#{heading10} for Objective 'INTERMEDIATE ask questions to initiate/ sustain conversation'"
      puts "#{heading11} for Objective 'NOVICE high-frequency'"
      puts "#{heading12} for Objective 'NOVICE formulaic'"
      puts "#{heading13} for Objective 'NOVICE practiced'"
      puts "#{heading14} for Objective 'INTERMEDIATE familiar themes/topics'"
      puts "#{heading15} for Objective 'INTERMEDIATE personalized'"
      puts "#{heading16} for Objective 'NOVICE imitate modeled words'"
      puts "#{heading17} for Objective 'NOVICE use facial expressions, gestures'"
      puts "#{heading18} for Objective 'NOVICE resort to first language'"
      puts "#{heading19} for Objective 'NOVICE repeat/request repetition'"
      puts "#{heading20} for Objective 'NOVICE indicate lack of understanding'"
      puts "#{heading21} for Objective 'INTERMEDIATE ask questions, request clarification'"
      puts "#{heading22} for Objective 'INTERMEDIATE self-correct or restate when not understood'"
      puts "#{heading23} for Objective 'INTERMEDIATE circumlocute'"
      puts " "
elsif verbose == 0
      puts "#{student}"
else
      puts "Something Broke!"
end


#changes to the output sheet
worksheet2 = workbook[1]
#outputs the crap

worksheet2.add_cell(0, row, "#{student}")
worksheet2.add_cell(1, row, "#{heading1}")
worksheet2.add_cell(2, row, "#{heading2}")
worksheet2.add_cell(3, row, "#{heading3}")
worksheet2.add_cell(4, row, "#{heading4}")
worksheet2.add_cell(5, row, "#{heading5}")
worksheet2.add_cell(6, row, "#{heading6}")
worksheet2.add_cell(7, row, "#{heading7}")
worksheet2.add_cell(8, row, "#{heading8}")
worksheet2.add_cell(9, row, "#{heading9}")
worksheet2.add_cell(10, row, "#{heading10}")
worksheet2.add_cell(11, row, "#{heading11}")
worksheet2.add_cell(12, row, "#{heading12}")
worksheet2.add_cell(13, row, "#{heading13}")
worksheet2.add_cell(14, row, "#{heading14}")
worksheet2.add_cell(15, row, "#{heading15}")
worksheet2.add_cell(16, row, "#{heading16}")
worksheet2.add_cell(17, row, "#{heading17}")
worksheet2.add_cell(18, row, "#{heading18}")
worksheet2.add_cell(19, row, "#{heading19}")
worksheet2.add_cell(20, row, "#{heading20}")
worksheet2.add_cell(21, row, "#{heading21}")
worksheet2.add_cell(22, row, "#{heading22}")
worksheet2.add_cell(23, row, "#{heading23}")

#writes the whole book
workbook.write("./speaking.xlsx")
row = row + 1
end
