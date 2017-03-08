# rubyspanishxlsx
Custom ruby for parsing a spreadsheet for my wife's Spanish class.

This will take an xlsx file that my wife has had her students fill with google forms and parse it based on some criteria so that she can glance at what objective the students' submissions are meeting.

It requires RubyXL and it's dependencies.

It can only parse *.XLSX files.
The XLSX file is expected to be in the same folder as the *.rb.
The filename is hardcoded at the moment.
The particular objectives the program is looking for are in a particular place in the spreadsheet and are currently hardcoded.
This was my first Ruby project and it's pretty messy and not very efficient.
