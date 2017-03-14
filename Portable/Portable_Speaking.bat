echo off
cd Ruby2.3.0\bin
move %~dp0\speaking.xlsx %~dp0\Ruby2.3.0\bin
ruby %~dp0\speaking_spreadsheet_conv_2.0_portable.rb
move %~dp0\Ruby2.3.0\bin\speaking.xlsx  %~dp0\ 
echo Process complete. Press ENTER to open the spreadsheet.
pause
start excel %~dp0\speaking.xlsx
