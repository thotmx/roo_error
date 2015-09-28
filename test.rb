require 'roo'
file = File.open("#{File.dirname(__FILE__)}/example.xlsx")
# file = File.open("#{File.dirname(__FILE__)}/example.ods")
# file = File.open("#{File.dirname(__FILE__)}/example.csv")

spreadsheet = case File.extname(file.path)
	when ".csv" then Roo::CSV.new(file.path, csv_options: {col_sep: ";"})
	when ".xls"  then Roo::Excel.new(file.path)
	when ".xlsx" then Roo::Excelx.new(file.path)
	when ".ods"  then Roo::OpenOffice.new(file.path)
	else raise "Unknown file type: #{file.original_filename}"
end

header = spreadsheet.row(1)

(2..spreadsheet.last_row).map do |i|
  row = Hash[[header, spreadsheet.row(i)].transpose]

  puts row
end


########################################################
#
# First row result of .xlsx file:
#
# "OE"=>34012.0
########################################################
#
# First row result .ods file:
#
# "OE"=>34012.0
########################################################
#
# First row result .ods file:
#
# "OE"=>"340120"