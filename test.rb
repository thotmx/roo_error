require 'roo'
%w{ example.xlsx example.ods example.csv }.each do |file_name|
  file = File.open(File.join(File.dirname(__FILE__), file_name))

  spreadsheet = case File.extname(file.path)
                when ".csv" then Roo::CSV.new(file.path, csv_options: {col_sep: ";"})
                when ".xls"  then Roo::Excel.new(file.path)
                when ".xlsx" then Roo::Excelx.new(file.path)
                when ".ods"  then Roo::OpenOffice.new(file.path)
                else raise "Unknown file type: #{file.original_filename}"
                end

  header = spreadsheet.row(1)

  puts "*"*50
  puts file_name
  puts "*"*50

  (2..3).map do |i|
    row = Hash[[header, spreadsheet.row(i)].transpose]

    puts row
  end
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
# First row result .csv file:
#
# "OE"=>"340120"
