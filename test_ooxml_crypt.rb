require 'ooxml_crypt'
$LOAD_PATH.unshift File.expand_path('lib', __dir__)
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
  sheet.add_row ['First', 'Second', 'Third']
  sheet.add_row [1, 2, 3]
end

p.serialize('unencrypted-caxlsx.xlsx')

# This file gives an error in Excel
OoxmlCrypt.encrypt_file('unencrypted-caxlsx.xlsx', 'abc123', 'encrypted-caxlsx.xlsx')

require 'rubyXL'
require 'rubyXL/convenience_methods'

def standardize_xlsx_with_rubyxl(input_file, output_file)
  workbook = RubyXL::Parser.parse(input_file)
  workbook.write(output_file)
end

standardize_xlsx_with_rubyxl('unencrypted-caxlsx.xlsx', 'unencrypted-rubyxl.xlsx')

# This file works in Excel!
OoxmlCrypt.encrypt_file('unencrypted-rubyxl.xlsx', 'abc123', 'encrypted-rubyxl.xlsx')
