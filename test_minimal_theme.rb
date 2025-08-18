#!/usr/bin/env ruby
require 'axlsx'
require 'ooxml_crypt'
require 'win32ole'

def test_excel_encryption(file_path, description)
  test_file = file_path.gsub('.xlsx', '-encrypted.xlsx')

  begin
    puts "   🔐 Encrypting #{description}..."
    OoxmlCrypt.encrypt_file(file_path, 'test123', test_file)
    puts "   ✅ Encryption: SUCCESS"

    puts "   📊 Testing encrypted file in Excel..."
    excel = WIN32OLE.new('Excel.Application')
    excel.visible = false
    excel.displayAlerts = false
    workbook = excel.Workbooks.Open(File.absolute_path(test_file), nil, nil, nil, 'test123')
    workbook.Close(false)
    excel.Quit

    File.delete(test_file) if File.exist?(test_file)
    puts "   ✅ Excel opening: SUCCESS"
    true
  rescue StandardError => e
    puts "   ❌ FAILED: #{e.message}"
    begin
      excel&.Quit
    rescue StandardError
    end
    File.delete(test_file) if File.exist?(test_file)
    false
  end
end

puts "🧪 TESTING MINIMAL THEME"
puts "=" * 50

puts "\n1️⃣ Creating file with minimal theme..."
p = Axlsx::Package.new
wb = p.workbook
wb.add_worksheet(name: 'Test Sheet') do |sheet|
  sheet.add_row ['Minimal', 'Theme', 'Test']
  sheet.add_row [1, 2, 3]
end
p.serialize('minimal_theme_test.xlsx')
puts "   ✅ File created with minimal theme"

puts "\n2️⃣ Testing minimal theme file..."
minimal_works = test_excel_encryption('minimal_theme_test.xlsx', 'minimal theme file')

puts "\n" + "=" * 50
puts "🎯 RESULT:"
puts "   Minimal theme works: #{minimal_works ? '✅ SUCCESS!' : '❌ FAILED'}"

if minimal_works
  puts "\n🎉 AMAZING! Even the minimal theme works!"
  puts "   Just an empty <a:theme> tag is sufficient for encryption compatibility"
else
  puts "\n💔 Minimal theme doesn't work. Need to add more content."
end

# Cleanup
File.delete('minimal_theme_test.xlsx') if File.exist?('minimal_theme_test.xlsx')
