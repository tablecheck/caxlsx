#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'
require 'ooxml_crypt'
require 'win32ole'

def test_excel_encryption(file_path)
  test_file = file_path.gsub('.xlsx', '-test-encrypted.xlsx')

  begin
    OoxmlCrypt.encrypt_file(file_path, 'test123', test_file)

    excel = WIN32OLE.new('Excel.Application')
    excel.visible = false
    excel.displayAlerts = false
    workbook = excel.Workbooks.Open(File.absolute_path(test_file), nil, nil, nil, 'test123')
    workbook.Close(false)
    excel.Quit

    File.delete(test_file) if File.exist?(test_file)
    puts "✅ Encryption + Excel: SUCCESS"
    true
  rescue StandardError => e
    puts "❌ Encryption + Excel: FAILED - #{e.message}"
    begin
      excel&.Quit
    rescue StandardError
    end
    File.delete(test_file) if File.exist?(test_file)
    false
  end
end

def files_identical?(file1, file2)
  return false unless File.exist?(file1) && File.exist?(file2)

  hash1 = Digest::MD5.hexdigest(File.read(file1))
  hash2 = Digest::MD5.hexdigest(File.read(file2))
  puts "File1 hash: #{hash1}"
  puts "File2 hash: #{hash2}"
  hash1 == hash2
end

puts "🧪 SIMPLE TEST: Direct replacement verification"
puts "=" * 60

# Test the known working RubyXL file first
puts "\n1️⃣ Testing original RubyXL file..."
test_excel_encryption('unencrypted-rubyxl.xlsx')

# Test the known failing caxlsx file
puts "\n2️⃣ Testing original caxlsx file..."
test_excel_encryption('base-caxlsx.xlsx')

# Now try the simplest possible replacement: just copy the RubyXL file
puts "\n3️⃣ Creating perfect copy of RubyXL file..."
FileUtils.copy('unencrypted-rubyxl.xlsx', 'perfect_copy.xlsx')

puts "   Checking if copy is identical..."
identical = files_identical?('perfect_copy.xlsx', 'unencrypted-rubyxl.xlsx')
puts "   Identical: #{identical ? '✅' : '❌'}"

puts "\n   Testing perfect copy..."
test_excel_encryption('perfect_copy.xlsx')

puts "\n" + "=" * 60
puts "🎯 ANALYSIS: If the perfect copy works, the issue is in our ZIP reconstruction logic"
puts "            If it fails, there might be a deeper issue"
