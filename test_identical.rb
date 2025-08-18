#!/usr/bin/env ruby
require 'ooxml_crypt'
require 'win32ole'

def test_excel_compatibility(file_path, password = nil)
  excel = WIN32OLE.new('Excel.Application')
  excel.visible = false
  excel.displayAlerts = false

  workbook = if password
               excel.Workbooks.Open(File.absolute_path(file_path), nil, nil, nil, password)
             else
               excel.Workbooks.Open(File.absolute_path(file_path))
             end

  workbook.Close(false)
  excel.Quit
  true
rescue StandardError => e
  puts "Excel error for #{file_path}: #{e.message}"
  begin
    excel&.Quit
  rescue StandardError
  end
  false
end

puts "🧪 Testing truly identical RubyXL file..."
puts "📁 File: truly-identical.xlsx"

# Test unencrypted version first
puts "\n1️⃣ Testing unencrypted file..."
unencrypted_works = test_excel_compatibility('truly-identical.xlsx')
puts "✅ Unencrypted opens in Excel: #{unencrypted_works}"

# Now try to encrypt it
puts "\n2️⃣ Attempting to encrypt..."
begin
  OoxmlCrypt.encrypt_file('truly-identical.xlsx', 'test123', 'truly-identical-encrypted.xlsx')
  puts "✅ Encryption: SUCCESS"

  # Test if encrypted file opens in Excel
  puts "\n3️⃣ Testing encrypted file in Excel..."
  encrypted_works = test_excel_compatibility('truly-identical-encrypted.xlsx', 'test123')
  puts "📊 Encrypted file opens in Excel: #{encrypted_works}"

  if encrypted_works
    puts "\n🎉 BREAKTHROUGH: Truly identical RubyXL file WORKS with encryption!"
    puts "   This confirms that RubyXL's complete file reconstruction IS the solution."
  else
    puts "\n😱 MYSTERY: Even the truly identical file fails encryption!"
    puts "   This suggests something even deeper is at play."
  end
rescue StandardError => e
  puts "💥 Encryption FAILED: #{e.message}"
  puts "   Even the identical RubyXL file cannot be encrypted!"
end
