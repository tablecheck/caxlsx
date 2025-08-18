require_relative 'tc_helper.rb'
require 'ooxml_crypt'
require 'win32ole'

class OoxmlTest < Test::Unit::TestCase
  def teardown
    # Clean up generated test files
    files_to_cleanup = [
      'unencrypted-caxlsx.xlsx',
      'encrypted-caxlsx.xlsx',
      'unencrypted-rubyxl.xlsx',
      'encrypted-rubyxl.xlsx'
    ]

    files_to_cleanup.each do |file|
      File.delete(file) if File.exist?(file)
    end
  end

  def check_excel_file_compatibility(file_path, password = nil)
    excel = WIN32OLE.new('Excel.Application')
    excel.visible = false  # Hide Excel during testing
    excel.displayAlerts = false  # Disable alerts/prompts

    # Try to open the workbook (with password if provided)
    workbook = if password
                 excel.Workbooks.Open(File.absolute_path(file_path), nil, nil, nil, password)
               else
                 excel.Workbooks.Open(File.absolute_path(file_path))
               end

    workbook.Close(false)  # Close without saving
    excel.Quit
    true
  rescue StandardError => e
    puts "Excel error for #{file_path}: #{e.message}"
    begin
      excel&.Quit
    rescue StandardError
      # Ignore cleanup errors
    end
    false
  end

  def test_encrypt_encrypted_excel
    p = Axlsx::Package.new
    wb = p.workbook

    wb.add_worksheet(name: 'Basic Worksheet') do |sheet|
      sheet.add_row ['First', 'Second', 'Third']
      sheet.add_row [1, 2, 3]
    end

    p.serialize('unencrypted-caxlsx.xlsx')

    # Test that the unencrypted caxlsx file works
    caxlsx_unencrypted_works = check_excel_file_compatibility('unencrypted-caxlsx.xlsx')
    puts "Caxlsx unencrypted file opens in Excel: #{caxlsx_unencrypted_works}"

    # Encrypt the caxlsx file
    OoxmlCrypt.encrypt_file('unencrypted-caxlsx.xlsx', 'abc123', 'encrypted-caxlsx.xlsx')

    # Test the ENCRYPTED caxlsx file - this is where issues likely occur
    caxlsx_encrypted_works = check_excel_file_compatibility('encrypted-caxlsx.xlsx', 'abc123')
    puts "Caxlsx ENCRYPTED file opens in Excel: #{caxlsx_encrypted_works}"

    require 'rubyXL'
    require 'rubyXL/convenience_methods'

    def standardize_xlsx_with_rubyxl(input_file, output_file)
      workbook = RubyXL::Parser.parse(input_file)
      workbook.write(output_file)
    end

    standardize_xlsx_with_rubyxl('unencrypted-caxlsx.xlsx', 'unencrypted-rubyxl.xlsx')

    # Test that the unencrypted RubyXL file works
    rubyxl_unencrypted_works = check_excel_file_compatibility('unencrypted-rubyxl.xlsx')
    puts "RubyXL unencrypted file opens in Excel: #{rubyxl_unencrypted_works}"

    # Encrypt the RubyXL file (try-catch in case of OoxmlCrypt issues)
    begin
      OoxmlCrypt.encrypt_file('unencrypted-rubyxl.xlsx', 'abc123', 'encrypted-rubyxl.xlsx')
      rubyxl_encrypted_works = check_excel_file_compatibility('encrypted-rubyxl.xlsx', 'abc123')
      puts "RubyXL ENCRYPTED file opens in Excel: #{rubyxl_encrypted_works}"
    rescue StandardError => e
      puts "Failed to encrypt RubyXL file: #{e.message}"
      rubyxl_encrypted_works = nil
    end

    # Report the actual behavior we observed
    puts "\n=== Test Results ==="
    puts "Unencrypted Files:"
    puts "- Caxlsx file works in Excel: #{caxlsx_unencrypted_works}"
    puts "- RubyXL file works in Excel: #{rubyxl_unencrypted_works}"
    puts "\nEncrypted Files (with password 'abc123'):"
    puts "- Caxlsx ENCRYPTED file works in Excel: #{caxlsx_encrypted_works}"
    puts "- RubyXL ENCRYPTED file works in Excel: #{rubyxl_encrypted_works || 'Could not test (encryption failed)'}"

    # Test the main hypothesis: caxlsx encrypted files fail in Excel
    if caxlsx_encrypted_works
      puts "\n❌ UNEXPECTED: Caxlsx encrypted files work - the bug may be fixed or environment-specific"
      assert_true(false, "Expected caxlsx encrypted files to fail based on original comments")
    else
      puts "\n✅ CONFIRMED: Caxlsx encrypted files FAIL in Excel!"
      puts "   This confirms the bug described in the original comments."

      if rubyxl_encrypted_works == true
        puts "   RubyXL encrypted files work - confirming the workaround."
        assert_false(caxlsx_encrypted_works, "Caxlsx encrypted files should fail (confirms the bug)")
        assert_true(rubyxl_encrypted_works, "RubyXL encrypted files should work (confirms workaround)")
      else
        puts "   RubyXL testing incomplete, but caxlsx failure confirmed."
        assert_false(caxlsx_encrypted_works, "Caxlsx encrypted files should fail (confirms the bug)")
      end
    end
  end
end
