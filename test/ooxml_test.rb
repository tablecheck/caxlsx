require_relative 'tc_helper.rb'
require 'ooxml_crypt'
require 'win32ole'
require 'zip'
require 'digest'
require 'fileutils'
require 'stringio'

class OoxmlTest < Test::Unit::TestCase
  def teardown
    # Skip cleanup to allow file inspection
    # files_to_cleanup = [
    #   'unencrypted-caxlsx.xlsx',
    #   'encrypted-caxlsx.xlsx',
    #   'unencrypted-rubyxl.xlsx',
    #   'encrypted-rubyxl.xlsx'
    # ]

    # files_to_cleanup.each do |file|
    #   File.delete(file) if File.exist?(file)
    # end
  end

  def compare_xlsx_files(file1, file2)
    puts "\n=== XLSX File Comparison: #{File.basename(file1)} vs #{File.basename(file2)} ==="

    # Basic file comparison
    size1 = File.size(file1)
    size2 = File.size(file2)
    puts "File sizes: #{size1} bytes vs #{size2} bytes"
    puts "Size difference: #{size2 - size1} bytes" if size1 != size2

    # Checksums
    hash1 = Digest::MD5.hexdigest(File.read(file1))
    hash2 = Digest::MD5.hexdigest(File.read(file2))
    puts "MD5 hashes: #{file1 == file2 ? 'IDENTICAL' : 'DIFFERENT'}"
    puts "  #{File.basename(file1)}: #{hash1}"
    puts "  #{File.basename(file2)}: #{hash2}"

    # ZIP structure comparison
    begin
      entries1 = {}
      entries2 = {}

      Zip::File.open(file1) do |zip|
        zip.each do |entry|
          entries1[entry.name] = {
            size: entry.size,
            compressed_size: entry.compressed_size,
            crc: entry.crc
          }
        end
      end

      Zip::File.open(file2) do |zip|
        zip.each do |entry|
          entries2[entry.name] = {
            size: entry.size,
            compressed_size: entry.compressed_size,
            crc: entry.crc
          }
        end
      end

      puts "\nZIP entries comparison:"
      all_entries = (entries1.keys + entries2.keys).uniq.sort

      all_entries.each do |entry_name|
        entry1 = entries1[entry_name]
        entry2 = entries2[entry_name]

        if entry1.nil?
          puts "  #{entry_name}: MISSING in #{File.basename(file1)}"
        elsif entry2.nil?
          puts "  #{entry_name}: MISSING in #{File.basename(file2)}"
        elsif entry1[:crc] != entry2[:crc]
          puts "  #{entry_name}: DIFFERENT (CRC: #{entry1[:crc]} vs #{entry2[:crc]})"
          puts "    Sizes: #{entry1[:size]} vs #{entry2[:size]} bytes"
        else
          puts "  #{entry_name}: identical"
        end
      end

      # Extract and compare key XML files for detailed analysis
      key_files = ['[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']
      key_files.each do |key_file|
        next unless entries1[key_file] && entries2[key_file] && entries1[key_file][:crc] != entries2[key_file][:crc]

        puts "\n--- Detailed comparison of #{key_file} ---"

        content1 = nil
        content2 = nil

        Zip::File.open(file1) do |zip|
          entry = zip.find_entry(key_file)
          content1 = entry.get_input_stream.read if entry
        end

        Zip::File.open(file2) do |zip|
          entry = zip.find_entry(key_file)
          content2 = entry.get_input_stream.read if entry
        end

        next unless content1 && content2

        puts "Length: #{content1.length} vs #{content2.length} chars"
        # Show first few lines of difference
        lines1 = content1.split("\n")[0..5]
        lines2 = content2.split("\n")[0..5]
        puts "#{File.basename(file1)} content (first 6 lines):"
        lines1.each_with_index { |line, i| puts "  #{i + 1}: #{line[0..100]}" }
        puts "#{File.basename(file2)} content (first 6 lines):"
        lines2.each_with_index { |line, i| puts "  #{i + 1}: #{line[0..100]}" }
      end
    rescue StandardError => e
      puts "Error comparing ZIP structures: #{e.message}"
    end
  end

  def extract_key_xlsx_files(xlsx_file, prefix)
    puts "\n=== Extracting key files from #{xlsx_file} ==="

    key_files = [
      '[Content_Types].xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
      'xl/theme/theme1.xml'
    ]

    Zip::File.open(xlsx_file) do |zip|
      key_files.each do |file_path|
        entry = zip.find_entry(file_path)
        if entry
          output_filename = "#{prefix}_#{file_path.gsub('/', '_')}"
          puts "Extracting #{file_path} → #{output_filename}"
          entry.extract(output_filename) { true }  # overwrite if exists
        else
          puts "❌ MISSING: #{file_path}"
        end
      end
    end
  rescue StandardError => e
    puts "Error extracting files from #{xlsx_file}: #{e.message}"
  end

  def create_truly_identical_xlsx(source_file, output_file)
    puts "\n=== Creating TRULY identical XLSX by copying ALL RubyXL files ==="
    
    # Start with source file as base
    FileUtils.copy(source_file, output_file)
    
    # Replace ALL content with content from RubyXL version
    Zip::File.open(output_file, create: false) do |target_zip|
      Zip::File.open('unencrypted-rubyxl.xlsx') do |source_zip|
        source_zip.each do |source_entry|
          puts "  📋 Copying #{source_entry.name} from RubyXL version"
          
          # Remove existing entry if present
          target_zip.remove(source_entry.name) if target_zip.find_entry(source_entry.name)
          
          # Add the content from RubyXL file
          target_zip.add(source_entry.name, source_entry.get_input_stream) { true }
        end
      end
    end
    
    puts "  ✅ Created truly identical file"
  rescue StandardError => e
    puts "Error creating identical XLSX: #{e.message}"
    puts e.backtrace[0..3].join("\n")
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

    # COMPARE THE TWO UNENCRYPTED FILES TO SEE WHAT RUBYXL CHANGED
    compare_xlsx_files('unencrypted-caxlsx.xlsx', 'unencrypted-rubyxl.xlsx')

    # Extract key files for detailed inspection
    extract_key_xlsx_files('unencrypted-caxlsx.xlsx', 'caxlsx')
    extract_key_xlsx_files('unencrypted-rubyxl.xlsx', 'rubyxl')

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

  def test_targeted_encryption_fixes
    puts "\n" + "="*80
    puts "🧪 TARGETED ENCRYPTION COMPATIBILITY TESTING"
    puts "="*80

    # First create the base caxlsx file
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_worksheet(name: 'Test Sheet') do |sheet|
      sheet.add_row ['Test', 'Data']
      sheet.add_row [1, 2]
    end
    p.serialize('base-caxlsx.xlsx')

    # Test 1: Original caxlsx file (baseline - should fail)
    puts "\n🔬 Test 1: Original caxlsx file (baseline)"
    test_encryption_compatibility('base-caxlsx.xlsx', 'original')

    # Test 2: Add only theme file
    puts "\n🔬 Test 2: Add ONLY theme file"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-with-theme.xlsx', {
      theme: { action: :add_theme }
    })
    test_encryption_compatibility('caxlsx-with-theme.xlsx', 'theme-only')

    # Test 3: Add theme + update content types
    puts "\n🔬 Test 3: Add theme + update content types"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-theme-contenttypes.xlsx', {
      theme: { action: :add_theme },
      content_types: { action: :update_content_types }
    })
    test_encryption_compatibility('caxlsx-theme-contenttypes.xlsx', 'theme+contenttypes')

    # Test 4: Add only standalone="yes" to XML declarations
    puts "\n🔬 Test 4: Add ONLY standalone='yes' declarations"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-standalone.xlsx', {
      '[Content_Types].xml' => { action: :add_standalone },
      'xl/workbook.xml' => { action: :add_standalone },
      'xl/worksheets/sheet1.xml' => { action: :add_standalone }
    })
    test_encryption_compatibility('caxlsx-standalone.xlsx', 'standalone-only')

    # Test 5: Add only missing namespaces
    puts "\n🔬 Test 5: Add ONLY missing XML namespaces"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-namespaces.xlsx', {
      namespaces: { action: :add_namespaces }
    })
    test_encryption_compatibility('caxlsx-namespaces.xlsx', 'namespaces-only')

    # Test 6: Combination - theme + content types + standalone
    puts "\n🔬 Test 6: Theme + content types + standalone declarations"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-theme-standalone.xlsx', {
      theme: { action: :add_theme },
      content_types: { action: :update_content_types },
      '[Content_Types].xml' => { action: :add_standalone },
      'xl/workbook.xml' => { action: :add_standalone },
      'xl/worksheets/sheet1.xml' => { action: :add_standalone }
    })
    test_encryption_compatibility('caxlsx-theme-standalone.xlsx', 'theme+standalone')

    # Test 7: Full RubyXL-like treatment
    puts "\n🔬 Test 7: ALL modifications (full RubyXL-like treatment)"
    create_modified_xlsx('base-caxlsx.xlsx', 'caxlsx-full-treatment.xlsx', {
      theme: { action: :add_theme },
      content_types: { action: :update_content_types },
      '[Content_Types].xml' => { action: :add_standalone },
      'xl/workbook.xml' => { action: :add_standalone },
      'xl/worksheets/sheet1.xml' => { action: :add_standalone },
      namespaces: { action: :add_namespaces }
    })
    test_encryption_compatibility('caxlsx-full-treatment.xlsx', 'full-treatment')

    puts "\n" + "="*80
    puts "🎯 ENCRYPTION COMPATIBILITY TEST RESULTS SUMMARY"
    puts "="*80
  end

  private

  def test_encryption_compatibility(file_path, test_name)
    puts "  📁 Testing: #{file_path}"
    
    # First verify the unencrypted file opens in Excel
    unencrypted_works = check_excel_file_compatibility(file_path)
    puts "  📊 Unencrypted opens in Excel: #{unencrypted_works}"
    
    # Try to encrypt the file
    encrypted_file = file_path.gsub('.xlsx', '-encrypted.xlsx')
    begin
      OoxmlCrypt.encrypt_file(file_path, 'test123', encrypted_file)
      puts "  🔐 Encryption: SUCCESS"
      
      # Test if encrypted file opens in Excel
      encrypted_works = check_excel_file_compatibility(encrypted_file, 'test123')
      puts "  📊 Encrypted opens in Excel: #{encrypted_works}"
      
      if encrypted_works
        puts "  ✅ #{test_name.upcase}: ENCRYPTION COMPATIBILITY FIXED!"
      else
        puts "  ❌ #{test_name}: Encrypted file still fails in Excel"
      end
      
    rescue => e
      puts "  💥 Encryption: FAILED (#{e.message})"
      puts "  ❌ #{test_name}: Cannot encrypt file"
    end
    
    puts ""
  end
end
