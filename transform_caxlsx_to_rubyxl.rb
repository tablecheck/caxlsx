#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'
require 'ooxml_crypt'
require 'win32ole'

class CaxlsxToRubyXLTransformer
  def initialize(caxlsx_file, rubyxl_file)
    @caxlsx_file = caxlsx_file
    @rubyxl_file = rubyxl_file
    @working_file = 'transform_working.xlsx'
  end

  def test_excel_encryption(file_path)
    # Test if file can be encrypted and opened in Excel
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
      true
    rescue StandardError => e
      begin
        excel&.Quit
      rescue StandardError
      end
      File.delete(test_file) if File.exist?(test_file)
      false
    end
  end

  def extract_all_files(xlsx_file, output_dir)
    FileUtils.rm_rf(output_dir) if Dir.exist?(output_dir)
    FileUtils.mkdir_p(output_dir)

    files = []
    Zip::File.open(xlsx_file) do |zip|
      zip.each do |entry|
        next if entry.directory?

        file_path = File.join(output_dir, entry.name)
        FileUtils.mkdir_p(File.dirname(file_path))
        entry.extract(file_path) { true }
        files << entry.name
      end
    end
    files.sort
  end

  def create_xlsx_from_files(source_dir, output_file)
    File.delete(output_file) if File.exist?(output_file)

    Zip::File.open(output_file, create: true) do |zip|
      Dir.glob("#{source_dir}/**/*").each do |file_path|
        next if File.directory?(file_path)

        relative_path = file_path.gsub("#{source_dir}/", "").gsub("#{source_dir}\\", "").gsub('\\', '/')
        zip.add(relative_path, file_path)
      end
    end
  end

  def files_identical?(file1, file2)
    return false unless File.exist?(file1) && File.exist?(file2)

    Digest::MD5.hexdigest(File.read(file1)) == Digest::MD5.hexdigest(File.read(file2))
  end

  def transform_step_by_step
    puts "🔧 SYSTEMATIC TRANSFORMATION: caxlsx → RubyXL-compatible"
    puts "=" * 80

    # Extract both files
    puts "\n📂 Extracting files..."
    caxlsx_files = extract_all_files(@caxlsx_file, 'caxlsx_extracted')
    rubyxl_files = extract_all_files(@rubyxl_file, 'rubyxl_extracted')

    # Start with original caxlsx file
    FileUtils.copy(@caxlsx_file, @working_file)

    puts "\n🧪 Testing original caxlsx file..."
    original_works = test_excel_encryption(@working_file)
    puts "   Encryption works: #{original_works ? '✅' : '❌'}"

    # Find all files that differ
    all_files = (caxlsx_files + rubyxl_files).uniq.sort
    different_files = []

    all_files.each do |file_name|
      caxlsx_file_path = "caxlsx_extracted/#{file_name}"
      rubyxl_file_path = "rubyxl_extracted/#{file_name}"

      if !File.exist?(caxlsx_file_path)
        different_files << { name: file_name, type: :missing_in_caxlsx }
      elsif !File.exist?(rubyxl_file_path)
        different_files << { name: file_name, type: :missing_in_rubyxl }
      elsif !files_identical?(caxlsx_file_path, rubyxl_file_path)
        different_files << { name: file_name, type: :different }
      end
    end

    puts "\n📊 Found #{different_files.length} files that differ:"
    different_files.each do |file_info|
      status = case file_info[:type]
               when :missing_in_caxlsx then "❌ MISSING in caxlsx"
               when :missing_in_rubyxl then "➕ EXTRA in caxlsx"
               when :different then "🔄 DIFFERENT"
               end
      puts "   #{file_info[:name]}: #{status}"
    end

    # Test replacing each file one by one
    puts "\n🔬 Testing systematic replacement..."
    successful_changes = []

    different_files.each_with_index do |file_info, index|
      puts "\n#{index + 1}/#{different_files.length}: Testing #{file_info[:name]}"

      # Create working directory with current state
      FileUtils.rm_rf('working_extracted') if Dir.exist?('working_extracted')
      extract_all_files(@working_file, 'working_extracted')

      case file_info[:type]
      when :missing_in_caxlsx
        # Add missing file from RubyXL
        source_path = "rubyxl_extracted/#{file_info[:name]}"
        dest_path = "working_extracted/#{file_info[:name]}"
        FileUtils.mkdir_p(File.dirname(dest_path))
        FileUtils.copy(source_path, dest_path)
        puts "   ➕ Added #{file_info[:name]} from RubyXL"

      when :missing_in_rubyxl
        # Remove extra file from caxlsx
        file_path = "working_extracted/#{file_info[:name]}"
        File.delete(file_path) if File.exist?(file_path)
        puts "   ➖ Removed #{file_info[:name]} (not in RubyXL)"

      when :different
        # Replace with RubyXL version
        source_path = "rubyxl_extracted/#{file_info[:name]}"
        dest_path = "working_extracted/#{file_info[:name]}"
        FileUtils.copy(source_path, dest_path)
        puts "   🔄 Replaced #{file_info[:name]} with RubyXL version"
      end

      # Rebuild XLSX and test
      test_file = "test_step_#{index}.xlsx"
      create_xlsx_from_files('working_extracted', test_file)

      encryption_works = test_excel_encryption(test_file)
      puts "   🧪 Encryption works: #{encryption_works ? '✅' : '❌'}"

      if encryption_works
        successful_changes << file_info
        FileUtils.copy(test_file, @working_file)
        puts "   ✅ BREAKTHROUGH! This change fixes encryption!"

        # Check if we're now byte-identical to RubyXL
        if files_identical?(@working_file, @rubyxl_file)
          puts "   🎉 FILE IS NOW BYTE-IDENTICAL TO RUBYXL!"
          break
        end
      end

      File.delete(test_file) if File.exist?(test_file)
    end

    puts "\n" + "=" * 80
    puts "🎯 TRANSFORMATION RESULTS:"

    if successful_changes.empty?
      puts "❌ No single file replacement fixed the encryption issue"
      puts "   The problem may require multiple simultaneous changes"
    else
      puts "✅ Found #{successful_changes.length} changes that enable encryption:"
      successful_changes.each_with_index do |change, i|
        puts "   #{i + 1}. #{change[:name]} (#{change[:type]})"
      end

      final_works = test_excel_encryption(@working_file)
      final_identical = files_identical?(@working_file, @rubyxl_file)

      puts "\nFinal status:"
      puts "   Encryption works: #{final_works ? '✅' : '❌'}"
      puts "   Byte-identical to RubyXL: #{final_identical ? '✅' : '❌'}"

      if final_works && final_identical
        puts "\n🏆 SUCCESS! Created transformation that makes caxlsx encryption-compatible!"
        FileUtils.copy(@working_file, 'fixed_caxlsx.xlsx')
        puts "   Saved result as: fixed_caxlsx.xlsx"
      end
    end

    # Cleanup
    FileUtils.rm_rf('caxlsx_extracted')
    FileUtils.rm_rf('rubyxl_extracted')
    FileUtils.rm_rf('working_extracted')
    File.delete(@working_file) if File.exist?(@working_file)
  end
end

# Run the transformation
if ARGV.length != 2
  puts "Usage: ruby #{$0} <caxlsx_file> <rubyxl_file>"
  exit 1
end

transformer = CaxlsxToRubyXLTransformer.new(ARGV[0], ARGV[1])
transformer.transform_step_by_step
