#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'
require 'ooxml_crypt'
require 'win32ole'

class ReverseEngineer
  def initialize(caxlsx_file, rubyxl_file)
    @caxlsx_file = caxlsx_file
    @rubyxl_file = rubyxl_file
    @different_files = []
  end

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

  def apply_replacements(base_dir, files_to_replace)
    files_to_replace.each do |file_info|
      case file_info[:type]
      when :missing_in_caxlsx
        source_path = "rubyxl_extracted/#{file_info[:name]}"
        dest_path = "#{base_dir}/#{file_info[:name]}"
        FileUtils.mkdir_p(File.dirname(dest_path))
        FileUtils.copy(source_path, dest_path)

      when :different
        source_path = "rubyxl_extracted/#{file_info[:name]}"
        dest_path = "#{base_dir}/#{file_info[:name]}"
        FileUtils.copy(source_path, dest_path)
      end
    end
  end

  def find_minimal_changes
    puts "🔧 REVERSE ENGINEERING: Finding minimal changes for encryption compatibility"
    puts "=" * 80

    # Extract both files
    puts "\n📂 Extracting files..."
    caxlsx_files = extract_all_files(@caxlsx_file, 'caxlsx_extracted')
    rubyxl_files = extract_all_files(@rubyxl_file, 'rubyxl_extracted')

    # Find different files
    all_files = (caxlsx_files + rubyxl_files).uniq.sort

    all_files.each do |file_name|
      caxlsx_file_path = "caxlsx_extracted/#{file_name}"
      rubyxl_file_path = "rubyxl_extracted/#{file_name}"

      if !File.exist?(caxlsx_file_path)
        @different_files << { name: file_name, type: :missing_in_caxlsx }
      elsif !File.exist?(rubyxl_file_path)
        @different_files << { name: file_name, type: :missing_in_rubyxl }
      elsif !files_identical?(caxlsx_file_path, rubyxl_file_path)
        @different_files << { name: file_name, type: :different }
      end
    end

    puts "\n📊 Found #{@different_files.length} files that differ:"
    @different_files.each_with_index do |file_info, i|
      status = case file_info[:type]
               when :missing_in_caxlsx then "❌ MISSING"
               when :different then "🔄 DIFFERENT"
               end
      puts "   #{i + 1}. #{file_info[:name]}: #{status}"
    end

    # Step 1: Test with ALL replacements (should work - creates identical file)
    puts "\n🧪 STEP 1: Testing with ALL #{@different_files.length} file replacements..."
    FileUtils.rm_rf('working_extracted') if Dir.exist?('working_extracted')
    FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

    apply_replacements('working_extracted', @different_files)
    create_xlsx_from_files('working_extracted', 'all_replacements.xlsx')

    all_work = test_excel_encryption('all_replacements.xlsx')
    is_identical = files_identical?('all_replacements.xlsx', @rubyxl_file)

    puts "   Encryption works: #{all_work ? '✅' : '❌'}"
    puts "   Byte-identical to RubyXL: #{is_identical ? '✅' : '❌'}"

    unless all_work
      puts "\n💥 ERROR: Even with ALL replacements, encryption doesn't work!"
      puts "   This suggests a problem with our replacement logic."
      cleanup
      return nil
    end

    puts "\n✅ Confirmed: All replacements work. Now finding minimal set..."

    # Step 2: Remove files one at a time to find minimal set
    current_replacements = @different_files.dup
    essential_files = []

    puts "\n🔬 STEP 2: Removing files one at a time to find minimal set..."

    @different_files.each_with_index do |file_to_remove, index|
      puts "\n#{index + 1}/#{@different_files.length}: Testing without #{file_to_remove[:name]}"

      # Create test set without this file
      test_replacements = current_replacements.reject { |f| f[:name] == file_to_remove[:name] }

      # Apply replacements without this file
      FileUtils.rm_rf('working_extracted') if Dir.exist?('working_extracted')
      FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

      apply_replacements('working_extracted', test_replacements)

      test_file = "test_without_#{index}.xlsx"
      create_xlsx_from_files('working_extracted', test_file)

      encryption_works = test_excel_encryption(test_file)
      puts "   Without #{file_to_remove[:name]}: #{encryption_works ? '✅ Still works' : '❌ Breaks encryption'}"

      if encryption_works
        # This file is not needed
        current_replacements = test_replacements
        puts "   ⚪ #{file_to_remove[:name]} is not needed"
      else
        # This file is essential
        essential_files << file_to_remove
        puts "   🔴 #{file_to_remove[:name]} is ESSENTIAL for encryption!"
      end

      File.delete(test_file) if File.exist?(test_file)
    end

    puts "\n" + "=" * 80
    puts "🎯 RESULTS: Minimal changes needed for encryption compatibility"
    puts "=" * 80

    if essential_files.empty?
      puts "🤯 SHOCKING: No files are essential! Encryption should work without any changes!"
    else
      puts "✅ Found #{essential_files.length} essential file(s):"
      essential_files.each_with_index do |file_info, i|
        action = case file_info[:type]
                 when :missing_in_caxlsx then "ADD"
                 when :different then "REPLACE"
                 end
        puts "   #{i + 1}. #{action}: #{file_info[:name]}"
      end

      # Create the minimal fix
      puts "\n🔧 Creating minimal fix file..."
      FileUtils.rm_rf('working_extracted') if Dir.exist?('working_extracted')
      FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

      apply_replacements('working_extracted', essential_files)
      create_xlsx_from_files('working_extracted', 'minimal_fix.xlsx')

      final_test = test_excel_encryption('minimal_fix.xlsx')
      puts "   Final test: #{final_test ? '✅' : '❌'}"

      if final_test
        puts "\n🏆 SUCCESS! Created minimal fix with only #{essential_files.length} file changes."
        puts "💾 Saved as: minimal_fix.xlsx"
      end
    end

    cleanup
    essential_files
  end

  def cleanup
    FileUtils.rm_rf('caxlsx_extracted')
    FileUtils.rm_rf('rubyxl_extracted')
    FileUtils.rm_rf('working_extracted')
    File.delete('all_replacements.xlsx') if File.exist?('all_replacements.xlsx')
  end
end

# Run the reverse engineering
if ARGV.length != 2
  puts "Usage: ruby #{$0} <caxlsx_file> <rubyxl_file>"
  exit 1
end

engineer = ReverseEngineer.new(ARGV[0], ARGV[1])
result = engineer.find_minimal_changes

if result && !result.empty?
  puts "\n🎉 Found minimal solution!"
else
  puts "\n😞 Could not find minimal changes."
end
