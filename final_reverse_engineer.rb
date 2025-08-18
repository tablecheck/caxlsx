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

def extract_files_simple(xlsx_file, output_dir)
  FileUtils.rm_rf(output_dir) if Dir.exist?(output_dir)
  FileUtils.mkdir_p(output_dir)

  file_order = []
  Zip::File.open(xlsx_file) do |zip|
    zip.each do |entry|
      next if entry.directory?

      file_path = File.join(output_dir, entry.name)
      FileUtils.mkdir_p(File.dirname(file_path))
      entry.extract(file_path) { true }
      file_order << entry.name
    end
  end
  file_order
end

def create_xlsx_simple(source_dir, output_file, file_order)
  File.delete(output_file) if File.exist?(output_file)

  Zip::File.open(output_file, create: true) do |zip|
    file_order.each do |file_name|
      file_path = File.join(source_dir, file_name)
      if File.exist?(file_path)
        zip.add(file_name, file_path)
      end
    end
  end
end

def files_identical?(file1, file2)
  return false unless File.exist?(file1) && File.exist?(file2)

  Digest::MD5.hexdigest(File.read(file1)) == Digest::MD5.hexdigest(File.read(file2))
end

puts "🔧 FINAL REVERSE ENGINEERING: Finding minimal changes"
puts "=" * 60

# Extract both files
puts "\n📂 Extracting files..."
caxlsx_files = extract_files_simple('base-caxlsx.xlsx', 'caxlsx_extracted')
rubyxl_files = extract_files_simple('unencrypted-rubyxl.xlsx', 'rubyxl_extracted')

# Find different files
different_files = []
all_files = (caxlsx_files + rubyxl_files).uniq.sort

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

puts "📊 Found #{different_files.length} files that differ:"
different_files.each_with_index do |file_info, i|
  status = case file_info[:type]
           when :missing_in_caxlsx then "❌ MISSING"
           when :different then "🔄 DIFFERENT"
           end
  puts "   #{i + 1}. #{file_info[:name]}: #{status}"
end

# Step 1: Test with ALL replacements
puts "\n🧪 STEP 1: Testing with ALL file replacements..."
FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

different_files.each do |file_info|
  case file_info[:type]
  when :missing_in_caxlsx
    source_path = "rubyxl_extracted/#{file_info[:name]}"
    dest_path = "working_extracted/#{file_info[:name]}"
    FileUtils.mkdir_p(File.dirname(dest_path))
    FileUtils.copy(source_path, dest_path)
  when :different
    source_path = "rubyxl_extracted/#{file_info[:name]}"
    dest_path = "working_extracted/#{file_info[:name]}"
    FileUtils.copy(source_path, dest_path)
  end
end

create_xlsx_simple('working_extracted', 'all_replacements.xlsx', rubyxl_files)

all_work = test_excel_encryption('all_replacements.xlsx')
puts "   All replacements work: #{all_work ? '✅' : '❌'}"

unless all_work
  puts "\n💥 ERROR: Even with ALL replacements, encryption doesn't work!"
  exit 1
end

# Step 2: Remove files one at a time to find minimal set
puts "\n🔬 STEP 2: Finding minimal set by removing files one at a time..."

current_replacements = different_files.dup
essential_files = []

different_files.each_with_index do |file_to_remove, index|
  print "\r#{index + 1}/#{different_files.length}: Testing without #{file_to_remove[:name]}..."

  # Create test set without this file
  test_replacements = current_replacements.reject { |f| f[:name] == file_to_remove[:name] }

  # Apply replacements without this file
  FileUtils.rm_rf('working_extracted')
  FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

  test_replacements.each do |file_info|
    case file_info[:type]
    when :missing_in_caxlsx
      source_path = "rubyxl_extracted/#{file_info[:name]}"
      dest_path = "working_extracted/#{file_info[:name]}"
      FileUtils.mkdir_p(File.dirname(dest_path))
      FileUtils.copy(source_path, dest_path)
    when :different
      source_path = "rubyxl_extracted/#{file_info[:name]}"
      dest_path = "working_extracted/#{file_info[:name]}"
      FileUtils.copy(source_path, dest_path)
    end
  end

  test_file = "test_without_#{index}.xlsx"
  create_xlsx_simple('working_extracted', test_file, rubyxl_files)

  encryption_works = test_excel_encryption(test_file)

  if encryption_works
    # This file is not needed
    current_replacements = test_replacements
    puts " ⚪ Not needed"
  else
    # This file is essential
    essential_files << file_to_remove
    puts " 🔴 ESSENTIAL!"
  end

  File.delete(test_file) if File.exist?(test_file)
end

puts "\n" + "=" * 60
puts "🎯 MINIMAL CHANGES NEEDED FOR ENCRYPTION COMPATIBILITY:"

if essential_files.empty?
  puts "🤯 SHOCKING: No files are essential!"
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
  puts "\n🔧 Creating minimal fix..."
  FileUtils.rm_rf('working_extracted')
  FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

  essential_files.each do |file_info|
    case file_info[:type]
    when :missing_in_caxlsx
      source_path = "rubyxl_extracted/#{file_info[:name]}"
      dest_path = "working_extracted/#{file_info[:name]}"
      FileUtils.mkdir_p(File.dirname(dest_path))
      FileUtils.copy(source_path, dest_path)
    when :different
      source_path = "rubyxl_extracted/#{file_info[:name]}"
      dest_path = "working_extracted/#{file_info[:name]}"
      FileUtils.copy(source_path, dest_path)
    end
  end

  create_xlsx_simple('working_extracted', 'minimal_fix.xlsx', rubyxl_files)

  final_test = test_excel_encryption('minimal_fix.xlsx')
  if final_test
    puts "✅ Minimal fix verified!"
    puts "💾 Saved as: minimal_fix.xlsx"
  else
    puts "❌ Minimal fix failed - unexpected!"
  end
end

# Cleanup
FileUtils.rm_rf('caxlsx_extracted')
FileUtils.rm_rf('rubyxl_extracted')
FileUtils.rm_rf('working_extracted')
File.delete('all_replacements.xlsx') if File.exist?('all_replacements.xlsx')
