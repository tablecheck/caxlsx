#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'

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

  hash1 = Digest::MD5.hexdigest(File.read(file1))
  hash2 = Digest::MD5.hexdigest(File.read(file2))
  puts "Original RubyXL hash: #{hash1}"
  puts "Reconstructed hash:   #{hash2}"
  hash1 == hash2
end

puts "🧪 ZIP RECONSTRUCTION TEST"
puts "=" * 50

puts "\n1️⃣ Extracting RubyXL files..."
files = extract_all_files('unencrypted-rubyxl.xlsx', 'rubyxl_extracted')
puts "   Extracted #{files.length} files"

puts "\n2️⃣ Reconstructing XLSX from extracted files..."
create_xlsx_from_files('rubyxl_extracted', 'reconstructed_rubyxl.xlsx')

puts "\n3️⃣ Comparing original vs reconstructed..."
identical = files_identical?('unencrypted-rubyxl.xlsx', 'reconstructed_rubyxl.xlsx')

if identical
  puts "✅ SUCCESS: ZIP reconstruction works perfectly!"
  puts "   The issue is in our file replacement logic, not ZIP creation"
else
  puts "❌ PROBLEM: ZIP reconstruction is not identical!"
  puts "   This explains why our file replacements don't work"
  puts "   Need to fix ZIP creation logic first"
end

puts "\n4️⃣ File size comparison:"
size1 = File.size('unencrypted-rubyxl.xlsx')
size2 = File.size('reconstructed_rubyxl.xlsx')
puts "   Original: #{size1} bytes"
puts "   Reconstructed: #{size2} bytes"
puts "   Difference: #{size2 - size1} bytes"

# Cleanup
FileUtils.rm_rf('rubyxl_extracted')
File.delete('reconstructed_rubyxl.xlsx') if File.exist?('reconstructed_rubyxl.xlsx')
