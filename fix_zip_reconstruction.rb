#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'

def extract_files_with_metadata(xlsx_file, output_dir)
  FileUtils.rm_rf(output_dir) if Dir.exist?(output_dir)
  FileUtils.mkdir_p(output_dir)
  
  file_metadata = {}
  Zip::File.open(xlsx_file) do |zip|
    zip.each do |entry|
      next if entry.directory?
      
      file_path = File.join(output_dir, entry.name)
      FileUtils.mkdir_p(File.dirname(file_path))
      entry.extract(file_path) { true }
      
      # Store metadata
      file_metadata[entry.name] = {
        method: entry.compression_method,
        level: entry.compression_level,
        time: entry.time,
        comment: entry.comment
      }
    end
  end
  file_metadata
end

def create_xlsx_preserving_structure(source_dir, output_file, original_xlsx, file_metadata = {})
  File.delete(output_file) if File.exist?(output_file)
  
  # Get the original file order from the source XLSX
  original_order = []
  Zip::File.open(original_xlsx) do |zip|
    zip.each do |entry|
      next if entry.directory?
      original_order << entry.name
    end
  end
  
  Zip::File.open(output_file, create: true) do |zip|
    # Add files in the same order as original
    original_order.each do |file_name|
      file_path = File.join(source_dir, file_name)
      next unless File.exist?(file_path)
      
      # Use original metadata if available
      metadata = file_metadata[file_name] || {}
      
      zip.add(file_name, file_path) do |src, dest|
        # Try to preserve compression settings
        if metadata[:method]
          dest.compression_method = metadata[:method] 
        end
        if metadata[:time]
          dest.time = metadata[:time]
        end
        true
      end
    end
  end
end

def files_identical?(file1, file2)
  return false unless File.exist?(file1) && File.exist?(file2)
  
  hash1 = Digest::MD5.hexdigest(File.read(file1))
  hash2 = Digest::MD5.hexdigest(File.read(file2))
  puts "Original hash: #{hash1}"
  puts "Fixed hash:    #{hash2}"
  hash1 == hash2
end

puts "🔧 FIXING ZIP RECONSTRUCTION"
puts "="*50

puts "\n1️⃣ Extracting with metadata preservation..."
file_metadata = extract_files_with_metadata('unencrypted-rubyxl.xlsx', 'rubyxl_extracted')
puts "   Extracted #{file_metadata.keys.length} files with metadata"

puts "\n2️⃣ Reconstructing with structure preservation..."
create_xlsx_preserving_structure('rubyxl_extracted', 'fixed_reconstructed.xlsx', 'unencrypted-rubyxl.xlsx', file_metadata)

puts "\n3️⃣ Testing fixed reconstruction..."
identical = files_identical?('unencrypted-rubyxl.xlsx', 'fixed_reconstructed.xlsx')

if identical
  puts "✅ SUCCESS: Fixed ZIP reconstruction works!"
  puts "   Now we can properly test file replacements"
else
  puts "❌ Still not identical, but let's check if it's closer..."
  
  size1 = File.size('unencrypted-rubyxl.xlsx')
  size2 = File.size('fixed_reconstructed.xlsx')
  puts "   Original: #{size1} bytes"
  puts "   Fixed: #{size2} bytes"
  puts "   Difference: #{size2 - size1} bytes (was -328)"
end

# Cleanup
FileUtils.rm_rf('rubyxl_extracted')
File.delete('fixed_reconstructed.xlsx') if File.exist?('fixed_reconstructed.xlsx')
