#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'

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
    # Add files in the original order
    file_order.each do |file_name|
      file_path = File.join(source_dir, file_name)
      if File.exist?(file_path)
        zip.add(file_name, file_path)
      end
    end
  end
end

def compare_zip_details(file1, file2)
  puts "\n=== DETAILED ZIP COMPARISON ==="

  entries1 = {}
  entries2 = {}

  Zip::File.open(file1) do |zip|
    zip.each do |entry|
      next if entry.directory?

      entries1[entry.name] = {
        size: entry.size,
        compressed_size: entry.compressed_size,
        crc: entry.crc,
        method: entry.compression_method
      }
    end
  end

  Zip::File.open(file2) do |zip|
    zip.each do |entry|
      next if entry.directory?

      entries2[entry.name] = {
        size: entry.size,
        compressed_size: entry.compressed_size,
        crc: entry.crc,
        method: entry.compression_method
      }
    end
  end

  puts "File 1 entries: #{entries1.keys.length}"
  puts "File 2 entries: #{entries2.keys.length}"

  entries1.each do |name, info1|
    info2 = entries2[name]
    if info2.nil?
      puts "❌ #{name}: MISSING in file2"
    elsif info1[:crc] != info2[:crc]
      puts "❌ #{name}: CRC differs (#{info1[:crc]} vs #{info2[:crc]})"
      puts "   Size: #{info1[:size]} vs #{info2[:size]}"
      puts "   Compressed: #{info1[:compressed_size]} vs #{info2[:compressed_size]}"
      puts "   Method: #{info1[:method]} vs #{info2[:method]}"
    else
      puts "✅ #{name}: identical"
    end
  end
end

puts "🔧 SIMPLE ZIP RECONSTRUCTION TEST"
puts "=" * 50

puts "\n1️⃣ Extracting files..."
file_order = extract_files_simple('unencrypted-rubyxl.xlsx', 'test_extracted')
puts "   Extracted #{file_order.length} files in order: #{file_order}"

puts "\n2️⃣ Reconstructing XLSX..."
create_xlsx_simple('test_extracted', 'test_reconstructed.xlsx', file_order)

puts "\n3️⃣ Comparing files..."
hash1 = Digest::MD5.hexdigest(File.read('unencrypted-rubyxl.xlsx'))
hash2 = Digest::MD5.hexdigest(File.read('test_reconstructed.xlsx'))

puts "Original hash:      #{hash1}"
puts "Reconstructed hash: #{hash2}"

size1 = File.size('unencrypted-rubyxl.xlsx')
size2 = File.size('test_reconstructed.xlsx')
puts "Original size:      #{size1} bytes"
puts "Reconstructed size: #{size2} bytes"

if hash1 == hash2
  puts "✅ SUCCESS: Perfect reconstruction!"
else
  puts "❌ DIFFERENT: Let's see what's wrong..."
  compare_zip_details('unencrypted-rubyxl.xlsx', 'test_reconstructed.xlsx')
end

# Cleanup
FileUtils.rm_rf('test_extracted')
File.delete('test_reconstructed.xlsx') if File.exist?('test_reconstructed.xlsx')
