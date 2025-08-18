#!/usr/bin/env ruby
require 'zip'
require 'digest'

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
  puts "MD5 hashes: #{hash1 == hash2 ? 'IDENTICAL' : 'DIFFERENT'}"
  puts "  #{File.basename(file1)}: #{hash1}"
  puts "  #{File.basename(file2)}: #{hash2}"
  
  # ZIP structure comparison
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
  
  different_files = []
  
  all_entries.each do |entry_name|
    entry1 = entries1[entry_name]
    entry2 = entries2[entry_name]
    
    if entry1.nil?
      puts "  #{entry_name}: MISSING in #{File.basename(file1)}"
      different_files << entry_name
    elsif entry2.nil?
      puts "  #{entry_name}: MISSING in #{File.basename(file2)}"
      different_files << entry_name
    elsif entry1[:crc] != entry2[:crc]
      puts "  #{entry_name}: DIFFERENT (CRC: #{entry1[:crc]} vs #{entry2[:crc]})"
      puts "    Sizes: #{entry1[:size]} vs #{entry2[:size]} bytes"
      different_files << entry_name
    else
      puts "  #{entry_name}: identical"
    end
  end
  
  puts "\n=== DIFFERENCES SUMMARY ==="
  if different_files.empty?
    puts "✅ FILES ARE IDENTICAL!"
  else
    puts "❌ #{different_files.length} files differ: #{different_files.join(', ')}"
  end
end

compare_xlsx_files('caxlsx-full-treatment.xlsx', 'unencrypted-rubyxl.xlsx')
