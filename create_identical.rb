#!/usr/bin/env ruby
require 'fileutils'
require 'zip'
require 'digest'

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

# Create the identical file
create_truly_identical_xlsx('base-caxlsx.xlsx', 'truly-identical.xlsx')

# Compare hashes to verify
puts "\n=== VERIFICATION ==="
hash1 = Digest::SHA256.hexdigest(File.read('unencrypted-rubyxl.xlsx'))
hash2 = Digest::SHA256.hexdigest(File.read('truly-identical.xlsx'))
puts "RubyXL file hash:      #{hash1}"
puts "Identical file hash:   #{hash2}"
puts "Files are identical: #{hash1 == hash2 ? '✅ YES' : '❌ NO'}"
