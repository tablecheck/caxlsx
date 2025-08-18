#!/usr/bin/env ruby
require 'zip'
require 'fileutils'
require 'digest'
require 'ooxml_crypt'
require 'win32ole'

class CombinationTransformer
  def initialize(caxlsx_file, rubyxl_file)
    @caxlsx_file = caxlsx_file
    @rubyxl_file = rubyxl_file
    @different_files = []
    @test_count = 0
  end

  def test_excel_encryption(file_path)
    @test_count += 1
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

  def apply_combination(base_dir, combination)
    # Apply a combination of file replacements
    combination.each do |file_info|
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

  def test_combinations
    puts "🔧 COMBINATION TESTING: Finding minimal changes for encryption compatibility"
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

    puts "\n📊 Found #{@different_files.length} files that differ"

    # Test combinations of increasing size
    1.upto(@different_files.length) do |combo_size|
      puts "\n🔬 Testing combinations of #{combo_size} files..."

      @different_files.combination(combo_size).each_with_index do |combination, index|
        print "\r   Progress: #{index + 1}/#{@different_files.combination(combo_size).count} "

        # Apply this combination
        FileUtils.rm_rf('working_extracted') if Dir.exist?('working_extracted')
        FileUtils.cp_r('caxlsx_extracted', 'working_extracted')

        apply_combination('working_extracted', combination)

        # Test this combination
        test_file = "combo_test_#{combo_size}_#{index}.xlsx"
        create_xlsx_from_files('working_extracted', test_file)

        if test_excel_encryption(test_file)
          puts "\n\n🎉 FOUND WORKING COMBINATION! (#{combo_size} files)"
          puts "Required changes:"
          combination.each_with_index do |file_info, i|
            action = case file_info[:type]
                     when :missing_in_caxlsx then "ADD"
                     when :different then "REPLACE"
                     end
            puts "   #{i + 1}. #{action}: #{file_info[:name]}"
          end

          # Check if byte-identical
          if files_identical?(test_file, @rubyxl_file)
            puts "\n✅ This creates a byte-identical file to RubyXL!"
          else
            puts "\n⚠️  Still not byte-identical to RubyXL, but encryption works!"
          end

          FileUtils.copy(test_file, 'minimal_fix.xlsx')
          puts "\n💾 Saved working file as: minimal_fix.xlsx"

          File.delete(test_file) if File.exist?(test_file)
          cleanup
          return combination
        end

        File.delete(test_file) if File.exist?(test_file)
      end

      puts ""
    end

    puts "\n❌ No combination of file replacements fixed the encryption issue!"
    puts "Total tests performed: #{@test_count}"
    cleanup
    nil
  end

  def cleanup
    FileUtils.rm_rf('caxlsx_extracted')
    FileUtils.rm_rf('rubyxl_extracted')
    FileUtils.rm_rf('working_extracted')
  end
end

# Run the combination testing
if ARGV.length != 2
  puts "Usage: ruby #{$0} <caxlsx_file> <rubyxl_file>"
  exit 1
end

transformer = CombinationTransformer.new(ARGV[0], ARGV[1])
result = transformer.test_combinations

if result
  puts "\n🏆 SUCCESS! Found minimal changes needed for encryption compatibility."
else
  puts "\n💔 Could not find a combination that enables encryption."
end
