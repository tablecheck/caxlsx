$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'test/unit'
require "timecop"
require 'webmock/test_unit'
require "axlsx.rb"

def windows_platform?
  RUBY_PLATFORM =~ /mswin|mingw|cygwin/
end

def excel_available?
  return @excel_available unless @excel_available.nil?

  @excel_available = begin
    require 'win32ole'
    excel = WIN32OLE.new('Excel.Application')
    excel.Quit
    true
  rescue StandardError
    false
  end
end
