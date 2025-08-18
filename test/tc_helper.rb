$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'simplecov'
SimpleCov.start do
  add_filter "/test/"
  add_filter "/vendor/"
end

require 'minitest/autorun'
require 'timecop'
require 'webmock/minitest'
require 'axlsx'

module Minitest
  class Test
    def assert_false(value)
      assert_equal(false, value)
    end

    def refute_raises
      yield
    rescue StandardError => e
      raise Minitest::Assertion, "Expected no exception, but raised: #{e.class.name} with message '#{e.message}'"
    end

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
  end
end
