require 'rubygems'
require 'bundler/setup'
require 'json2xls'
require 'pp'
require 'fileutils'
require 'date'
require 'multi_json'
require 'spreadsheet'


RSpec.configure do |config|
  # Enable flags like --only-failures and --next-failure
  config.example_status_persistence_file_path = ".rspec_status"

  # Disable RSpec exposing methods globally on `Module` and `main`
  config.disable_monkey_patching!

  config.expect_with :rspec do |c|
    c.syntax = :expect
  end
end
