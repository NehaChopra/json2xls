##
#
#   author: neha chopra,
#   Reads the files containing the JSON and convert them to .XLS file
#
#

require 'pp'
require 'fileutils'
require 'date'
require 'multi_json'
require 'spreadsheet'

module Json2xls
	class Generator
		def initialize(paths, options={})

			@name             = options[:name] || 'Worksheet'
			@path             = options[:path] || ((Dir.exists?("#{ENV['HOME']}/Json2xls/#{Time.now.to_s}")) ? "#{ENV['HOME']}/Json2xls/" : FileUtils::mkdir_p("#{ENV['HOME']}/Json2xls/#{Time.now.to_s}").first)
			@header_row       = options[:header_row] || 0
			@data_row         = @header_row + 1
			@current_data_row = @data_row
			@header_format    = options[:header_format] || (Spreadsheet::Format.new :color => :black, :weight => :bold, :pattern => 1, :pattern_fg_color => :orange, :bottom => :double, :top => :double)
			@bold_format      = Spreadsheet::Format.new :weight => :bold, :color => :black
			@wrap_text_format = Spreadsheet::Format.new :text_wrap => true, :vertical_align => 'top'
			@book             = Spreadsheet::Workbook.new

			json_loader(paths)

		end

		##
		#
		# Builds the sheets for a Worksheet
		#
		#
		def build_sheet
			sheet = @book.create_worksheet :name => @name
			sheet
		end


		##
		#
		# writing to .XLS file
		#
		#
		def write file
			@book.write file
		end

		##
		#
		# Loads the file from given path
		# and initiate the processing of the JSON object
		#
		#
		def json_loader paths
			paths = [paths] unless paths.is_a?(Array)

			paths.map do |path|
				@current_data_row = @data_row
				json_objects = file_loader path

				process_json json_objects
			end
		end

		##
		#
		# Reads the given file and convert the JSON
		# provided into the sheet to Ruby object
		#
		#
		def file_loader path
			raw_data = IO.read path
			begin
				return MultiJson.load raw_data
			rescue Exception => e

			end
		end

		##
		#
		# After processing of the objects ..XLS file is build
		#
		#
		def process_json json_objects
			sheet = build_sheet
			json_objects.keys.each do |key|
				build_xls json_objects[key], key, sheet
			end
		end

		##
		#
		# Building the ..XLS files
		# build_headers method build the headers for .XLS file
		# build_values method ensures the values get generated for .XLS file
		#
		def build_xls json_objects, key, sheet

			excel_headers = build_headers json_objects, sheet

			build_values json_objects, excel_headers, key, sheet

			write "#{@path}/#{@name}.xls"
		end

		##
		#
		# building headers for .XLS file
		#
		#
		def build_headers json_objects, sheet
			excel_headers = generate_keys json_objects
			excel_headers.unshift("Identifier")

			(excel_headers).length.times do |index|
				sheet[@header_row,index] = excel_headers[index].to_s
				sheet.column(index).width = 30
				sheet.row(@header_row).set_format(index, @header_format)
			end

			excel_headers
		end

		##
		#
		# building result set row values for .XLS file
		#
		#
		def build_values json_objects, excel_headers, key, sheet
			contained_values =
			excel_headers.map do |head|
				if(head.eql?'Identifier')
					value = key.to_s
				else
					value = generate_value json_objects, head
				end
				value
			end


			(excel_headers).length.times do |index|
				sheet[@current_data_row,index] = contained_values[index].to_s
				sheet.column(index).width = 30
			end
			@current_data_row = @current_data_row +1
		end

		def generate_value json_objects, head
			segmented_heads = head.split(' ')
			segmented_heads.each do |obj|
				return nil if obj.nil?
				json_objects = json_objects[obj]
			end
			json_objects
		end

		def generate_keys(obj, prefixed_header=nil)
			keys = obj.keys

			result_set =
			  keys.map do |key|

				  untangled_key = untangled_key(key)
				  value = obj[key]
				  if value.is_a?Hash

					  generate_keys(obj[key], (prefixed_header ? "#{prefixed_header} #{untangled_key}" : untangled_key))
				  else
					  if !prefixed_header.nil?
						  "#{prefixed_header} #{untangled_key}"
					  else
						  untangled_key
					  end
				  end
			  end

			result_set.compact.flatten.uniq
		end

		def untangled_key key
			key
		end

	end
end

