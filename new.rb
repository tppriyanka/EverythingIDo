#!/usr/bin/env ruby
require 'rubygems'
require 'mysql'
require 'active_support/all'
require 'active_record'
require "date"
require 'time'
require 'faster_csv'
require 'yaml/store'
require 'rubyXL'
require 'zip/zip' #rubyzip

class Logs_Expiry < ActiveRecord::Base
  set_table_name :logs_expiry
end

$current_path=File.expand_path(File.dirname(__FILE__))
@startdate = 1.days.ago.strftime('%Y-%m-%d')

def naming_files()
  circles = ["chennai","rotn"]
  files = []
  circles.each do |circle|
    filename = "ussd-vodafone-bonus-wise-lapser-#{circle}-#{@startdate}.xlsx"
    files << filename
    end
return(files)
end

def db_operation()
	
	dbconfig = YAML::load(File.open('db_settings.yml'))['dev']
   	ActiveRecord::Base.establish_connection(dbconfig)
        log_data = Logs_Expiry.find_by_sql("select distinct (left(input_file, length(input_file)-22)) from logs_expiry")
        return log_data
        
end

def fetch_data()
dbconfig = YAML::load(File.open('db_settings.yml'))['dev']
   	ActiveRecord::Base.establish_connection(dbconfig)
results=Logs_Expiry.find_by_sql("select date(created_at),left(input_file, length(input_file)-22),sum(input_file_size),sum(expired_file_size) from  logs_expiry where created_at between '2011-12-01 00:00:00' and '2011-12-29 23:59:59' and mode= 'D+1' and circle ='Chennai' group by date(created_at), left(input_file, length(input_file)-22),input_file_size,expired_file_size limit 80; ")
return results

end

def write_data(results)
  results.each do |row|
   
  
  row.attributes.each do |key, value|
    puts "Key: #{key} | Value: #{value}"
  end

   #row.attributes().each do |attr|
   
   #  p attr.length
   #  p "------------------"
     
   #  end
   end

end

def make_file(workbook,log_data,row_num)
      log_data.each do |row|
           row.attributes().each do |attr|
		
	      workbook.worksheets[0].add_cell(row_num,0,attr[1])	
	      workbook.worksheets[0].merge_cells(row_num,0,row_num,3)
	      workbook.worksheets[1].add_cell(row_num,0,attr[1])	
	      workbook.worksheets[1].merge_cells(row_num,0,row_num,3)
		row_num += 1
		
	end
end
return row_num
end

def run(filename)
  workbook = RubyXL::Workbook.new
  workbook.worksheets = []
  workbook.worksheets << RubyXL::Worksheet.new(workbook,'D+1')
  workbook.worksheets << RubyXL::Worksheet.new(workbook,'D+5')
  workbook.worksheets[0].sheet_name = 'D+1'
  workbook.worksheets[1].sheet_name = 'D+5'
  
  row_num = 2
  row_num = make_file(workbook,db_operation(),row_num)
  row_num +=2
  row_num =make_file(workbook,db_operation(),row_num)
  row_num +=2
  row_num =make_file(workbook,db_operation(),row_num)  
  write_data(fetch_data)
  workbook.write("#{$current_path}/downloads/#{filename}")
end

def zipping(files)
Zip::ZipFile.open("#{$current_path}/downloads/ussd-vodafone-bonus-wise-lapser_#{@startdate}.zip", Zip::ZipFile::CREATE) {|zipfile|
    files.each {|filename|
     zipfile.add("#{filename}","#{$current_path}/downloads/#{filename}")
      }
    }
    files.each {|filename|
    File.delete("#{$current_path}/downloads/#{filename}")
   }
end

files = naming_files()
run(files[0])
run(files[1])
zipping(files)
