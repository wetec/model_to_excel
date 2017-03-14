require "model_to_excel/version"

module ModelToExcel
  # Your code goes here...
  def self.to_excel(file_name)
    Spreadsheet.client_encoding = "UTF-8"
    book = Spreadsheet::Workbook.new
    first_row = ["编号","字段","类型","注释"]

    sheet = book.create_worksheet :name => "表注释"
    sheet.row(0).concat ["编号","表名","中文名","注释"]
    num = 1
	ActiveRecord::Base.connection.tables.each do |table|
	      next if table.match(/\Aschema_migrations\z/)
	      begin
	      	  class_name = table.singularize.classify
		      columns = class_name.constantize.columns
		      sheet[num, 0] = num
	          sheet[num, 1] = table.to_s
	          sheet[num, 2] = ""
	          sheet[num, 3] = ""
	          num += 1
          rescue Exception => e
	        puts "table"
	        puts table

	        next
	      end
	end

    ActiveRecord::Base.connection.tables.each do |table|
      next if table.match(/\Aschema_migrations\z/)
      begin
      	  class_name = table.singularize.classify
	      columns = class_name.constantize.columns

	      sheet = book.create_worksheet :name => table
	      #第一横行
	      sheet.row(0).concat first_row
	      num = 1
	      columns.each do |column|
	          puts column 
	          puts column.type
	          sheet[num, 0] = num
	          sheet[num, 1] = column.name
	          sheet[num, 2] = column.type.to_s
	          sheet[num, 3] = ""
	          num += 1
	          
	        end
      rescue Exception => e
        puts "table"
        puts table

        next
      end
    end

    book.write "#{Rails.root}/public/#{file_name}.xls"
  end
end
