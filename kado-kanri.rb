# -*- encoding: cp932 -*-
require 'win32ole'
require 'date'
require 'json'
require 'pp'


class WorkProject
	def initialize(id, client, pj_name)
		@id      = id     
		@client  = client   
		@pj_name = pj_name
	end
end

class PrivacyUtil

	def initialize(propaties_file)
		@propaties_file = propaties_file[0]
	end


    ### usage ###
    # 	mail = json_hash["mail"]
    # 	password = json_hash["password"]
	def make_hash_from_json
		if File.exist?(@propaties_file)
			f = File.open(@propaties_file)
			json_hash = JSON.parse(f.read)

			return json_hash

		else 
			puts "ERROR: please put #{@propaties_file} in your dir"
			exit
		end
	end
end


def select_sheet(book, year, month)
	book.WorkSheets.each do |st|
		yyyymm = year.to_s + month.to_s
		if st.Name == yyyymm
			puts "found sheet: " + st.Name
			return st
		end
	end
end

def find_row(sheet, pj)
	cells = sheet.cells
	### row scan...
	row_num = 1
	catch(:last){
		sheet.UsedRange.Rows.each do |row|
		record = []

		### column scan...
		col_num = 1
		row.Columns.each do |cell|
			print "DEBUG: "; puts cell.Address

			### check Project code...
			if cell.Value == pj
				puts "XXXXXXXXX HIT XXXXXXXXXXX"
				@target_row_num = row_num
				puts "TARGET_ROW::: " + @target_row_num.to_s
				throw :last
			end
		col_num += 1
		end
		row_num += 1
		end
	}

	return @target_row_num

end

#def insert_hour(sheet, day, pj, hour)
def insert_hour(sheet, pj, hour)

	@target_row_num = find_row(sheet, pj)
	#TODO: @target_col_num = find_col(sheet, day)

	cells = sheet.cells
	puts "CHANGE CELL::: " +  "(" + @target_row_num.to_s + "," + @target_col_num.to_s +  ")"
	print "before: "
	pp cells.item(@target_row_num, 7).value

	cells.item(@target_row_num, 7).value = hour.to_s

	print "after: "
	pp cells.item(@target_row_num, 7).value
	#cells.item(1, 7).value = hour.to_s
end


class EditExcel

	today_full = Date.today
	today = today_full.day
	print "today's [day] is..." + today.to_s + "\n"
	print "target [day] is..." + "NOTDEFINED" + "\n"


	excel = WIN32OLE.new('Excel.Application')
	excel.visible = true

	s = PrivacyUtil.new(ARGV)
	res = s.make_hash_from_json
	path = res["path_test"]
	puts path
	#book = excel.Workbooks.Open(excel.GetOpenFilename)
	book = excel.Workbooks.Open(path)

	sheet = select_sheet(book, "2011", "04")

	insert_hour(sheet, "E100802T01", 3)

	#book.close(false)
	#excel.quit
end
