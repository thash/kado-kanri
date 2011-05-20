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

		### column scan...
		row.Columns.each do |cell|
			#print "DEBUG: "; puts cell.Address

			### check Project code...
			if cell.Value == pj
				puts "XXXXXXXXX HIT - ROW XXXXXXXXX"
				@target_row_num = row_num
				puts "TARGET_ROW::: " + @target_row_num.to_s
				throw :last
			end
		end
		row_num += 1
		end
	}

	return @target_row_num

end

def find_col(sheet, day)

	cells = sheet.cells

	### column scan...
	col_num = 1
	catch(:last){
		### FORMAT: date need to be placed in 4th row
		sheet.UsedRange.Rows[4].Columns.each do |col|

		col.Rows.each do |cell|
			#print "DEBUG: "; puts cell.Address

	#puts "DEBUG: " + "day: " + day
	#puts "DEBUG: " + "cell.Value: " + cell.Value.to_i.to_s
			### check Project code...
			if cell.Value.to_i.to_s == day
				puts "XXXXXXXXX HIT - COL XXXXXXXXX"
				@target_col_num = col_num
				puts "TARGET_COL::: " + @target_col_num.to_s
				throw :last
			end
			col_num += 1
		end
		end
	}

	return @target_col_num
end

def insert_hour(sheet, day, pj, hour)

	@target_row_num = find_row(sheet, pj)
	@target_col_num = find_col(sheet, day)

	cells = sheet.cells
	puts "CHANGE CELL::: " +  "(" + @target_row_num.to_s + "," + @target_col_num.to_s +  ")"
	print "before: "
	pp cells.item(@target_row_num, @target_col_num).value

	cells.item(@target_row_num, @target_col_num).value = hour.to_s

	print "after: "
	pp cells.item(@target_row_num, @target_col_num).value
end



### MAIN PROCESS ###
class EditExcel

	s = PrivacyUtil.new(ARGV)
	res = s.make_hash_from_json
	path = res["path_test"]
	arg_year = res["year"] #TODO: Date.today.year
	arg_month = res["month"] #TODO: Date.today.month
	arg_day = res["day"] # arg_day = Date.today.day
	arg_pj = res["pj"]
	arg_hour = res["hour"]

	puts path
	print "TARGET [pj]  is..." + arg_pj.to_s + "\n"
	print "TARGET [day] is..." + arg_day.to_s + "\n"

	excel = WIN32OLE.new('Excel.Application')
	excel.visible = true

	#book = excel.Workbooks.Open(excel.GetOpenFilename)
	book = excel.Workbooks.Open(path)

	sheet = select_sheet(book, arg_year, arg_month)

	insert_hour(sheet, arg_day, arg_pj, arg_hour)

	#book.close(false)
	#excel.quit
end
