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

	def make_hash_from_json
		if File.exist?(@propaties_file)
			f = File.open(@propaties_file)
			json_hash = JSON.parse(f.read)

			return json_hash

			### usage
			# 	mail = json_hash["mail"]
			# 	password = json_hash["password"]

		else 
			puts "ERROR: please put #{@propaties_file} in your dir"
			exit
		end
	end
end

class ShowExcel

		today_full = Date.today
		today = today_full.day
		print "today's [day] is..." + today.to_s + "\n"
		print "target [day] is..." + "NOTDEFINED" + "\n"


	excel = WIN32OLE.new('Excel.Application')

    s = PrivacyUtil.new(ARGV)
	res = s.make_hash_from_json
	path = res["path_test_jp"]
	puts path
	#book = excel.Workbooks.Open(app.GetOpenFilename)
	book = excel.Workbooks.Open(path)

	flag = 0
	book.WorkSheets.each do |sheet|
		if sheet.Name == "201104"
			puts sheet.Name
			puts ""

			### row scan...
			sheet.UsedRange.Rows.each do |row|
				record = []

				### column scan...
				row.Columns.each do |cell|
					puts cell.Address
					
					### check Project code...
					if cell.Value == "E100802T01"
							puts "XXXXXXXXXXXXXXXXXXXXXXXXXX"
							target_row = row.Address
							puts target_row
							exit
						flag = 1
						record << cell.Value
					#	if cell.Value == "ID"
					#		tmpAd = cell.Address 
					#		puts tmpAd
					#	end
					end
				end
			#	puts record.join(",")
			end
		end
	end

#	#使っているワークシート範囲を一行ずつ取り出す
#	for row in book.ActiveSheet.UsedRange.Rows do
#		#取り出した行から、セルを一つづつ取り出す
#		for cell in row.Columns do
#			p cell.Address
#			p (cell.Value.class) == String ? \
#				cell.Value.encode("cp932") : \
#				cell.Value
#
#			p '-------'
#		end
#	end

	book.close(false)
	app.quit
end

# s = PrivacyUtil.new(ARGV)
# puts s
# pp s
# puts s.make_hash_from_json
