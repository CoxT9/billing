require 'spreadsheet'
require 'mysql2'
require 'Client'

class SheetReader

	def initialize(filename,month,year)
		
		#DCL is where we find most of our info (name,services)
		@DCL_START_ROW = 7

		#Client info is where we find the email, address , and parent names
		@CLIENT_INFO_START_ROW = 5

		#This block of constants are referring to locations in the DCL sheet
		@NAME = 0 
		@THERAPIST = 1
		@SERVICE_TYPE = 2
		@DATE = 3
		@DURATION = 4
		@RATE = 7
		@TRAVEL_COST = 9
		@SERVICE_TOTAL = 10
		@TOTAL_COST = 11 # this is present on the same row as name

		#these contain locations in the client info sheet
		@PARENT_NAMES = 2
		@ADDRESS = 3
		@PROV = 4
		@POSTAL_CODE = 5
		@EMAIL = 7
		@book = Spreadsheet.open(filename)
		@dcl = @book.worksheet(0)
		@clientInfo = @book.worksheet("Client Info")
		@curDclRow = @dcl.row(@DCL_START_ROW)
		@curDclRowN = @DCL_START_ROW

		#need to set month and year for the database
		@month = month
		@year = year
	end

	def process_sheet
		c=get_next_client
		clientNum = 0
		db = get_database

		until(c.nil?) do
			c=get_next_client
			db.query("INSERT INTO clients VALUES (#{clientNum},'#{c.name}','#{c.address}','#{c.prov}','#{c.postal}','#{c.email}','#{c.totalCost}');")
			clientNum++;
		end 
	end

	#get information for, and return, the client
	def get_next_client
		name = @curDclRow[@NAME].strip unless @curDclRow[@NAME].nil?
		if name.nil?
			nil
		else
			client = Client.new(name)
			client.totalCost = @curDclRow[@TOTAL_COST]
			services = get_services
			services.each {|service| client.add_service(service)} 
			update_row
			get_info(client)
			client
		end
	end 

	def get_services
		services = []

		until(@curDclRow[@SERVICE_TYPE].nil?) do 
			s=Service.new(@curDclRow[@SERVICE_TYPE],@curDclRow[@THERAPIST],@curDclRow[@DATE],@curDclRow[@DURATION],@curDclRow[@RATE],@curDclRow[@TRAVEL_COST],@curDclRow[@SERVICE_TOTAL])
			services << s 
			update_row
		end

		services
	end 

	def get_info(client)
		curRow = @clientInfo.row(@CLIENT_INFO_START_ROW)
		curPos = @CLIENT_INFO_START_ROW

		while(!curRow[@NAME].nil? && !curRow[@NAME].strip.include?(client.name.strip)) do
			curPos+=1
			curRow=@clientInfo.row(curPos)
		end

		#we either have the row we want or a row with nothing in it (we didnt find what we were looking for)
		if(curRow[@NAME].nil?)
			client.hasInfo = false
		else
			client.parents = curRow[@PARENT_NAMES]
			client.address = curRow[@ADDRESS]
			client.prov = curRow[@PROV]
			client.postal = curRow[@POSTAL_CODE]
			client.email = curRow[@EMAIL]
		end
	end

	def update_row
		@curDclRowN+=1
		@curDclRow = @dcl.row(@curDclRowN)
	end

	#These finctions relate to the database that SheetReader will store client info in

	def get_database
		db = Mysql2::Client(:host => 'localhost',:user => 'root',:password => 'abcd0311')

		begin
			db.query("DROP DATABASE #{@month}_#{@year}_clients")
			db.query("CREATE DATABASE #{@month}_#{@year}_clients")
		rescue Exception => e
			db.query("CREATE DATABASE #{@month}_#{@year}_clients")
		end

		db.query("USE #{@month}_#{@year}_clients")
		db.query("CREATE TABLE users(cID integer, name varchar(50), address varchar(50), prov varchar(50), postal varchar(50), email varchar(50), cost float);")
		db
	end
end 
