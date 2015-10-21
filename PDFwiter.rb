#this class is responsable for writing the PDF invoices and reciepts

require 'prawn'
require 'mysql2'
require_relative "./Client"

class PDFWriter

	def initialize(month,year)
		#these costants have to do with writing the list of services
		@LINE_SIZE = 13.875
		@DATE_WIDTH = 75
		@SPACE = 5
		@SERVICE_WIDTH = 95
		@TIME_WIDTH = 75
		@HRATE_WIDTH = 75
		@TRAVEL_WIDTH = 75
		@COST_WIDTH = 70	
		@TOTAL_WIDTH = 80

		@month = month
		@year = year

		@db = get_database
	end

	def generate_all
		clients = @db.query("SELECT cID FROM clients")

		clients.each do |id|
			generatePDFs(id)
		end
	end

	def generatePDFs(cID)

		c=get_client(cID)

		invoice = Prawn::Document.new
		receipt = Prawn::Document.new

		invoice.float{invoice.image "./logo.png", :width => 110}
		receipt.float{receipt.image "./logo.png", :width => 110}

		companyData = File.new('companyData.txt', 'r')

		addHeader(invoice, companyData)
		companyData.seek(0)
		addHeader(receipt, companyData)

		invoice.move_down 20
		receipt.move_down 20

		invoice.text "<b>#{@month.upcase} INVOICE</b>", :align => :center,:inline_format => true
		receipt.text "<b>#{@month.upcase} RECEIPT</b>", :align => :center,:inline_format => true
		
		invoice.move_down(30)
		receipt.move_down(30)
		
		dateAndInfo(invoice,c)
		dateAndInfo(receipt,c)

		invoice.move_down(20)
		receipt.move_down(20)

		addData(invoice, c)
		addData(receipt, c)

		invoice.pad(10) {invoice.text "Please make cheque payable to Tanis Friesen", :align => :center}
		receipt.pad(10) {receipt.text "Thank you for payment recieved in full", :align => :center}
		invoice.encrypt_document(:permissions => {:modify_contents => false}, :user_password => pdf_password(c.name))
		receipt.encrypt_document(:permissions => {:modify_contents => false}, :user_password => pdf_password(c.name))
		
		#this section creates the folder structure we need
		dir = "#{@year}/#{@month}/#{c.name.gsub("/","-")}"
		Dir.mkdir(dir) unless File.exist?(dir)



		#now that the folders are set up, we can store the file
		invoice.render_file("#{dir}/#{get_client_initials(c)}-invoice.pdf")
		receipt.render_file("#{dir}/#{get_client_initials(c)}-receipt.pdf")

	end

	def pdf_password(name)
		name[0,name.index(" ")] 
	end

	def get_date
		date = Time.now.strftime("%d/%m/%Y")
		date.to_s
	end

	def addHeader(pdf, text)
		while line = text.gets
			pdf.text line,:align => :center, :inline_format => true
		end
	end

	def dateAndInfo(pdf, c)
		pdf.text "Client(s): #{c.name}" 
		pdf.text "Parent(s): #{c.parents}"
		pdf.text c.address
		pdf.text c.prov
		pdf.text c.postal
		pdf.text c.email unless (c.email.nil? || c.email == "" || c.email == " ")
	end

	def addData(pdf, c)
		x = 0 #where the bounding box starts (in terms of x value)
		height = (c.services.size+1) * @LINE_SIZE #the height of the boxes (all the same)

		#first make the bounding box that stores the dates
		pdf.float do
			pdf.bounding_box([x,pdf.cursor], :width => @DATE_WIDTH , :height => height) do
				pdf.text "<u>Service date</u>", :inline_format => true
				c.services.each {|service| pdf.text "#{service.date}"}
			end
		end
		x+= @DATE_WIDTH+@SPACE

		#bounding box that stores the service types
		pdf.float do
			pdf.bounding_box([x,pdf.cursor], :width => @SERVICE_WIDTH , :height => height) do
				pdf.text "<u>Service</u>", :inline_format => true
				c.services.each {|service| pdf.text "#{service.type}"}
			end
			x+= @DATE_WIDTH+@SPACE
		end 

		pdf.float do
			pdf.bounding_box([x,pdf.cursor], :width => @TIME_WIDTH, :height => height) do
				pdf.text "<u>Service Time</u>", :inline_format => true
				c.services.each {|service| pdf.text "%.2f" % [service.time]}
			end
			x+= @TIME_WIDTH + @SPACE
		end

		pdf.float do
			pdf.bounding_box([x,pdf.cursor], :width => @HRATE_WIDTH, :height => height) do
				pdf.text "<u>Hourly Rate</u>", :inline_format => true
				c.services.each {|service| pdf.text "$%.2f" % [service.rate]}
			end
			x+= @HRATE_WIDTH + @SPACE
		end

		pdf.float do
			pdf.bounding_box([x,pdf.cursor], :width => @TRAVEL_WIDTH, :height => height) do
				pdf.text "<u>Travel Charge</u>", :inline_format => true
				c.services.each {|service| pdf.text "$%.2f" % [service.travelCharge]}
			end
			x+= @HRATE_WIDTH + @SPACE
		end

		pdf.bounding_box([x,pdf.cursor], :width => @COST_WIDTH, :height => height) do
				pdf.text "<u>Service Cost</u>", :inline_format => true
				c.services.each {|service| pdf.text "$%.2f" % [service.total]}
		end

		pdf.move_down 5
		
		pdf.bounding_box([x-32,pdf.cursor], :width => @TOTAL_WIDTH, :height => @LINE_SIZE) do
			pdf.text "Total: $%.2f" % [c.totalCost]
		end
	end

	def get_client_initials(c)
		c.name[0]+c.name[c.name.index(" ")+1]
	end

	def get_month
		months = [ "January" , "February" , "March" , "April" , "May" , "June" , "July" , "August" , "September" , "October" , "November" , "December" ]
		date = Time.now.strftime("%m").to_i
		months[date - 1]
	end

	def get_year
		Time.now.strftime("%Y")
	end

	def get_client(cID)
		c = @db.query("SELECT * FROM clients WHERE cID=#{cID};")

		c.each do |cInfo|
			client = Client.new(cInfo["name"])
			client.address = cInfo["address"]
			client.prov = cInfo["prov"]
			client.postal = cInfo["postal"]
			client.address = cInfo["address"]
			client.email = cInfo["email"]
			client.totalCost = cInfo["cost"]
		end

		client
	end

	def get_database
		db = Mysql2::Client(:host => 'localhost',:user => 'root',:password => 'abcd0311')
		db.query("USE #{@month}_#{@year}_clients")
		db
	end
end
