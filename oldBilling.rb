Shoes.setup do
	gem 'prawn'
	gem 'spreadsheet'
	gem 'mail'
end

require 'spreadsheet'
require 'mail'
require 'prawn'


####################################################################################
#This is the Service writer section
####################################################################################

class Service

	attr_accessor :type,:therapist,:date,:time,:rate,:travelCharge,:total

	def initialize(type,therapist,date,time,rate,travelCharge,total)
		@type = type
		@therapist = therapist
		@date = convertDate(date)
		@time = time 
		@rate = rate
		@travelCharge = travelCharge
		@total = total
	end

	def to_s
		"%s     %-12s      %.2f              $%.2f           $%.2f              $%.2f" % [@date,@type,@time,@rate,@travelCharge,@total]
	end  

	def convertDate(date)
		s=date.to_s
		
		unless(date.to_s.index('T').nil?)
			s=date.to_s[0,date.to_s.index('T')] unless date.nil?
		end

		s
	end
end    

####################################################################################
#This is the client section
####################################################################################

class Client
	attr_accessor :name,:parents,:address,:prov,:postal,:services,:totalCost,:email,:hasInfo

	def initialize(name)
		@name = name
		@parents=nil 
		@address = nil
		@prov = nil
		@postal = nil
		@services = []
		@totalCost = nil
		@email=nil
		@hasInfo = true
	end

	def add_service(service)
		services << service
	end

	def to_s
		s=@name.to_s+"\n"
		s+=@parents.to_s
		s+="\n"
		s+=@address.to_s
		s+="\n"
		s+=@prov.to_s
		s+="\n"
		s+=@postal.to_s
		s+="\n"
		s+=@email.to_s
		s+="\n"
		s+=@totalCost.to_s
		s+="\n"
		s+=@services.to_s
		s 
	end 

	def email=(mail)
		@email=mail.to_s
	end

	def service_string
		s=""
		@services.each {|service| s+=(service.to_s+"\n")}
		s+="                            Total: $%.2f" % [@totalCost]
		s
	end
end

####################################################################################
#This is the SheetReader section
####################################################################################


class SheetReader

	def initialize(filename)
		
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
	end

	def get_spreadsheet
		puts "Enter the name of the spreadsheet file: "
		filename = gets.chomp.strip
		puts "Trying to open #{filename}"
		book = Spreadsheet.open(filename) if File.exist?(filename) 

		while(book.nil? && filename!="") do
			puts "File not found. Enter with format <filename>.xls (or input nothing to exit): "
			filename = gets.chomp.strip
			puts "Trying to open #{filename}"
			book = Spreadsheet.open(filename) if File.exist?(filename)
		end 

		if(filename == "")
			puts "Program ended (user input)"
			raise "Program Ended"
		else 
			puts "File opened successfully"
			book
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
end 

####################################################################################
#This is the PDF writer section
####################################################################################

#these costants have to do with writing the list of services
$LINE_SIZE = 13.875
$DATE_WIDTH = 75
$SPACE = 5
$SERVICE_WIDTH = 95
$TIME_WIDTH = 75
$HRATE_WIDTH = 75
$TRAVEL_WIDTH = 75
$COST_WIDTH = 70
$TOTAL_WIDTH = 80

def generatePDFs(c,year,month)
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

	invoice.text "<b>#{month.upcase} INVOICE</b>", :align => :center,:inline_format => true
	receipt.text "<b>#{month.upcase} RECEIPT</b>", :align => :center,:inline_format => true
	
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
	dir = "#{year}/#{month}/#{c.name.gsub("/","-")}"
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
	height = (c.services.size+1) * $LINE_SIZE #the height of the boxes (all the same)

	#first make the bounding box that stores the dates
	pdf.float do
		pdf.bounding_box([x,pdf.cursor], :width => $DATE_WIDTH , :height => height) do
			pdf.text "<u>Service date</u>", :inline_format => true
			c.services.each {|service| pdf.text "#{service.date}"}
		end
	end
	x+= $DATE_WIDTH+$SPACE

	#bounding box that stores the service types
	pdf.float do
		pdf.bounding_box([x,pdf.cursor], :width => $SERVICE_WIDTH , :height => height) do
			pdf.text "<u>Service</u>", :inline_format => true
			c.services.each {|service| pdf.text "#{service.type}"}
		end
		x+= $DATE_WIDTH+$SPACE
	end 

	pdf.float do
		pdf.bounding_box([x,pdf.cursor], :width => $TIME_WIDTH, :height => height) do
			pdf.text "<u>Service Time</u>", :inline_format => true
			c.services.each {|service| pdf.text "%.2f" % [service.time]}
		end
		x+= $TIME_WIDTH + $SPACE
	end

	pdf.float do
		pdf.bounding_box([x,pdf.cursor], :width => $HRATE_WIDTH, :height => height) do
			pdf.text "<u>Hourly Rate</u>", :inline_format => true
			c.services.each {|service| pdf.text "$%.2f" % [service.rate]}
		end
		x+= $HRATE_WIDTH + $SPACE
	end

	pdf.float do
		pdf.bounding_box([x,pdf.cursor], :width => $TRAVEL_WIDTH, :height => height) do
			pdf.text "<u>Travel Charge</u>", :inline_format => true
			c.services.each {|service| pdf.text "$%.2f" % [service.travelCharge]}
		end
		x+= $HRATE_WIDTH + $SPACE
	end

	pdf.bounding_box([x,pdf.cursor], :width => $COST_WIDTH, :height => height) do
			pdf.text "<u>Service Cost</u>", :inline_format => true
			c.services.each {|service| pdf.text "$%.2f" % [service.total]}
	end

	pdf.move_down 5
	
	pdf.bounding_box([x-32,pdf.cursor], :width => $TOTAL_WIDTH, :height => $LINE_SIZE) do
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


####################################################################################
#This is the GUI section
####################################################################################

$email = ""
$password = ""

def loginView
	Shoes.app :height => 250, :width => 300, :title => "Billing" do
		background gainsboro

		stack :margin => 20 do
			background white
			
			flow :margin => 5 do
				para "Email: "
				@emailLine = edit_line
			end

			flow :margin => 5 do
				para "Password: "
				@passwordLine = edit_line
				@passwordLine.secret = true
			end

			button "Enter",:displace_left => 80 do
				if(@passwordLine == "" || @emailLine.text == "")
					@errorText.show
				else
					$email = @emailLine.text
					$password = @passwordLine.text
					settings = File.open("settings.txt")

					options = { :address              => settings.gets.chomp, # "real world will" be Smtp.mail.yahoo.com (Rogers is powered by Yahoo)
            					:port                 => settings.gets.chomp.to_i, #25, 2525, 465, 587
            					:domain               => settings.gets.chomp, #40180
            					:user_name            => $email,
            					:password             => $password, #Careful here !!
            					:authentication       => 'plain',
	            				:enable_starttls_auto => true  }
            

								settings.close

								Mail.defaults do
  								delivery_method :smtp, options
								end

								selectionView
								close
				end
			end

			@errorText = para "Enter email and password"
			@errorText.hide
		end
	end
end

def selectionView
	Shoes.app :height => 158, :title => "Billing" do

			background white
			flow do
				stack :width => "25%" do	
					pdfLogo=image "pdflogo.png"
					pdfLogo.height = 118
					button "Generate PDFs" do
						generationView
					end
				end

				stack :width => "25%" do
					image "maillogo.png"
					button "Mail Invoices" do
						mailSelectView("invoice")
					end
				end

				stack :width => "25%" do
					image "maillogo.png"
					button "Mail Receipts" do
						mailSelectView("receipt")
					end
				end

				stack :width => "25%" do
					settingsLogo = image "settingsLogo.png"
					settingsLogo.height = 118
					button "Settings" do
						settingsView
					end
				end
			end#flow
		end #app
end#selectionView

def generationView 
	Shoes.app :title => "Billing",:height => 295 do
			@filename = ""
			background white
			stack do
				tagline "Generate PDF files" , underline: "single"
				
				flow margin: 20 do
					background gainsboro
					button "Choose file",:displace_top => 5 do 
						@filename = ask_open_file()
						@fileLine.text=@filename
					end
					@fileLine = edit_line
					@fileLine.displace(0,5)
				end

				flow margin: 20 ,:displace_top => "-10" do
					para "Enter the Year     "
					@yearLine = edit_line
					@yearLine.text=get_year
				end

				flow margin: 20,:displace_top => "-30" do
					background gainsboro
					para "Enter the Month  ",:displace_top => 5
					@monthLine = edit_line
					@monthLine.text=get_month
					@monthLine.displace(0,5)
				end

				button "Generate", :displace_top => -20 do
					generate(@filename,@yearLine.text,@monthLine.text)
				end 
			end
		end
end

def generate(filename,year,month)
		para "beginning generation process: "
		s=SheetReader.new(filename)
		c=s.get_next_client
		para c.inspect
		curClient=0
		infoMsg = ""
		emailMsg = ""

		Dir.mkdir(year) unless File.exist?(year)
		Dir.mkdir("#{year}/#{month}") unless File.exist?("#{year}/#{month}")
		infoFile = File.new("#{year}/#{month}/info.txt","w")

		while(!c.nil?) do
			para "beginning generation for #{c.name}"
			generatePDFs(c,year,month)
			infoFile.write ("%s %s" % [c.name.gsub(" ","-").gsub("/","-"),c.email]) #blank spaces and slashes are replaced with dashes to make searching easier
			infoFile.write("\n")

			c=s.get_next_client
			para "done"
		end

		infoFile.close
		alert "PDF generation complete\n"+infoMsg+emailMsg , :title => "Generate PDFs"
	end

def mailSelectView(type)
	Shoes.app :height => 150, :width => 300, :title => "Billing" do
		background white
		stack do
			para "Select the year and month for emailing"
			flow do
				para "Year     "
				@years = list_box items: (2015..2030).to_a
				@years.choose(get_year)
				@years.change do 
					if(File.exist? "./#{@years.text}")
						@months.items=Dir.entries("./#{@years.text}").delete_if {|s| s == "." || s == ".."}
					else
						@months.items=[]
					end
				end
			end

			flow do 
				para "Month  "
				@months = list_box items: Dir.entries("./#{get_year}").delete_if {|s| s == "." || s == ".." || s == ".DS_Store"}
			end

			button "ok", :displace_left => 130 do
				mailView(type,@years.text,@months.text)
				close
			end
		end
	end
end

def mailView(type,year,month)
	Shoes.app :height => 200 , :title => "Billing" do
		stack :width => "40%" do
			stack :height => 200,:scroll => true do
				@clients = Dir.entries("./#{year}/#{month}").sort.delete_if {|s| s == "." || s==".." || s== "info.txt" || s == ".DS_Store"}
				@mailList = Array.new
				@clients.each do |client|
					flow do
						curClient = check
						para "#{client}"
						curClient.click do
							if curClient.checked?
								@mailList << client
							else
								@mailList.delete(client)
							end 
						end
					end
				end
			end
		end

		stack :width => "60%" do 
			background gainsboro
			button "send all", :displace_top => 10 do
				sendEmails(type,@clients,year,month)
			end

			button "send selected",:displace_top => 30 do
				sendEmails(type,@mailList,year,month)
			end

			button "send custom" ,:displace_top => 50 do
				if(@mailList.size != 1)
					@errorText.show
				else
					@errorText.hide
					customMailView(@mailList[0],year,month,type)
				end
			end

			@errorText = para "Only one client can be selected for custom email",:displace_top => 50
			@errorText.hide
		end
	end
end

def customMailView(clientName,year,month,type)
	box_line = 49
	if(type == "invoice")
		mailBody = "Hello,\n\nPlease see attached below your #{month} #{type} for #{pdf_password(clientName)}'s speech language services provided by Active Communication Therapy. In order to open this attachment, please use your child's first name as the password."
		mailBody += "\nPlease remit payment by email money transfer to this email address (using your child's name as password), or send cheque made out to Tanis Friesen at 71 Caroline Avenue, Ottawa, ON, K1Y 0S8.\nPlease confirm that you have received this email.\n\nThank you,"
	else 
		mailBody = "Hello,\n\nThank you for your payment for #{pdf_password(clientName)}'s #{month} services. Please see your #{month} #{type} attached below.\n"
	end
	mailBody += "\nTanis Friesen, Reg. CASLPO\nSpeech Language Pathologist\nActive Communication Therapy\nPhone: (613)728-8878\nFax: (613)667-9757"
	
	mailSubject = "#{pdf_password(clientName)}'s #{month} #{type}"

	Shoes.app :height => 400 , :title => "Billing" do
		background gainsboro
		stack :margin => 20,:height => 400 do
			background white

			flow :margin => 5 do
				para "To: "
				@emailLine = edit_line
				@emailLine.text=getEmail(clientName,year,month)
			end

			flow :margin => 5 do
				para "Subject: "
				@subjectLine = edit_line
				@subjectLine.text=mailSubject
			end

			flow :margin => 5 do
				para "Attachment: "
				@attachmentLine = edit_line
				@attachmentLine.text = "#{year}/#{month}/#{clientName}/#{get_initials(clientName)}-#{type}.pdf"
			end

			@body = edit_box :width => 400, :height => 200
			@body.displace(50,0)
			@body.text=mailBody

			button "send", :displace_left => 180 do
				para "foo"
				sendCustom(@emailLine.text, @subjectLine.text, @body.text, @attachmentLine.text)
			end
		end
	end
end

def settingsView  #this needs to be fixed
	Shoes.app :height => 190 , :width => 350 , :title => "Billing" do
		@settings = File.open("settings.txt","r")
		@address = @settings.gets.chomp
		@port = @settings.gets.chomp
		@domain = @settings.gets.chomp
		@settings.close

		background gainsboro

		stack :margin => 20 do
			background white

			flow do
				para "Address: "
				@addressLine =  edit_line
				@addressLine.text = @address
			end 

			flow do
				para "Port: "
				@portLine =  edit_line
				@portLine.text = @port
			end 

			flow do
				para "Domain: "
				@domainLine = edit_line
				@domainLine.text = @domain
			end
		end

		button "Done", :displace_top => 5, :displace_left => 140 do
		File.delete("settings.txt")
			@settings = File.open("settings.txt","w")
			@settings.write(@addressLine.text.chomp)
			@settings.write("\n")
			@settings.write(@portLine.text.chomp)
			@settings.write("\n")
			@settings.write(@domainLine.text.chomp)
			@settings.write("\n")
			@settings.close
			close
		end
	end

end

def getEmail(clientName,year,month)
	searchName = clientName.gsub(" ","-").gsub("/","-")
	email = nil

	File.foreach("#{year}/#{month}/info.txt") do |line|
		tok = line.split(" ")
		email = tok[1] if (tok[0] == searchName)
	end

	email
end

def get_initials(name) 
	name[0]+name[name.index(" ")+1]
end



def sendEmails(type, clientList, year, month)

   clientList.each do |name|

   	############################################
	#the messages must be changed here (ONLY MAKE CHANGES WITHIN THE QUOTATION MARKS)
	############################################
	if(type == "invoice")
		mailBody = "Hello,\n\nPlease see attached below your #{month} #{type} for #{pdf_password(name)}'s speech language services provided by Active Communication Therapy. In order to open this document, please use your child's first name as the password.\n"
		mailBody += "\nPlease remit payment by email money transfer to this email address (using your child's name as password), or send cheque made out to Tanis Friesen at 71 Caroline Avenue, Ottawa, ON, K1Y 0S8.\nPlease confirm that you have received this email.\n\nThank you,"
	else 
		mailBody = "Hello,\n\nThank you for your payment for #{pdf_password(name)}'s #{month} services. Please see your #{month} #{type} attached below.\n"
	end
	mailBody += "\nTanis Friesen, Reg. CASLPO\nSpeech Language Pathologist\nActive Communication Therapy\nPhone: (613)728-8878\nFax: (613)667-9757"
	
	mailSubject = "#{pdf_password(name)}'s #{month} #{type}"
	#############################################
	#############################################
	#############################################
	

   		clientEmail = getEmail(name,year,month)

   		if !clientEmail.nil?

   		clientEmail=clientEmail.chomp
	    	mail = Mail.new do
	          from $email.chomp
	          to getEmail(name, year, month).chomp 
	       end

	       mail.subject(mailSubject) # Consider adding the name of the company in the subject field
	       mail.body(mailBody)
	       fileName = "#{get_initials(name)}-#{type}.pdf"
	       fileDir = "#{year}/#{month}/#{name}/#{fileName}" #remember to sub stuff in here
	       mail.add_file(:filename => fileName, :content => File.read(fileDir))
	       mail.deliver
   		end #if
    end

    para "END"
    alert "All messages sent", :title => "Billing"
end

def sendCustom(clientEmail, sub, bod, file) # All strings
    Mail.deliver do # Streamlined delivery of a custom email
        from $email
        to clientEmail
        subject sub
        body bod
        add_file(file)
    end
    alert "Mail sent", :title => "Billing"
end

loginView