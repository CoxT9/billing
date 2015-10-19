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