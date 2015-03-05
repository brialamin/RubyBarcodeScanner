=begin
Author: Samuel Adams
Class: CSCI 3333
Date: 18 April 2013
Build Instructions: Run the code through cmd with ruby and everything should execute as intended.
=end

#needed for using the db
require 'win32ole' 
#creating the database
class AccessDb
    attr_accessor :mdb, :connection, :data, :fields

    def initialize(mdb=nil)
        @mdb = mdb
        @connection = nil
        @data = nil
        @fields = nil
    end

    def open
        connection_string =  'Provider=Microsoft.ACE.OLEDB.12.0;Data Source='
		connection_string << @mdb
        @connection = WIN32OLE.new('ADODB.Connection')
        @connection.Open(connection_string)
    end

    def query(sql)
        recordset = WIN32OLE.new('ADODB.Recordset')
        recordset.Open(sql, @connection)
        @fields = []
        recordset.Fields.each do |field|
            @fields << field.Name
        end
        begin
            @data = recordset.GetRows
        rescue
            @data = []
        end
        recordset.Close
    end

    def execute(sql)
        @connection.Execute(sql)
    end

    def close
        @connection.Close
    end
end
#basically the start of the actual code, some file error checking and such, making sure the file is there
def check_file
	if File.exists?("inventory.accdb") or File.exists?("inventory.mdb")
		print "The inventory database file is there.\n"
		db = AccessDb.new('inventory.accdb')
		db.open
		scan_barcode(db)			
	else
		print "There is no inventory file in your default directory, would you\n"
		print "like to specify a file location? (Y/N)\n"
		answer = gets.chomp
		if answer == "Y" or answer == "y"
			print "Enter the full file location, including the inventory file at the end: "
			dir = gets.chomp
			while !File.exists?(dir)
				print "\nThis file still isn't there.  Try entering it again: "
				dir = gets.chomp
			end
			print "\nFound the database!\n"
			db = AccessDb.new(dir)
			db.open
			scan_barcode(db)
		end
		if answer == "N" or answer == "n"
			print "Fine then, I didn't want to help you anyway :|.\n"
		end
	end
end
#barcode scanning function
def scan_barcode(db)
	#infinite while loop to keep the program running
	while(true)
		print "Enter a barcode number: "
		barcode = gets.chomp
			#only accepts numbers
			if (barcode =~ /^[0-9]{1,20}$/)
				#executes the SQL command
				db.query("SELECT * FROM items WHERE Barcode = '#{barcode}';")
				field_names = db.fields
				rows = db.data
				#checks to see if there's a result
				if(rows.length > 0)
					raw_string = rows.join(",")
					raw_string.gsub(/"/,'')
					print_rows = raw_string.split(",")
					#checks to see if the quantity is not 0
					if print_rows[3].to_i < 1 
						answer = nil
						#loop to ensure a y/n answer
						while answer != "y" and answer != "Y" and answer != "n" and answer != "N"
							print "Barcode #{barcode} found in the database but has a zero quantity.  Do you want to update quantity? [Y/N]: "
							answer = gets.chomp
							if answer == "y" or answer == "Y"
								#gets a new quantity
								print "\nEnter the new quantity: "
								quantity = gets.chomp
								#number error checking
								while !(quantity =~ /^[0-9]{1,20}$/)
									print "\nEnter the new quantity (numbers only): "
									quantity = gets.chomp
								end
								#updates the quantity
								db.execute("UPDATE items SET Quantity = '#{quantity}' WHERE Barcode = '#{barcode}';")
								print "\nQuantity updated!\n"
							end
							#goes back through the function
							if answer == "n" or answer == "N"
								scan_barcode(db)
							end
						end
						#goes back through the function after the quantity has been changed
						scan_barcode(db)
					end			
					#decreasing the quantity
					db.execute("UPDATE items SET Quantity = Quantity-1 WHERE Barcode = '#{barcode}';")
					#printing the barcode information
					print "\nBarcode #{barcode} found in the database. Details are given\n"
					print "below: \n"
					for i in (1..5) do
						print "   "
						print field_names[i]
						print ": "
						print print_rows[i]
						print "\n"
					end
					print "\n"	
			else
				#ask if the user wants to add the barcode
				print "Cannot find barcode #{barcode}, would you like to add it to the database? [Y/N]: "
				answer1 = nil
				while answer1 != "y" and answer1 != "Y" and answer1 != "n" and answer1 != "N"
					answer1 = gets.chomp
					#collects the information for the new barcode
					if answer1 == "y" or answer1 == "Y"
						print "\nWhat is the item name? "
						item_name = gets.chomp
						print "\nWhat is the Item Category? "
						item_category = gets.chomp
						print "\nWhat is the Quantity? "
						quantity = gets.chomp
						#error checking...
						while !(quantity =~ /^[0-9]{1,20}$/)
							print "\nEnter the quantity (numbers only): "
							quantity = gets.chomp
						end
						print "\nWhat is the price? "
						price = gets.chomp
						print "\nWhat is the Description? "
						description = gets.chomp
						#adds the item to the database
						db.execute("INSERT INTO items VALUES(#{barcode},'#{item_name}','#{item_category}',#{quantity},#{price},'#{description}');")
						print "\nDatabase updated!\n"
					end
					#if they don't want to, go back to the beginning of the function
					if answer1 == "n" or answer1 == "N"
						scan_barcode(db)
					end
				end
			end
		else
			#end of error checking for the first error check for numbers
			print "You may only enter numbers.\n"
			scan_barcode(db)
		end
	end
end
#cmd line parameters, not finished but I ran out of time
def cmd_check
	if defined?(ARGV)
		#help function
		if ARGV.first == "-h" or ARGV.first == "-H" or ARGV.first == "?" or ARGV.first == "help" or ARGV.first == "Help"
			print "Usage: ruby inventory.rb [?|-h|help|[-u|-o|-z <infile>|[<outfile>]]]\n"
			print "\n"
			print "Parameters:\n"
			print "   ?                	displays this usage information\n"
			print "   -h               	displays this usage information\n"
			print "   Help			displays this usage information\n"
			print "   -u <infile>      	update the inventory using the  file  <infile>.\n"
			print "			The   filename  <infile>  must  have   a   .csv\n"
			print "                    	extension and it must be a  text file in  comma\n"
			print "                    	separated value  (CSV)  format.  Note that  the\n"
			print "			values must be in double quote.\n"
			print "   -z|-o [<outfile>]    output  either  the  entire  content   of   the\n"
			print "			database  (-o) or only those records for  which\n"
			print " 			the quantity is zero  (-z). If no <outifle>  is\n"
			print "			specified  then output on the console otherwise\n"
			print "                        output  in the text file named  <outfile>.  The\n"
			print "			output in both cases must be in a tab separated\n"
			print "			value (tsv) format.\n"
			exit(0)
		end
		if ARGV.first == "-u" or ARGV.first == "-U" 
			if ARGV[1].nil?
				print "no input file defined.\n"
			else
				print "input file " + ARGV[1] + " defined.\n"
			end
		end
		#this isn't working quite as intended
		if ARGV.first == "-z" or ARGV.first == "-Z"
			db = AccessDb.new('inventory.accdb')
			db.open
			if ARGV[1].nil?
				db.query("SELECT * FROM items WHERE Quantity = 0;")
				field_names = db.fields
				rows = db.data
				count = db.query("SELECT COUNT(*) FROM items WHERE Quantity = 0;")
				if(rows.length > 0)
					raw_string = rows.join(",")
					raw_string.gsub(/"/,'')
					print_rows = raw_string.split(",")
					for j in (0..count.to_i)
						for i in (0..5) do
							print field_names[i]
							print ": "
							print print_rows[i]
							print "\n"
						end
						print "----------------------------------------------------------\n"
					end
				end
			else
			
			end
			exit(0)
		end
	else
		main
	end
end
def main
	cmd_check
	check_file
end
main