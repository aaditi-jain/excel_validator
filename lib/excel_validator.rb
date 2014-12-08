require "excel_validator/version"
module ExcelValidator
  class ValidateFile
    require "roo"   	
    
    def initialize(file_name,file_path)
      case File.extname(file_name)      
        when ".xls"
          @file =  Roo::Excel.new(file_path)
        when ".xlsx"
          @file = Roo::Excelx.new(file_path)
        else raise "Not a Excel / Excelx File"
      end
  	end 

    #counts the number of sheets in the file.
    #sheet_number: should be an integer
  	def number_of_sheets_is?(sheet_number)
      @file.sheets.count == sheet_number ? true : false  
  	end
  	  
    #checks the names of the sheets in the file
    #sheets_hash: should be a hash with key, value as position, name respectively. Example: {1: 'Bank', 2: 'Metric', 7: 'Country'}
  	def names_of_sheets_are?(sheets_hash, case_sensitive = false)  	  
  	  sheets_hash.each do |key, value|
        if ( case_sensitive ? (@file.sheets[(key.to_i - 1)] != value) : (@file.sheets[(key.to_i - 1)].to_s.downcase != value.to_s.downcase ))        
          return false
        end
  	  end
  	  true  	
    end
     
    #checks that the first non_empty row in the sheet is as specified
    def first_row_is?(sheet_number, row_number)
      set_default_sheet(sheet_number)
      @file.first_row == row_number ? true : false
    end
 
    #checks that the last non_empty row in the sheet is as specified
    def last_row_is?(sheet_number, row_number)
      set_default_sheet(sheet_number)
      @file.last_row == row_number ? true : false
    end 
    
    #checks that the first non_empty column in the sheet is as specified
    def first_column_is?(sheet_number, column_number)
      set_default_sheet(sheet_number)
      @file.first_column_as_letter == column_number ? true : false
    end
 
    #checks that the last non_empty column in the sheet is as specified
    def last_column_is?(sheet_number, column_number)
      set_default_sheet(sheet_number)
      @file.last_column_as_letter == column_number ? true : false
    end 

    #checks if the content at the specified row and column is as required
    def content_of_cell_is?(sheet_number, row, column, content, case_sensitive=false)  
      set_default_sheet(sheet_number)
      return ( (case_sensitive) ? (@file.cell(row,column).to_s == content.to_s) : (@file.cell(row,column).to_s.downcase == content.to_s.downcase))
    end


    #checks if the specified column is empty
    def column_is_empty?(sheet_number, column)
      set_default_sheet(sheet_number)  
      @file.first_row.upto @file.last_row do |row|   
        return false if @file.cell(row,column).nil? != true
      end
      true
    end

    #checks if the specified row is empty
    def row_is_empty?(sheet_number, row)
      @file.first_column.upto @file.last_column do |column|
        return false if @file.cell(row,column).nil? != true
      end 
      true       
    end

    #Returns true if there is an empty sheet in the file
    def empty_sheets_present?
      @file.sheets.count.times do |sheet_number|
        set_default_sheet(sheet_number + 1)
        @file.first_row ? next : (return true)
      end   
      false  
    end

    #Checks if a keyword is present in the sheet
    def keyword_present?(sheet_number, keyword, case_sensitive=false )
      set_default_sheet(sheet_number)
      @file.first_row.upto @file.last_row do |row|
        @file.first_column.upto @file.last_column do |column|
          return true if ( (case_sensitive) ? @file.cell(row, column) == keyword : @file.cell(row, column).to_s.downcase == keyword.to_s.downcase )
        end
      end
      false
    end

    private
      def set_default_sheet(sheet_number)
        @file.default_sheet = @file.sheets[sheet_number - 1]
      end

  end #class end
end #module end
