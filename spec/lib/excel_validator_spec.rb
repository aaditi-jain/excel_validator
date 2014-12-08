require 'spec_helper'

describe ExcelValidator do    
  before (:all) do   	 
    @input_file = ExcelValidator::ValidateFile.new("test_excel.xlsx", "spec/fixtures/test_excel.xlsx")
  end

  describe '#number_of_sheets_is?' do 
    it 'should return true if input integer is equal to number sheets in the file' do
      @input_file.number_of_sheets_is?(3).should eq(true)
    end 	

    it 'should return false if input integer is not equal to number sheets in the file' do
      @input_file.number_of_sheets_is?(2).should eq(false)
   	end 
  end

  describe '#first_row_is?' do 
    it 'should return true if start row of the specified sheet is as in the arguments' do
      @input_file.first_row_is?(1, 2).should eq(true)
    end 	

    it 'should return false if start row of the specified sheet is not as in the arguments' do
      @input_file.first_row_is?(1, 6).should eq(false)
   	end 
  end

  describe '#last_row_is?' do 
    it 'should return true if last row of the specified sheet is as in the arguments' do
      @input_file.last_row_is?(1, 5).should eq(true)
    end   

    it 'should return false if last row of the specified sheet is not as in the arguments' do
      @input_file.last_row_is?(1, 2).should eq(false)
    end     
  end

  describe '#first_column_is?' do 
    it 'should return true if start column of the specified sheet is as in the arguments' do
      @input_file.first_column_is?(1, 'B').should eq(true)
    end 	

    it 'should return false if start column of the specified sheet is not as in the arguments' do
      @input_file.first_column_is?(1, 'C').should eq(false)
   	end 
  end

  describe '#last_column_is?' do 
    it 'should return true if last column of the specified sheet is as in the arguments' do
      @input_file.last_column_is?(1, 'E').should eq(true)
    end   

    it 'should return false if last column of the specified sheet is not as in the arguments' do
      @input_file.last_column_is?(1, 'G').should eq(false)
    end     
  end  

  describe '#column_is_empty?' do 
    it 'should return true if specified column in the specified sheet is empty' do
      @input_file.column_is_empty?(1, 'A').should eq(true)
    end   

    it 'should return false if specified column in the specified sheet is not empty' do
      @input_file.column_is_empty?(1, 'B').should eq(false)
    end     
  end

  describe '#row_is_empty?' do 
    it 'should return true if specified row in the specified sheet is empty' do
      @input_file.row_is_empty?(1, 6).should eq(true)
    end   

    it 'should return false if specified row in the specified sheet is not empty' do
      @input_file.row_is_empty?(1, 2).should eq(false)
    end     
  end

  describe '#empty_sheets_present?' do 
    it 'should return true if there are empty sheets present in the file' do
      @input_file.empty_sheets_present?.should eq(true)
    end   
  end

  describe '#keyword_present?' do 
    it 'should return true if keyword (case insensitve) is present in sheet ' do
      @input_file.keyword_present?(1,"Name").should eq(true)
    end   

    it 'should return false if keyword (case insensitve) is not present in sheet ' do
      @input_file.keyword_present?(1, "Absent").should eq(false)
    end  

    it 'should return true if keyword(case sensitve) is present in sheet ' do
      @input_file.keyword_present?(1,"Name", true).should eq(true)
    end   

    it 'should return false if keyword (case sensitve) is not present in sheet' do
      @input_file.keyword_present?(1, "name", true).should eq(false)
    end        
  end

  describe "#names_of_sheets_are?" do
    it 'should return true if names of sheets in file are as specified in the input hash(case insensitve)' do
      @input_file.names_of_sheets_are?({1 => 'users', 2 => 'country'}).should eq(true)
    end  

    it 'should return false if names of sheets in file are not as specified in the input hash(case insensitve)' do
      @input_file.names_of_sheets_are?({1 => 'Users', 2 => 'Metric'}).should eq(false)
    end

    it 'should return true if names of sheets in file are as specified in the input hash(case sensitve)' do
      @input_file.names_of_sheets_are?({1 => 'Users', 2 => 'Country'}, true).should eq(true)
    end  

    it 'should return false if names of sheets in file are not as specified in the input hash(case sensitve)' do
      @input_file.names_of_sheets_are?({1 => 'users', 2 => 'country'}, true).should eq(false)
    end   
  end

  describe "#content_of_cell_is?" do
  	it 'should return true if content of the cell is as specified(case insensitve)' do
      @input_file.content_of_cell_is?(2, 3, 'B', 'India').should eq(true)
  	end

  	it 'should return false if content of the cell is not as specified(case insensitve)' do
      @input_file.content_of_cell_is?(2, 3, 'B', 'Countries').should eq(false)
  	end  	

    it 'should return true if content of the cell is as specified(case sensitve)' do
      @input_file.content_of_cell_is?(2, 3, 'B', 'India', true).should eq(true)
    end

    it 'should return false if content of the cell is not as specified(case sensitve)' do
      @input_file.content_of_cell_is?(2, 3, 'B', 'india', true).should eq(false)
    end   
  end
  
end