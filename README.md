# ExcelValidator

  Validates the format of your imported excel (xls, xlsx) file.

## Installation

Add this line to your application's Gemfile:
  `gem 'excel_validator'`

And then execute:
  `$ bundle`

Or install it yourself as:
  `$ gem install excel_validator`

## Usage
  **require 'excel_validator'**

  **f = ExcelValidator::ValidateFile.new(file_name, file_path)**

  #returns true if the number of sheets in the file is 5.
  
  **f.number_of_sheets_is?(5)**
  
  #returns true if the names of the sheets at each position is as specified. If the names are case sensitive , do  f.names_of_sheets_are?({1: 'Banks', 2: 'Countries', 3: 'Metrics'}, true)
  
  **f.names_of_sheets_are?({1: 'Banks', 2: 'Countries', 3: 'Metrics'})**
 

  #returns true if the first non-empty row of 2nd sheet in the file is 12

  **f.first_row_is?(2, 5)**

  
  #returns true if the first non-empty row of 2nd sheet in the file is 100
  
  **f.last_row_is?(2,100)**

  
  #returns true if the first non-empty row of 2nd sheet in the file is 'B'
  
  **f.first_column_is?(2, 'B')**
  
  
  #returns true if the first non-empty row of 2nd sheet in the file is 'CK'
  
  **f.last_column_is?(2, 'CK')**

  
  #returns true if the 'B' column of 2nd sheet in the file is completely empty
  
  **f.column_is_empty?(2, 'B')**

  
  #returns true if the 12th row of 2nd sheet in the file is completely empty
  
  **f.row_is_empty?(2, 12)**

  
  #returns true if at 2nd sheet in the file Row3, Column C, content is 'Location'. If the content is case sensitive , do  
  
  **f.content_of_cell_is?(2, 3, 'C', 'Location', true)**

  
  #returns true if there are empty sheets present in the file.
  
  **f.empty_sheets_present?**
  
  
  #returns true if keyword 'ClassA' is present the 2nd sheet of the file. If the keyword is case sensitive, do f.keyword_present?(2, 'ClassA', true)
  
  **f.keyword_present?(2, 'ClassA')**




## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
