# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'excel_validator/version'

Gem::Specification.new do |spec|
  spec.name          = "excel_validator"
  spec.version       = ExcelValidator::VERSION
  spec.authors       = ["Aaditi Jain"]
  spec.email         = ["aaditi2290@gmail.com"]
  spec.description   = %q{provides methods to validates the content of excel file}
  spec.summary       = %q{provides methods to validates the content of excel file}
  spec.homepage      = "https://github.com/Aaditi/excel_validator.git"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rspec"
  
  spec.add_runtime_dependency "roo", "~> 1.13.2"
end
