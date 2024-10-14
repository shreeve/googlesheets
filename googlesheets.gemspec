# encoding: utf-8

Gem::Specification.new do |s|
  s.name        = "googlesheets"
  s.version     = `grep -m 1 '^\s*VERSION' lib/googlesheets.rb | head -1 | cut -f 2 -d '"'`
  s.author      = "Steve Shreeve"
  s.email       = "steve.shreeve@gmail.com"
  s.summary     = "Ruby library for Google Sheets"
  s.description = "This gem allows easy access to Google Sheets API V4."
  s.homepage    = "https://github.com/shreeve/googlesheets"
  s.license     = "MIT"
  s.files       = `git ls-files`.split("\n") - %w[.gitignore]
  s.executables = `cd bin && git ls-files .`.split("\n")
  s.add_runtime_dependency "censive", "~> 1.1"
  s.add_runtime_dependency "google-api-client", "~> 0.53.0"
  s.add_runtime_dependency "googleauth", "~> 1.11"
  s.add_runtime_dependency "pstore", "~> 0.1.3"
end
