# encoding: utf-8

Gem::Specification.new do |s|
  s.name        = "googlesheets"
  s.version     = `grep -m 1 '^\s*VERSION' lib/googlesheets.rb | head -1 | cut -f 2 -d '"'`
  s.author      = "Steve Shreeve"
  s.email       = "steve.shreeve@gmail.com"
  s.summary     = "Ruby gem for Google Sheets"
  s.description = "This gem allows easy access to Google Sheets API V4."
  s.homepage    = "https://github.com/shreeve/googlesheets"
  s.license     = "MIT"
  s.files       = `git ls-files`.split("\n") - %w[.gitignore]
  s.add_runtime_dependency "google-api-client", "~> 0.53.0"
end
