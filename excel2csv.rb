# frozen_string_literal: true

require 'bundler/inline'

gemfile(true) do
  source 'https://rubygems.org'

  git_source(:github) { |repo| "https://github.com/#{repo}.git" }

  gem 'roo'
  gem 'base64'
  gem 'csv'
  gem 'optparse'
end

require 'roo'
require 'optparse'

def output_csv(sheet, sheet_name)
  puts "processing #{sheet}"
  header = (list = sheet.entries.to_a).shift.map{ |e| e.gsub /\s+/, '' }.map { |str| str.unicode_normalize(:nfkc)}
  CSV.open("#{sheet_name}.csv", "w") do |csv|
    csv << header
    list.each { |r| csv << r.map { |str| str.to_s.unicode_normalize(:nfkc)} }
  end
end

args = ARGV.getopts(nil, 'in:', 'sheets:--').transform_keys!(&:to_sym).map do |k, v|
  [
    k,
    {
      in: ->(value) { value },
      sheets: ->(value) { value.split(',') }
    }[k].call(v),
  ]
end.to_h

(wb = Roo::Excelx.new(args[:in])).sheets.each do |sh|
  output_csv(wb.sheet(sh), sh) if args[:sheets].include?(sh) || args[:sheets].include?('--')
end
