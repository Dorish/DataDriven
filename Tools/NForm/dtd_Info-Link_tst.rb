=begin
This script expects to be executed in the same directory where the NFrom DTD
files are located.  This file can be copied to the DTD dir  -or- the DTD files
can be copied to this directory temporarily.

DTD file location in Nform is C:\nform\bin\mds\ddr

=end


#$:.unshift File.dirname(__FILE__)#.chomp('driver/web')<<'lib' # add library to path
require 'rubygems'
require 'nokogiri'
require 'uri'
require 'net/http'


def fetch(uri_str, limit = 10)
  # You should choose a better exception.
  raise ArgumentError, 'too many HTTP redirects' if limit == 0

  response = Net::HTTP.get_response(URI(uri_str))

  case response
  when Net::HTTPSuccess then
    response
  when Net::HTTPRedirection then
    location = response['location']
    puts "redirected to #{location}"
    fetch(location, limit - 1)
  else
    response.value
  end
end

Dir.chdir(File.dirname(__FILE__)) # change to directory of this file
Dir.foreach(".") do |file|
  case file
  when /^\.|\S+\.rb/  # do nothing with the dots or *.rb
  else
  doc = Nokogiri::XML.parse File.read(file)

    linkname = doc.css('LinkName').map{|x| x.text}  # <LinkName>
    
    puts "-----------  #{file}  ------------------"
    url = doc.css('Uri').map do |y| url = y.text    # <Uri>
    puts "   "+linkname.shift  # pull link name that goes with url
    puts url

    site = Net::HTTP.get_response(URI.parse(url.to_s))

    if site.code =~ /4\d+|302/   # don't fetch (redirect) if code = 4xx or 302
      p site
    else
      p site if site.code =~ /301/  # Show
      p fetch(url.to_s)
    end
  end
  puts"-----------------------------------------------\n\n"
  end
end

#r.code = 200 | 404 | 500, etc
#r.body = *text of page*
# http://www.checkupdown.com/status/E301.html
# https://github.com/augustl/net-http-cheat-sheet
# 
