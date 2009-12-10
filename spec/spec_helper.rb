$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'ews-api'
require 'spec'
require 'spec/autorun'

config_file = File.dirname(__FILE__) + '/test-config.yml'

if File.exist?(config_file)
  unless defined? EWS_CONFIG
    EWS_CONFIG = YAML.load_file config_file
  end
  
  EWS::Service.endpoint EWS_CONFIG['endpoint']
  EWS::Service.set_auth EWS_CONFIG['username'], EWS_CONFIG['password']
else
  unless defined? EWS_CONFIG
    EWS_CONFIG = nil
  end
  
  puts <<-EOS

=================================================================
Create 'spec/test-config.yml' to automatically configure
the endpoint and credentials. The file is ignored via
.gitignore so it will not be committed.

endpoint:
  :uri: 'https://localhost/ews/exchange.asmx'
  :version: 1
username: testuser
password: xxxxxx

=================================================================

EOS
end

module EWS::SpecHelper
  def response_to_doc(name)
    to_doc response(name)
  end
  
  def response(name)
    @responses ||= YAML.load_file File.dirname(__FILE__) + '/response_fixtures.yml'
    @responses[name.to_s]
  end

  def to_doc(xml)    
    doc = Handsoap::XmlQueryFront.parse_string xml, :nokogiri
    EWS::Service.apply_namespaces! doc
    doc
  end
end

Spec::Runner.configure do |config|
  config.include(EWS::SpecHelper)
end
