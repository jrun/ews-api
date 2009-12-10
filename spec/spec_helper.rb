$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'ews-api'
require 'spec'
require 'spec/autorun'

EWS::Service.logger = $stdout

config_file = File.dirname(__FILE__) + '/test-config.yml'

if File.exist?(config_file)
  EWS_CONFIG = YAML.load_file config_file
  
  EWS::Service.endpoint EWS_CONFIG['endpoint']
  EWS::Service.set_auth EWS_CONFIG['username'], EWS_CONFIG['password']
else
  EWS_CONFIG = nil
  
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

Spec::Runner.configure do |config|
end
