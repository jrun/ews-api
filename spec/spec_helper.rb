$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'ews-api'
require 'spec'
require 'spec/autorun'

module EWS::SpecHelper
  def response_to_doc(name)
    to_doc response(name)
  end
  
  def response(name)
    @responses ||= YAML.load_file File.dirname(__FILE__) + '/response_fixtures.yml'
    wrap_in_soap_response @responses[name.to_s]
  end

  def to_doc(xml)    
    doc = Handsoap::XmlQueryFront.parse_string xml, :nokogiri
    EWS::Service.apply_namespaces! doc
    doc
  end

  def mock_response(soap_body, headers = {}, status = 200)
    Handsoap::Http.drivers[:mock] =
      Handsoap::Http::Drivers::MockDriver.new(:status => status,
                                              :headers => default_headers,
                                              :content => soap_body)
    Handsoap.http_driver = :mock
  end

  def default_headers
    @default_headers ||= {
      'Content-Type' => 'text/xml; charset=utf-8'
    }
  end
  
  def wrap_in_soap_response(soap_body)
    <<-EOS
<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Header>
    <t:ServerVersionInfo xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" MajorVersion="1" MinorVersion="1" MajorBuildNumber="1" MinorBuildNumber="1"/>
  </soap:Header>
  <soap:Body>
  #{soap_body}
  </soap:Body>
</soap:Envelope>
EOS
  end
end

Spec::Runner.configure do |config|
  config.include(EWS::SpecHelper)
end
