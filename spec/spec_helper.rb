$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'ews-api'
require 'spec'
require 'spec/autorun'

TNS = %q|xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"|
MNS = %q|xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages"|  

module EWS::SpecHelper  
  def fixtures
    @fixtures ||= Dir[File.dirname(__FILE__) + '/fixtures/*.xml'].inject({}) do |fixtures, f|
      fixtures[File.basename(f).chomp('.xml')] = File.expand_path(f)
      fixtures
    end
  end
  
  def response_to_doc(name)
    to_doc response(name)
  end
  
  def response(name)
    response = File.read fixtures[name.to_s]
    wrap_in_soap_response response
  end

  def to_doc(xml)    
    doc = Handsoap::XmlQueryFront.parse_string xml, :nokogiri
    EWS::Service.apply_namespaces! doc
    doc
  end

  def mock_response(soap_body, headers = {}, status = 200)
    Handsoap::Http.drivers[:mock] =
      Handsoap::Http::Drivers::MockDriver.new(:status => status,
                                              :headers => default_headers.merge(headers),
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
  
  def new_document
    doc = Handsoap::XmlMason::Document.new
    doc.xml_header = nil
    EWS::Service.register_aliases! doc
    doc
  end
end

Spec::Runner.configure do |config|
  config.include(EWS::SpecHelper)
end
