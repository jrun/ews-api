require File.dirname(__FILE__) + '/../spec_helper'

describe EWS::Service do
  DEFAULT_HEADERS = {
    'Content-Type' => 'text/xml; charset=utf-8'
  }

  def mock_body(soap_body)
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
  
  def mock_response(soap_body, headers = {})
    Handsoap::Http.drivers[:mock] =
      Handsoap::Http::Drivers::MockDriver.new(:status => 200,
                                              :headers => DEFAULT_HEADERS.merge(headers),
                                              :content => mock_body(soap_body))
    Handsoap.http_driver = :mock
  end
  
  context '#get_item' do    
    it "should raise an error when the id is not found" do
      mock_response <<-EOS
<m:GetItemResponse xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <m:ResponseMessages>
    <m:GetItemResponseMessage ResponseClass="Error">
      <m:MessageText>Id must be non-empty.</m:MessageText>
      <m:ResponseCode>ErrorInvalidIdEmpty</m:ResponseCode>
      <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
      <m:Items/>
    </m:GetItemResponseMessage>
  </m:ResponseMessages>
</m:GetItemResponse>
EOS
      lambda do
        EWS::Service.get_item nil
      end.should raise_error(EWS::ResponseError, 'Id must be non-empty.')
    end
  end

end
