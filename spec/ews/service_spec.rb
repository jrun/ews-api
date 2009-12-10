require File.dirname(__FILE__) + '/../spec_helper'

describe EWS::Service do
  DEFAULT_HEADERS = {
    'Content-Type' => 'text/xml; charset=utf-8'
  }
  
  def mock_response(soap_body, headers = {})
    Handsoap::Http.drivers[:mock] =
      Handsoap::Http::Drivers::MockDriver.new(:status => 200,
                                              :headers => DEFAULT_HEADERS.merge(headers),
                                              :content => soap_body)
    Handsoap.http_driver = :mock
  end
  
  context '#get_item' do    
    it "should raise an error when the id is not found" do
      mock_response response(:get_item_with_error)
      lambda do
        EWS::Service.get_item nil
      end.should raise_error(EWS::ResponseError, 'Id must be non-empty.')
    end
  end
end
