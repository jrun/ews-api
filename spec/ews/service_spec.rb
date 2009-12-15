require File.dirname(__FILE__) + '/../spec_helper'

EWS::Service.endpoint 'http://localhost/ews/exchange.aspx'

describe EWS::Service do
  context '#get_item' do    
    it "should raise an error when the id is not found" do
      mock_response response(:get_item_with_error)
      lambda do
        EWS::Service.get_item nil
      end.should raise_error(EWS::ResponseError, 'Id must be non-empty.')
    end
  end
end
