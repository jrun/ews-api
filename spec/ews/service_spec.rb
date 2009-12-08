require File.dirname(__FILE__) + '/../spec_helper'

EWS::Service.logger = $stdout

describe EWS::Service do
  it "something simple to see green" do
    lambda do 
      EWS::Service.endpoint
    end.should raise_error(/Missing option :uri/)
  end

  context '#get_folder' do
    it "should successfully retrieve the inbox" do
      lambda do
        EWS::Service.get_folder('inbox')
      end.should_not raise_error
    end
  end
  
end
