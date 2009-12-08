require File.dirname(__FILE__) + '/../spec_helper'

describe EWS::Service do
  it "something simple to see green" do
    lambda do 
      EWS::Service.endpoint
    end.should raise_error(/Missing option :uri/)
  end
end
