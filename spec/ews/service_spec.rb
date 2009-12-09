require File.dirname(__FILE__) + '/../spec_helper'

EWS::Service.logger = $stdout

describe EWS::Service do
  it "something simple to see green" do
    lambda do 
      EWS::Service.endpoint
    end.should raise_error(/Missing option :uri/)
  end

  context '#find_folder' do
    it "should successfully retrieve the inbox" do
      lambda do
        EWS::Service.find_folder
      end.should_not raise_error
    end
  end

  context '#get_folder' do
    it "should successfully retrieve the inbox" do
      lambda do
        EWS::Service.get_folder
      end.should_not raise_error
    end
  end

  context '#find_item' do
    it "should successfully find the item" do
      lambda do
        EWS::Service.find_item(:inbox)
      end.should_not raise_error
    end
  end
  
  context '#get_item' do
    it "should be successfull" do
      lambda do
        EWS::Service.get_item EWS_CONFIG['item_id'] # lame
      end.should_not raise_error
    end
    
    it "should raise an error when the id is not found" do
      pending do
        lambda do
          EWS::Service.get_item nil
        end.should raise_error
      end
    end
  end

  context '#get_attachment' do
    it "should successfully get the attachment" do
      lambda do
        EWS::Service.get_attachment EWS_CONFIG['attachment_id'] # lame 
      end.should_not raise_error
    end
  end
  
end
