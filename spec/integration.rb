require File.dirname(__FILE__) + '/spec_helper'

EWS::Service.logger = $stdout

describe 'Integration Tests' do
  context 'find_folder' do
    it "should find the folder without errors" do
      lambda do
        EWS::Service.find_folder(:inbox)
      end.should_not raise_error
    end

    it "should raise a Fault when the item does not exist" do
      error_msg = /The value 'does-not-exist' is invalid according to its datatype/
      lambda do
        EWS::Service.find_folder('does-not-exist')
      end.should raise_error(Handsoap::Fault,  error_msg)
    end
  end

  context 'get_folder' do
    it "should get the folder without errors" do
      EWS::Service.get_folder(:inbox).should be_instance_of(EWS::Folder)
    end
    
    it "should raise a SoapError when the ResponseMessage is an Error" do
      error_msg = /The value 'does-not-exist' is invalid according to its datatype/
      lambda do
        EWS::Service.get_folder('does-not-exist')
      end.should raise_error(Handsoap::Fault,  error_msg)
    end
  end

  context 'find_item' do
    it "should find the item without errors" do
      lambda do
        EWS::Service.find_item(:inbox)
      end.should_not raise_error
    end

    it "should raise a Fault when the item does not exist" do
      error_msg = /The value 'does-not-exist' is invalid according to its datatype/
      lambda do
        EWS::Service.find_item('does-not-exist')
      end.should raise_error(Handsoap::Fault,  error_msg)
    end
  end
  
  context 'get_item' do
    it "should get the item without errors" do
      lambda do
        EWS::Service.get_item EWS_CONFIG['item_id'], :base_shape => :AllProperties
      end.should_not raise_error
    end

    it "should raise a SoapError when the ResponseMessage is an Error" do
      lambda do
        EWS::Service.get_item(nil)
      end.should raise_error(EWS::ResponseError, /Id must be non-empty/)
    end
  end

  context 'get_attachment' do
    it "should get the attachment without errors" do
      lambda do
        rtn = EWS::Service.get_attachment EWS_CONFIG['attachment_id'] # lame
      end.should_not raise_error
    end

    it "should raise a SoapError when the ResponseMessage is an Error" do
      lambda do
        EWS::Service.get_attachment(nil)
      end.should raise_error(EWS::ResponseError, /Id must be non-empty/)
    end
  end

end
