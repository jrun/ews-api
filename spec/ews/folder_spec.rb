require File.dirname(__FILE__) + '/../spec_helper'

module EWS
  describe Folder do
    
    before(:each) do
      @folder = Folder.new(:name => 'Inbox')
    end

    it "should iterate through each message" do
      mock_response response(:find_item_all_properties)
      
      message_count = 0
      subjects = ['test', 'Regarding Brandon Stark', 'Re: Regarding Brandon Stark']
      
      @folder.each_message do |message|
        message_count += 1
        message.subject.should == subjects.shift
      end
  
      message_count.should == 3
    end
    
  end
end
