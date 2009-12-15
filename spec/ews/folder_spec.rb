require File.dirname(__FILE__) + '/../spec_helper'

module EWS
  describe Folder do
    before(:each) do
      @parser = Parser.new
      @folder = @parser.parse_get_folder response_to_doc(:get_folder)
    end

    it "#id should be the folder_id id" do
      @folder.id.should == 'AAA'
    end

    it "#change_key should be the folder_id change_key" do
      @folder.change_key.should == 'BBB'
    end

    it "should have a folder_id" do
      @folder.folder_id.should == {:id => 'AAA', :change_key => 'BBB'}
    end

    it "should have a name" do
      @folder.name.should == 'Inbox'
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
