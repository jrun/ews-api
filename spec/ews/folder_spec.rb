require File.dirname(__FILE__) + '/../spec_helper'

module EWS
  describe Folder do
    before(:each) do
      @parser = Parser.new
    end

    context "#iems" do
      it "should call service#find_item with the folder id" do
        service = mock(:service)
        service.should_receive(:find_item).with('maid of tarth', anything())
        
        folder = Folder.new(:folder_id => {:id => 'maid of tarth'})
        folder.service = service
        
        folder.items
      end      
    end
    
    context '#folders' do
      it "should be indexed by display name" do
        mock_response response(:find_folder)
        folder = Folder.new        
        folder.folders['subfolder1'].should_not be_nil
        folder.folders['subfolder2'].should_not be_nil
      end
    end

    context '#each_message' do
      it "should iterate through the messages" do
        mock_response response(:find_item_all_properties)
        
        message_count = 0
        subjects = ['test', 'Regarding Brandon Stark', 'Re: Regarding Brandon Stark']
        folder = Folder.new(:display_name => 'Inbox', :folder_id => {:id => 'maid of tarth'})
        
        folder.each_message do |message|
          message_count += 1
          message.subject.should == subjects.shift
        end
        
        message_count.should == 3
      end
    end
    
    context 'get' do
      before(:each) do
        @folder = @parser.parse_get_folder response_to_doc(:get_folder)
      end
      
      it "id should be the folder id id" do
        @folder.id.should == 'AAA'
      end
      
      it "change ley should be the folder id change key" do
        @folder.change_key.should == 'BBB'
      end

      it "should have a folder_id" do
        @folder.folder_id.should == {:id => 'AAA', :change_key => 'BBB'}
      end

      it "should have a display name" do
        @folder.display_name.should == 'Inbox'
      end

      it "name should be the same as display name" do
        @folder.name.should == @folder.display_name
      end

      it "should have a total count" do
        @folder.total_count.should == 43
      end

      it "should have a child folder count" do
        @folder.child_folder_count.should == 2
      end

      it "should have a unread count" do
        @folder.unread_count.should == 0
      end
    end

    context 'find' do
      before(:each) do
        @folders = @parser.parse_find_folder response_to_doc(:find_folder)
      end

      it "should build an array of folders" do
        @folders.size.should > 0
        @folders.first.should be_instance_of(Folder)
      end
    end
    
  end
end
