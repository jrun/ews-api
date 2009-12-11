require File.dirname(__FILE__) + '/../spec_helper'

module EWS
   
  describe Parser do
    before(:each) do
      @parser = Parser.new
    end
    
    context 'parsing get_folder should build a folder with' do
      before(:each) do
        @folder = @parser.parse_get_folder response_to_doc(:get_folder)
      end

      { :folder_id => 'AAA',
        :change_key => 'BBB',
        :name => 'Inbox' }.each do |attr, value|

        it attr do
          @folder.send(attr).should == value
        end
      end      
    end

    context 'parsing get_item with all properites should build a message with' do
      MESSAGE_BODY = "<html><body>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut ullamcorper egestas accumsan. Nulla rhoncus porttitor dictum. Curabitur ac nulla a orci sollicitudin blandit at quis odio. Integer eleifend fringilla bibendum. Sed condimentum, lectus vitae suscipit sollicitudin, quam tortor faucibus nisl, ac tristique lectus justo in leo. Aliquam placerat arcu erat, vel lobortis dui. Nulla posuere sodales sapien eu interdum. Cras in pretium ante. Pellentesque ut velit est. Quisque nisl risus, molestie vel porta in, pellentesque a odio. Curabitur sit amet faucibus lectus. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Maecenas et fermentum enim. Nulla ac dolor elit. Quisque eu dui sapien, ac tincidunt orci.</body></html>"

      before(:each) do
        @message = @parser.parse_get_item  response_to_doc(:get_item_all_properties)
      end
      
      { :item_id          => {:id => 'HRlZ', :change_key => '0Tk4V'},
        :parent_folder_id => {:id => 'AQAUAAA', :change_key => 'AQAAAA=='},
        :subject          => 'A Clash of Kings',
        :body             => MESSAGE_BODY,
        :body_type        => 'HTML',
        :has_attachments  => true
      }.each do |attr, value|
        
        it attr do
          @message.send(attr).should == value
        end
      end

      it "attachments" do
        @message.attachments.size.should == 1
        
        attachment = @message.attachments.first
        attachment.attachment_id.should == "UAHR"
        attachment.name.should == "A Clash of Kings"
        attachment.content_type.should == "message/rfc822"
      end
    end

    context 'parsing get_item with no attachments' do
      before(:each) do
        @message = @parser.parse_get_item response_to_doc(:get_item_no_attachments)
      end

      it "should create a message with an empty array of attachments" do
        @message.attachments.should == []
      end
    end
    
  end
end
