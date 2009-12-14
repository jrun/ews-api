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
    
    context 'parsing find_item with the base_shape of "AllProperties"' do
      before(:each) do
        @items = @parser.parse_find_item response_to_doc(:find_item_all_properties)
        @item = @items.first
      end

      it "should build an array of items" do
        @items.size.should == 3
      end

      it "should set the item id" do              
        @item.item_id.should == {:id => 'BUEm5G5U', :change_key => 'LABU0Ut45'}
      end

      it "should set the message body to nil" do
        @item.body.should be_nil
      end

      it "should set the attachments to nil because they are not in the response" do
        @item.attachments.should be_nil
      end

      it "should set the header to nil because the internet message headers are not in the response" do
        @item.header.should be_nil
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

      it "header" do        
        @message.header.size.should == 7

        @message.header['received'] = ['from example.com (0.0.0.0) by example.com (0.0.0.0) with SMTP Server id 0.0.0.0; Wed, 2 Dec 2009 10:38:47 -0600',
                    'from example.com ([0.0.0.0])  by example.com with SMTP; 02 Dec 2009 10:38:47 -0600',
                    'from example.com (0.0.0.0) by example.com (0.0.0.0) with SMTP Server id 0.0.0.0; Wed, 2 Dec 2009 10:38:47 -0600',
                    'from localhost by example.com;  02 Dec 2009 10:38:47 -0600']
        
        @message.header['subject'].should == ['A Clash of Kings']
        @message.header['message-id'].should == ['<847mtq@example.com>']
        @message.header['date'].should == ['Wed, 2 Dec 2009 10:38:47 -0600']
        @message.header['mime-version'].should == ['1.0']
        @message.header['content-type'].should == ['multipart/report']
        @message.header['return-path'].should == ['<>']
      end
    end

    context 'parsing get_item with no attachments' do
      before(:each) do
        @message = @parser.parse_get_item response_to_doc(:get_item_no_attachments)
      end

      it "should create an item with an empty array of attachments" do
        @message.attachments.should == []
      end
    end

    context 'parsing get_item with base_shape of "IdOnly"' do
      before(:each) do
        @message = @parser.parse_get_item response_to_doc(:get_item_id_only)        
      end

      it "should set the item_id" do
        @message.item_id[:id].should == 'AFTRNedwA'
        @message.item_id[:change_key].should == 'FwAAA'
      end
    end
    
    context 'parsing get_item with base_shape of "Default"' do
      before(:each) do
        @message = @parser.parse_get_item response_to_doc(:get_item_default)
      end
      
      it "should set the parent_folder_id to nil" do
        @message.parent_folder_id be_nil
      end

      it "should set the header to nil when the internet message headers are not present" do
        @message.header.should be_nil
      end
    end
    
  end
end
