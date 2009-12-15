require File.dirname(__FILE__) + '/../spec_helper'

module EWS

  describe Attachment do
    before(:each) do
      @parser = Parser.new
      @attachment = @parser.parse_get_attachment response_to_doc(:get_attachment)
    end
    
    it "should have an attachment_id" do
      @attachment.attachment_id.should == 'CCC'
    end
    
    it "id should be the same as the attachment id" do
      @attachment.id.should == @attachment.attachment_id
    end

    it "should have a content_type" do
      @attachment.content_type.should == 'message/rfc822'
    end
    
    it "item should be a Message" do
      @attachment.item.should be_instance_of(Message)
      @attachment.item.subject.should == 'I am subject. Hear me roar.'
    end
  end
  
end
