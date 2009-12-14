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
  end
  
end
