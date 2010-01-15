require File.join(File.dirname(__FILE__), '..', '..', 'spec_helper')

module EWS
  describe ShapeBuilder do
    before(:each) do
      @doc = new_document
    end

    def builder(opts = {})
      ShapeBuilder.new(@doc, opts)
    end
    
    context "#item_shape!  should build an ItemShape node" do
      def expected_item_shape(base_shape)
      ( <<-EOS
<tns:ItemShape #{MNS}>
  <t:BaseShape #{TNS}>#{base_shape}</t:BaseShape>
</tns:ItemShape>
EOS
        ).chomp
      end
    
      def expected_include_mime_type(include_mime_type)
        ( <<-EOS
<tns:ItemShape #{MNS}>
  <t:BaseShape #{TNS}>Default</t:BaseShape>
  <t:IncludeMimeContent #{TNS}>#{include_mime_type}</t:IncludeMimeContent>
</tns:ItemShape>
EOS
          ).chomp
      end
      
      it "with the BaseShape of 'Default' when no :base_shape option is given" do
        builder.item_shape!
        @doc.to_s.should == expected_item_shape('Default')
      end
    
      it "with the BaseShape of 'Default'" do
        builder.item_shape!
        @doc.to_s.should == expected_item_shape('Default')
      end

      it "with the BaseShape of 'AllProperties'" do
        builder(:base_shape => 'AllProperties').item_shape!
        @doc.to_s.should == expected_item_shape('AllProperties')      
      end
      
      it "with the BaseShape of  'IdOnly'" do
        builder(:base_shape => 'IdOnly').item_shape!
        @doc.to_s.should == expected_item_shape('IdOnly')      
      end

      it "with IncludeMimeContent false" do
        builder(:include_mime_content => false).item_shape!
        @doc.to_s.should == expected_include_mime_type(false)
      end

      it "with IncludeMimeContent true" do
        builder(:include_mime_content => true).item_shape!
        @doc.to_s.should == expected_include_mime_type(true)
      end      
    end

    context "#folder_shape! should build an FolderShape node" do
      def expected_folder_shape(base_shape)
        ( <<-EOS
<tns:FolderShape #{MNS}>
  <t:BaseShape #{TNS}>#{base_shape}</t:BaseShape>
</tns:FolderShape>
EOS
          ).chomp
      end

      it "with the BaseShape of 'Default' when no :base_shape option is given" do
        builder.folder_shape!
        @doc.to_s.should == expected_folder_shape('Default')
      end
    
      it "with the BaseShape of 'Default'" do
        builder.folder_shape!
        @doc.to_s.should == expected_folder_shape('Default')
      end

      it "with the BaseShape of 'AllProperties'" do
        builder(:base_shape => 'AllProperties').folder_shape!
        @doc.to_s.should == expected_folder_shape('AllProperties')      
      end

      it "with the BaseShape of  'IdOnly'" do
        builder(:base_shape => 'IdOnly').folder_shape!
        @doc.to_s.should == expected_folder_shape('IdOnly')      
      end
    end  
  
    context "#attachment_shape! should build an AttachmentShape node" do
      def expected_empty_shape
        "<tns:AttachmentShape #{MNS} />"
      end
                
      def expected_include_mime_type(include_mime_type)
        ( <<-EOS
<tns:AttachmentShape #{MNS}>
  <t:IncludeMimeContent #{TNS}>#{include_mime_type}</t:IncludeMimeContent>
</tns:AttachmentShape>
EOS
          ).chomp
      end

      def expected_body_type(body_type)
        ( <<-EOS
<tns:AttachmentShape #{MNS}>
  <t:BodyType #{TNS}>#{body_type}</t:BodyType>
</tns:AttachmentShape>
EOS
          ).chomp  
      end
      
      it "an empty shape when no options are given" do
        builder.attachment_shape!
        @doc.to_s.should == expected_empty_shape
    end
    
      it "with IncludeMimeContent false" do
        builder(:include_mime_content => false).attachment_shape!
        @doc.to_s.should == expected_include_mime_type(false)
      end

      it "with IncludeMimeContent true" do
        builder(:include_mime_content => true).attachment_shape!
        @doc.to_s.should == expected_include_mime_type(true)
      end
    
      it "with BodyType Best" do
        builder(:body_type => :Best).attachment_shape!
        @doc.to_s.should == expected_body_type(:Best)      
      end
    end
    
  end
end
