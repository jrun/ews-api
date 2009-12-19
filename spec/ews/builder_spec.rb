require File.join(File.dirname(__FILE__), '..', 'spec_helper')

describe EWS::Builder do
  before(:each) do
    @builder = EWS::Builder.new
    @doc = Handsoap::XmlMason::Document.new
    @doc.xml_header = nil
    EWS::Service.register_aliases! @doc
  end
  
  TNS = %q|xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"|
  MNS = %q|xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages"|
    
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
      @builder.item_shape!(@doc)
      @doc.to_s.should == expected_item_shape('Default')
    end
    
    it "with the BaseShape of 'Default'" do
      @builder.item_shape!(@doc)      
      @doc.to_s.should == expected_item_shape('Default')
    end

    it "with the BaseShape of 'AllProperties'" do
      @builder.item_shape!(@doc, :base_shape => 'AllProperties')      
      @doc.to_s.should == expected_item_shape('AllProperties')      
    end

    it "with the BaseShape of  'IdOnly'" do
      @builder.item_shape!(@doc, :base_shape => 'IdOnly')      
      @doc.to_s.should == expected_item_shape('IdOnly')      
    end

    it "with IncludeMimeContent false" do
      @builder.item_shape!(@doc, :include_mime_content => false)
      @doc.to_s.should == expected_include_mime_type(false)
    end

    it "with IncludeMimeContent true" do
      @builder.item_shape!(@doc, :include_mime_content => true)
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
      @builder.folder_shape!(@doc)
      @doc.to_s.should == expected_folder_shape('Default')
    end
    
    it "with the BaseShape of 'Default'" do
      @builder.folder_shape!(@doc)      
      @doc.to_s.should == expected_folder_shape('Default')
    end

    it "with the BaseShape of 'AllProperties'" do
      @builder.folder_shape!(@doc, :base_shape => 'AllProperties')      
      @doc.to_s.should == expected_folder_shape('AllProperties')      
    end

    it "with the BaseShape of  'IdOnly'" do
      @builder.folder_shape!(@doc, :base_shape => 'IdOnly')      
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
      @builder.attachment_shape!(@doc)
      @doc.to_s.should == expected_empty_shape
    end
    
    it "with IncludeMimeContent false" do
      @builder.attachment_shape!(@doc, :include_mime_content => false)
      @doc.to_s.should == expected_include_mime_type(false)
    end

    it "with IncludeMimeContent true" do
      @builder.attachment_shape!(@doc, :include_mime_content => true)
      @doc.to_s.should == expected_include_mime_type(true)
    end
    
    it "with BodyType Best" do
      @builder.attachment_shape!(@doc, :body_type => :Best)
      @doc.to_s.should == expected_body_type(:Best)      
    end
  end
  
  context '#item_id! should build ItemId' do
    def expected_item_id
      %Q|<t:ItemId #{TNS} Id="XYZ" />|
    end

    def expected_item_id_with_change_key
      %Q|<t:ItemId #{TNS} ChangeKey="123" Id="XYZ" />|
    end
    
    it "given an id" do
      @builder.item_id!(@doc, 'XYZ')
      @doc.to_s.should == expected_item_id
    end

    it "with a ChangeKey whe the :change_key option is given" do
      @builder.item_id!(@doc, 'XYZ', :change_key => '123')
      @doc.to_s.should == expected_item_id_with_change_key
    end
  end
  
  context '#folder_id!' do
    context 'should build FolderId' do
      def expected_folder_id
        %Q|<t:FolderId #{TNS} Id="XYZ" />|
      end

      def expected_folder_id_with_change_key
        %Q|<t:FolderId #{TNS} ChangeKey="123" Id="XYZ" />|
      end
      
      it "given an id" do
        @builder.folder_id!(@doc, 'XYZ')
        @doc.to_s.should == expected_folder_id
      end
      
      it "with a ChangeKey whe the :change_key option is given" do
        @builder.folder_id!(@doc, 'XYZ', :change_key => '123')
        @doc.to_s.should == expected_folder_id_with_change_key
      end
    end
  
    context "should build a DistinguishedFolderId" do
      def expected_distinguished_folder_id(id)
        %Q|<t:DistinguishedFolderId #{TNS} Id="#{id}" />|
      end
      
      def expected_distinguished_folder_id_with_change_key
        %Q|<t:DistinguishedFolderId #{TNS} ChangeKey="222" Id="inbox" />|
      end
    
      EWS::DistinguishedFolders.each do |folder|
        it "given #{folder.inspect}" do
          @builder.folder_id!(@doc, folder)
          @doc.to_s.should == expected_distinguished_folder_id(folder)
        end
      end
      
      it "with a ChangeKey when given the :change_key option" do
        @builder.folder_id!(@doc, :inbox, :change_key => '222')
        @doc.to_s.should == expected_distinguished_folder_id_with_change_key
      end
    end
  end
  
  context "#parent_folder_ids should bulid ParentFolderIds" do
    def expected_parent_folder_ids_with_one
      (<<-EOS
<tns:ParentFolderIds #{MNS}>
  <t:FolderId #{TNS} Id="111" />
</tns:ParentFolderIds>
EOS
      ).chomp
    end

    def expected_parent_folder_ids_with_one_distinguished
          (<<-EOS
<tns:ParentFolderIds #{MNS}>
  <t:DistinguishedFolderId #{TNS} Id="inbox" />
</tns:ParentFolderIds>
EOS
      ).chomp  
    end
    
    def expected_parent_folder_ids_with_more
      (<<-EOS
<tns:ParentFolderIds #{MNS}>
  <t:FolderId #{TNS} Id="111" />
  <t:DistinguishedFolderId #{TNS} Id="inbox" />
</tns:ParentFolderIds>
EOS
       ).chomp
    end
        
    it "when given one id" do
      @builder.parent_folder_ids!(@doc, '111')
      @doc.to_s.should == expected_parent_folder_ids_with_one
    end

    it "when given one distinguished folder" do
      @builder.parent_folder_ids!(@doc, :inbox)
      @doc.to_s.should == expected_parent_folder_ids_with_one_distinguished
    end

    it "when given an array of ids" do
      @builder.parent_folder_ids!(@doc, ['111', :inbox])
      @doc.to_s.should == expected_parent_folder_ids_with_more
    end
  end
  
end
