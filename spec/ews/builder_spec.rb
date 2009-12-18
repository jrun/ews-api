require File.join(File.dirname(__FILE__), '..', 'spec_helper')

describe EWS::Builder do
  before(:each) do
    @builder = EWS::Builder.new
    @doc = Handsoap::XmlMason::Document.new
    @doc.xml_header = nil
    EWS::Service.register_aliases! @doc
  end
  
  TNS = %q|xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"|
     
  context "#base_shape!" do
    def expected_base_shape(shape)
      %Q|<t:BaseShape #{TNS}>#{shape}</t:BaseShape>|
    end

    it "should build a BaseShape node with 'Default'" do
      @builder.base_shape!(@doc)      
      @doc.to_s.should == expected_base_shape('Default')
    end

    it "should build a BaseShape node with 'AllProperties'" do
      @builder.base_shape!(@doc, :base_shape => 'AllProperties')      
      @doc.to_s.should == expected_base_shape('AllProperties')      
    end

    it "should build a BaseShape node with 'IdOnly'" do
      @builder.base_shape!(@doc, :base_shape => 'IdOnly')      
      @doc.to_s.should == expected_base_shape('IdOnly')      
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
  
  context '#folder_id! should build FolderId' do
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
  
  context "#distinguished_folder_id! should build a DistinguishedFolderId" do
    def expected_distinguished_folder_id
      %Q|<t:DistinguishedFolderId #{TNS} Id="inbox" />|
    end

    def expected_distinguished_folder_id_with_change_key
      %Q|<t:DistinguishedFolderId #{TNS} ChangeKey="222" Id="inbox" />|
    end
    
    it "given a symbol as the folder name " do
      @builder.distinguished_folder_id!(@doc, :inbox)
      @doc.to_s.should == expected_distinguished_folder_id
    end
    
    it "givne  a capitalized string as the folder name " do
      @builder.distinguished_folder_id!(@doc, 'Inbox')
      @doc.to_s.should == expected_distinguished_folder_id
    end
    
    it "with a ChangeKey when the :change_key option is given" do
      @builder.distinguished_folder_id!(@doc, :inbox, :change_key => '222')
      @doc.to_s.should == expected_distinguished_folder_id_with_change_key
    end
  end
  

end
