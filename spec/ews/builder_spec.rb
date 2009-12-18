require File.join(File.dirname(__FILE__), '..', 'spec_helper')

describe EWS::Builder do
  before(:each) do
    @builder = EWS::Builder.new
    @doc = Handsoap::XmlMason::Document.new
    @doc.xml_header = nil
    EWS::Service.register_aliases! @doc
  end

  context "#base_shape!" do
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

  def expected_base_shape(shape)
    %Q|<t:BaseShape xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">#{shape}</t:BaseShape>|
  end
end
