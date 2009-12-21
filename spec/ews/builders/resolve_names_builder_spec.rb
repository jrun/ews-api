require File.join(File.dirname(__FILE__), '..', '..', 'spec_helper')

module EWS
  describe ResolveNamesBuilder do
    before(:each) do
      @doc = Handsoap::XmlMason::Document.new
      @doc.xml_header = nil
      EWS::Service.register_aliases! @doc
      
      @action_node = Handsoap::XmlMason::Element.new(@doc, 'tns', 'ResolveNames')
      @builder = ResolveNamesBuilder.new(@action_node)
    end

    context "#unresolved_entry!" do
      def expected_unresolved_entry_xml
        (<<-EOS
<tns:ResolveNames #{MNS}>
  <tns:UnresolvedEntry>Stark</tns:UnresolvedEntry>
</tns:ResolveNames>
EOS
         ).chomp
      end
      
      it "should build a UnresolvedEntry node" do
        @builder.unresolved_entry!('Stark')       
        @action_node.to_s.should == expected_unresolved_entry_xml
      end
    end

    context "#return_full_contact_data!" do
      it "should add attribute when true" do
        @builder.return_full_contact_data!(true)

        expected_xml = %Q|<tns:ResolveNames ReturnFullContactData="true" #{MNS} />|
        @action_node.to_s.should == expected_xml
      end

      it "should add attribute when false" do
        @builder.return_full_contact_data!(false)

        expected_xml = %Q|<tns:ResolveNames ReturnFullContactData="false" #{MNS} />|
        @action_node.to_s.should == expected_xml
      end
    end
  end
end
