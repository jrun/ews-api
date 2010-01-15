module EWS
  class ResolveNamesBuilder
    def initialize(action_node)
      @action_node = action_node
    end

    def unresolved_entry!(entry)
      @action_node.add('tns:UnresolvedEntry', entry)
    end

    def return_full_contact_data!(return_full_contact_data)
      @action_node.set_attr 'ReturnFullContactData', return_full_contact_data
    end    
  end
end
