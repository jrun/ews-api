module EWS
  
  class Builder
    def base_shape!(shape, opts = {})
      opts[:base_shape] ||= :Default
      shape.add 't:BaseShape', opts[:base_shape]
    end
    
    # @see http://msdn.microsoft.com/en-us/library/aa580234(EXCHG.80).aspx
    # ItemId
    def item_id!(parent, item_id, opts = {})
      id_element! parent, 't:ItemId', item_id, opts
    end
    
    # @see http://msdn.microsoft.com/en-us/library/aa579461(EXCHG.80).aspx
    # FolderId
    def folder_id!(parent, folder_id, opts = {})
      id_element! parent, 't:FolderId', folder_id, opts
    end
    
    # @param parent [Handsoap::XmlMason::Node]
    #
    # @param folder_name Can only be Exchange named folders. See
    # referenced docs for details.
    #
    # @param [Hash] opts
    # @option opts [Symbol] :change_key
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580808(EXCHG.80).aspx
    # DistinguishedFolderId
    def distinguished_folder_id!(parent, folder_name, opts = {})
      normalized_id = folder_name.to_s.downcase
      id_element! parent, 't:DistinguishedFolderId', normalized_id, opts
    end
    
    private
    def id_element!(parent, id_name, id, opts = {})
      parent.add(id_name) do |folder_id|
        folder_id.set_attr 'Id', id        
        folder_id.set_attr 'ChangeKey', opts[:change_key] if opts[:change_key]
      end      
    end
    
  end
  
end
