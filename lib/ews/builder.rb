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

    # @see http://msdn.microsoft.com/en-us/library/aa565998(EXCHG.80).aspx
    # ParentFolderIds
    def parent_folder_ids!(action, folder_ids, opts = {})
      folder_ids_element! action, 'tns:ParentFolderIds', folder_ids, opts
    end

    # @see http://msdn.microsoft.com/en-us/library/aa580509(EXCHG.80).aspx
    # FolderIds
    def folder_ids!(action, folder_ids, opts = {})
      folder_ids_element! action, 'tns:FolderIds', folder_ids, opts
    end
    
    # @param parent [Handsoap::XmlMason::Node]
    #
    # @param folder_id [String, Symbol] When a EWS::DistinguishedFolder a
    # DistinguishedFolderId is created otherwise FolderId is used
    #
    # @param [Hash] opts
    # @option opts [Symbol] :change_key
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580808(EXCHG.80).aspx
    # DistinguishedFolderId
    #   
    # @see http://msdn.microsoft.com/en-us/library/aa579461(EXCHG.80).aspx
    # FolderId
    def folder_id!(parent, folder_id, opts = {})
      folder_id_element_name = if DistinguishedFolders.include?(folder_id)
        't:DistinguishedFolderId'
      else
        't:FolderId'
      end
      id_element! parent, folder_id_element_name, folder_id, opts
    end
        
    private
    def folder_ids_element!(parent, ids_element_name, folder_ids, opts = {})
      ids = case folder_ids
      when Enumerable
        folder_ids
      else
        [folder_ids]
      end
      parent.add(ids_element_name) do |ids_element|
        ids.each do |folder_id|
          folder_id! ids_element, folder_id, opts
        end
      end
    end
    
    def id_element!(parent, id_name, id, opts = {})
      parent.add(id_name) do |id_element|
        id_element.set_attr 'Id', id        
        id_element.set_attr 'ChangeKey', opts[:change_key] if opts[:change_key]
      end 
    end
    
  end
  
end
