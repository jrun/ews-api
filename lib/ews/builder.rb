module EWS
  
  # @see http://msdn.microsoft.com/en-us/library/aa563810(EXCHG.80).aspx
  # AdditionalProperties
  # @see http://msdn.microsoft.com/en-us/library/aa580545(EXCHG.80).aspx
  # BaseShape  
  # @see http://msdn.microsoft.com/en-us/library/aa565622(EXCHG.80).aspx
  # BodyType
  # @see http://msdn.microsoft.com/en-us/library/aa580499(EXCHG.80).aspx
  # IncludeMimeContent  
  class Builder
    def initialize(action_node, opts, &block)
      @action_node, @opts = action_node, opts
      instance_eval(&block) if block_given?
    end
    
    # @param [Hash] opts
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    # @option opts [true, false] :include_mime_content
    # @option opts [String, Symbol] :body_type Best, HTML or Text
    # @option opts [Symbol] :additional_properties  
    def item_shape!
      @action_node.add('tns:ItemShape') do |shape_node|
        ShapeBuilder.new(shape_node, @opts) do
          base_shape!
          include_mime_content!
          body_type!
          additional_properties!
        end
      end
    end
    
    # @param [Hash] opts
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    # @option opts [Symbol] :additional_properties
    def folder_shape!
      @action_node.add('tns:FolderShape') do |shape_node|
        ShapeBuilder.new(shape_node, @opts)  do
          base_shape!
          additional_properties!
        end
      end
    end
    
    # @param [Hash] opts
    # @option opts [true, false] :include_mime_type
    # @option opts [String, Symbol] :body_type Best, HTML or Text
    # @option opts [Symbol] :additional_properties
    #
    # @see http://msdn.microsoft.com/en-us/library/aa563727(EXCHG.80).aspx
    # AttachmentShape
    def attachment_shape!
      @action_node.add('tns:AttachmentShape') do |shape|
        ShapeBuilder.new(shape, @opts) do
          include_mime_content!
          body_type!
          additional_properties!
        end
      end
    end
    
    # @see http://msdn.microsoft.com/en-us/library/aa563525(EXCHG.80).aspx
    # ItemIds
    def item_ids!(item_ids)
      @action_node.add('tns:ItemIds') do |ids_element|
        to_a(item_ids).each do |item_id|
          item_id! ids_element, item_id, @opts
        end
      end
    end

    # @see http://msdn.microsoft.com/en-us/library/aa565998(EXCHG.80).aspx
    # ParentFolderIds
    def parent_folder_ids!(folder_ids)
      folder_container! 'tns:ParentFolderIds', folder_ids
    end

    # @see http://msdn.microsoft.com/en-us/library/aa580509(EXCHG.80).aspx
    # FolderIds
    def folder_ids!(folder_ids)
      folder_container! 'tns:FolderIds', folder_ids
    end
        
    def to_folder_id!(folder_id)
      folder_container! 'tns:ToFolderId', folder_id
    end
        
    def folder_container!(container_node_name, folder_ids)
      @action_node.add(container_node_name) do |container_node|
        to_a(folder_ids).each do |folder_id|
          folder_id! container_node, folder_id, @opts
        end
      end
    end

    def attachment_ids!(attachment_ids)
      @action_node.add('tns:AttachmentIds') do |ids_node|
        to_a(attachment_ids).each do |attachment_id|
          id_element! ids_node, 't:AttachmentId', attachment_id
        end
      end
    end

    def traversal!
      @action_node.set_attr 'Traversal', (@opts[:traversal] || :Shallow) 
    end        
        
    # @see http://msdn.microsoft.com/en-us/library/aa580234(EXCHG.80).aspx
    # ItemId
    def item_id!(parent, item_id, opts = {})
      id_element! parent, 't:ItemId', item_id, opts
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

    # @param opts [Hash] Still an argument so opts can remain optional
    def id_element!(parent, id_name, id, opts = {})
      parent.add(id_name) do |id_element|
        id_element.set_attr 'Id', id        
        id_element.set_attr 'ChangeKey', opts[:change_key] if opts[:change_key]
      end 
    end

    # TODO: core_ext?
    def to_a(ids)
      case ids
      when Enumerable
        ids
      else
        [ids]
      end
    end

  end
  
end
