require 'forwardable'

module EWS
  
  # AdditionalProperties
  # @see http://msdn.microsoft.com/en-us/library/aa580545(EXCHG.80).aspx
  # BaseShape  
  # @see http://msdn.microsoft.com/en-us/library/aa565622(EXCHG.80).aspx
  # BodyType
  # @see http://msdn.microsoft.com/en-us/library/aa580499(EXCHG.80).aspx
  # IncludeMimeContent  
  class Builder
    extend Forwardable

    def_delegators :@resolve_names_builder,
                   :unresolved_entry!,
                   :return_full_contact_data!

    def_delegators :@shape_builder,
                   :item_shape!,
                   :folder_shape!,
                   :attachment_shape!
    
    def initialize(action_node, opts, &block)
      @action_node, @opts = action_node, opts
      @resolve_names_builder = ResolveNamesBuilder.new(action_node)
      @shape_builder = ShapeBuilder.new(action_node, opts)
      instance_eval(&block) if block_given?
    end
        
    def item_id_container!(container_node_name, item_ids)
      id_container!(container_node_name, item_ids) do |container_node, id|
        item_id! container_node, id, @opts
      end
    end

    def folder_id_container!(container_node_name, folder_ids)
      id_container!(container_node_name, folder_ids) do |container_node, id|
        folder_id! container_node, id, @opts
      end
    end

    def attachment_ids!(attachment_ids)
      id_container!('tns:AttachmentIds', attachment_ids) do |container_node, id|
        id_node! container_node, 't:AttachmentId', id
      end
    end
    
    def id_container!(container_node_name, ids)
      @action_node.add(container_node_name) do |container_node|
        Array(ids).each {|id| yield container_node, id }
      end
    end

    def traversal!
      @action_node.set_attr 'Traversal', (@opts[:traversal] || :Shallow) 
    end        
        
    # @see http://msdn.microsoft.com/en-us/library/aa580234(EXCHG.80).aspx
    # ItemId
    def item_id!(container_node, item_id, opts = {})
      id_node! container_node, 't:ItemId', item_id, opts
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
    def folder_id!(container_node, folder_id, opts = {})
      node_name = if DistinguishedFolders.include?(folder_id)
        't:DistinguishedFolderId'
      else
        't:FolderId'
      end
      id_node! container_node, node_name, folder_id, opts
    end

    # @param opts [Hash] Still an argument so opts can remain optional
    def id_node!(container_node, id_node_name, id, opts = {})
      container_node.add(id_node_name) do |id_node|
        id_node.set_attr 'Id', id        
        id_node.set_attr 'ChangeKey', opts[:change_key] if opts[:change_key]
      end 
    end
  end
  
end
