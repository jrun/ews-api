module EWS
  class ShapeBuilder
    def initialize(action_node, opts = {})
      @action_node, @opts = action_node, opts
      @shape_node = nil
    end

    # @param [Hash] opts
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    # @option opts [true, false] :include_mime_content
    # @option opts [String, Symbol] :body_type Best, HTML or Text
    # @option opts [Symbol] :additional_properties  
    def item_shape!
      set_shape_node 'tns:ItemShape'
      base_shape!
      include_mime_content!
      body_type!
      additional_properties!      
    end
    
    # @param [Hash] opts
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    # @option opts [Symbol] :additional_properties
    def folder_shape!
      set_shape_node 'tns:FolderShape'
      base_shape!
      additional_properties!
    end
    
    # @param [Hash] opts
    # @option opts [true, false] :include_mime_type
    # @option opts [String, Symbol] :body_type Best, HTML or Text
    # @option opts [Symbol] :additional_properties
    #
    # @see http://msdn.microsoft.com/en-us/library/aa563727(EXCHG.80).aspx
    # AttachmentShape
    def attachment_shape!
      set_shape_node 'tns:AttachmentShape'
      include_mime_content!
      body_type!
      additional_properties!
    end
    
    private
    def set_shape_node(name)
      @action_node.add(name) do |shape_node|
        @shape_node = shape_node
      end      
    end
    
    def base_shape!
      @opts[:base_shape] ||= :Default      
      @shape_node.add 't:BaseShape', @opts[:base_shape]
    end
    
    def include_mime_content!
      if @opts.has_key?(:include_mime_content)
        @shape_node.add 't:IncludeMimeContent', @opts[:include_mime_content]
      end
    end    

    def body_type!
      if @opts.has_key?(:body_type)
        @shape_node.add 't:BodyType', @opts[:body_type]
      end
    end
    
    def additional_properties!
    end
  end
end
