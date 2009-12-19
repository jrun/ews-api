module EWS
  class ShapeBuilder
    def initialize(shape, opts = {}, &block)
      @shape, @opts = shape, opts
      instance_eval(&block) if block_given?
    end
    
    def base_shape!
      @opts[:base_shape] ||= :Default      
      @shape.add 't:BaseShape', @opts[:base_shape]
    end
    
    def include_mime_content!
      if @opts.has_key?(:include_mime_content)
        @shape.add 't:IncludeMimeContent', @opts[:include_mime_content]
      end
    end    

    def body_type!
      if @opts.has_key?(:body_type)
        @shape.add 't:BodyType', @opts[:body_type]
      end
    end
    
    def additional_properties!
    end
  end
end
