module EWS
    
  class Model
    def initialize(attrs)
      @attrs = attrs.dup
    end
    
    def method_missing(method_name, *args)
      if @attrs.has_key?(method_name)
        @attrs[method_name]
      else
        super method_name, *args
      end
    end
  end
  
end
