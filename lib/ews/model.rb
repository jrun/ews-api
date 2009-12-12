module EWS
    
  class Model
    def initialize(attrs = {})
      @attrs = attrs.dup
    end
    
    def method_missing(meth, *args)
      method_name = meth.to_s
      if method_name.end_with?('=')
        attr = method_name.chomp('=').to_sym
        @attrs[attr] = args.first        
      elsif method_name.end_with?('?')        
        attr = method_name.chomp('?').to_sym
        @attrs[attr] == true        
      elsif @attrs.has_key?(meth)        
        @attrs[meth]        
      else
        super method_name, *args
      end
    end
    
  end
  
end
