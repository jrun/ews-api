module EWS
    
  class Model
    attr_accessor :service
    
    def initialize(attrs = {})
      @attrs = attrs.dup
      @service = EWS::Service
    end

    def shallow?
      raise NotImplementedError, "Each model must determine when it is shallow."
    end
    
    protected
    attr_reader :attrs
    
    public
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
        super meth, *args
      end
    end
    
  end
  
end
