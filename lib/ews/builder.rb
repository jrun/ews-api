module EWS
  
  class Builder
    def base_shape!(shape, opts = {})
      opts[:base_shape] ||= :Default
      shape.add 't:BaseShape', opts[:base_shape]
    end
  end
  
end
