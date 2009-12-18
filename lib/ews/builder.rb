module EWS
  
  class Builder
    def base_shape!(shape, opts = {})
      opts[:base_shape] ||= :Default
      shape.add 't:BaseShape', opts[:base_shape]
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
      parent.add('t:DistinguishedFolderId') do |folder_id|
        folder_id.set_attr 'Id', folder_name.to_s.downcase
        if opts[:change_key]
          folder_id.set_attr 'ChangeKey', opts[:change_key]
        end
      end      
    end
    
  end
  
end
