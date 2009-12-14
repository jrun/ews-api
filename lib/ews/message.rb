module EWS
  
  class Message < Model
    def id
      @id ||= attrs[:item_id][:id]
    end

    def change_key
      @change_key ||= attrs[:item_id][:change_key]
    end
    
    def shallow?
      self.body.nil? || self.header.nil?
    end
  end
  
end
