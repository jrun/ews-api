module EWS
  
  class Message < Model
    def shallow?
      self.body.nil? || self.header.nil?
    end
  end
  
end
