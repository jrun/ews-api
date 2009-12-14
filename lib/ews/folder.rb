module EWS  

  class Folder < Model
    def each_message
      find_folder_items.each do |message|
        yield message
      end
    end

    private
    def find_folder_items
      # NOTE: This assumes Service#find_item only returns
      # Messages. That is true now but will change as more
      # of the parser is implemented.
      service.find_item(self.name.downcase, :base_shape => :AllProperties)
    end
  end
  
end
