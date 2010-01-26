module EWS  

  class Folder < Model
    def id
      attrs[:folder_id][:id]
    end

    def change_key
      attrs[:folder_id][:change_key]
    end

    def name
      attrs[:display_name]
    end

    def each_message
      items.each {|message| yield message }
    end

    def folders
      @folders ||= find_folders.inject({}) do |folders, folder|
        folders[folder.name] = folder
        folders
      end
    end

    def items
      @items ||= find_folder_items
    end
    
    private
    def find_folder_items
      # NOTE: This assumes Service#find_item only returns
      # Messages. That is true now but will change as more
      # of the parser is implemented.
      service.find_item(self.id, :base_shape => :AllProperties)
    end

    def find_folders
      service.find_folder(self.name)
    end
  end
  
end
