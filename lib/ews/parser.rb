module EWS
  
  class Parser
    def parse_get_folder(doc)
      folder = doc.xpath('//t:Folder')

      attrs = {
        :folder_id => folder.xpath('t:FolderId/@Id').to_s,
        :change_key => folder.xpath('t:FolderId/@ChangeKey').to_s,
        :name => folder.xpath('t:DisplayName/text()').to_s
      }
       
      Folder.new attrs
    end
  end
end
