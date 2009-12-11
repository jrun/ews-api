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

    def parse_get_item(doc)
      parse_message doc.xpath('//t:Message').first
    end

    private
    def parse_message(message_node)
      attrs = {
        :item_id          => parse_id(message_node.xpath('t:ItemId')),
        :parent_folder_id => parse_id(message_node.xpath('t:ParentFolderId')),
        :subject          => message_node.xpath('t:Subject/text()').to_s,
        :body             => message_node.xpath('t:Body/text()').to_s,
        :body_type        => message_node.xpath('t:Body/@BodyType').to_s,
        :has_attachments  => parse_bool(message_node.xpath('t:HasAttachments')),
        :attachments      => parse_attachments(message_node.xpath('t:Attachments'))
      }
      Message.new attrs
    end

    def parse_attachments(attachments_node)
      return [] unless attachments_node
      attachments_node.xpath('t:ItemAttachment').map do |node|
        attrs = {
          :attachment_id => node.xpath('t:AttachmentId/@Id').to_s,
          :name          => node.xpath('t:Name/text()').to_s,
          :content_type  => node.xpath('t:ContentType/text()').to_s
        }
        Attachment.new attrs
      end
    end
    
    def parse_id(id_node)
      { :id         => id_node.xpath('@Id').to_s,
        :change_key => id_node.xpath('@ChangeKey').to_s }
    end

    def parse_bool(val)
      case val.to_s.downcase
      when 'true'
        true
      when 'false'
        false
      else
        nil
      end
    end

  end
end
