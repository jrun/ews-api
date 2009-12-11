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
      # TODO: support all of the types of items
      parse_message doc.xpath('//t:Message')
    end

    private
    def parse_message(message_node)
      return nil if message_node.empty?
      attrs = {
        :item_id          => parse_id(message_node.xpath('t:ItemId')),
        :parent_folder_id => parse_id(message_node.xpath('t:ParentFolderId')),
        :subject          => message_node.xpath('t:Subject/text()').to_s,
        :body             => message_node.xpath('t:Body/text()').to_s,
        :body_type        => message_node.xpath('t:Body/@BodyType').to_s,
        :has_attachments  => parse_bool(message_node.xpath('t:HasAttachments')),
        :attachments      => parse_attachments(message_node.xpath('t:Attachments')),
        :header           => parse_header(message_node.xpath('t:InternetMessageHeaders'))
      }
      Message.new attrs
    end

    def parse_attachments(attachments_node)
      return [] if attachments_node.empty?
      attachments_node.xpath('t:ItemAttachment').map do |node|
        attrs = {
          :attachment_id => node.xpath('t:AttachmentId/@Id').to_s,
          :name          => node.xpath('t:Name/text()').to_s,
          :content_type  => node.xpath('t:ContentType/text()').to_s
        }
        Attachment.new attrs
      end
    end

    def parse_header(header_node)
      return {} if header_node.empty?
      header_node.xpath('t:InternetMessageHeader').inject({}) do |header, node|
        name = node.xpath('@HeaderName').to_s.downcase
        
        header[name] = [] unless header.has_key?(name)          
        header[name] << node.xpath('text()').to_s        
        header
      end
    end
    
    def parse_id(id_node)
      return nil if id_node.empty?
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
