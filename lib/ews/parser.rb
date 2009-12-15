module EWS
  
  class Parser
    def parse_find_folder(doc)
      doc.xpath('//t:Folders/child::*').map do |node|
        parse_exchange_folder node.xpath('.') # force NodeSelection
      end.compact
    end
    
    def parse_get_folder(doc)
      parse_exchange_folder doc.xpath('//m:Folders/child::*[1]')
    end
    
    def parse_find_item(doc)
      doc.xpath('//t:Items/child::*').map do |node|
        parse_exchange_item node.xpath('.') # force NodeSelection
      end.compact
    end
    
    def parse_get_item(doc)      
      parse_exchange_item doc.xpath('//m:Items/child::*[1]')
    end

    def parse_get_attachment(doc)
      parse_attachment doc.xpath('//m:Attachments/child::*[1]')
    end
    
    # Checks the ResponseMessage for errors.
    #
    # @see http://msdn.microsoft.com/en-us/library/aa494164%28EXCHG.80%29.aspx
    # Exhange 2007 Valid Response Messages
    def parse_response_message(doc)      
      error_node = doc.xpath('//m:ResponseMessages/child::*[@ResponseClass="Error"]')
      unless error_node.empty?
        error_msg = error_node.xpath('m:MessageText/text()').to_s
        response_code = error_node.xpath('m:ResponseCode/text()').to_s
        raise EWS::ResponseError.new(error_msg, response_code)
      end
    end

    private
    def parse_exchange_folder(folder_node)
      case folder_node.node_name
      when 'Folder'
        parse_folder folder_node
      when 'CalendarFolder'
      when 'ContactsFolder'
      when 'SearchFolder'
      when 'TasksFolder'
      else
        nil
      end
    end
    
    def parse_folder(folder_node)
      attrs = {
        :folder_id          => parse_id(folder_node.xpath('t:FolderId')),
        :display_name       => folder_node.xpath('t:DisplayName/text()').to_s,
        :total_count        => folder_node.xpath('t:TotalCount/text()').to_i,
        :child_folder_count => folder_node.xpath('t:ChildFolderCount/text()').to_i,
        :unread_count       => folder_node.xpath('t:UnreadCount/text()').to_i
      }   
      Folder.new attrs
    end
    
    def parse_exchange_item(item_node)
      case item_node.node_name
      when 'Item'
      when 'Message'
        parse_message item_node
      when 'CalendarItem'
      when 'Contact'
      when 'Task'
      when 'MeetingMessage'
      when 'MeetingRequest'
      when 'MeetingResponse'
      when 'MeetingCancellation'
      else
        nil
      end
    end
      
    def parse_message(message_node)
      attrs = {
        :item_id          => parse_id(message_node.xpath('t:ItemId')),
        :parent_folder_id => parse_id(message_node.xpath('t:ParentFolderId')),
        :subject          => message_node.xpath('t:Subject/text()').to_s,
        :body             => message_node.xpath('t:Body/text()').to_s,
        :body_type        => message_node.xpath('t:Body/@BodyType').to_s
      }

      nodeset = message_node.xpath('t:HasAttachments')
      attrs[:has_attachments] = if not nodeset.empty?
        parse_bool(nodeset)
      end
      
      nodeset = message_node.xpath('t:Attachments')
      attrs[:attachments] = if not nodeset.empty?
        nodeset.xpath('t:ItemAttachment|t:FileAttachment').map do |node|
          parse_attachment node
        end
      end

      nodeset = message_node.xpath('t:InternetMessageHeaders')
      attrs[:header] = if not nodeset.empty?
        parse_header nodeset
      end
   
      Message.new attrs
    end

    EXCHANGE_ITEM_XPATH = ['t:Item',
                           't:Message',
                           't:CalendarItem',
                           't:Contact',
                           't:Task',
                           't:MeetingMessage',
                           't:MeetingRequest',
                           't:MeetingResponse',
                           't:MeetingCancellation'].join('|').freeze
    
    def parse_attachment(attachment_node)
      attrs = {
        :attachment_id => attachment_node.xpath('t:AttachmentId/@Id').to_s,
        :name          => attachment_node.xpath('t:Name/text()').to_s,
        :content_type  => attachment_node.xpath('t:ContentType/text()').to_s,
        :content_id    => attachment_node.xpath('t:ContentId/text()').to_s,
        :content_location => attachment_node.xpath('t:ContentLocation/text()').to_s
      }
      
      case attachment_node.node_name
      when 'ItemAttachment'
        attrs[:item] = parse_exchange_item attachment_node.xpath(EXCHANGE_ITEM_XPATH)
      when 'FileAttachment'
      end
      
      Attachment.new attrs
    end

    def parse_header(header_node)
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
