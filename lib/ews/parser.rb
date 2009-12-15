module EWS
  
  class Parser
    def parse_get_folder(doc)
      folder = doc.xpath('//t:Folder')
      attrs = {
        :folder_id  => folder.xpath('t:FolderId/@Id').to_s,
        :change_key => folder.xpath('t:FolderId/@ChangeKey').to_s,
        :name       => folder.xpath('t:DisplayName/text()').to_s
      }      
      Folder.new attrs
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
    
    RESPONSE_MSG_XPATH = ['//m:FindFolderResponseMessage',
                          '//m:GetFolderResponseMessage',
                          '//m:FindItemResponseMessage',
                          '//m:GetItemResponseMessage',
                          '//m:GetAttachmentResponseMessage'].join('|')
    
    # Parses the ResponseMessage looking for errors.
    #
    # @see http://msdn.microsoft.com/en-us/library/aa494164%28EXCHG.80%29.aspx
    # Exhange 2007 Valid Response Messages
    #    
    # CopyFolderResponseMessage
    # CopyItemResponseMessage
    # CreateAttachmentResponseMessage
    # CreateFolderResponseMessage
    # CreateItemResponseMessage
    # CreateManagedFolderResponseMessage
    # DeleteAttachmentResponseMessage
    # DeleteFolderResponseMessage
    # DeleteItemResponseMessage
    # ExpandDLResponseMessage
    # FindFolderResponseMessage
    # FindItemResponseMessage
    # GetAttachmentResponseMessage
    # GetEventsResponseMessage
    # GetFolderResponseMessage
    # GetItemResponseMessage
    # MoveFolderResponseMessage
    # MoveItemResponseMessage
    # ResolveNamesResponseMessage
    # SendItemResponseMessage
    # SendNotificationResponseMessage
    # SubscribeResponseMessage
    # SyncFolderHierarchyResponseMessage
    # SyncFolderItemsResponseMessage
    # UnsubscribeResponseMessage
    # UpdateFolderResponseMessage
    # UpdateItemResponseMessage
    # ConvertIdResponseMessage
    def parse_response_message(doc)      
      response_msg = doc.xpath(RESPONSE_MSG_XPATH)
      if response_msg.xpath('@ResponseClass').to_s == 'Error'
        error_msg = (response_msg / 'm:MessageText/text()').to_s
        response_code = (response_msg / 'm:ResponseCode/text()').to_s
        raise EWS::ResponseError.new(error_msg, response_code)
      end
    end

    private
    def parse_exchange_item(item_node)
      case item_node.node_name
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
