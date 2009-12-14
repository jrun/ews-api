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
    
    def parse_find_item(doc)
      doc.xpath('//t:Items/*').map do |node|
        case node.node_name
        when 'Message'
          parse_message node.xpath('.') # force NodeSelection
        else
          nil
        end
      end.compact    
    end
    
    def parse_get_item(doc)
      # TODO: support all of the types of items
      parse_message doc.xpath('//t:Message')
    end

    def parse_get_attachment(doc)
      parse_attachments(doc.xpath('//m:Attachments')).first
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
    def parse_message(message_node)
      return nil if message_node.empty?
      attrs = {
        :item_id          => parse_id(message_node.xpath('t:ItemId')),
        :parent_folder_id => parse_id(message_node.xpath('t:ParentFolderId')),
        :subject          => message_node.xpath('t:Subject/text()').to_s,
        :body             => message_node.xpath('t:Body/text()').to_s,
        :body_type        => message_node.xpath('t:Body/@BodyType').to_s
      }

      nodeset = message_node.xpath('t:HasAttachments')
      attrs[:has_attachments] = if nodeset.empty?
        nil
      else
        parse_bool(nodeset)
      end

      nodeset = message_node.xpath('t:Attachments')
      attrs[:attachments] = if nodeset.empty?
        nil
      else
        parse_attachments nodeset
      end

      nodeset = message_node.xpath('t:InternetMessageHeaders')
      attrs[:header] = if nodeset.empty?
        nil
      else
        parse_header nodeset
      end
      
      Message.new attrs
    end

    def parse_attachments(attachments_node)
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
