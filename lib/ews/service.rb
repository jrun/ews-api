# -*- coding: utf-8 -*-

# ntlm authentication works using httpclient and rubyntlm
Handsoap.http_driver = :http_client
Handsoap.xml_query_driver = :nokogiri

module EWS
  # Implementation of Exchange Web Services
  #
  # @see http://msdn.microsoft.com/en-us/library/bb409286.aspx Exchange Web Services Operations
  # @see http://msdn.microsoft.com/en-us/library/aa580545.aspx BaseShape
  class Service < Handsoap::Service
    
    @@username, @@password = nil, nil
          
    def set_auth(username, password)
      @@username, @@password = username, password
    end
    
    def on_after_create_http_request(req)
      if @@username && @@password 
        req.set_auth @@username, @@password
      end
    end
    
    def on_create_document(doc)
      # register namespaces for the request
      doc.alias 'tns', 'http://schemas.microsoft.com/exchange/services/2006/messages'
      doc.alias 't', 'http://schemas.microsoft.com/exchange/services/2006/types'
    end
    
    def on_response_document(doc)
      # register namespaces for the response
      doc.add_namespace 'soap', '"http://schemas.xmlsoap.org/soap/envelope'
      doc.add_namespace 't', 'http://schemas.microsoft.com/exchange/services/2006/types'
      doc.add_namespace 'm', 'http://schemas.microsoft.com/exchange/services/2006/messages'
      parse_response_message doc
    end
    # public methods
  
    def resolve_names!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/ResolveNames'
      response = invoke('tns:ResolveNames', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def expand_dl!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/ExpandDL'
      response = invoke('tns:ExpandDL', soap_action) do |message|
        raise "TODO"
      end
    end
    
    # Finds folders for the given parent folder.
    #
    # @example Request
    #
    #  <FindFolder Traversal="Shallow" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
    #    <FolderShape>
    #      <t:BaseShape>Default</t:BaseShape>
    #    </FolderShape>
    #    <ParentFolderIds>
    #      <t:DistinguishedFolderId Id="inbox"/>
    #    </ParentFolderIds>
    #  </FindFolder>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa563918.aspx
    # FindFolder
    #
    # @see http://msdn.microsoft.com/en-us/library/aa494311.aspx
    # FolderShape
    #
    # @todo Support options
    #   Traversal: +Shallow, Deep, SoftDeleted+
    #   FolderShape: +IdOnly, Default, AllProperties+
    def find_folder(parent_folder_name = :root)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder'
      response = invoke('tns:FindFolder', soap_action) do |find_folder|
        find_folder.set_attr 'Traversal', 'Deep'
        find_folder.add('tns:FolderShape') do |shape|
          shape.add('t:BaseShape', 'Default')
        end
        find_folder.add('tns:ParentFolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') do |id|
            id.set_attr 'Id', parent_folder_name
          end
        end
      end
      
    end
        
    # @example Request
    #
    #  <GetFolder xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
    #             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    #    <FolderShape>
    #      <t:BaseShape>Default</t:BaseShape>
    #    </FolderShape>
    #    <FolderIds>
    #      <t:DistinguishedFolderId Id="inbox"/>
    #    </FolderIds>
    #  </GetFolder>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580274.aspx
    # MSDN - GetFolder operation
    #
    def get_folder(name = :root)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder'
      response = invoke('tns:GetFolder', soap_action) do |get_folder|
        get_folder.add('tns:FolderShape') do |shape|
          shape.add('t:BaseShape', 'Default')
        end
        get_folder.add('tns:FolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') { |id| id.set_attr 'Id', name }
        end
      end
      
    end
    
    def convert_id!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/ConvertId'
      response = invoke('tns:ConvertId', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def create_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateFolder'
      response = invoke('tns:CreateFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def delete_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/DeleteFolder'
      response = invoke('tns:DeleteFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def update_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/UpdateFolder'
      response = invoke('tns:UpdateFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def move_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/MoveFolder'
      response = invoke('tns:MoveFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def copy_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CopyFolder'
      response = invoke('tns:CopyFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def subscribe!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/Subscribe'
      response = invoke('tns:Subscribe', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def unsubscribe!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/Unsubscribe'
      response = invoke('tns:Unsubscribe', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def get_events
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetEvents'
      response = invoke('tns:GetEvents', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def sync_folder_hierarchy!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/SyncFolderHierarchy'
      response = invoke('tns:SyncFolderHierarchy', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def sync_folder_items!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/SyncFolderItems'
      response = invoke('tns:SyncFolderItems', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def create_managed_folder!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateManagedFolder'
      response = invoke('tns:CreateManagedFolder', soap_action) do |message|
        raise "TODO"
      end
    end

    # @example Request
    #  <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
    #            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
    #          Traversal="Shallow">
    #    <ItemShape>
    #      <t:BaseShape>IdOnly</t:BaseShape>
    #    </ItemShape>
    #    <ParentFolderIds>
    #      <t:DistinguishedFolderId Id="deleteditems"/>
    #    </ParentFolderIds>
    #  </FindItem>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa566107.aspx
    # FindItem
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580545.aspx
    # BaseShape
    def find_item(parent_folder_name = :root)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindItem'
      response = invoke('tns:FindItem', soap_action) do |find_item|
        find_item.set_attr 'Traversal', 'Shallow'
        find_item.add('tns:ItemShape') do |shape|
          shape.add('t:BaseShape', 'IdOnly')
        end
        find_item.add('tns:ParentFolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') do |folder_id|
            folder_id.set_attr 'Id', parent_folder_name
          end
        end
      end
      
    end
    
    # @example Request for getting a mail message
    #  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
    #           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    #    <ItemShape>
    #      <t:BaseShape>Default</t:BaseShape>
    #      <t:IncludeMimeContent>true</t:IncludeMimeContent>
    #    </ItemShape>
    #    <ItemIds>
    #      <t:ItemId Id="AAAlAF" ChangeKey="CQAAAB" />
    #    </ItemIds>
    #  </GetItem>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa566013.aspx
    # GetItem (E-mail Message)
    #
    # @see http://msdn.microsoft.com/en-us/library/aa565261.aspx
    # ItemShape
    #
    # @see http://msdn.microsoft.com/en-us/library/aa563525.aspx
    # ItmeIds
    def get_item(item_id, change_key = nil)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetItem'
      response = invoke('tns:GetItem', soap_action) do |get_item|
        get_item.add('tns:ItemShape') do |shape|
          shape.add('t:BaseShape', 'AllProperties')
          shape.add('t:IncludeMimeContent', false)
        end
        get_item.add('tns:ItemIds') do |ids|
          ids.add('t:ItemId') do |item|
            item.set_attr 'Id', item_id
            item.set_attr 'ChangeKey', change_key if change_key
          end
        end
      end
    end
    
    def create_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
      response = invoke('tns:CreateItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def delete_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/DeleteItem'
      response = invoke('tns:DeleteItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def update_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem'
      response = invoke('tns:UpdateItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def send_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/SendItem'
      response = invoke('tns:SendItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def move_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem'
      response = invoke('tns:MoveItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def copy_item!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CopyItem'
      response = invoke('tns:CopyItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def create_attachment!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateAttachment'
      response = invoke('tns:CreateAttachment', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def delete_attachment!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/DeleteAttachment'
      response = invoke('tns:DeleteAttachment', soap_action) do |message|
        raise "TODO"
      end
    end

    # @example Request
    #  <GetAttachment xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
    #                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    #    <AttachmentShape/>
    #    <AttachmentIds>
    #      <t:AttachmentId Id="AAAtAEFkbWluaX..."/>
    #    </AttachmentIds>
    #  </GetAttachment>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa494316.aspx
    # GetAttachment
    def get_attachment(attachment_id)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetAttachment'
      response = invoke('tns:GetAttachment', soap_action) do |get_attachment|
        get_attachment.add('tns:AttachmentShape')
        get_attachment.add('tns:AttachmentIds') do |ids|
          ids.add('t:AttachmentId') do |attachment|
            attachment.set_attr 'Id', attachment_id
          end
        end
      end
    end
    
    def get_delegate
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetDelegate'
      response = invoke('tns:GetDelegate', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def add_delegate!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/AddDelegate'
      response = invoke('tns:AddDelegate', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def remove_delegate!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/RemoveDelegate'
      response = invoke('tns:RemoveDelegate', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def update_delegate!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/UpdateDelegate'
      response = invoke('tns:UpdateDelegate', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def get_user_availability
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetUserAvailability'
      response = invoke('tns:GetUserAvailability', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def get_user_oof_settings
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetUserOofSettings'
      response = invoke('tns:GetUserOofSettings', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def set_user_oof_settings!
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/SetUserOofSettings'
      response = invoke('tns:SetUserOofSettings', soap_action) do |message|
        raise "TODO"
      end
    end
       
    private    
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
      if (response_msg / '@ResponseClass').to_s == 'Error'
        error_msg = (response_msg / 'm:MessageText/text()').to_s
        response_code = (response_msg / 'm:ResponseCode/text()').to_s
        raise EWS::ResponseError.new(error_msg, response_code)
      end
    end
    
    # helpers
    
    # TODO
  end
end
