# -*- coding: utf-8 -*-

# ntlm authentication works using httpclient and rubyntlm
Handsoap.http_driver = :http_client

module EWS
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
      doc.add_namespace 'ns', 'http://schemas.microsoft.com/exchange/services/2006/messages'
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
    
    def find_folder
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder'
      response = invoke('tns:FindFolder', soap_action) do |message|
        raise "TODO"
      end
    end
    
    def find_item
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindItem'
      response = invoke('tns:FindItem', soap_action) do |message|
        raise "TODO"
      end
    end
    
    # @param The name of the folder to retrieve
    # 
    # @example Request
    #  Source: http://msdn.microsoft.com/en-us/library/aa580263.aspx
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
    def get_folder(name)
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
    
    def get_item
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetItem'
      response = invoke('tns:GetItem', soap_action) do |message|
        raise "TODO"
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
    
    def get_attachment
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetAttachment'
      response = invoke('tns:GetAttachment', soap_action) do |message|
        raise "TODO"
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
    # helpers
    # TODO
  end
end
