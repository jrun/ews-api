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
    def self.endpoint(uri)
      super :uri => uri, :version => 1
    end
                 
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
      apply_namespaces! doc
      parser.parse_response_message doc
    end

    def apply_namespaces!(doc)
      doc.add_namespace 'soap', '"http://schemas.xmlsoap.org/soap/envelope'
      doc.add_namespace 't', 'http://schemas.microsoft.com/exchange/services/2006/types'
      doc.add_namespace 'm', 'http://schemas.microsoft.com/exchange/services/2006/messages'
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
    # @param [Hash] opts Options to manipulate the FindFolder response
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    #
    # @example Request
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
    def find_folder(parent_folder_name = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder'
      opts[:base_shape] ||= :Default
      
      response = invoke('tns:FindFolder', soap_action) do |find_folder|
        find_folder.set_attr 'Traversal', 'Deep'
        find_folder.add('tns:FolderShape') do |shape|
          shape.add 't:BaseShape', opts[:base_shape]
        end
        find_folder.add('tns:ParentFolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') do |id|
            id.set_attr 'Id', parent_folder_name.to_s.downcase
          end
        end
      end
      parser.parse_find_folder response.document
    end

    # @param [Hash] opts Options to manipulate the FindFolder response
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    #
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
    def get_folder(name = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder'
      opts[:base_shape] ||= :Default
      
      response = invoke('tns:GetFolder', soap_action) do |get_folder|
        get_folder.add('tns:FolderShape') do |shape|
          shape.add 't:BaseShape', opts[:base_shape]
        end
        get_folder.add('tns:FolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') { |id| id.set_attr 'Id', name }
        end
      end
      parser.parse_get_folder response.document
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

    # @param [Hash] opts Options to manipulate the FindItem response
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    #
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
    #
    # @see http://msdn.microsoft.com/en-us/library/aa566107(EXCHG.80).aspx
    # FindItem - Exchange 2007
    #
    # @see http://msdn.microsoft.com/en-us/library/aa566107.aspx
    # FindItem - Exchange 2010
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580545.aspx
    # BaseShape    
    def find_item(parent_folder_name = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindItem'      
      opts[:base_shape] ||= :IdOnly
      
      response = invoke('tns:FindItem', soap_action) do |find_item|
        find_item.set_attr 'Traversal', 'Shallow'
        find_item.add('tns:ItemShape') do |shape|
          shape.add 't:BaseShape', opts[:base_shape]
        end
        find_item.add('tns:ParentFolderIds') do |ids|
          ids.add('t:DistinguishedFolderId') do |folder_id|
            folder_id.set_attr 'Id', parent_folder_name.to_s.downcase
          end
        end
      end
      parser.parse_find_item response.document
    end

    # @param [Hash] opts Options to manipulate the FindFolder response
    # @option opts [String, Symbol] :base_shape (Default) IdOnly, Default, AllProperties
    # @option opts [String] :change_key
    #
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
    def get_item(item_id, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetItem'
      opts[:base_shape] ||= :Default
      
      response = invoke('tns:GetItem', soap_action) do |get_item|
        get_item.add('tns:ItemShape') do |shape|
          shape.add 't:BaseShape', opts[:base_shape]
          shape.add 't:IncludeMimeContent', false
        end
        get_item.add('tns:ItemIds') do |ids|
          ids.add('t:ItemId') do |item|
            item.set_attr 'Id', item_id
            if opts[:change_key]
              item.set_attr 'ChangeKey', opts[:change_key]
            end
          end
        end
      end
      parser.parse_get_item response.document
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

    # @param folder_id Name of the destination folder
    # @param item_ids [Array] List of item ids to be moved
    #
    # @example Request
    #   <MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
    #            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    #    <ToFolderId>
    #      <t:FolderId Id="drafts"/>
    #    </ToFolderId>
    #    <ItemIds>
    #      <t:ItemId Id="AAAtAEF/swbAAA=" ChangeKey="EwAAABYA/s4b"/>
    #    </ItemIds>
    #  </MoveItem>
    #
    # @see http://msdn.microsoft.com/en-us/library/aa565781%28EXCHG.80%29.aspx
    # MoveItem
    #
    # @see http://msdn.microsoft.com/en-us/library/aa580808%28EXCHG.80%29.aspx
    # DistinguishedFolderId
    #
    def move_item!(folder_id, item_ids)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem'
      response = invoke('tns:MoveItem', soap_action) do |move_item|
        move_item.add('tns:ToFolderId') do |to_folder|

          # TODO: Support both FolderID and DistinguishedFolderId
          to_folder.add('t:FolderId') do |folder_id_node|
            folder_id_node.set_attr 'Id', folder_id
          end          
        end
        move_item.add('tns:ItemIds') do |ids|
          item_ids.each do |item_id|
            ids.add('t:ItemId') {|item_node| item_node.set_attr 'Id', item_id }
          end
        end
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
      parser.parse_get_attachment response.document
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
    def parser
      @parser ||= Parser.new      
    end
    
    # helpers
    
    # TODO
  end
end
