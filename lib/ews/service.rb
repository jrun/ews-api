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
      register_aliases! doc    
    end
    
    def on_response_document(doc)
      apply_namespaces! doc
      parser.parse_response_message doc
    end
    
    def register_aliases!(doc)
      doc.alias 'tns', 'http://schemas.microsoft.com/exchange/services/2006/messages'
      doc.alias 't', 'http://schemas.microsoft.com/exchange/services/2006/types'
    end
    
    def apply_namespaces!(doc)
      doc.add_namespace 'soap', '"http://schemas.xmlsoap.org/soap/envelope'
      doc.add_namespace 't', 'http://schemas.microsoft.com/exchange/services/2006/types'
      doc.add_namespace 'm', 'http://schemas.microsoft.com/exchange/services/2006/messages'
    end
    # public methods
  
    def resolve_names!(unresolved_entry, return_full_contact_data = true)
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/ResolveNames'
      response = invoke('tns:ResolveNames', soap_action) do |resolve_names|
        builder(resolve_names) do
          unresolved_entry! unresolved_entry
          return_full_contact_data! return_full_contact_data
        end
      end
      parser.parse_resolve_names response.document
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
    def find_folder(folder_ids = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder'     
      response = invoke('tns:FindFolder', soap_action) do |find_folder|
        builder(find_folder, opts) do
          traversal!
          folder_shape!
          folder_id_container! 'tns:ParentFolderIds', folder_ids
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
    def get_folder(folder_id = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder'      
      response = invoke('tns:GetFolder', soap_action) do |get_folder|
        builder(get_folder, opts) do 
          folder_shape!
          folder_id_container! 'tns:FolderIds', folder_id
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
    def find_item(folder_ids = :root, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/FindItem'
      response = invoke('tns:FindItem', soap_action) do |find_item|
        builder(find_item, opts) do
          traversal!
          item_shape!
          folder_id_container! 'tns:ParentFolderIds', folder_ids         
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
      raise PreconditionFailed, "Id must be non-empty" if item_id.nil?
      
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetItem'
      response = invoke('tns:GetItem', soap_action) do |get_item|
        builder(get_item, opts) do
          item_shape!
          item_id_container! 'tns:ItemIds', item_id
        end
      end
      parser.parse_get_item response.document
    end
    
    def create_item!(item, destination_folder_id, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
      response = invoke('tns:CreateItem', soap_action) do |create_item|
        builder(create_item, opts) do
          folder_id_container! 'SavedItemFolderId', destination_folder_id
        end
        raise 'TODO'
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
    def move_item!(folder_id, item_id, opts = {})
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem'
      response = invoke('tns:MoveItem', soap_action) do |move_item|
        builder(move_item, opts) do
          folder_id_container! 'tns:ToFolderId', folder_id
          item_id_container! 'tns:ItemIds', item_id
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
    def get_attachment(attachment_id, opts = {})
      raise PreconditionFailed, "Id must be non-empty" if attachment_id.nil?
      
      soap_action = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetAttachment'
      response = invoke('tns:GetAttachment', soap_action) do |get_attachment|
        builder(get_attachment, opts) do
          attachment_shape!
          attachment_ids! attachment_id
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

    def builder(action_node, opts = {}, &block)
      Builder.new(action_node, opts, &block)
    end
  end
end
