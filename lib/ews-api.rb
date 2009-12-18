require 'net/ntlm'
require 'handsoap'

require 'ews/error'
require "ews/distinguished_folders"
require 'ews/model'
require 'ews/attachment'
require 'ews/message'
require 'ews/folder'
require 'ews/builder'
require 'ews/parser'
require 'ews/service'

module EWS
  def self.inbox
    Service.get_folder(:inbox, :base_shape => :AllProperties)
  end
end
