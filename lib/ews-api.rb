require 'net/ntlm'
require 'handsoap'

require 'ews/error'
require 'ews/model'
require 'ews/attachment'
require 'ews/message'
require 'ews/folder'
require 'ews/parser'
require 'ews/service'

module EWS
  def self.folder(name)
    Service.get_folder(name, :base_shape => :AllProperties)
  end
end
