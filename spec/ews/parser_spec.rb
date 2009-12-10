require File.dirname(__FILE__) + '/../spec_helper'

module EWS
   
  describe Parser do
    before(:each) do
      @parser = Parser.new
    end
    
    context 'parsing get_folder should build a folder with' do
      before(:each) do
        @folder = @parser.parse_get_folder response_to_doc(:get_folder)
      end

      { :folder_id => 'AAA',
        :change_key => 'BBB',
        :name => 'Inbox' }.each do |attr, value|

        it attr do
          @folder.send(attr).should == value
        end
      end
      
    end
    
  end
end
