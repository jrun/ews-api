require File.dirname(__FILE__) + '/../spec_helper'

describe EWS::Model do
  context 'with a method ending in ?' do
    it "should return true when the attribute value is true" do
      model = EWS::Model.new(:has_attachment => true)
      model.should have_attachment
    end

    it "should return false when the attribute value is not true" do
      model = EWS::Model.new(:has_attachment => false)
      model.should_not have_attachment
      
      model = EWS::Model.new(:has_attachment => nil)
      model.should_not have_attachment

      model = EWS::Model.new(:has_attachment => 'no')
      model.should_not have_attachment      
    end
  end

  context 'with a method ending in =' do
    it "should set the attribute's value" do
      model = EWS::Model.new
      model.subject = 'A Feast for Crows'
      
      model.subject.should == 'A Feast for Crows'
    end
  end

end
