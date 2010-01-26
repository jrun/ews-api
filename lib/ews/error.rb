module EWS
  
  class Error < StandardError
  end
  
  class PreconditionFailed < Error
  end
  
  class ResponseError < Error
    attr_reader :response_code

    def initialize(message, response_code)
      super message
      @response_code = response_code
    end
  end
end
