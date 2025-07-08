class CellBorder

  def initialize(style:, color:)
    @style = style
    @color = color
    
  end

  attr_accessor :style, :color

end

module Enumerable
  def first_result
   block_given? ? each {|item| result = (yield item) and return result} && nil : find {|item| item}
  end
end