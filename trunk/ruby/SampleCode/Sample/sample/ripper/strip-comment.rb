# $Id: strip-comment.rb 24 2012-11-23 10:13:10Z TatsuyaEndo $

require 'ripper/filter'

class CommentStripper < Ripper::Filter
  def CommentStripper.strip(src)
    new(src).parse(nil)
  end

  def on_default(event, token, data)
    print token
  end

  def on_comment(token, data)
    puts
  end
end

CommentStripper.strip(ARGF)
