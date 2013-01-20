#
#   nop.rb -
#   	$Release Version: 0.9.6$
#   	$Revision: 24 $
#   	by Keiju ISHITSUKA(keiju@ruby-lang.org)
#
# --
#
#
#
module IRB
  module ExtendCommand
    class Nop

      @RCS_ID='-$Id: nop.rb 24 2012-11-23 10:13:10Z TatsuyaEndo $-'

      def self.execute(conf, *opts)
	command = new(conf)
	command.execute(*opts)
      end

      def initialize(conf)
	@irb_context = conf
      end

      attr_reader :irb_context

      def irb
	@irb_context.irb
      end

      def execute(*opts)
	#nop
      end
    end
  end
end

