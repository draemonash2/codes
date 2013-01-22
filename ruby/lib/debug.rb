#! /usr/bin/env ruby
# =================================================
#	$Brief	���̑��֐���񋟂��� $
#	
#	$Date:: 2013-01-07 00:30:23 +0900#$
#	$Rev: 28 $
#	$Author: TatsuyaEndo $
#	$HeadURL: file:///C:/Repo/trunk/ruby/lib/debug.rb $
#	
# =================================================

	# �f�o�b�O�p
	def check_fixnum(fixDebugNum, strCheckFixnum)
		
		puts ""
		puts "fixDebugNum                   = #{fixDebugNum}"
		puts "strCheckFixnum                = #{strCheckFixnum}"
	end

	# �f�o�b�O�p
	def check_string(fixDebugNum, strCheckString)
		
		puts ""
		puts "fixDebugNum                   = #{fixDebugNum}"
		puts "strCheckString                = #{strCheckString}"
	end

	# �f�o�b�O�p
	def check_array(fixDebugNum, strCheckArray)
		
		puts ""
		puts "fixDebugNum                   = #{fixDebugNum}"
		puts "strCheckArray                 = #{strCheckArray}"
		puts "strCheckArray.class           = #{strCheckArray.class}"
		puts "strCheckArray[0].class        = #{strCheckArray[0].class}"
		puts "strCheckArray[0][0].class     = #{strCheckArray[0][0].class}"
	end

	# �p�����[�^�`�F�b�N
	def check_param(*args)
		for i in 0 .. (args.length - 1)
			case args[i]
				when ""
					puts "Parameter error!"
					puts "ArgsNum   = #{i}"
					puts caller(1)
				when nil
					puts "Parameter error!"
					puts "ArgsNum   = #{i}"
					puts caller(1)
				else
					# None
			end
		end
	end
