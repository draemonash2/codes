class Reji
  SHOUHIZEI = 0.05

  def initialize(init=0)
    @sum = init
  end
  
  def kounyuu(kingaku)
    @sum += kingaku
    print("‚¨”ƒ‚¢ã‚°:", kingaku, "\n")
  end
  
  def goukei()
    return @sum * (1 + SHOUHIZEI)
  end
end

reji = Reji.new(0)
reji.kounyuu(100)
reji.kounyuu(80)
print("‡Œv‹àŠz:", reji.goukei(), "\n")

print("Á”ïÅ—¦:", Reji::SHOUHIZEI)
