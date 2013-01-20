class Car
  def accele(acceletime)
    print(acceletime, "•bŠÔƒAƒNƒZƒ‹‚ğ“¥‚İ‚Ü‚µ‚½\n")
  end
end

class Soarer < Car
  def accele(acceletime)
    super
    print("‰Á‘¬‚µ‚Ü‚µ‚½")
  end
end

soarer = Soarer.new
soarer.accele(10)
