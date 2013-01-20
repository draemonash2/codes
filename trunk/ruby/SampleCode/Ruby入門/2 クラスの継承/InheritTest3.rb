class Car
  def accele
    print("ƒAƒNƒZƒ‹‚ğ“¥‚İ‚Ü‚µ‚½\n")
  end
end

class Soarer < Car
  def accele
    super
    print("‰Á‘¬‚µ‚Ü‚µ‚½\n")
  end
end

soarer = Soarer.new
soarer.accele
