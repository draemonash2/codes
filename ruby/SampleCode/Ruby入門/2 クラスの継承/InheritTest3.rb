class Car
  def accele
    print("�A�N�Z���𓥂݂܂���\n")
  end
end

class Soarer < Car
  def accele
    super
    print("�������܂���\n")
  end
end

soarer = Soarer.new
soarer.accele
