class Car
  def accele(acceletime)
    print(acceletime, "�b�ԃA�N�Z���𓥂݂܂���\n")
  end
end

class Soarer < Car
  def accele(acceletime)
    super
    print("�������܂���")
  end
end

soarer = Soarer.new
soarer.accele(10)
