class Car
  def accele(acceletime=1)
    print("�A�N�Z���𓥂݂܂���\n")
    print("�X�s�[�h��", calcSpeed(acceletime), "Km�ł�\n")
  end

  public :accele

  def brake
    print("�u���[�L�𓥂݂܂���\n")
  end

  public :brake

  def calcSpeed(acceletime)
    return acceletime * 10
  end

  private :calcSpeed
end

car = Car.new
car.accele(10)
