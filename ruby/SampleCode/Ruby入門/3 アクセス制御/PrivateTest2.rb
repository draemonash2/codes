class Car
  private

  def calcSpeed(acceletime)
    return acceletime * 10
  end

  public

  def accele(acceletime=1)
    print("�A�N�Z���𓥂݂܂���\n")
    print("�X�s�[�h��", calcSpeed(acceletime), "Km�ł�\n")
  end

  def brake
    print("�u���[�L�𓥂݂܂���\n")
  end

end

car = Car.new
car.accele(10)
