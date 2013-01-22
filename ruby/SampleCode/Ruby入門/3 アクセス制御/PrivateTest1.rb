class Car
  def accele(acceletime=1)
    print("アクセルを踏みました\n")
    print("スピードは", calcSpeed(acceletime), "Kmです\n")
  end

  public :accele

  def brake
    print("ブレーキを踏みました\n")
  end

  public :brake

  def calcSpeed(acceletime)
    return acceletime * 10
  end

  private :calcSpeed
end

car = Car.new
car.accele(10)
