#
# アクセス制御とは
#

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

=begin
アクセルを踏みました
スピードは100Kmです
=end



#
# アクセス制御をまとめて設定する
#

class Car
  private

  def calcSpeed(acceletime)
    return acceletime * 10
  end

  public

  def accele(acceletime=1)
    print("アクセルを踏みました\n")
    print("スピードは", calcSpeed(acceletime), "Kmです\n")
  end

  def brake
    print("ブレーキを踏みました\n")
  end

end

car = Car.new
car.accele(10)

=begin
アクセルを踏みました
スピードは100Kmです
=end
