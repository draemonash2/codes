class Car
  @@count = 0

  def initialize(carname="未定義")
    @name = carname
    @@count += 1
  end

  def getCount()
    return @@count;
  end

end

car1 = Car.new("crown")
car2 = Car.new("civic")
car3 = Car.new("alto")

print("現在生成されたオブジェクト数:", car1.getCount())
