class Car
  @@count = 0

  def initialize(carname="����`")
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

print("���ݐ������ꂽ�I�u�W�F�N�g��:", car1.getCount())
