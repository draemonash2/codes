class Car
  def initialize(carname="����`")
    @name = carname
  end
  
  attr_accessor :name
end

car = Car.new()
car.name = "civic"
print(car.name)
