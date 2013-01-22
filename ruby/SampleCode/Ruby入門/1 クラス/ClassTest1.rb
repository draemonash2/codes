class Car
  def initialize(carname)
    @name = carname
  end
  
  def dispName
    print(@name)
  end
end

car = Car.new("crown")
car.dispName

