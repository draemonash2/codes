class Car
  def dispClassname
    print("Car class\n")
  end

  def dispString(str)
    print(str, "\n")
  end
end

car = Car.new()
car.dispClassname
car.dispString("crown")
