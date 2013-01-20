#
# クラスとは
#

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

=begin
crown
=end



#
# クラスの定義とオブジェクトの作成
#

class Car
  def initialize(carname)
    @name = carname
  end
  
  def dispName
    print(@name, "\n")
  end
end

car1 = Car.new("crown")
car1.dispName

car2 = Car.new("civic")
car2.dispName

=begin
crown
civic
=end



#
# インスタンスメソッド
#

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

=begin
Car class
crown
=end



#
# インスタンス変数
#

class Car
  def setName(str)
    @name = str
  end
  
  def dispName()
    print(@name, "\n")
  end
end

car1 = Car.new()
car1.setName("crown")

car2 = Car.new()
car2.setName("civic")

car1.dispName()
car2.dispName()




=begin
crown
civic
=end



#
# initializeメソッド
#

class Car
  def initialize(carname="未定義")
    @name = carname
  end

  def dispName()
    print(@name, "\n")
  end
end

car1 = Car.new("civic")
car2 = Car.new()

car1.dispName()
car2.dispName()

=begin
civic
未定義
=end



#
# アクセスメソッド
#

class Car
  def initialize(carname="未定義")
    @name = carname
  end
  
  attr_accessor :name
end

car = Car.new()
car.name = "civic"
print(car.name)

=begin
civic
=end



#
# 定数
#

class Reji
  SHOUHIZEI = 0.05

  def initialize(init=0)
    @sum = init
  end
  
  def kounyuu(kingaku)
    @sum += kingaku
    print("お買い上げ:", kingaku, "\n")
  end
  
  def goukei()
    return @sum * (1 + SHOUHIZEI)
  end
end

reji = Reji.new(0)
reji.kounyuu(100)
reji.kounyuu(80)
print("合計金額:", reji.goukei(), "\n")

print("消費税率:", Reji::SHOUHIZEI)

=begin
お買い上げ:100
お買い上げ:80
合計金額:189.0
消費税率:0.05
=end



#
# クラス変数
#

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

=begin
現在生成されたオブジェクト数:3
=end
