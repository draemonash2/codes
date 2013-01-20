class Car
  def accele
    print("アクセルを踏みました\n")
  end
end

class Soarer < Car
  def accele
    super
    print("加速しました\n")
  end
end

soarer = Soarer.new
soarer.accele
