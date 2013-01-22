class Car
  def accele
    print("アクセルを踏みました\n")
  end

  def brake
    print("ブレーキを踏みました\n")
  end
end

class Soarer < Car
  def openRoof
    print("ルーフを開けました\n")
  end

  def accele
    print("アクセルを踏んで加速しました\n")
  end
end

class Crown < Car
  def reclining
    print("シートをリクライニングしました\n")
  end
end

soarer = Soarer.new
soarer.accele

crown = Crown.new
crown.accele
