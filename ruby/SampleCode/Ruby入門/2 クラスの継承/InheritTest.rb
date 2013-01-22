#
# クラスを継承する
#

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
end

class Crown < Car
  def reclining
    print("シートをリクライニングしました\n")
  end
end

soarer = Soarer.new
soarer.openRoof
soarer.accele

crown = Crown.new
crown.reclining
crown.brake

=begin
ルーフを開けました
アクセルを踏みました
シートをリクライニングしました
ブレーキを踏みました
=end



#
# メソッドのオーバーライド
#

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

=begin
アクセルを踏んで加速しました
アクセルを踏みました
=end



#
# スーパークラスのメソッドを呼び出す
#

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

=begin
アクセルを踏みました
加速しました
=end



#
# 引数があるスーパークラスのメソッドを呼び出す
#

class Car
  def accele(acceletime)
    print(acceletime, "秒間アクセルを踏みました\n")
  end
end

class Soarer < Car
  def accele(acceletime)
    super
    print("加速しました")
  end
end

soarer = Soarer.new
soarer.accele(10)

=begin
10秒間アクセルを踏みました
加速しました
=end
