#
# モジュールの定義
#

module HeikinModule
  def heikin(x, y)
    kekka = (x + y) / 2
    return kekka
  end
end

class Test
  include HeikinModule

  def dispHeikin(x, y)
    kekka = heikin(x, y)
    print(x, "と", y, "の平均は", kekka, "です\n")
  end
end

test = Test.new
test.dispHeikin(10, 8)

=begin

=end



#
# モジュールを関数のように使う
#

module SuuchiModule
  def minValue(x, y)
    if x < y
      return x
    else
      return y
    end
  end

  def maxValue(x, y)
    if x > y
      return x
    else
      return y
    end
  end

  module_function :minValue
  module_function :maxValue
end

include SuuchiModule
print(minValue(10, 8), "\n")
print(maxValue(10, 8), "\n")

=begin

=end



#
# クラスの中にモジュールをインクルードする
#

module SuuchiModule
  def minValue(x, y)
    if x < y
      return x
    else
      return y
    end
  end
end

class Test
  include SuuchiModule

  def dispValue(x, y)
    min = minValue(x, y)
    print("2つの値", x, "と", y, "の中で小さい値は", min, "です\n")
  end
end

test = Test.new
test.dispValue(10, 8)

=begin

=end
