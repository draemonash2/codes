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
