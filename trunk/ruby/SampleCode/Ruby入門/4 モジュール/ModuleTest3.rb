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
    print("2�̒l", x, "��", y, "�̒��ŏ������l��", min, "�ł�\n")
  end
end

test = Test.new
test.dispValue(10, 8)
