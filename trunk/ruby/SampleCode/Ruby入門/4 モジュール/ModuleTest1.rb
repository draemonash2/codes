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
    print(x, "��", y, "�̕��ς�", kekka, "�ł�\n")
  end
end

test = Test.new
test.dispHeikin(10, 8)
