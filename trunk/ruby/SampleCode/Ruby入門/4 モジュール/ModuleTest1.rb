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
    print(x, "‚Æ", y, "‚Ì•½‹Ï‚Í", kekka, "‚Å‚·\n")
  end
end

test = Test.new
test.dispHeikin(10, 8)
