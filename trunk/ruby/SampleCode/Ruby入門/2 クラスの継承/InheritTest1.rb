class Car
  def accele
    print("�A�N�Z���𓥂݂܂���\n")
  end

  def brake
    print("�u���[�L�𓥂݂܂���\n")
  end
end

class Soarer < Car
  def openRoof
    print("���[�t���J���܂���\n")
  end
end

class Crown < Car
  def reclining
    print("�V�[�g�����N���C�j���O���܂���\n")
  end
end

soarer = Soarer.new
soarer.openRoof
soarer.accele

crown = Crown.new
crown.reclining
crown.brake
