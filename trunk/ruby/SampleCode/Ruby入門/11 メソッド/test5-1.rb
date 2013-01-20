#! ruby -Ku
require "kconv"

def printHello(msg, *names)
  allname = ""
  names.each do |name|
    allname << name << " "
  end
  print(Kconv.tosjis(msg + "," + allname + "\n"))
end

printHello("こんにちは")
printHello("こんにちは", "山田")
printHello("こんにちは", "山田", "遠藤")
printHello("こんにちは", "山田", "遠藤", "加藤", "高橋")
