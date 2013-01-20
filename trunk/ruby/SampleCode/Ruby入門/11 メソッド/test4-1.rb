#! ruby -Ku
require "kconv"

def printHello(msg="No msg", name="No name")
  print(Kconv.tosjis(msg + "," + name + "\n"))
end

printHello("こんにちは", "佐藤")
printHello("お元気ですか")
printHello()
