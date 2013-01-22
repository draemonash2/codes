#! ruby -Ku
require "kconv"

def printHello(msg, name)
  print(msg + "," + name + "\n")
end

def addString(str)
  str << ",Japan"
end

printHello("Hello", "Yamada")
print("\n")

address = "Tokyo"
print(Kconv.tosjis("呼び出し前:") + address + "\n")

addString(address)
print(Kconv.tosjis("呼び出し前:") + address + "\n")
