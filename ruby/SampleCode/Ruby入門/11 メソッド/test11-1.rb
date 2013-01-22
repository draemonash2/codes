#
# Rubyにおけるトップレベル
#

#! ruby -Ku

print(self.to_s + "\n")
print(self.class.to_s + "\n")

=begin
main
Object
=end



#
# メソッドの定義と呼び出し
#

#! ruby -Ku
require "kconv"

def printHello
  print("Hello\n")
end

print(Kconv.tosjis("メソッドを呼び出します\n"))
printHello
print(Kconv.tosjis("メソッドを呼び出しました\n"))

=begin
メソッドを呼び出します
Hello
メソッドを呼び出しました
=end



#
# 引数を付けたメソッド呼び出し
#

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

=begin
Hello,Yamada

呼び出し前:Tokyo
呼び出し前:Tokyo,Japan
=end



#
# 引数のデフォルト値
#

#! ruby -Ku
require "kconv"

def printHello(msg="No msg", name="No name")
  print(Kconv.tosjis(msg + "," + name + "\n"))
end

printHello("こんにちは", "佐藤")
printHello("お元気ですか")
printHello()

=begin
こんにちは,佐藤
お元気ですか,No name
No msg,No name
=end



#
# 引数を配列として受け取る
#

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

=begin
こんにちは,
こんにちは,山田
こんにちは,山田 遠藤
こんにちは,山田 遠藤 加藤 高橋
=end



#
# メソッドの戻り値
#

#! ruby -Ku
require "kconv"

def hikaku(num1, num2)
  print("num1 = ", num1, "\n")
  print("num2 = ", num2, "\n")
  if num1 > num2 then
    return num1
  else
    return num2
  end
end

num = hikaku(10, 18)
print(Kconv.tosjis("大きい値は"), num, Kconv.tosjis("です\n"))

num = hikaku(27, 5)
print(Kconv.tosjis("大きい値は"), num, Kconv.tosjis("です\n"))

=begin
num1 = 10
num2 = 18
大きい値は18です
num1 = 27
num2 = 5
大きい値は27です
=end



#
# 多重代入を使って複数の戻り値を取得
#

#! ruby -Ku
require "kconv"

def keisan(num1, num2)
  print("num1 = ", num1, "\n")
  print("num2 = ", num2, "\n")
  return num1 + num2, num1 - num2
end

plus, minus = keisan(10, 25)
print(Kconv.tosjis("加算の結果:"), plus, "\n")
print(Kconv.tosjis("減算の結果:"), minus, "\n")

=begin
num1 = 10
num2 = 25
加算の結果:35
減算の結果:-15
=end
