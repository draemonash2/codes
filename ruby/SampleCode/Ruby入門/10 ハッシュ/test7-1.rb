#! ruby -Ku
require "kconv"

hash = {"Lemon" => 100, "Orange" => 150, "Banana" => 250}

print(Kconv.tosjis("eachメソッド\n"));
hash.each{|key, value|
  print(key + "=>", value, "\n")
}

print(Kconv.tosjis("each_keyメソッド\n"));
hash.each_key{|key|
  print("key = " + key + "\n")
}

print(Kconv.tosjis("each_valueメソッド\n"));
hash.each_value{|value|
  print("value = ", value, "\n")
}
