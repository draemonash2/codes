#! ruby -Ku
require "kconv"

hash = {"Lemon" => 100, "Orange" => 150}
p hash
print(Kconv.tosjis("配列の要素数 = "), hash.length, "\n");

hash["Banana"] = 80
p hash
print(Kconv.tosjis("配列の要素数 = "), hash.size, "\n");
