#! ruby -Ku
require "kconv"

hash = {"Lemon" => 100, "Orange" => 150}
print("Lemon = ", hash["Lemon"], "\n");
print("Orange = ", hash["Orange"], "\n");

print("\n");

hash["Lemon"] = 120

print("Lemon = ", hash["Lemon"], "\n");
print("Orange = ", hash["Orange"], "\n");
print(Kconv.tosjis("現在の要素数:"), hash.length, "\n")

hash["Banana"] = 90
print("Banana = ", hash["Banana"], "\n");
print(Kconv.tosjis("追加後の要素数:"), hash.length, "\n")

print("\n");

hash.store("Peach", 210)
hash.store("Banana", 80)
print("Peach = ", hash["Peach"], "\n");
print("Banana = ", hash["Banana"], "\n");
