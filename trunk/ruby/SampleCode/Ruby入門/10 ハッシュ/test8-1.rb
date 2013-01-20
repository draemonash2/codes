#! ruby -Ku
require "kconv"

hash = {"Lemon" => 100, "Orange" => 150, "Banana" => 250}
p hash
print("\n")

print(Kconv.tosjis("keysメソッド\n"));
key_array = hash.keys
p key_array

print(Kconv.tosjis("valuesメソッド\n"));
value_array = hash.values
p value_array

print(Kconv.tosjis("to_aメソッド\n"));
array = hash.to_a
p array
