#
# 値の取得
#

#! ruby -Ku
require "kconv"

hash = {"Yamada" => 34, "Katou" => 28, "Endou" => 18}

print(Kconv.tosjis("年齢 : "), hash["Katou"], "\n")
print(Kconv.tosjis("年齢 : "), hash.fetch("Yamada"), "\n")
print(Kconv.tosjis("年齢 : "), hash["Takahashi"], "\n")
print(Kconv.tosjis("年齢 : "), hash.fetch("Takahashi"), "\n")

=begin
年齢 : 28
年齢 : 34
年齢 :
test2-1.rb:9:in `fetch': key not found: "Takahashi" (KeyError)
        from test2-1.rb:9:in `<main>'
=end



#
# Hashクラス
#

#! ruby -Ku
require "kconv"

hash = Hash["Yamada" => "Tokyo", "Katou" => "Osaka", "Endou" => "Fukuoka"]

print(Kconv.tosjis("コピー元\n"))
hash.each do |key, val|
  print("key:", key.object_id, ", val:", val.object_id, "\n")
end

print(Kconv.tosjis("コピー先\n"))
copyhash = Hash[hash]
copyhash.each do |key, val|
  print("key:", key.object_id, ", val:", val.object_id, "\n")
end

=begin
コピー元
key:15264540, val:15264612
key:15264528, val:15264588
key:15264516, val:15264564
コピー先
key:15264540, val:15264612
key:15264528, val:15264588
key:15264516, val:15264564
=end



#
# デフォルトの設定
#

#! ruby -Ku
require "kconv"

addressh1 = Hash.new("none")
addressh1["Itou"] = "Tokyo"

print(Kconv.tosjis("住所 : "), addressh1["Itou"], "\n")
print(Kconv.tosjis("住所 : "), addressh1["Yamada"], "\n")

print("\n");

addressh2 = Hash.new{|hash, key|
  hash[key] = key
}

print(Kconv.tosjis("住所 : "), addressh2["Yamada"], "\n")
print(Kconv.tosjis("住所 : "), addressh2["Takahashi"], "\n")
print(Kconv.tosjis("要素数 : "),addressh2.length(), "\n")

addressh3 = Hash.new{|hash, key|
  key
}

print("\n");

print(Kconv.tosjis("住所 : "), addressh3["Yamada"], "\n")
print(Kconv.tosjis("住所 : "), addressh3["Takahashi"], "\n")
print(Kconv.tosjis("要素数 : "),addressh3.length(), "\n")

print("\n");

addressh4 = Hash.new("none")

print(Kconv.tosjis("住所 : "), addressh4["Itou"], "\n")
print(Kconv.tosjis("住所 : "), addressh4.fetch("Yamada", "nothing"), "\n")
print(Kconv.tosjis("住所 : "), addressh4.fetch("Endou"){|key|key}, "\n")

print("\n");

addressh5 = Hash.new()
addressh5.default = "none"
print(Kconv.tosjis("住所 : "), addressh5["Yamada"], "\n")

=begin
住所 : Tokyo
住所 : none

住所 : Yamada
住所 : Takahashi
要素数 : 2

住所 : Yamada
住所 : Takahashi
要素数 : 0

住所 : none
住所 : nothing
住所 : Endou

住所 : none
=end



#
# 要素の追加と値の変更
#

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

=begin
Lemon = 100
Orange = 150

Lemon = 120
Orange = 150
現在の要素数:2
Banana = 90
追加後の要素数:3

Peach = 210
Banana = 80
=end



#
# ハッシュのサイズの取得
#

#! ruby -Ku
require "kconv"

hash = {"Lemon" => 100, "Orange" => 150}
p hash
print(Kconv.tosjis("配列の要素数 = "), hash.length, "\n");

hash["Banana"] = 80
p hash
print(Kconv.tosjis("配列の要素数 = "), hash.size, "\n");

=begin
{"Lemon"=>100, "Orange"=>150}
配列の要素数 = 2
{"Lemon"=>100, "Orange"=>150, "Banana"=>80}
配列の要素数 = 3
=end



#
# ハッシュに対する繰り返し
#

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

=begin
eachメソッド
Lemon=>100
Orange=>150
Banana=>250
each_keyメソッド
key = Lemon
key = Orange
key = Banana
each_valueメソッド
value = 100
value = 150
value = 250
=end



#
# ハッシュに含まれるキーや値を配列として取得
#

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

=begin
{"Lemon"=>100, "Orange"=>150, "Banana"=>250}

keysメソッド
["Lemon", "Orange", "Banana"]
valuesメソッド
[100, 150, 250]
to_aメソッド
[["Lemon", 100], ["Orange", 150], ["Banana", 250]]
=end
