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
