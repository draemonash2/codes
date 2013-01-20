#! ruby -Ku
require "kconv"

hash = {"Yamada" => 34, "Katou" => 28, "Endou" => 18}

print(Kconv.tosjis("年齢 : "), hash["Katou"], "\n")
print(Kconv.tosjis("年齢 : "), hash.fetch("Yamada"), "\n")
print(Kconv.tosjis("年齢 : "), hash["Takahashi"], "\n")
print(Kconv.tosjis("年齢 : "), hash.fetch("Takahashi"), "\n")
