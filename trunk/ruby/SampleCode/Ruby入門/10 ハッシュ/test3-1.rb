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
