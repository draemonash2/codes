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
