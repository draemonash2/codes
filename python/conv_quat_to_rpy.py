#!/usr/bin/env python3

# usage : python3 conv_quat_to_rpy.py <x> <y> <z> <w>

import sys
import math

def main():
    args = sys.argv
    if len(args) != 5:
        print('[error] wrong number of arguments')
        print('  usage : python3 conv_quat_to_rpy.py <x> <y> <z> <w>')
        return 0
    
    x = float(args[1])
    y = float(args[2])
    z = float(args[3])
    w = float(args[4])
    
    q0q0 = w * w;
    q1q1 = x * x;
    q2q2 = y * y;
    q3q3 = z * z;
    q0q1 = w * x;
    q0q2 = w * y;
    q0q3 = w * z;
    q1q2 = x * y;
    q1q3 = x * z;
    q2q3 = y * z;
    
    roll = math.atan2((2.0 * (q2q3 + q0q1)), (q0q0 - q1q1 - q2q2 + q3q3));
    pitch = -math.asin((2.0 * (q1q3 - q0q2)));
    yaw = math.atan2((2.0 * (q1q2 + q0q3)), (q0q0 + q1q1 - q2q2 - q3q3));
    
    print("roll  : " + str(roll))
    print("pitch : " + str(pitch))
    print("yaw   : " + str(yaw))

if __name__ == "__main__":
    main()


