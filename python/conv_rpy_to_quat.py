#!/usr/bin/env python3

# usage : python3 conv_rpy_to_quat.py <roll> <pitch> <yaw>

import sys
import math

def main():
    args = sys.argv
    if len(args) != 4:
        print('[error] wrong number of arguments')
        print('  usage : python3 conv_rpy_to_quat.py <roll> <pitch> <yaw>')
        return 0
    
    roll = float(args[1])
    pitch = float(args[2])
    yaw = float(args[3])
    
    cy = math.cos(yaw * 0.5);
    sy = math.sin(yaw * 0.5);
    cp = math.cos(pitch * 0.5);
    sp = math.sin(pitch * 0.5);
    cr = math.cos(roll * 0.5);
    sr = math.sin(roll * 0.5);
    
    w = cr * cp * cy + sr * sp * sy;
    x = sr * cp * cy - cr * sp * sy;
    y = cr * sp * cy + sr * cp * sy;
    z = cr * cp * sy - sr * sp * cy;
    
    print("x: " + str(x) + ",")
    print("y: " + str(y) + ",")
    print("z: " + str(z) + ",")
    print("w: " + str(w))

if __name__ == "__main__":
    main()


