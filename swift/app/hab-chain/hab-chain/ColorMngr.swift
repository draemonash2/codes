//
//  ColorMngr.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/31.
//

import Foundation
import SwiftUI

func getColorString(color: Color, continuation_count: Int) -> String {
    var color_name: String = ""
    switch color {
        case Color.red: color_name = "color_red"
        case Color.blue: color_name = "color_blue"
        case Color.green: color_name = "color_green"
        default: return ""
    }
    
    var color_index: Int = 0
    let min: Int = 0
    let max: Int = 5
    if continuation_count < min {
        color_index = min
    } else if min <= continuation_count && continuation_count <= max {
        color_index = continuation_count
    } else {
        color_index = max
    }
    
    return String(color_name) + String(color_index)
}
func _test_getColorString() {
    print("### getColorString() test start ###")
    print(getColorString(color: Color.red, continuation_count: 0))
    print(getColorString(color: Color.red, continuation_count: 1))
    print(getColorString(color: Color.red, continuation_count: 3))
    print(getColorString(color: Color.red, continuation_count: 5))
    print(getColorString(color: Color.red, continuation_count: 6))
    print(getColorString(color: Color.blue, continuation_count: 3))
    print(getColorString(color: Color.green, continuation_count: 3))
    print(getColorString(color: Color.white, continuation_count: 3))
    print("### getColorString() test finished ###")
    print("")
}
