//
//  Test.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import Foundation
import SwiftUI

struct Item {
    let item_id: String = UUID().uuidString
    var item_name: String = ""
    var status:Dictionary<Date, Int> = [:]
    var skip_num: Int = 999
    var color: Color = Color.red
    var is_archived: Bool = false
}

struct HabChainData {
    var item_id_list: [String] = []
    var items: [Item] = []
}

var hab_chain_data: HabChainData = HabChainData()

func Test2() -> String
{
    var dicStatus:Dictionary<Date, Int> = [:]
    
    let today = Date()
    let yesterday = Calendar.current.date(byAdding: .day,value: -1, to: Date())!
    let yesterday_1 = Calendar.current.date(byAdding: .day,value: -2, to: Date())!

    let dateFormatter = DateFormatter()
    dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))

    dicStatus.updateValue(1, forKey: today)
    dicStatus.updateValue(2, forKey: yesterday)
    for (key,value) in dicStatus {
        print("\(dateFormatter.string(from: key)) : \(value)")
    }
    print("\(dicStatus[today]!)")
    print("\(dicStatus[yesterday]!)")
    //print("\(dicStatus[yesterday_1]!)")
    print("\(dicStatus.keys.contains(today))")
    print("\(dicStatus.keys.contains(yesterday))")
    print("\(dicStatus.keys.contains(yesterday_1))")
    return "1"
}

func Test3() -> String
{
    //let item_id_tmp: String = UUID().uuidString
    //let item_id_tmp2: String = UUID().uuidString
    let item: Item = Item(item_name: "aaa", skip_num: 10)
    hab_chain_data.item_id_list.append(UUID().uuidString)
    hab_chain_data.item_id_list.append(UUID().uuidString)
    hab_chain_data.items.append(item)
    print(hab_chain_data.item_id_list[0])
    print(hab_chain_data.item_id_list[1])
    print(hab_chain_data.items[0].item_id)
    print(hab_chain_data.items[0].skip_num)
    return hab_chain_data.item_id_list[0]
}
