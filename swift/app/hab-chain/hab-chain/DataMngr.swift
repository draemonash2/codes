//
//  Test.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import Foundation
import SwiftUI

enum ItemStatus {
    case NotYet
    case Finish
    case Skip
}

struct Item {
    var item_name: String = ""
    var status:Dictionary<Date, ItemStatus> = [:]
    var skip_num: Int = 999
    var color: Color = Color.red
    var is_archived: Bool = false
}

struct HabChainData {
    var item_id_list: [String] = []
    var items: Dictionary<String, Item> = [:]
    var id = UUID()

    init() {
        self.setValueForTest()
        self.printAll()
        self._test_calcContinuationCount()
    }
    func generateItemId() -> String
    {
        return UUID().uuidString
    }
    mutating func addItem(
        new_item_id: String,
        new_item: Item
    )
    {
        self.items.updateValue(new_item, forKey: new_item_id)
        self.item_id_list.append(new_item_id)
    }
    mutating func removeItem(
        trgt_item_id: String
    )
    {
        //remove from items
        self.items.removeValue(forKey:trgt_item_id)
        
        //remove item_id_list
        var remove_index: Int = 0
        var is_matched: Bool = false
        for (cur_index, item_id) in self.item_id_list.enumerated() {
            if item_id == trgt_item_id {
                remove_index = cur_index
                is_matched = true
                break
            }
        }
        if is_matched == true {
            self.item_id_list.remove(at: remove_index)
        }
        self.id = UUID() // for refresh view
    }
    func getItem(item_id: String) -> Item {
        return self.items[item_id]!
    }
    func existItemName(item_name: String) -> Bool {
        var is_exist: Bool = false
        for (_,value) in self.items {
            if value.item_name == item_name {
                is_exist = true
                break
            }
        }
        return is_exist
    }
    func getItemId(
        item_name: String
    ) -> String
    {
        var ret: String = ""
        for (item_id, item) in self.items {
            if item.item_name == item_name {
                ret = item_id
                break
            }
        }
        return ret
    }
    mutating func setItemStatus(
        item_id: String,
        date: Date,
        item_status: ItemStatus
    )
    {
        self.id = UUID() // for refresh view
        self.items[item_id]!.status = [date : item_status]
    }
    mutating func toggleItemStatus(
        item_id: String,
        date: Date
    )
    {
        self.id = UUID() // for refresh view
        let cur_itemstatus = self.items[item_id]!.status[date]
        switch cur_itemstatus {
            case .NotYet: self.items[item_id]!.status = [date : .Finish]
            case .Finish: self.items[item_id]!.status = [date : .Skip]
            case .Skip:   self.items[item_id]!.status = [date : .NotYet]
            default:      print("[error] unknown itemstatus.")
        }
    }
    func calcContinuationCount(
        base_date: Date,
        item_id: String
    ) -> Int
    {
        var date_offset: Int = 0
        var continuation_count: Int = 0
        var is_coutinue: Bool = true
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))
        if let unwrapped_item = self.items[item_id] {
            for (item_date, value) in unwrapped_item.status.sorted(by: { $0.key > $1.key }) {
                let date = Calendar.current.date(byAdding: .day,value: date_offset, to: base_date)!
                //print("\(dateFormatter.string(from: date))")
                if item_date <= date {
                    if dateFormat(date: item_date) == dateFormat(date: date) {
                        switch value {
                            case .Finish:
                                continuation_count += 1
                                is_coutinue = true
                            case .Skip:
                                is_coutinue = true
                            case .NotYet:
                                is_coutinue = false
                        }
                    } else {
                        is_coutinue = false
                    }
                    if is_coutinue == true {
                        date_offset -= 1
                    } else {
                        break
                    }
                }
            }
        }
        return continuation_count
    }
    func _test_calcContinuationCount() {
        var offset: Int = 0
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1
    }
    mutating func setValueForTest() {
        let item1: Item = Item(item_name: "aa", skip_num: 10)
        let item3: Item = Item(item_name: "cccc", skip_num: 30)
        let item2: Item = Item(
            item_name: "bbb",
            status: [
                Calendar.current.date(byAdding: .day,value: 0, to: Date())!     : .Finish,
                Calendar.current.date(byAdding: .day,value: -1, to: Date())!    : .NotYet,
                Calendar.current.date(byAdding: .day,value: -2, to: Date())!    : .Skip,
                Calendar.current.date(byAdding: .day,value: -3, to: Date())!    : .Finish,
                Calendar.current.date(byAdding: .day,value: -4, to: Date())!    : .Finish,
                Calendar.current.date(byAdding: .day,value: -5, to: Date())!    : .Finish,
                Calendar.current.date(byAdding: .day,value: -7, to: Date())!    : .NotYet,
                Calendar.current.date(byAdding: .day,value: -8, to: Date())!    : .Skip,
                Calendar.current.date(byAdding: .day,value: -9, to: Date())!    : .Finish
            ],
            skip_num: 20
        )

        self.addItem(new_item_id: generateItemId(), new_item: item2)
        self.addItem(new_item_id: generateItemId(), new_item: item3)
        self.addItem(new_item_id: generateItemId(), new_item: item1)
    }
    func printAll() {
        print("### item_id_list ###")
        for (cur_index, item_id) in self.item_id_list.enumerated() {
            print("\(cur_index) : \(item_id) : \(self.items[item_id]!.item_name)")
        }
        print("### items ###")
        for (key,value) in self.items {
            print("\(key) : \(value.item_name)")
        }
        print("")
    }
}

func dateFormat(date: Date) -> String {
    let f = DateFormatter()
    f.dateStyle = .long
    f.timeStyle = .none
    return f.string(from: date)
}

func Test2() -> String
{
    var dicStatus:Dictionary<Date, ItemStatus> = [:]
    
    let today = Date()
    let yesterday = Calendar.current.date(byAdding: .day,value: -1, to: Date())!
    let yesterday_1 = Calendar.current.date(byAdding: .day,value: -2, to: Date())!

    let dateFormatter = DateFormatter()
    dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))

    dicStatus.updateValue(.Skip, forKey: today)
    dicStatus.updateValue(.Finish, forKey: yesterday)
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

#if false
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
#endif

