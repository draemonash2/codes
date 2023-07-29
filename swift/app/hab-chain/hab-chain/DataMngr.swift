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
    case Done
    case Skip
}

struct Item {
    var item_name: String = ""
    var status:Dictionary<String, ItemStatus> = [:]
    var skip_num: Int = 999
    var color: Color = Color.red
    var is_archived: Bool = false
}

struct HabChainData {
    var item_id_list: [String] = []
    var items: Dictionary<String, Item> = [:]

    init() {
        self.setValueForTest()
        self.printAll()
        //self._test_calcContinuationCount()
        //_test_convDateToStr()
        //self._test_toggleItemStatus()
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
        self.items[item_id]!.status = [convDateToStr(date: date) : item_status]
    }
    mutating func toggleItemStatus(
        item_id: String,
        date: Date
    )
    {
        let date_str = convDateToStr(date: date)
        let cur_itemstatus = self.items[item_id]!.status[date_str]
        switch cur_itemstatus {
            case .NotYet: self.items[item_id]!.status.updateValue(.Done, forKey: date_str)
            case .Done:   self.items[item_id]!.status.updateValue(.Skip, forKey: date_str)
            case .Skip:   self.items[item_id]!.status.updateValue(.NotYet, forKey: date_str)
            default:      self.items[item_id]!.status.updateValue(.Done, forKey: date_str)
        }
    }
    mutating func _test_toggleItemStatus() {
        let item_id: String = self.getItemId(item_name: "bbb")
        self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
    }
    func calcContinuationCount(
        base_date: Date,
        item_id: String
    ) -> Int
    {
        var date_offset: Int = 0
        var continuation_count: Int = 0
        var is_coutinue: Bool = true
        if let unwrapped_item = self.items[item_id] {
            for (item_date, item_status) in unwrapped_item.status.sorted(by: { $0.key > $1.key }) {
                let date = Calendar.current.date(byAdding: .day,value: date_offset, to: base_date)!
                let date_str = convDateToStr(date: date)
                //print("\(date_str)")
                if item_date <= date_str {
                    if item_date == date_str {
                        switch item_status {
                            case .Done:
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
    func convDateToStr(date: Date) -> String {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))
        return dateFormatter.string(from: date)
    }
    func _test_convDateToStr() {
        print(convDateToStr(date: Date()))
    }
    mutating func setValueForTest() {
        let item1: Item = Item(item_name: "aa", skip_num: 10)
        let item3: Item = Item(
            item_name: "cccc",
            status: [
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: 0, to: Date())!)     : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -1, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -6, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -7, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -8, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -9, to: Date())!)    : .Done
            ],
            skip_num: 30
        )
        let item2: Item = Item(
            item_name: "bbb",
            status: [
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: 0, to: Date())!)     : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -1, to: Date())!)    : .NotYet,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -7, to: Date())!)    : .NotYet,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -8, to: Date())!)    : .Skip,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -9, to: Date())!)    : .Done
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
            for (key,value) in value.status.sorted(by: { $0.key > $1.key }) {
                print("\(key) : \(value)")
            }
        }
        print("")
    }
}


#if false
func Test2() -> String
{
    var dicStatus:Dictionary<String, ItemStatus> = [:]
    
    let today = Date()
    let yesterday = Calendar.current.date(byAdding: .day,value: -1, to: Date())!
    let yesterday_1 = Calendar.current.date(byAdding: .day,value: -2, to: Date())!

    dicStatus.updateValue(.Skip, forKey: today)
    dicStatus.updateValue(.Done, forKey: yesterday)
    for (key,value) in dicStatus {
        print("\(convDateToStr(key)) : \(value)")
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
#endif

