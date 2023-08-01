//
//  DataMngr.swift
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
        //self._test_calcTotalItemStatus()
        //self._test_calcContinuationCountAll()
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
    func getItemStatusStr(
        item_id: String,
        date: Date
    ) -> String {
        var ret: String = ""
        if let unwrapped_item_id = self.items[item_id] {
            let date_str: String = self.convDateToStr(date: date)
            if let unwrapped_status = unwrapped_item_id.status[date_str] {
                ret = convItemStatusToStr(item_status: unwrapped_status)
            }
        }
        return ret
    }
    func convItemStatusToStr(
        item_status: ItemStatus
    ) -> String
    {
        var ret: String = ""
        switch item_status {
            case .NotYet: ret = "NotYet"
            case .Done: ret = "Done"
            case .Skip: ret = "Skip"
        }
        return ret
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
        print("### toggleItemStatus() test start ###")
        let item_id: String = self.getItemId(item_name: "bbb")
        self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        self.toggleItemStatus(item_id: item_id, date: Date()); self.printAll()
        print("### toggleItemStatus() test finished ###")
        print("")
    }
    func calcContinuationCount(
        base_date: Date,
        item_id: String
    ) -> Int
    {
        var date_offset: Int = 0
        var continuation_count: Int = 0
        if let unwrapped_item = self.items[item_id] {
            while true {
                var is_continue: Bool = true
                let date = Calendar.current.date(byAdding: .day,value: date_offset, to: base_date)!
                let date_str = convDateToStr(date: date)
                if unwrapped_item.status.keys.contains(date_str) {
                    switch unwrapped_item.status[date_str]! {
                        case .Done:
                            continuation_count += 1
                            is_continue = true
                        case .Skip:
                            is_continue = true
                        case .NotYet:
                            is_continue = false
                    }
                } else {
                    is_continue = false
                }
                if is_continue == true {
                    date_offset -= 1
                } else {
                    break
                }
            }
        }
        return continuation_count
    }
    func _test_calcContinuationCount() {
        var offset: Int = 0
        print("### calcContinuationCount() test start ###")
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 0
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 3
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 3
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 2
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 0
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 0
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 1
        print("\(offset) : \(String(calcContinuationCount(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!, item_id: self.getItemId(item_name: "bbb"))))"); offset -= 1 // 1
        print("### calcContinuationCount() test finished ###")
        print("")
    }
    private func calcTotalItemStatus(
        date: Date
    ) -> ItemStatus
    {
        let date_str: String = convDateToStr(date: date)
        var is_continue: Bool = true
        var is_skip_all: Bool = true
        for (_, item) in self.items {
            if item.status.keys.contains(date_str) {
                switch item.status[date_str]! {
                    case .Done:
                        is_continue = true
                        is_skip_all = false
                    case .Skip:
                        is_continue = true
                    case .NotYet:
                        is_continue = false
                        is_skip_all = false
                }
            } else {
                is_continue = false
                is_skip_all = false
            }
            if is_continue == false {
                break
            }
        }
        var total_item_status: ItemStatus = .Done
        if is_continue == true {
            if is_skip_all == true {
                total_item_status = .Skip
            } else {
                total_item_status = .Done
            }
        } else {
            total_item_status = .NotYet
        }
        return total_item_status
    }
    private func _test_calcTotalItemStatus() {
        var offset: Int = 0
        print("### calcTotalItemStatus() test start ###")
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("\(offset) : \(convItemStatusToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
        print("### calcTotalItemStatus() test finished ###")
        print("")
    }
    func calcContinuationCountAll(
        base_date: Date
    ) -> Int
    {
        var date_offset: Int = 0
        var continuation_count: Int = 0
        while true {
            var is_continue: Bool = true
            let date = Calendar.current.date(byAdding: .day,value: date_offset, to: base_date)!
            let date_item_status: ItemStatus = calcTotalItemStatus(date: date)
            switch date_item_status {
                case .Done:
                    continuation_count += 1
                    is_continue = true
                case .Skip:
                    is_continue = true
                case .NotYet:
                    is_continue = false
            }
            if is_continue == true {
                date_offset -= 1
            } else {
                break
            }
        }
        return continuation_count
    }
    private func _test_calcContinuationCountAll()
    {
        var offset: Int = 0
        print("### calcContinuationCountAll() test start ###")
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("\(offset) : \(calcContinuationCountAll(base_date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!))"); offset -= 1 // 1
        print("### calcContinuationCountAll() test finished ###")
        print("")
    }
    func convDateToStr(date: Date) -> String {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))
        return dateFormatter.string(from: date)
    }
    func _test_convDateToStr() {
        print("### convDateToStr() test start ###")
        print(convDateToStr(date: Date()))
        print("### convDateToStr() test finished ###")
        print("")
    }
    func convDateToMmdd(date: Date) -> String {
        let formatMd = DateFormatter()
        let formatEee = DateFormatter()
        formatMd.dateFormat = DateFormatter.dateFormat(fromTemplate: "Md", options: 0, locale: Locale(identifier: "ja_JP"))
        formatEee.dateFormat = DateFormatter.dateFormat(fromTemplate: "EEE", options: 0, locale: Locale(identifier: "ja_JP"))
        return formatMd.string(from: date) + "\n" + formatEee.string(from: date)
    }
    mutating func setValueForTest() {
        let item1: Item = Item(
            item_name: "aa",
            status: [
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Skip,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .NotYet
            ],
            skip_num: 10,
            color: Color.blue
        )
        let item3: Item = Item(
            item_name: "cccc",
            status: [
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: 0, to: Date())!)     : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -1, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -6, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -7, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -8, to: Date())!)    : .Done,
                convDateToStr(date: Calendar.current.date(byAdding: .day,value: -9, to: Date())!)    : .Done
            ],
            skip_num: 30,
            color: Color.green
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
