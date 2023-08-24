//
//  DataMngr.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.aaa
//

import Foundation
import SwiftUI

enum ItemStatus: String {
    case NotYet
    case Done
    case Skip
}

struct Item {
    var item_name: String = ""
    var daily_statuses: Dictionary<String, ItemStatus> = [:]
    var skip_num: Int = 999
    var color: Color = Color.red
    var is_archived: Bool = false
}

struct HabChainData {
    var item_id_list: [String] = []
    var items: Dictionary<String, Item> = [:]
    var whole_color: Color = Color.green
    var is_show_status_popup: Bool = true

    /* for Json <TOP> */
    struct ItemJson: Codable {
        var item_name: String = ""
        var daily_statuses: Dictionary<String, String> = [:]
        var skip_num: Int = 10
        var color: String = "red"
        var is_archived: String = "false"
    }

    struct HabChainDataJson: Codable {
        var item_id_list: [String] = []
        var items: Dictionary<String, ItemJson> = [:]
        var whole_color: String = "red"
        var is_show_status_popup: String = "true"
    }
    /* for Json <TOP> */

    init() {
        //self.setValueForTest()
        //self.printAll()
        //self._test_calcContinuationCount()
        //self._test_convToStr()
        //self._test_toggleItemStatus()
        //self._test_calcTotalItemStatus()
        //self._test_calcContinuationCountAll()
        //writeJson()
        //readJson()
        //testJsonDict()
        //testJsonDict2()
        //self._test_formatDateD()
    }
    mutating func clear()
    {
        self.item_id_list.removeAll()
        self.items.removeAll()
        self.is_show_status_popup = true
        self.whole_color = Color.red
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
    func existItemName(
        item_name: String
    ) -> Bool
    {
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
    func convToStr(
        date: Date
    ) -> String
    {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = DateFormatter.dateFormat(fromTemplate: "yyyyMMdd", options: 0, locale: Locale(identifier: "ja_JP"))
        return dateFormatter.string(from: date)
    }
        private func _test_convToStr()
        {
            print("### convToStr() test start ###")
            print(convToStr(date: Date()))
            print("### convToStr() test finished ###")
            print("")
        }
    func convToStr(
        item_status: ItemStatus
    ) -> String
    {
        var ret: String = ""
        ret = item_status.rawValue
        return ret
    }
    func convToStr(
        color: Color
    ) -> String
    {
        var ret: String = ""
        switch color {
            case Color.red: ret = "red"
            case Color.green: ret = "green"
            case Color.blue: ret = "blue"
            default: fatalError("[error] convToStr()へ不明なcolorが指定されました。")
        }
        return ret
    }
    func convToStr(
        variable: Bool
    ) -> String
    {
        var ret: String = ""
        ret = variable.description
        return ret
    }
    func convFromStr(
        item_status: String
    ) -> ItemStatus
    {
        var ret: ItemStatus
        switch item_status {
            case "NotYet": ret = .NotYet
            case "Done": ret = .Done
            case "Skip": ret = .Skip
            default: fatalError("[error] convFromStr()へ不明なitem_statusが指定されました。")
        }
        return ret
    }
    func convFromStr(
        color: String
    ) -> Color
    {
        var ret: Color
        switch color {
            case "red": ret = Color.red
            case "green": ret = Color.green
            case "blue": ret = Color.blue
            default: fatalError("[error] convFromStr()へ不明なcolorが指定されました。")
        }
        return ret
    }
    func convFromStr(
        variable: String
    ) -> Bool
    {
        var ret: Bool
        switch variable {
            case "true": ret = true
            case "false": ret = false
            default: fatalError("[error] convFromStr()へ不明なvariableが指定されました。")
        }
        return ret
    }
    mutating func toggleItemStatus(
        item_id: String,
        date: Date
    )
    {
        let date_str = convToStr(date: date)
        var new_itemstatus: ItemStatus = .Done
        if let unwrapped_item = self.items[item_id] {
            if let unwrapped_itemstatus = unwrapped_item.daily_statuses[date_str] {
                switch unwrapped_itemstatus {
                    case .NotYet: new_itemstatus = .Done
                    case .Done:   new_itemstatus = .Skip
                    case .Skip:   new_itemstatus = .NotYet
                }
            }
        }
        self.items[item_id]!.daily_statuses.updateValue(new_itemstatus, forKey: date_str)
    }
        mutating func _test_toggleItemStatus()
        {
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
    func getVisibleItemIdList() -> [String]
    {
        var ret_item_id_list: [String] = []
        for (_, item_id) in self.item_id_list.enumerated() {
            if let unwraped_item: Item = self.items[item_id] {
                if unwraped_item.is_archived == false {
                    ret_item_id_list.append(item_id)
                }
            }
        }
        return ret_item_id_list
    }
    func calcContinuationCount(
        base_date: Date,
        item_id: String
    ) -> Int
    {
        let CNT_MAX: Int = 999
        var date_offset: Int = 0
        var continuation_count: Int = 0
        if let unwrapped_item = self.items[item_id] {
            if unwrapped_item.is_archived == false {
                while true {
                    var is_continue: Bool = true
                    let date = Calendar.current.date(byAdding: .day,value: date_offset, to: base_date)!
                    let date_str = convToStr(date: date)
                    if unwrapped_item.daily_statuses.keys.contains(date_str) {
                        switch unwrapped_item.daily_statuses[date_str]! {
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
        }
        if continuation_count >= CNT_MAX {
            continuation_count = CNT_MAX
        }
        return continuation_count
    }
        private func _test_calcContinuationCount()
        {
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
        var total_item_status: ItemStatus = .Done
        if self.items.count > 0 {
            let date_str: String = convToStr(date: date)
            var is_continue: Bool = true
            var is_skip_all: Bool = true
            for (_, item) in self.items {
                if item.is_archived == false {
                    if item.daily_statuses.keys.contains(date_str) {
                        switch item.daily_statuses[date_str]! {
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
            }
            if is_continue == true {
                if is_skip_all == true {
                    total_item_status = .Skip
                } else {
                    total_item_status = .Done
                }
            } else {
                total_item_status = .NotYet
            }
        } else {
            total_item_status = .NotYet
        }
        return total_item_status
    }
        private func _test_calcTotalItemStatus()
        {
            var offset: Int = 0
            print("### calcTotalItemStatus() test start ###")
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
            print("\(offset) : \(convToStr(item_status: calcTotalItemStatus(date: Calendar.current.date(byAdding: .day,value: offset, to: Date())!)))"); offset -= 1 // 1
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
    func formatDateMmdd(
        date: Date,
        delimiter: String = ""
    ) -> String
    {
        let formatMd = DateFormatter()
        let formatEee = DateFormatter()
        formatMd.dateFormat = DateFormatter.dateFormat(fromTemplate: "Md", options: 0, locale: Locale(identifier: "ja_JP"))
        formatEee.dateFormat = DateFormatter.dateFormat(fromTemplate: "EEE", options: 0, locale: Locale(identifier: "ja_JP"))
        return formatMd.string(from: date) + delimiter + formatEee.string(from: date)
    }
    func formatDateD(
        date: Date
    ) -> String
    {
        let formatD = DateFormatter()
        formatD.dateFormat = DateFormatter.dateFormat(fromTemplate: "d", options: 0, locale: Locale(identifier: "en_US"))
        return formatD.string(from: date)
    }
        private func _test_formatDateD()
        {
            print("### formatDateD() test start ###")
            print(formatDateD(date: Date()))
            print("### formatDateD() test finished ###")
            print("")
        }
    /* for Json <TOP> */
    func convRawStruct2JsonStruct() -> HabChainDataJson
    {
        var hab_chain_data_json: HabChainDataJson = HabChainDataJson()
        for (_, item_id_self) in self.item_id_list.enumerated() {
            hab_chain_data_json.item_id_list.append(item_id_self)
        }
        for (item_id_self, item_self) in self.items {
            var item_json: ItemJson = ItemJson()
            item_json.item_name = item_self.item_name
            item_json.skip_num = item_self.skip_num
            item_json.color = convToStr(color: item_self.color)
            item_json.is_archived = convToStr(variable: item_self.is_archived)
            for (date_self, status_self) in item_self.daily_statuses {
                item_json.daily_statuses.updateValue(convToStr(item_status: status_self), forKey: date_self)
            }
            hab_chain_data_json.items.updateValue(item_json, forKey: item_id_self)
        }
        hab_chain_data_json.whole_color = self.convToStr(color: self.whole_color)
        hab_chain_data_json.is_show_status_popup = self.convToStr(variable: self.is_show_status_popup)
        return hab_chain_data_json
    }
    mutating func convJsonStruct2RawStruct(
        hab_chain_data_jsonstruct: HabChainDataJson
    )
    {
        self.clear()
        
        for (_, item_id_json) in hab_chain_data_jsonstruct.item_id_list.enumerated() {
            self.item_id_list.append(item_id_json)
        }
        for (item_id_json, item_json) in hab_chain_data_jsonstruct.items {
            var item_self: Item = Item()
            item_self.item_name = item_json.item_name
            item_self.skip_num = item_json.skip_num
            item_self.color = convFromStr(color: item_json.color)
            item_self.is_archived = convFromStr(variable: item_json.is_archived)
            for (date_json, status_json) in item_json.daily_statuses {
                item_self.daily_statuses.updateValue(convFromStr(item_status: status_json), forKey: date_json)
            }
            self.items.updateValue(item_self, forKey: item_id_json)
        }
        self.whole_color = self.convFromStr(color: hab_chain_data_jsonstruct.whole_color)
        self.is_show_status_popup = self.convFromStr(variable: hab_chain_data_jsonstruct.is_show_status_popup)
    }
    mutating func convJsonStruct2JsonString(
        hab_chain_data_jsonstruct: HabChainDataJson
    ) -> String
    {
        var ret_json_string: String = ""
        let encoder = JSONEncoder()
        encoder.outputFormatting = .prettyPrinted //JSONデータを整形する
        guard let json_value = try? encoder.encode(hab_chain_data_jsonstruct) else {
            //fatalError("JSONエンコードエラー")
            return ret_json_string
        }
        
        let json_string = String(data: json_value, encoding: .utf8)
        if let unwrapped_jsonstring = json_string {
            print(unwrapped_jsonstring)
            ret_json_string = unwrapped_jsonstring
        }
        return ret_json_string
    }
    mutating func convJsonString2JsonStruct(
        hab_chain_data_jsonstring: String
    ) -> HabChainDataJson
    {
        if hab_chain_data_jsonstring == "" {
            return HabChainDataJson()
        }
        let json_data = hab_chain_data_jsonstring.data(using: .utf8)!
        let decoder = JSONDecoder()
        guard let hab_chain_data_jsonstruct = try? decoder.decode(HabChainDataJson.self, from: json_data) else {
            print("JSONデコードエラー")
            return HabChainDataJson()
        }
        return hab_chain_data_jsonstruct
    }
    mutating func getRawStruct2JsonString() -> String
    {
        let hab_chain_data_jsonstruct: HabChainDataJson = self.convRawStruct2JsonStruct()
        return self.convJsonStruct2JsonString(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
    }
    mutating func setJsonString2RawStruct(
        json_string: String
    )
    {
        let hab_chain_data_jsonstruct = self.convJsonString2JsonStruct(hab_chain_data_jsonstring: json_string)
        self.convJsonStruct2RawStruct(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
    }
    mutating func saveJsonString()
    {
        let hab_chain_data_jsonstring: String = self.getRawStruct2JsonString()
        
        guard let dirURL = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first else {
            fatalError("フォルダURL取得エラー")
        }
        let fileURL = dirURL.appendingPathComponent("hab_chain_data.json")
        print("saveJsonString() URL = " + fileURL.path)

        do {
            try hab_chain_data_jsonstring.write(toFile: fileURL.path, atomically: true, encoding: .utf8)
        } catch {
            fatalError("JSON書き込みエラー")
        }
    }
    mutating func loadJsonString()
    {
        guard let dirURL = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first else {
            fatalError("フォルダURL取得エラー")
        }
        let fileURL = dirURL.appendingPathComponent("hab_chain_data.json")
        print("loadJsonString() URL = " + fileURL.path)

        if !FileManager.default.fileExists(atPath: NSHomeDirectory() + "/Documents/" + "hab_chain_data.json"){
            fatalError("JSONが存在しない")
        }
        guard let json_string = try? String(contentsOf: fileURL, encoding: .utf8) else {
            fatalError("JSON読み込みエラー")
        }
        //print(json_string)
        
        self.setJsonString2RawStruct(json_string: json_string)
    }
    /* for Json <END> */
    
    mutating func setValueForTest()
    {
        #if false
        let item1: Item = Item(
            item_name: "aa",
            status: [
                convToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Skip,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .NotYet
            ],
            skip_num: 10,
            color: Color.blue
        )
        let item3: Item = Item(
            item_name: "cccc",
            status: [
                convToStr(date: Calendar.current.date(byAdding: .day,value: 0, to: Date())!)     : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -1, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -6, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -7, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -8, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -9, to: Date())!)    : .Done
            ],
            skip_num: 30,
            color: Color.green
        )
        let item2: Item = Item(
            item_name: "bbb",
            status: [
                convToStr(date: Calendar.current.date(byAdding: .day,value: 0, to: Date())!)     : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -1, to: Date())!)    : .NotYet,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -2, to: Date())!)    : .Skip,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -3, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -4, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -5, to: Date())!)    : .Done,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -7, to: Date())!)    : .NotYet,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -8, to: Date())!)    : .Skip,
                convToStr(date: Calendar.current.date(byAdding: .day,value: -9, to: Date())!)    : .Done
            ],
            skip_num: 20
        )

        self.addItem(new_item_id: generateItemId(), new_item: item2)
        self.addItem(new_item_id: generateItemId(), new_item: item3)
        self.addItem(new_item_id: generateItemId(), new_item: item1)
        #else
        let item1: Item = Item(item_name: "aa")
        self.addItem(new_item_id: generateItemId(), new_item: item1)
        #endif
    }
    func printAll()
    {
        print("### item_id_list ###")
        for (cur_index, item_id) in self.item_id_list.enumerated() {
            print("\(cur_index) : \(item_id) : \(self.items[item_id]!.item_name)")
        }
        print("### items ###")
        for (key,value) in self.items {
            print("\(key) : \(value.item_name)")
            for (key,value) in value.daily_statuses.sorted(by: { $0.key > $1.key }) {
                print("\(key) : \(value)")
            }
        }
        print("")
    }
}



