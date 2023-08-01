//
//  JsonMngr.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/31.
//

import Foundation

struct ItemStr: Codable {
    var item_name: String = ""
    var status: Dictionary<String, String> = [:]
    var skip_num: Int = 10
    var color: String = "red"
    var is_archived: String = "false"
}

struct HabChainDataStr: Codable {
    var item_id_list: [String] = []
    var items: Dictionary<String, ItemStr> = [:]
}


func readJson() {
    guard let dirURL = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first else {
        fatalError("フォルダURL取得エラー")
    }

    if !FileManager.default.fileExists(atPath: NSHomeDirectory() + "/Documents/" + "hab_chain_data.json"){
        fatalError("JSONが存在しない")
    }

    let fileURL = dirURL.appendingPathComponent("hab_chain_data.json")

    guard let data = try? Data(contentsOf: fileURL) else
    {
        fatalError("JSON読み込みエラー")
    }
         
    let decoder = JSONDecoder()
    guard let hab_chain_data_str = try? decoder.decode(HabChainDataStr.self, from: data) else {
        fatalError("JSONデコードエラー")
    }
    print(hab_chain_data_str.item_id_list[0])
    print(hab_chain_data_str.items["ABC"]!.item_name)
    print(hab_chain_data_str.items["BCD"]!.status["2023/08/01"]!)
}

func writeJson() {
    let hab_chain_data_str: HabChainDataStr =
    HabChainDataStr(
        item_id_list:[
          "ABC",
          "BCD"
        ],
        items: [
            "ABC": ItemStr(
                item_name: "aa",
                status: [
                    "2023/07/25": "NotYet",
                    "2023/07/29": "Done",
                    "2023/08/01": "Skip"
                ],
                skip_num: 999,
                color: "red",
                is_archived: "false"
            ),
            "BCD": ItemStr(
                item_name: "bbb",
                status: [
                    "2023/07/25": "NotYet",
                    "2023/07/29": "Done",
                    "2023/08/01": "Skip"
                ],
                skip_num: 10,
                color: "green",
                is_archived: "false"
            )
        ]
    )
    guard let dirURL = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first else {
        fatalError("フォルダURL取得エラー")
    }

    let fileURL = dirURL.appendingPathComponent("hab_chain_data.json")

    let encoder = JSONEncoder()
    encoder.outputFormatting = .prettyPrinted //JSONデータを整形する
    guard let jsonValue = try? encoder.encode(hab_chain_data_str) else {
        fatalError("JSONエンコードエラー")
    }
    
    do {
        try jsonValue.write(to: fileURL)
    } catch {
        fatalError("JSON書き込みエラー")
    }
}

#if false
//https://softmoco.com/swift/swift-how-to-parse-json-with-dictionary.php
func testJsonDict() {
    let jsonString = """
    {
       "taxRateInfo":[
          {
             "rate":0.0775,
             "jurisdiction":"ANAHEIM",
             "city":"ANAHEIM",
             "county":"ORANGE",
             "tac":"300110370000"
          },
          {
             "rate":0.0776,
             "jurisdiction":"ANAHEIM2",
             "city":"ANAHEIM2",
             "county":"ORANGE2",
             "tac":"300110370002"
          }
       ],
       "geocodeInfo":{
          "bufferDistance":50
       },
       "termsOfUse":"https://www.cdtfa.ca.gov/dataportal/policy.htm",
       "disclaimer":"https://www.cdtfa.ca.gov/dataportal/disclaimer.htm"
    }
    """

    do {
        let jsonDict = try JSONSerialization.jsonObject(with: Data(jsonString.utf8)) as? [String: Any]
        let taxRateInfos = jsonDict?["taxRateInfo"] as? [[String: Any]]
        if taxRateInfos != nil && taxRateInfos!.count > 0 {
            let rate = taxRateInfos![1]["rate"]
            let city = taxRateInfos![1]["city"]
            
            print("City: \(city ?? ""), Tax Rate: \(rate ?? "")")
        }
    } catch {
        print("Unexpected error: \(error).")
    }
}

func testJsonDict2() {
    //let jsonObj = ["Name":"Taro",
    //               "Age": 1,
    //               "dict": ["Name": "Taro",
    //                        "Age": ["Name": "Taro",
    //                                "Age": 1
    //                               ]
    //                       ],
    //               "nulls": nil
    //] as [String : Any?]
    let jsonObj = [
        "Name": ["item_name": "aaaa", "status": ["2017/11/11": ".Skip", "2017/11/12": ".Skip"], "skip_num": 999],
        "Name2": ["item_name": "aaaa", "status": ["2017/11/11": ".Skip", "2017/11/13∫": ".Skip"], "skip_num": 999]
    ] as [String : Any?]

    do {
        let jsonData = try JSONSerialization.data(withJSONObject: jsonObj, options: [])
        let jsonStr = String(bytes: jsonData, encoding: .utf8)!
        print(jsonStr)  // 生成されたJSON文字列 => {"Name":"Taro"}
    } catch let error {
        print(error)
    }
}
#endif
