//
//  StatisticsView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/08/26.
//

import SwiftUI

struct StatisticsViewSetting {
    let BUTTON_SIZE_PX :CGFloat? = 10
    let BUTTON_SPACING_PX :CGFloat? = 1
    let BUTTON_NUM_MIN :Int = 3
    let LIST_PADDING_PX :CGFloat? = 5
    let WIDTH_PADDING_PX :CGFloat? = 5
}

struct StatisticsView: View {
    @Binding var hab_chain_data: HabChainData
    private let VIEW_SETTING: StatisticsViewSetting = StatisticsViewSetting()
    var body: some View {
        GeometryReader { geometry in
            VStack {
                let button_num_tmp: Int = Int((geometry.size.width - VIEW_SETTING.LIST_PADDING_PX!*2 - VIEW_SETTING.WIDTH_PADDING_PX!*2) / (VIEW_SETTING.BUTTON_SIZE_PX! + VIEW_SETTING.BUTTON_SPACING_PX!))
                let button_num: Int = button_num_tmp > VIEW_SETTING.BUTTON_NUM_MIN ? button_num_tmp : VIEW_SETTING.BUTTON_NUM_MIN
                let latest_date: Date = Calendar.current.date(byAdding: .day,value: 0, to: Date())!
                let old_date: Date = Calendar.current.date(byAdding: .day,value: -button_num, to: Date())!
                //Text("button num is \(button_num)")
                Text(hab_chain_data.formatDateMmdd(date: old_date) + "〜" + hab_chain_data.formatDateMmdd(date: latest_date))
                    .font(.title)
                List {
                    let item_id_list: [String] = hab_chain_data.getVisibleItemIdList()
                    ForEach(Array(item_id_list.enumerated()), id: \.element) { list_idx, item_id in
                        let background_color: Color = list_idx % 2 == 0 ? Color("list_background1") : Color("list_background2")
                        if let unwraped_item: Item = hab_chain_data.items[item_id] {
                            if unwraped_item.is_archived == false {
                                Group {
                                    HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                                        Text(unwraped_item.item_name)
                                        Spacer()
                                        AchieveRateView(hab_chain_data: $hab_chain_data, item_id: item_id)
                                    }
                                    .padding([.leading, .trailing], VIEW_SETTING.WIDTH_PADDING_PX )
                                    HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                                        Spacer()
                                        ForEach(-(button_num - 1)..<1, id: \.self) { date_offset in
                                            IndivItemStatusTextCircle(hab_chain_data: $hab_chain_data, item_id: item_id, date_offset: date_offset)
                                        }
                                    }
                                    .padding([.leading, .trailing], VIEW_SETTING.WIDTH_PADDING_PX )
                                }
                                .contentShape(Rectangle())
                                .listRowInsets(EdgeInsets())
                                .listRowBackground(background_color)
                            }
                        }
                    }
                }
                .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                .listStyle(.plain)
                .environment(\.editMode, .constant(.active))
            }
        }
    }
}

struct AchieveRateView: View {
    @Binding var hab_chain_data: HabChainData
    let item_id: String
    private let VIEW_SETTING: StatisticsViewSetting = StatisticsViewSetting()

    var body:some View {
        let achive_dic : Dictionary<Int, String> = [
            7: "week",
            30: "month",
            365: "year"
        ]
        ForEach(achive_dic.sorted(by: <), id: \.key) { date_num, word in
            VStack {
                Text("\(hab_chain_data.calcAchievementRate(item_id: item_id, date_num: date_num))％")
                Text(" / " + word)
                    .font(.caption)
                    .foregroundColor(Color.gray)
            }
        }
    }
}

struct IndivItemStatusTextCircle: View {
    @Binding var hab_chain_data: HabChainData
    let item_id: String
    let date_offset: Int
    private let VIEW_SETTING: StatisticsViewSetting = StatisticsViewSetting()

    var body:some View {
        if let unwraped_item: Item = hab_chain_data.items[item_id] {
            let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
            let continuation_cnt: Int = hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
            let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
            Text("")
                .font(.caption)
                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                .multilineTextAlignment(.center)
                .foregroundColor(Color.white)
                .background(Color(color_str))
                .clipShape(Circle())
        }
    }
}

struct StatisticsView_Previews: PreviewProvider {
    @State static var dummy_hab_chain_data: HabChainData = HabChainData()
    static var previews: some View {
        StatisticsView(hab_chain_data: $dummy_hab_chain_data)
    }

}
