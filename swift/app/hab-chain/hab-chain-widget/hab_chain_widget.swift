//
//  hab_chain_widget.swift
//  hab-chain-widget
//
//  Created by Tatsuya Endo on 2023/08/04.
//

import WidgetKit
import SwiftUI

struct ContentViewSetting {
    let BUTTON_NUM :Int = 10
    let BUTTON_SIZE_PX :CGFloat? = 15
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 10
    let ICON_SIZE_PX :CGFloat? = 30
    let PADDING_SIZE_PX :CGFloat? = 10
    let LIST_NUM_MAX :Int = 9
}

struct Provider: TimelineProvider {
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    mutating func generateHabChainData() -> HabChainData
    {
        var hab_chain_data: HabChainData = HabChainData()
        let hab_chain_data_jsonstruct = hab_chain_data.convJsonString2JsonStruct(hab_chain_data_jsonstring: app_json_string)
        hab_chain_data.convJsonStruct2RawStruct(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
        return hab_chain_data
    }
    func placeholder(in context: Context) -> SimpleEntry
    {
        var hab_chain_data: HabChainData = HabChainData()
        let hab_chain_data_jsonstruct = hab_chain_data.convJsonString2JsonStruct(hab_chain_data_jsonstring: app_json_string)
        hab_chain_data.convJsonStruct2RawStruct(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
        return SimpleEntry(date: Date(), hab_chain_data: hab_chain_data)
    }

    func getSnapshot(in context: Context, completion: @escaping (SimpleEntry) -> ())
    {
        var hab_chain_data: HabChainData = HabChainData()
        let hab_chain_data_jsonstruct = hab_chain_data.convJsonString2JsonStruct(hab_chain_data_jsonstring: app_json_string)
        hab_chain_data.convJsonStruct2RawStruct(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
        let entry = SimpleEntry(date: Date(), hab_chain_data: hab_chain_data)
        completion(entry)
    }

    func getTimeline(in context: Context, completion: @escaping (Timeline<Entry>) -> ())
    {
        var entries: [SimpleEntry] = []

        let currentDate = Date()
        for hourOffset in 0 ..< 5 {
            let entryDate = Calendar.current.date(byAdding: .hour, value: hourOffset, to: currentDate)!
            var hab_chain_data: HabChainData = HabChainData()
            let hab_chain_data_jsonstruct = hab_chain_data.convJsonString2JsonStruct(hab_chain_data_jsonstring: app_json_string)
            hab_chain_data.convJsonStruct2RawStruct(hab_chain_data_jsonstruct: hab_chain_data_jsonstruct)
            let entry = SimpleEntry(date: entryDate, hab_chain_data: hab_chain_data)
            entries.append(entry)
        }

        let timeline = Timeline(entries: entries, policy: .atEnd)
        completion(timeline)
    }
}

struct SimpleEntry: TimelineEntry {
    let date: Date
    let hab_chain_data: HabChainData
}

struct hab_chain_widgetEntryView : View {
    var entry: Provider.Entry
    private let CVIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body: some View {
        VStack (spacing : 1) {
            #if false
            HStack (spacing : 1) {
                Spacer()
                ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                    let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                    Text(entry.hab_chain_data.convDateToD(date: date))
                        .font(.caption2)
                        .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                        .multilineTextAlignment(.center)
                }
            }
            .padding(.leading, CVIEW_SETTING.PADDING_SIZE_PX)
            .padding(.trailing, CVIEW_SETTING.PADDING_SIZE_PX)
            HStack (spacing : 1) {
                Spacer()
                ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                    let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                    let continuation_cnt: Int = entry.hab_chain_data.calcContinuationCountAll(base_date: date)
                    let color_str: String = getColorString(color: Color.red, continuation_count: continuation_cnt)
                    //Text(String(continuation_cnt))
                    Text("")
                        .font(.caption)
                        .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                        .multilineTextAlignment(.center)
                        .foregroundColor(Color.white)
                        .background(Color(color_str))
                        .clipShape(Circle())
                }
            }
            .padding(.leading, CVIEW_SETTING.PADDING_SIZE_PX)
            .padding(.trailing, CVIEW_SETTING.PADDING_SIZE_PX)
            #endif
            //ForEach(entry.hab_chain_data.item_id_list, id: \.self) { item_id in
            //ForEach(monsters.indexed(), id: \.index) { monsterIndex, monster in
            //ForEach(entry.hab_chain_data.item_id_list.indexed(), id: \.index) { item_id_idx, item_id in
            let item_id_list: [String] = entry.hab_chain_data.getVisibleItemIdList()
            ForEach((0...(CVIEW_SETTING.LIST_NUM_MAX-1)), id: \.self) { i in
                if i < item_id_list.count {
                    let item_id: String = item_id_list[i]
                    if let unwraped_item = entry.hab_chain_data.items[item_id] {
                        if unwraped_item.is_archived == false {
                            HStack (spacing : 1) {
                                Text(unwraped_item.item_name)
                                    .font(.caption)
                                
                                Spacer()
                                
                                ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                    let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                    let date_str: String = entry.hab_chain_data.convDateToStr(date: date)
                                    let continuation_cnt: Int = entry.hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
                                    let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
                                    ZStack {
                                        //Text(String(continuation_cnt))
                                        Text("")
                                            .font(.caption)
                                            .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                                            .multilineTextAlignment(.center)
                                            .foregroundColor(Color.white)
                                            .background(Color(color_str))
                                            .clipShape(Circle())
                                        
                                        #if false
                                        if let unwrapped_item_status = unwraped_item.status[date_str] {
                                            if unwrapped_item_status == .Done {
                                                Circle()
                                                    .stroke(Color.white, lineWidth: 1)
                                                    .frame(width: CVIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: CVIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                                            } else if unwrapped_item_status == .Skip {
                                                Circle()
                                                    .stroke(Color.white, style: StrokeStyle(lineWidth: 1, dash: [4]))
                                                    .frame(width: CVIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: CVIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                                            } else {
                                                // Do Nothing
                                            }
                                        }
                                        #endif
                                    }
                                }
                            }
                            .contentShape(Rectangle())
                            //.padding(.all, CVIEW_SETTING.PADDING_SIZE_PX)
                            .padding(.leading, CVIEW_SETTING.PADDING_SIZE_PX)
                            .padding(.trailing, CVIEW_SETTING.PADDING_SIZE_PX)
                            //.padding(.top, CVIEW_SETTING.PADDING_SIZE_PX)
                            //.padding(.bottom, CVIEW_SETTING.PADDING_SIZE_PX)
                        }
                    }
                }
            }
            //.padding(.top, CVIEW_SETTING.PADDING_SIZE_PX)
            //.padding(.bottom, CVIEW_SETTING.PADDING_SIZE_PX)
            Spacer()
        }
    }
}

@main
struct hab_chain_widget: Widget {
    let kind: String = "hab_chain_widget"

    var body: some WidgetConfiguration {
        StaticConfiguration(kind: kind, provider: Provider()) { entry in
            hab_chain_widgetEntryView(entry: entry)
        }
        .configurationDisplayName("hab-chain Widget")
        .description("This is hab-chain widget.")
        .supportedFamilies([.systemMedium, .systemLarge])
    }
}

#if false
struct hab_chain_widget_Previews: PreviewProvider {
    static var previews: some View {
        hab_chain_widgetEntryView(entry: SimpleEntry(date: Date(), json_string: "test"))
            .previewContext(WidgetPreviewContext(family: .systemSmall))
    }
}
#endif
