//
//  hab_chain_widget.swift
//  hab-chain-widget
//
//  Created by Tatsuya Endo on 2023/08/04.
//

import WidgetKit
import SwiftUI

struct ContentViewSetting {
    let BUTTON_NUM_MIN :Int = 3
    let BUTTON_SIZE_PX :CGFloat? = 15
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 10
    let BUTTON_INCIRCLE_LINEWIDTH :CGFloat? = 1
    let BUTTON_SPACING_PX: CGFloat? = 1
    let PADDING_SIZE_PX :CGFloat? = 2
    let LIST_SPACING_PX :CGFloat? = 1
    let ITEM_ICON_SIZE_PX :CGFloat? = 15
    let ITEM_TEXT_WIDTH_PX: CGFloat? = 100
    let LIST_TOP_SPACING_PX :CGFloat? = 10
}

struct Provider: TimelineProvider {
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    mutating func generateHabChainData() -> HabChainData
    {
        var hab_chain_data: HabChainData = HabChainData()
        hab_chain_data.setJsonString2RawStruct(json_string: app_json_string)
        return hab_chain_data
    }
    func placeholder(in context: Context) -> SimpleEntry
    {
        var hab_chain_data: HabChainData = HabChainData()
        hab_chain_data.setJsonString2RawStruct(json_string: app_json_string)
        return SimpleEntry(date: Date(), hab_chain_data: hab_chain_data)
    }

    func getSnapshot(in context: Context, completion: @escaping (SimpleEntry) -> ())
    {
        var hab_chain_data: HabChainData = HabChainData()
        hab_chain_data.setJsonString2RawStruct(json_string: app_json_string)
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
            hab_chain_data.setJsonString2RawStruct(json_string: app_json_string)
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
    @Environment(\ .colorScheme) var colorScheme
    private let VIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body: some View {
        GeometryReader { geometry in
            VStack (spacing : VIEW_SETTING.LIST_SPACING_PX) {
                let button_num_tmp: Int = Int((geometry.size.width - VIEW_SETTING.ITEM_TEXT_WIDTH_PX! - VIEW_SETTING.PADDING_SIZE_PX!*2 - VIEW_SETTING.ITEM_ICON_SIZE_PX! - VIEW_SETTING.BUTTON_SPACING_PX! - VIEW_SETTING.LIST_TOP_SPACING_PX!*2) / (VIEW_SETTING.BUTTON_SIZE_PX! + VIEW_SETTING.BUTTON_SPACING_PX!))
                let button_num: Int = button_num_tmp > VIEW_SETTING.BUTTON_NUM_MIN ? button_num_tmp : VIEW_SETTING.BUTTON_NUM_MIN
                let list_num_max_tmp: Int = Int((geometry.size.height - VIEW_SETTING.LIST_SPACING_PX!*2 - VIEW_SETTING.LIST_TOP_SPACING_PX!*2) / (VIEW_SETTING.BUTTON_SIZE_PX! + VIEW_SETTING.LIST_SPACING_PX! + VIEW_SETTING.BUTTON_SPACING_PX!))
                let list_num_max: Int = list_num_max_tmp > 1 ? list_num_max_tmp : 1
                #if false
                Group {
                    HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                        Spacer()
                        ForEach(-(VIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                            let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                            Text(entry.hab_chain_data.formatDateD(date: date))
                                .font(.caption2)
                                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                .multilineTextAlignment(.center)
                        }
                    }
                    HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                        Spacer()
                        ForEach(-(VIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                            let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                            let continuation_cnt: Int = entry.hab_chain_data.calcContinuationCountAll(base_date: date)
                            let color_str: String = getColorString(color: Color.red, continuation_count: continuation_cnt)
                            //Text(String(continuation_cnt))
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
                .padding([.leading, .trailing], VIEW_SETTING.PADDING_SIZE_PX)
                #endif
                let item_id_list: [String] = entry.hab_chain_data.getVisibleItemIdList()
                ForEach((0...(list_num_max-1)), id: \.self) { list_idx in
                    if list_idx < item_id_list.count {
                        let item_id: String = item_id_list[list_idx]
                        let background_color: Color = list_idx % 2 == 1 ? Color("widget_list_background1") : Color("widget_list_background2")
                        if let unwraped_item = entry.hab_chain_data.items[item_id] {
                            if unwraped_item.is_archived == false {
                                HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                                    let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                    if unwraped_item.icon_name != "" {
                                        Image(systemName: unwraped_item.icon_name)
                                            .resizable()
                                            .scaledToFit()
                                            .frame(width: VIEW_SETTING.ITEM_ICON_SIZE_PX, height: VIEW_SETTING.ITEM_ICON_SIZE_PX)
                                            .foregroundColor(icon_color)
                                    } else {
                                        Image(systemName: unwraped_item.icon_name)
                                            .resizable()
                                            .scaledToFit()
                                            .frame(width: VIEW_SETTING.ITEM_ICON_SIZE_PX, height: VIEW_SETTING.ITEM_ICON_SIZE_PX)
                                            .foregroundColor(icon_color)
                                            .opacity(0)
                                    }

                                    Text(" ")
                                        .font(.caption)
                                        .padding(0)
                                    Text(unwraped_item.item_name)
                                        .font(.caption)
                                        .fixedSize(horizontal: false, vertical: true)
                                        .frame(maxWidth: VIEW_SETTING.ITEM_TEXT_WIDTH_PX, alignment: .leading)
                                        .lineLimit(1)
                                        .edgesIgnoringSafeArea(.all)
                                    
                                    
                                    Spacer(minLength: 1)
                                    
                                    ForEach(-(button_num - 1)..<1, id: \.self) { btn_idx in
                                        let date: Date = Calendar.current.date(byAdding: .day,value: btn_idx, to: Date())!
                                        let date_str: String = entry.hab_chain_data.convToStr(date: date)
                                        let continuation_cnt: Int = entry.hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
                                        let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
                                        ZStack {
                                            Text("")
                                                .font(.caption)
                                                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                                .multilineTextAlignment(.center)
                                                .foregroundColor(Color.white)
                                                .background(Color(color_str))
                                                .clipShape(Circle())
                                            
                                            #if false
                                            if let unwrapped_item_status = unwraped_item.status[date_str] {
                                                if unwrapped_item_status == .Done {
                                                    Circle()
                                                        .stroke(Color.white, lineWidth: VIEW_SETTING.BUTTON_INCIRCLE_LINEWIDTH!)
                                                        .frame(width: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                                                } else if unwrapped_item_status == .Skip {
                                                    Circle()
                                                        .stroke(Color.white, style: StrokeStyle(lineWidth: VIEW_SETTING.BUTTON_INCIRCLE_LINEWIDTH!, dash: [4]))
                                                        .frame(width: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                                                } else {
                                                    // Do Nothing
                                                }
                                            }
                                            #endif
                                        }
                                    }
                                }
                                .contentShape(Rectangle())
                                .background(background_color)
                                .padding(.leading, VIEW_SETTING.PADDING_SIZE_PX)
                                .padding(.trailing, VIEW_SETTING.PADDING_SIZE_PX)
                            }
                        }
                    }
                }
                Spacer()
            }
            .padding(.top, VIEW_SETTING.LIST_TOP_SPACING_PX)
            .padding(.bottom, VIEW_SETTING.LIST_TOP_SPACING_PX)
            .widgetBackground(backgroundView: Color.clear)
        }
    }
}

struct ItemIcon: View {
    @Environment(\ .colorScheme) var colorScheme
    let icon_name: String = ""
    let VIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        VStack {
            let icon_color :Color = colorScheme == .light ? Color.black: Color.white
            if icon_name != "" {
                Image(systemName: icon_name)
                    .resizable()
                    .scaledToFit()
                    .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                    .foregroundColor(icon_color)
            } else {
                Image(systemName: icon_name)
                    .resizable()
                    .scaledToFit()
                    .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                    .foregroundColor(icon_color)
                    .opacity(0)
            }
        }
        .widgetBackground(backgroundView: Color.clear)
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
        .contentMarginsDisabled()
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

extension View {
    func widgetBackground(backgroundView: some View) -> some View {
        if #available(watchOS 10.0, iOSApplicationExtension 17.0, iOS 17.0, macOSApplicationExtension 14.0, *) {
            return containerBackground(for: .widget) {
                backgroundView
            }
        } else {
            return background(backgroundView)
        }
    }
}
