//
//  ContentView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI
import WidgetKit

struct ContentViewSetting {
    let BUTTON_SIZE_PX :CGFloat? = 35
    let BUTTON_SPACING_PX :CGFloat? = 8
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 30
    let BUTTON_INCIRCLE_LINEWIDTH :CGFloat? = 1
    let BUTTON_NUM_MIN :Int = 3
    let ITEM_TEXT_WIDTH_PX :CGFloat? = 170
    let NAV_ICON_SIZE_PX :CGFloat? = 30
    let ITEM_ICON_SIZE_PX :CGFloat? = 25
    let LIST_PADDING_PX :CGFloat? = 20
}

struct ContentView: View {
    @Environment(\ .colorScheme) var colorScheme
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @State var hab_chain_data: HabChainData = HabChainData()
    @State var is_overlay_presented: Bool = false
    @State var is_show_item_edit_view: Bool = false
    @State var trgt_status: String = ""
    @State var trgt_item_id: String = ""
    private let VIEW_SETTING: ContentViewSetting = ContentViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()
    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        NavigationView {
            GeometryReader { geometry in
                ZStack {
                    VStack {
                        let button_num_tmp: Int = Int((geometry.size.width - VIEW_SETTING.ITEM_TEXT_WIDTH_PX! - VIEW_SETTING.LIST_PADDING_PX!*2 - VIEW_SETTING.ITEM_ICON_SIZE_PX! - VIEW_SETTING.BUTTON_SPACING_PX!) / (VIEW_SETTING.BUTTON_SIZE_PX! + VIEW_SETTING.BUTTON_SPACING_PX!))
                        let button_num: Int = button_num_tmp > VIEW_SETTING.BUTTON_NUM_MIN ? button_num_tmp : VIEW_SETTING.BUTTON_NUM_MIN
                        Text("hab-chain")
                            .font(.largeTitle)
                            .padding()
                        Text("ハビットチェーンの力で習慣を継続させましょう！")
                            .font(.caption)
                        HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                            Spacer()
                            ForEach(-(button_num - 1)..<1, id: \.self) { i in
                                let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                Text(hab_chain_data.formatDateMmdd(date: date, delimiter: "\n"))
                                    .font(.caption)
                                    .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                    .multilineTextAlignment(.center)
                            }
                        }
                        .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                        HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                            Spacer()
                            ForEach(-(button_num - 1)..<1, id: \.self) { i in
                                WholeItemStatusTextCircle(
                                    hab_chain_data: $hab_chain_data,
                                    date_offset: i
                                )
                            }
                        }
                        .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                        List {
                            ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                                if let unwraped_item: Item = hab_chain_data.items[item_id] {
                                    if unwraped_item.is_archived == false {
                                        HStack (spacing: VIEW_SETTING.BUTTON_SPACING_PX) {
                                            ItemIcon(hab_chain_data: $hab_chain_data, item_id: item_id)
                                                .padding(0)

                                            Text(unwraped_item.item_name)
                                                .frame(width: VIEW_SETTING.ITEM_TEXT_WIDTH_PX, alignment: .leading)
                                                .padding(0)

                                            Spacer(minLength: 1)

                                            ForEach(-(button_num - 1)..<1, id: \.self) { i in
                                                IndivItemStatusChangeButton(
                                                    hab_chain_data: $hab_chain_data,
                                                    is_overlay_presented: $is_overlay_presented,
                                                    trgt_status: $trgt_status,
                                                    item_id: item_id,
                                                    date_offset: i
                                                )
                                                .padding(0)
                                            }
                                        }
                                        .contentShape(Rectangle())
                                        .listRowInsets(EdgeInsets())
                                        .onTapGesture {
                                            print("pressed \(unwraped_item.item_name) item")
                                            trgt_item_id = item_id
                                            is_show_item_edit_view = true
                                        }
                                    }
                                }
                            }
                        }
                        .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                        .listStyle(.plain)
                        .environment(\.editMode, .constant(.active))
                        .sheet(isPresented: $is_show_item_edit_view) {
                            ItemEditView(
                                hab_chain_data: $hab_chain_data,
                                is_show_item_edit_view: $is_show_item_edit_view,
                                trgt_item_id: $trgt_item_id
                            )
                        }
                    }
                    .toolbar {
                        ToolbarItem(placement: .navigationBarLeading) {
                            NavigationLink(
                                destination: AppSettingView(hab_chain_data: $hab_chain_data)
                            ) {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "gearshape")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.NAV_ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                        }
                        ToolbarItem(placement: .navigationBarTrailing) {
                            NavigationLink(
                                destination: StatisticsView(hab_chain_data: $hab_chain_data)
                            ) {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "chart.line.uptrend.xyaxis")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.NAV_ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                        }
                        #if false
                        ToolbarItem(placement: .navigationBarTrailing) {
                            NavigationLink(
                                destination: SelectIconView()
                            ) {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "info.circle")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.NAV_ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                        }
                        #endif
                        ToolbarItem(placement: .navigationBarTrailing) {
                            NavigationLink(
                                destination: InformationView()
                            ) {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "info.circle")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.NAV_ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                        }
                        ToolbarItem(placement: .navigationBarTrailing) {
                            NavigationLink(
                                destination: ItemSettingView(hab_chain_data: $hab_chain_data)
                            ) {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "pencil")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.NAV_ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                        }
                    }
                    if is_overlay_presented {
                        PopupView(is_presented: $is_overlay_presented, trgt_status: $trgt_status)
                    }
                }
            }
        }
        .navigationViewStyle(StackNavigationViewStyle()) // for iPad
    }
}

struct IndivItemStatusChangeButton: View {
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @Binding var hab_chain_data: HabChainData
    @Binding var is_overlay_presented: Bool
    @Binding var trgt_status: String
    let item_id: String
    let date_offset: Int
    let VIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        if let unwraped_item: Item = hab_chain_data.items[item_id] {
            let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
            let date_str: String = hab_chain_data.convToStr(date: date)
            let continuation_cnt: Int = hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
            let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
            ZStack {
                Button {
                    print("pressed \(unwraped_item.item_name) \(date_offset) day button")
                    hab_chain_data.toggleItemStatus(item_id: item_id, date: date)
                    // output popup message
                    if hab_chain_data.is_show_status_popup == true {
                        withAnimation(.easeIn(duration: 0.2)) {
                            if let unwrapped_item = hab_chain_data.items[item_id] {
                                //let date_str: String = hab_chain_data.convToStr(date: date)
                                if let unwrapped_status = unwrapped_item.daily_statuses[date_str] {
                                    trgt_status = hab_chain_data.convToStr(item_status: unwrapped_status)
                                }
                            }
                            is_overlay_presented = true
                        }
                        DispatchQueue.main.asyncAfter(deadline: .now() + 1.0) {
                            withAnimation(.easeOut(duration: 0.1)) {
                                is_overlay_presented = false
                            }
                        }
                    }
                    app_json_string = hab_chain_data.getRawStruct2JsonString()
                    WidgetCenter.shared.reloadAllTimelines()
                } label: {
                    Text(String(continuation_cnt))
                        .font(.caption)
                        .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                        .multilineTextAlignment(.center)
                        .foregroundColor(Color.white)
                        .background(Color(color_str))
                        .clipShape(Circle())
                }
                .buttonStyle(PlainButtonStyle())
                
                if let unwrapped_item_status: ItemStatus = unwraped_item.daily_statuses[date_str] {
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
            }
        }
    }
}

struct WholeItemStatusTextCircle: View {
    @Binding var hab_chain_data: HabChainData
    let date_offset: Int
    let VIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
        let continuation_cnt: Int = hab_chain_data.calcContinuationCountAll(base_date: date)
        let color_str: String = getColorString(color: hab_chain_data.whole_color, continuation_count: continuation_cnt)
        Button {
        } label: {
            Text(String(continuation_cnt))
                .font(.caption)
                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                .multilineTextAlignment(.center)
                .foregroundColor(Color.white)
                .background(Color(color_str))
                .clipShape(Circle())
        }
        .buttonStyle(PlainButtonStyle())
    }
}

struct ItemIcon: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    let item_id: String
    let VIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        if let unwraped_item: Item = hab_chain_data.items[item_id] {
            let icon_color :Color = colorScheme == .light ? Color.black: Color.white
            if unwraped_item.icon_name != "" {
                Image(systemName: unwraped_item.icon_name)
                    .resizable()
                    .scaledToFit()
                    .frame(width: VIEW_SETTING.ITEM_ICON_SIZE_PX, height: VIEW_SETTING.ITEM_ICON_SIZE_PX)
                    .foregroundColor(icon_color)
            } else {
                Image(systemName: sf_symbols[0])
                    .resizable()
                    .scaledToFit()
                    .frame(width: VIEW_SETTING.ITEM_ICON_SIZE_PX, height: VIEW_SETTING.ITEM_ICON_SIZE_PX)
                    .foregroundColor(icon_color)
                    .opacity(0)
            }
        }
    }
}

struct PopupView: View {
    @Binding var is_presented: Bool
    @Binding var trgt_status: String
    
    var body: some View {
        Text(trgt_status)
            .frame(width: 150, height: 50)
            .foregroundColor(.black)
            .background(Color(red: 0.9, green: 0.9, blue: 0.9, opacity: 0.9))
            .cornerRadius(10)
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}


