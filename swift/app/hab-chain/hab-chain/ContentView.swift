//
//  ContentView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI
import WidgetKit

struct ContentViewSetting {
    let BUTTON_NUM :Int = 4
    let BUTTON_SIZE_PX :CGFloat? = 35
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 30
    let ICON_SIZE_PX :CGFloat? = 30
}

struct ContentView: View {
    @Environment(\ .colorScheme) var colorScheme
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @State var hab_chain_data: HabChainData = HabChainData()
    @State var is_overlay_presented: Bool = false
    @State var trgt_status: String = ""
    var add_item: Item?
    private let CVIEW_SETTING: ContentViewSetting = ContentViewSetting()
    var body: some View {
        let _ = Self._printChanges()
        NavigationView {
            ZStack {
                VStack {
                    Text("hab-chain")
                        .font(.largeTitle)
                        .onAppear() {
                            hab_chain_data.setJsonString2RawStruct(json_string: app_json_string)
                            WidgetCenter.shared.reloadAllTimelines()
                            //hab_chain_data.printAll()
                        }
                        .padding()
                    //List {
                        HStack {
                            Spacer()
                            ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                Text(hab_chain_data.convDateToMmdd(date: date, delimiter: "\n"))
                                    .font(.caption)
                                    .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                                    .multilineTextAlignment(.center)
                            }
                        }
                        .padding([.trailing, .leading], 40)
                        HStack {
                            Spacer()
                            ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                WholeItemStatusTextCircle(
                                    hab_chain_data: $hab_chain_data,
                                    date_offset: i
                                )
                            }
                        }
                    //}
                        .padding([.trailing, .leading], 40)
                    List {
                        #if false
                        HStack {
                            Spacer()
                            ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                Text(hab_chain_data.convDateToMmdd(date: date, delimiter: "\n"))
                                    .font(.caption)
                                    .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                                    .multilineTextAlignment(.center)
                            }
                        }
                        HStack {
                            Spacer()
                            ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                WholeItemStatusTextCircle(
                                    hab_chain_data: $hab_chain_data,
                                    date_offset: i
                                )
                            }
                        }
                        #endif
                        ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                            if let unwraped_item: Item = hab_chain_data.items[item_id] {
                                if unwraped_item.is_archived == false {
                                    HStack {
                                        Text(unwraped_item.item_name)
                                        
                                        Spacer()
                                        
                                        ForEach(-(CVIEW_SETTING.BUTTON_NUM - 1)..<1, id: \.self) { i in
                                            IndivItemStatusChangeButton(
                                                hab_chain_data: $hab_chain_data,
                                                is_overlay_presented: $is_overlay_presented,
                                                trgt_status: $trgt_status,
                                                item_id: item_id,
                                                date_offset: i
                                            )
                                        }
                                    }
                                    .contentShape(Rectangle())
                                    .onTapGesture {
                                        print("pressed \(unwraped_item.item_name) item")
                                    }
                                }
                            }
                        }
                    }
                    .padding(0)
                    .environment(\.editMode, .constant(.active))
                }
                .toolbar {
                    ToolbarItem(placement: .navigationBarTrailing) {
                        NavigationLink(
                            destination: ItemSettingView(hab_chain_data: $hab_chain_data)
                        ) {
                            Image(colorScheme == .light ? "pencil_light": "pencil_dark")
                                .resizable()
                                .aspectRatio(contentMode: .fit)
                                .frame(height: CVIEW_SETTING.ICON_SIZE_PX)
                        }
                    }
                    ToolbarItem(placement: .navigationBarLeading) {
                        NavigationLink(
                            destination: AppSettingView(hab_chain_data: $hab_chain_data)
                        ) {
                            Image(colorScheme == .light ? "setting_light": "setting_dark")
                                .resizable()
                                .aspectRatio(contentMode: .fit)
                                .frame(height: CVIEW_SETTING.ICON_SIZE_PX)
                        }
                    }
                }
                if is_overlay_presented {
                    PopupView(is_presented: $is_overlay_presented, trgt_status: $trgt_status)
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
    let CVIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        if let unwraped_item: Item = hab_chain_data.items[item_id] {
            let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
            let date_str: String = hab_chain_data.convDateToStr(date: date)
            let continuation_cnt: Int = hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
            let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
            ZStack {
                Button {
                    print("pressed \(unwraped_item.item_name) \(date_offset) day button")
                    hab_chain_data.toggleItemStatus(item_id: item_id, date: date)
                    // output popup message
                    withAnimation(.easeIn(duration: 0.2)) {
                        trgt_status = hab_chain_data.getItemStatusStr(item_id: item_id, date: date)
                        is_overlay_presented = true
                    }
                    DispatchQueue.main.asyncAfter(deadline: .now() + 1.0) {
                        withAnimation(.easeOut(duration: 0.1)) {
                            is_overlay_presented = false
                        }
                    }
                    app_json_string = hab_chain_data.getRawStruct2JsonString()
                } label: {
                    Text(String(continuation_cnt))
                        .font(.caption)
                        .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
                        .multilineTextAlignment(.center)
                        .foregroundColor(Color.white)
                        .background(Color(color_str))
                        .clipShape(Circle())
                }
                .buttonStyle(PlainButtonStyle())
                
                if let unwrapped_item_status: ItemStatus = unwraped_item.status[date_str] {
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
            }
        }
    }
}

struct WholeItemStatusTextCircle: View {
    @Binding var hab_chain_data: HabChainData
    let date_offset: Int
    let CVIEW_SETTING: ContentViewSetting = ContentViewSetting()

    var body:some View {
        let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
        let continuation_cnt: Int = hab_chain_data.calcContinuationCountAll(base_date: date)
        let color_str: String = getColorString(color: Color.red, continuation_count: continuation_cnt)
        Text(String(continuation_cnt))
            .font(.caption)
            .frame(width: CVIEW_SETTING.BUTTON_SIZE_PX, height: CVIEW_SETTING.BUTTON_SIZE_PX)
            .multilineTextAlignment(.center)
            .foregroundColor(Color.white)
            .background(Color(color_str))
            .clipShape(Circle())
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


