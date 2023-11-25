//
//  ItemStatusEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/08/12.
//

import SwiftUI

struct ItemStatusEditViewSetting {
    let DATE_NUM :Int = 90
    let BUTTON_SIZE_PX :CGFloat? = 40
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 35
    let ICON_SIZE_PX :CGFloat? = 30
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let STATUS_TEXT_WIDTH_PX: CGFloat = 60
    let TITLE_PADDING_PX: CGFloat = 30
    let ICON_NAME_DIC: Dictionary<ItemStatus, String> = [
        .Done: "checkmark.square",
        .NotYet: "square",
        .Skip: "clear"
    ]
}
struct ItemStatusRangeEditViewSetting {
    let ICON_SIZE_PX: CGFloat = 25
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let TEXT_EDITER_HEIGHT_PX: CGFloat = 80
    let DATE_PICKER_WIDTH_PX: CGFloat = 120
    let TITLE_PADDING_PX: CGFloat = 30
}

struct ItemStatusEditView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_status_edit_view: Bool
    @Binding var trgt_item_id: String
    @State var shouldPresentPopUpDialog: Bool = false
    @State var is_show_item_status_range_edit_view: Bool = false
    private let VIEW_SETTING: ItemStatusEditViewSetting = ItemStatusEditViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        TabView{
            ItemStatusIndivDateEditView(
                hab_chain_data: $hab_chain_data,
                is_show_item_status_edit_view: $is_show_item_status_edit_view,
                trgt_item_id: $trgt_item_id
            )
               .tabItem {
                   Image(systemName: "square.grid.3x3.topleft.filled")
                   Text("個別入力")
               }
            ItemStatusRangeEditView(
                hab_chain_data: $hab_chain_data,
                is_show_item_status_edit_view: $is_show_item_status_edit_view,
                trgt_item_id: $trgt_item_id
            )
               .tabItem {
                   Image(systemName: "square.grid.3x3.fill")
                   Text("範囲入力")
               }
        }
        //.tabViewStyle(PageTabViewStyle(indexDisplayMode: .never))
    }
}

struct ItemStatusIndivDateEditView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_status_edit_view: Bool
    @Binding var trgt_item_id: String
    private let VIEW_SETTING: ItemStatusEditViewSetting = ItemStatusEditViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        NavigationView {
            VStack {
                if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                    Text(unwrapped_item.item_name)
                        .font(.title)
                        .frame(maxWidth: .infinity, alignment: .leading)
                        .padding([.leading, .trailing], VIEW_SETTING.TITLE_PADDING_PX)
                    List {
                        ForEach(0..<VIEW_SETTING.DATE_NUM, id: \.self) { i in
                            let date_offset: Int = -i
                            let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
                            HStack {
                                Text(hab_chain_data.formatDateYyyyMmdd(date: date))
                                Spacer()
                                Button(action:{
                                    hab_chain_data.toggleItemStatus(item_id: trgt_item_id, date: date)
                                }) {
                                    let date_str = hab_chain_data.convToStr(date: date)
                                    let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                    if unwrapped_item.daily_statuses.keys.contains(date_str) {
                                        if let unwrapped_item_status = unwrapped_item.daily_statuses[date_str] {
                                            Image(systemName: VIEW_SETTING.ICON_NAME_DIC[unwrapped_item_status]!)
                                                .resizable()
                                                .aspectRatio(contentMode: .fit)
                                                .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                                .foregroundColor(icon_color)
                                        }
                                    } else {
                                        Image(systemName: VIEW_SETTING.ICON_NAME_DIC[.NotYet]!)
                                            .resizable()
                                            .aspectRatio(contentMode: .fit)
                                            .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                            .foregroundColor(icon_color)
                                    }
                                }
                                let date_str = hab_chain_data.convToStr(date: date)
                                if unwrapped_item.daily_statuses.keys.contains(date_str) {
                                    if let unwrapped_item_status = unwrapped_item.daily_statuses[date_str] {
                                        Text(hab_chain_data.convToStr(item_status: unwrapped_item_status))
                                            .frame(width: VIEW_SETTING.STATUS_TEXT_WIDTH_PX, alignment: .leading)
                                    }
                                } else {
                                    Text(hab_chain_data.convToStr(item_status: .NotYet))
                                        .frame(width: VIEW_SETTING.STATUS_TEXT_WIDTH_PX, alignment: .leading)
                                }
                            }
                        }
                    }
                    //.listStyle(.plain)
                }
                Button(action: {
                    is_show_item_status_edit_view = false
                }) {
                    Text("Done")
                        .frame(maxWidth: .infinity)
                        .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                        .multilineTextAlignment(.center)
                        .background(Color.blue)
                        .foregroundColor(Color.white)
                }
                .padding()
            }
            .navigationTitle("ステータス編集")
        }
        .navigationViewStyle(StackNavigationViewStyle())
    }
}

struct ItemStatusRangeEditView: View {
    enum ErrorKind {
        case none
        case start_later_than_finish_date
        case start_later_than_current_date
        case finish_later_than_current_date
    }
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_status_edit_view: Bool
    @Binding var trgt_item_id: String
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    @State private var start_date: Date = Date()
    @State private var finish_date: Date = Date()
    @State private var item_status: ItemStatus = .Done
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()
    private let VIEW_SETTING: ItemStatusRangeEditViewSetting = ItemStatusRangeEditViewSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        NavigationView {
            VStack {
                if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                    Text(unwrapped_item.item_name)
                        .font(.title)
                        .frame(maxWidth: .infinity, alignment: .leading)
                        .padding([.leading, .trailing], VIEW_SETTING.TITLE_PADDING_PX)
                    List {
                        #if false
                        Section {
                            DatePicker("", selection: $start_date, displayedComponents: [.date])
                                .datePickerStyle(CompactDatePickerStyle())
                                .frame(height:)
                        } header: {
                            Text("開始日")
                        }
                        Section {
                            DatePicker("", selection: $finish_date, displayedComponents: [.date])
                                .datePickerStyle(CompactDatePickerStyle())
                        } header: {
                            Text("終了日")
                        }
                        #else
                        Section {
                            HStack {
                                DatePicker("", selection: $start_date, displayedComponents: [.date])
                                    .datePickerStyle(CompactDatePickerStyle())
                                Text("〜")
                                DatePicker("", selection: $finish_date, displayedComponents: [.date])
                                    .datePickerStyle(CompactDatePickerStyle())
                                Spacer()
                            }
                        } header: {
                            Text("日付範囲")
                        }
                        #endif
                        Section {
                            Picker("", selection: $item_status) {
                                Text("Done").tag(ItemStatus.Done)
                                Text("Skip").tag(ItemStatus.Skip)
                                Text("NotYet").tag(ItemStatus.NotYet)
                            }
                            //.pickerStyle(DefaultPickerStyle())
                        } header: {
                            Text("ステータス")
                        }
                    }
                    .listStyle(InsetGroupedListStyle())
                }
                Button(action: {
                    pressEditButtonAction()
                }) {
                    Text("Done")
                        .frame(maxWidth: .infinity)
                        .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                        .multilineTextAlignment(.center)
                        .background(Color.blue)
                        .foregroundColor(Color.white)
                }
                .alert(isPresented: $is_show_alert) {
                    switch error_kind {
                        case .start_later_than_finish_date:
                            return Alert(title: Text("終了日は開始日以降に設定してください"))
                        case .start_later_than_current_date:
                            return Alert(title: Text("開始日には現在か過去の日付を設定してください"))
                        case .finish_later_than_current_date:
                            return Alert(title: Text("終了日には現在か過去の日付を設定してください"))
                        default:
                            return Alert(title: Text("[内部エラー] 不明なエラー"))
                    }
                }
                .padding()
            }
            .navigationTitle("ステータス編集")
        }
        .navigationViewStyle(StackNavigationViewStyle())
    }
    func pressEditButtonAction() {
        let start_date_str :String = hab_chain_data.convToStr(date: start_date)
        let finish_date_str :String = hab_chain_data.convToStr(date: finish_date)
        let current_date_str :String = hab_chain_data.convToStr(date: Date())
        //print(start_date_str + " - " + finish_date_str)
        
        if start_date_str > finish_date_str {
            is_show_alert = true
            error_kind = .start_later_than_finish_date
        } else if current_date_str < start_date_str {
            is_show_alert = true
            error_kind = .start_later_than_current_date
        } else if current_date_str < finish_date_str {
            is_show_alert = true
            error_kind = .finish_later_than_current_date
        } else {
            //let date_num: Int = Int(finish_date.timeIntervalSince(start_date))
            let elapsed_days :Int = Calendar.current.dateComponents([.day], from: start_date, to: finish_date).day! + 1
            for date_offset in 0..<elapsed_days {
                let date: Date = Calendar.current.date(byAdding: .day,value: -date_offset, to: Date())!
                let date_str: String = hab_chain_data.convToStr(date: date)
                //print(date_str)
                hab_chain_data.items[trgt_item_id]!.daily_statuses.updateValue(item_status, forKey: date_str)
            }

            is_show_item_status_edit_view = false
            is_show_alert = false
        }
    }
}
