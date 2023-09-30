//
//  ItemEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/28.
//

import SwiftUI

struct ItemEditViewSetting {
    let ICON_SIZE_PX: CGFloat = 25
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let TEXT_EDITER_HEIGHT_PX: CGFloat = 80
}

struct ItemEditView: View {
    enum ErrorKind {
        case none
        case blank_item_name
        case invalid_date
        case invalid_weekday_all
    }
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_edit_view: Bool
    @Binding var trgt_item_id: String
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    @State var is_show_select_icon_view: Bool = false
    @State var icon_name: String = ""
    @State private var start_date: Date = Date()
    @State private var finish_date: Date = Date()
    @State private var selected_start_date: Date = Date()
    @State private var selected_finish_date: Date = Date()
    @State private var is_start_date_valid = false
    @State private var is_finish_date_valid = false
    @State private var is_show_item_status_edit_view: Bool = false
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()
    private let VIEW_SETTING: ItemEditViewSetting = ItemEditViewSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        
        NavigationView {
            Form {
                Group {
                    Section {
                        TextField("e.g. プログラミングの勉強", text: Binding($hab_chain_data.items[trgt_item_id])!.item_name)
                            .autocapitalization(.none)
                    } header: {
                        Text("項目名")
                    }
                    Section {
                        Button(action: {
                            is_show_item_status_edit_view = true
                        }) {
                            Text("ステータス編集")
                                .foregroundColor(Color.blue)
                        }
                    } header: {
                        Text("ステータス")
                    }
                    Section {
                        TextField("e.g. スキルアップして転職に成功するため", text: Binding($hab_chain_data.items[trgt_item_id])!.purpose)
                            .autocapitalization(.none)
                        //TextEditor(text: Binding($hab_chain_data.items[trgt_item_id])!.purpose)
                        //    .frame(height: VIEW_SETTING.TEXT_EDITER_HEIGHT_PX)
                    } header: {
                        Text("目的")
                    }
                    Section {
                        Button(action: {
                            is_show_select_icon_view = true
                        }) {
                            if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                if unwrapped_item.icon_name != "" {
                                    Image(systemName: unwrapped_item.icon_name)
                                        .resizable()
                                        .scaledToFit()
                                        .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                                        .foregroundColor(icon_color)
                                } else {
                                    Text("-")
                                }
                            }
                        }
                        .buttonStyle(PlainButtonStyle())
                    } header: {
                        Text("アイコン")
                    }
                    Section {
                        TextField("e.g. 10", value: Binding($hab_chain_data.items[trgt_item_id])!.skip_num, format: .number)
                    } header: {
                        Text("スキップ可能数")
                    }
                    if FUNC_SETTING.color_select_enable == true {
                        Section {
                            Picker("", selection: Binding($hab_chain_data.items[trgt_item_id])!.color) {
                                Text("red").tag(Color.red)
                                Text("green").tag(Color.green)
                                Text("blue").tag(Color.blue)
                            }
                            .pickerStyle(DefaultPickerStyle())
                        } header: {
                            Text("色")
                        }
                    }
                    Section {
                        HStack {
                            Toggle(isOn: Binding($hab_chain_data.items[trgt_item_id])!.is_start_date_enb) {}
                                .labelsHidden()
                            if hab_chain_data.items[trgt_item_id]!.is_start_date_enb {
                                DatePicker("", selection: Binding($hab_chain_data.items[trgt_item_id])!.start_date, displayedComponents: [.date])
                                    .datePickerStyle(CompactDatePickerStyle())
                            }
                            //Text(hab_chain_data.convToStr(date: hab_chain_data.items[trgt_item_id]!.start_date))
                        }
                    } header: {
                        Text("開始日")
                    }
                    Section {
                        HStack {
                            Toggle(isOn: Binding($hab_chain_data.items[trgt_item_id])!.is_finish_date_enb) {}
                                .labelsHidden()
                            if hab_chain_data.items[trgt_item_id]!.is_finish_date_enb {
                                DatePicker("", selection: Binding($hab_chain_data.items[trgt_item_id])!.finish_date, displayedComponents: [.date])
                                    .datePickerStyle(CompactDatePickerStyle())
                            }
                            //Text(hab_chain_data.convToStr(date: hab_chain_data.items[trgt_item_id]!.finish_date))
                        }
                    } header: {
                        Text("終了日")
                    }
                    Section {
                        HStack {
                            Spacer()
                            let weekdays :[String] = ["日", "月", "火", "水", "木", "金", "土"]
                            ForEach(0...6, id: \.self) { weekday_idx in
                                Toggle(isOn: Binding($hab_chain_data.items[trgt_item_id])!.trgt_weekday[weekday_idx]) {
                                    Text(weekdays[weekday_idx])
                                }
                                .toggleStyle(.button)
                            }
                            Spacer()
                        }
                    } header: {
                        Text("曜日")
                    }
                }
                Group {
                    Section {
                        //TextField("", text: Binding($hab_chain_data.items[trgt_item_id])!.note)
                        //    .autocapitalization(.none)
                        TextEditor(text: Binding($hab_chain_data.items[trgt_item_id])!.note)
                            .frame(height: VIEW_SETTING.TEXT_EDITER_HEIGHT_PX)
                    } header: {
                        Text("備考")
                    }
                    Section {
                        Toggle(isOn: Binding($hab_chain_data.items[trgt_item_id])!.is_archived) {
                        }
                    } header: {
                        Text("アーカイブ")
                    }
                }
            }
            .navigationTitle("アイテム編集")
        }
        .navigationViewStyle(StackNavigationViewStyle())
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
                case .blank_item_name:
                    return Alert(title: Text("項目名を入力してください"))
                case .invalid_date:
                    return Alert(title: Text("終了日は開始日以降に設定してください"))
                case .invalid_weekday_all:
                    return Alert(title: Text("全ての曜日が無効化されています"))
                default:
                    return Alert(title: Text("[内部エラー] 不明なエラー"))
            }
        }
        .sheet(isPresented: $is_show_select_icon_view) {
            SelectIconView(
                is_show_select_icon_view: $is_show_select_icon_view,
                icon_name: Binding($hab_chain_data.items[trgt_item_id])!.icon_name
            )
        }
        .sheet(isPresented: $is_show_item_status_edit_view) {
            ItemStatusEditView(
                hab_chain_data: $hab_chain_data,
                is_show_item_status_edit_view: $is_show_item_status_edit_view,
                trgt_item_id: $trgt_item_id
            )
        }
        .padding()
    }
    func pressEditButtonAction() {
        var start_date_str :String = ""
        var finish_date_str :String = ""
        if hab_chain_data.items[trgt_item_id]!.is_start_date_enb && hab_chain_data.items[trgt_item_id]!.is_finish_date_enb {
            start_date_str = hab_chain_data.convToStr(date: hab_chain_data.items[trgt_item_id]!.start_date)
            finish_date_str = hab_chain_data.convToStr(date: hab_chain_data.items[trgt_item_id]!.finish_date)
        } else {
            start_date_str = ""
            finish_date_str = ""
        }
        //print(start_date_str + " - " + finish_date_str)
        
        var is_exist_valid_weekday: Bool = false
        for weekday_idx in 0...6 {
            if hab_chain_data.items[trgt_item_id]!.trgt_weekday[weekday_idx] {
                is_exist_valid_weekday = true
                break
            }
        }
        
        if hab_chain_data.items[trgt_item_id]!.item_name == "" {
            is_show_alert = true
            error_kind = .blank_item_name
        } else if start_date_str > finish_date_str {
            is_show_alert = true
            error_kind = .invalid_date
        } else if is_exist_valid_weekday == false {
            is_show_alert = true
            error_kind = .invalid_weekday_all
        } else {
            is_show_item_edit_view = false
            is_show_alert = false
        }
    }
}
