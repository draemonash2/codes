//
//  ItemAddView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/27.
//

import SwiftUI

struct ItemAddViewSetting {
    let ICON_SIZE_PX: CGFloat = 25
    let BUTTON_HEIGHT_PX: CGFloat = 50
}

struct ItemAddView: View {
    enum ErrorKind {
        case none
        case blank_item_name
        case exist_item_name
        case invalid_date
        case invalid_weekday_all
    }
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_add_view: Bool
    @State var new_item_name: String = ""
    @State var new_skip_num: Int = 10
    @State var new_color: Color = Color.red
    @State var new_is_archived: Bool = false
    @State var new_icon_name: String = ""
    @State var new_start_date: Date = Date()
    @State var new_finish_date: Date = Date()
    @State var new_is_start_date_enb: Bool = false
    @State var new_is_finish_date_enb: Bool = false
    @State var new_weekday: [Bool] = [true, true, true, true, true, true, true]
    @State var new_purpose: String = ""
    @State var new_note: String = ""
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    @State var is_show_select_icon_view: Bool = false
    @State var item_id_child: String = ""
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()
    private let VIEW_SETTING: ItemAddViewSetting = ItemAddViewSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        NavigationView {
            Form {
                Section {
                    TextField("e.g. プログラミングの勉強", text: $new_item_name)
                        .autocapitalization(.none)
                } header: {
                    Text("項目名")
                }
                Section {
                    TextField("e.g. スキルアップして転職に成功するため", text: $new_purpose)
                        .autocapitalization(.none)
                } header: {
                    Text("目的")
                }
                Section {
                    Button(action: {
                        is_show_select_icon_view = true
                    }) {
                        let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                        if new_icon_name != "" {
                            Image(systemName: new_icon_name)
                                .resizable()
                                .scaledToFit()
                                .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                                .foregroundColor(icon_color)
                        } else {
                            Text("-")
                        }
                    }
                    .buttonStyle(PlainButtonStyle())
                } header: {
                    Text("アイコン")
                }
                Section {
                    TextField("e.g. 10", value: $new_skip_num, format: .number)
                } header: {
                    Text("スキップ可能数")
                }
                if FUNC_SETTING.color_select_enable == true {
                    Section {
                        Picker("", selection: $new_color) {
                            Text("red").tag(Color.red)
                            Text("green").tag(Color.green)
                            Text("blue").tag(Color.blue)
                        }
                    } header: {
                        Text("色")
                    }
                }
                Section {
                    HStack {
                        Toggle(isOn: $new_is_start_date_enb) {}
                            .labelsHidden()
                        if new_is_start_date_enb {
                            DatePicker("", selection: $new_start_date, displayedComponents: [.date])
                                .datePickerStyle(CompactDatePickerStyle())
                        }
                    }
                } header: {
                    Text("開始日")
                }
                Section {
                    HStack {
                        Toggle(isOn: $new_is_finish_date_enb) {}
                            .labelsHidden()
                        if new_is_finish_date_enb {
                            DatePicker("", selection: $new_finish_date, displayedComponents: [.date])
                                .datePickerStyle(CompactDatePickerStyle())
                        }
                    }
                } header: {
                    Text("終了日")
                }
                Section {
                    HStack {
                        Spacer()
                        let weekdays :[String] = ["日", "月", "火", "水", "木", "金", "土"]
                        ForEach(0...6, id: \.self) { weekday_idx in
                            Toggle(isOn: $new_weekday[weekday_idx]) {
                                Text(weekdays[weekday_idx])
                            }
                            .toggleStyle(.button)
                        }
                        Spacer()
                    }
                } header: {
                    Text("曜日")
                }
                Section {
                    TextField("", text: $new_note)
                        .autocapitalization(.none)
                } header: {
                    Text("備考")
                }
                Section {
                    Toggle(isOn: $new_is_archived) {
                    }
                } header: {
                    Text("アーカイブ")
                }
            }
            .navigationTitle("アイテム追加")
        }
        .navigationViewStyle(StackNavigationViewStyle())
        Button(action: {
            pressAddButtonAction()
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
                case .exist_item_name:
                    return Alert(title: Text("同じ名前の項目名が存在します"))
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
                icon_name: $new_icon_name
            )
        }
        .padding()
        Button(action: {
            pressCancelButtonAction()
        }) {
            Text("Cancel")
                .frame(maxWidth: .infinity)
                .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .padding()
        .onDisappear() {
            hab_chain_data.printAll()
        }
    }
    func pressAddButtonAction() {
        var new_start_date_str :String = ""
        var new_finish_date_str :String = ""
        if new_is_start_date_enb && new_is_finish_date_enb {
            new_start_date_str = hab_chain_data.convToStr(date: new_start_date)
            new_finish_date_str = hab_chain_data.convToStr(date: new_finish_date)
        } else {
            new_start_date_str = ""
            new_finish_date_str = ""
        }
        //print(new_start_date_str + " - " + new_finish_date_str)
        
        var is_exist_valid_weekday: Bool = false
        for weekday_idx in 0...6 {
            if new_weekday[weekday_idx] {
                is_exist_valid_weekday = true
                break
            }
        }
        
        if new_item_name == "" {
            is_show_alert = true
            error_kind = .blank_item_name
        } else if hab_chain_data.existItemName(item_name: new_item_name) {
            is_show_alert = true
            error_kind = .exist_item_name
        } else if new_start_date_str > new_finish_date_str {
            is_show_alert = true
            error_kind = .invalid_date
        } else if is_exist_valid_weekday == false {
            is_show_alert = true
            error_kind = .invalid_weekday_all
        } else {
            let item = Item(
                item_name: new_item_name,
                skip_num: new_skip_num,
                color: new_color,
                is_archived: new_is_archived,
                icon_name: new_icon_name,
                start_date: new_start_date,
                finish_date: new_finish_date,
                is_start_date_enb: new_is_start_date_enb,
                is_finish_date_enb: new_is_finish_date_enb,
                trgt_weekday: new_weekday,
                purpose: new_purpose,
                note: new_note
            )
            hab_chain_data.addItem(new_item_id: hab_chain_data.generateItemId(), new_item: item)
            is_show_item_add_view = false
            is_show_alert = false
        }
    }
    func pressCancelButtonAction() {
        is_show_item_add_view = false
    }
}
