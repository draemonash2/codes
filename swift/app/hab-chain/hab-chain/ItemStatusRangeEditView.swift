//
//  ItemStatusRangeEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/09/23.
//

import SwiftUI

struct ItemStatusRangeEditViewSetting {
    let ICON_SIZE_PX: CGFloat = 25
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let TEXT_EDITER_HEIGHT_PX: CGFloat = 80
    let DATE_PICKER_WIDTH_PX: CGFloat = 120
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
    @Binding var is_show_item_status_range_edit_view: Bool
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
            Form {
                //Text("指定した日付範囲のステータスを更新します")
                #if false
                Section {
                    DatePicker("", selection: $start_date, displayedComponents: [.date])
                        .datePickerStyle(CompactDatePickerStyle())
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
                            //.frame(width: VIEW_SETTING.DATE_PICKER_WIDTH_PX)
                        Text("〜")
                        DatePicker("", selection: $finish_date, displayedComponents: [.date])
                            .datePickerStyle(CompactDatePickerStyle())
                            //.frame(width: VIEW_SETTING.DATE_PICKER_WIDTH_PX)
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
                    .pickerStyle(DefaultPickerStyle())
                } header: {
                    Text("ステータス")
                }
            }
            .navigationTitle("日付範囲一括入力")
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

            is_show_item_status_range_edit_view = false
            is_show_alert = false
        }
    }
}

//struct ItemStatusRangeEditView_Previews: PreviewProvider {
//    static var previews: some View {
//        ItemStatusRangeEditView()
//    }
//}
