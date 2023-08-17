//
//  ItemEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/28.
//

import SwiftUI

struct ItemEditViewSetting {
    let BUTTON_HEIGHT_PX: CGFloat = 50
}

struct ItemEditView: View {
    enum ErrorKind {
        case none
        case blank_item_name
    }
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_edit_view: Bool
    @Binding var trgt_item_id: String
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()
    private let VIEW_SETTING: ItemEditViewSetting = ItemEditViewSetting()

    var body: some View {
        let _ = Self._printChanges()
        
        NavigationView {
            Form {
                Section {
                    TextField("e.g. プログラミングの勉強", text: Binding($hab_chain_data.items[trgt_item_id])!.item_name)
                        .autocapitalization(.none)
                } header: {
                    Text("項目名")
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
                    } header: {
                        Text("色")
                    }
                }
                Section {
                    Toggle(isOn: Binding($hab_chain_data.items[trgt_item_id])!.is_archived) {
                    }
                } header: {
                    Text("アーカイブ")
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
                default:
                    return Alert(title: Text("[内部エラー] 不明なエラー"))
            }
        }
        .padding()
    }
    func pressEditButtonAction() {
        if hab_chain_data.items[trgt_item_id]!.item_name == "" {
            is_show_alert = true
            error_kind = .blank_item_name
        } else {
            is_show_item_edit_view = false
            is_show_alert = false
        }
    }
}
