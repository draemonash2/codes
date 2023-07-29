//
//  ItemAddView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/27.
//

import SwiftUI

struct ItemAddView: View {
    enum ErrorKind {
        case none
        case blank_item_name
        case exist_item_name
    }
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_add_view: Bool
    @State var new_item_name: String = ""
    @State var new_skip_num: Int = 10
    @State var new_color: Color = Color.red
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    
    var body: some View {
        let _ = Self._printChanges()
        NavigationView {
            Form {
                Section {
                    TextField("e.g. プログラミングの勉強", text: $new_item_name)
                        .autocapitalization(.none)
                } header: {
                    Text("項目名")
                }
                Section {
                    TextField("e.g. 10", value: $new_skip_num, format: .number)
                } header: {
                    Text("スキップ可能数")
                }
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
            .navigationTitle("アイテム追加")
        }
        Button(action: {
            pressAddButtonAction()
        }) {
            Text("Done")
                .frame(maxWidth: .infinity)
                .frame(height: 50)
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
                default:
                    return Alert(title: Text("[内部エラー] 不明なエラー"))
            }
        }
        .padding()
        Button(action: {
            pressCancelButtonAction()
        }) {
            Text("Cancel")
                .frame(maxWidth: .infinity)
                .frame(height: 50)
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
        if new_item_name == "" {
            is_show_alert = true
            error_kind = .blank_item_name
        } else if hab_chain_data.existItemName(item_name: new_item_name) {
            is_show_alert = true
            error_kind = .exist_item_name
        } else {
            let item = Item(
                item_name: new_item_name,
                skip_num: new_skip_num,
                color: new_color
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
