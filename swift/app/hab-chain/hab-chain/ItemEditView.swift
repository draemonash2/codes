//
//  ItemEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/28.
//

import SwiftUI

struct ItemEditView: View {
    enum ErrorKind {
        case none
        case blank_item_name
    }
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_edit_view: Bool
    @Binding var trgt_item_id: String
    @FocusState private var focusState
    @State var new_item_name: String = "abc"
    @State var new_skip_num: Int = 10
    @State var new_color: Color = Color.red
    @State private var is_show_alert: Bool = false
    @State private var error_kind: ErrorKind = .none
    
    var body: some View {
        let BUTTON_HEIGHT_PX: CGFloat = 50
        let _ = Self._printChanges()
        
        NavigationView {
            Form {
                Section {
                    TextField("e.g. プログラミングの勉強", text: $new_item_name)
                        .autocapitalization(.none)
                        .onAppear {
                            if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                                self.new_item_name = unwrapped_item.item_name
                            }
                        }
                } header: {
                    Text("項目名")
                }
                Section {
                    TextField("e.g. 10", value: $new_skip_num, format: .number)
                        .onAppear {
                            if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                                self.new_skip_num = unwrapped_item.skip_num
                            }
                        }
                } header: {
                    Text("スキップ可能数")
                }
                Section {
                    Picker("", selection: $new_color) {
                        Text("red").tag(Color.red)
                        Text("green").tag(Color.green)
                        Text("blue").tag(Color.blue)
                    }
                    .onAppear {
                        if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                            self.new_color = unwrapped_item.color
                        }
                    }
                } header: {
                    Text("色")
                }
            }
            .navigationTitle("アイテム編集")
        }
        //.onDisappear() {
        //    hab_chain_data.printAll()
        //}
        Button(action: {
            pressEditButtonAction()
        }) {
            Text("Done")
                .frame(maxWidth: .infinity)
                .frame(height: BUTTON_HEIGHT_PX)
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
        Button(action: {
            pressCancelButtonAction()
        }) {
            Text("Cancel")
                .frame(maxWidth: .infinity)
                .frame(height: BUTTON_HEIGHT_PX)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .padding()
    }
    func pressEditButtonAction() {
        if new_item_name == "" {
            is_show_alert = true
            error_kind = .blank_item_name
        } else {
            hab_chain_data.items[trgt_item_id]!.item_name = new_item_name
            hab_chain_data.items[trgt_item_id]!.skip_num = new_skip_num
            hab_chain_data.items[trgt_item_id]!.color = new_color
            is_show_item_edit_view = false
            is_show_alert = false
        }
    }
    func pressCancelButtonAction() {
        is_show_item_edit_view = false
    }
}
