//
//  ItemAddView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/27.
//

import SwiftUI

struct ItemAddView: View {
    @Binding var hab_chain_data: HabChainData
    @Binding var isShowItemAddView: Bool
    @State var is_cancel: Bool = false
    @State var new_item_name: String = ""
    @State var new_skip_num: Int = 10
    @State var new_color: Color = Color.red

    var body: some View {
        NavigationView {
            Form {
            //VStack {
                Section {
                    TextField("e.g. 勉強", text: $new_item_name)
                } header: {
                    Text("項目名")
                }
                Section {
                    TextField("e.g. 10", value: $new_skip_num, format: .number)
                } header: {
                    Text("スキップ可能数")
                }
                Section {
                    Picker("色", selection: $new_color) {
                        Text("red").tag(Color.red)
                        Text("green").tag(Color.green)
                        Text("blue").tag(Color.blue)
                    }
                } header: {
                    Text("色")
                }
            }
        }
        .navigationTitle("新規アイテム")
        Button(action: {
            isShowItemAddView = false
        }) {
            Text("Add")
        }
        .padding()
        Button(action: {
            is_cancel = true
            isShowItemAddView = false
        }) {
            Text("Cancel")
        }
        //.padding(20)
        //.frame(width: 350, height: 500)
        .onDisappear() {
            if is_cancel == false {
                let item = Item(item_name: new_item_name, skip_num: new_skip_num, color: new_color)
                hab_chain_data.addItem(new_item_id: hab_chain_data.generateItemId(), new_item: item)
            }
        }
        
    }
}

//struct ItemAddView_Previews: PreviewProvider {
//    static var previews: some View {
//        ItemAddView(item_id: "aaa")
//    }
//}
