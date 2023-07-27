//
//  SettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI

struct ItemEditView: View {
    @Binding var hab_chain_data: HabChainData
    @State var isShowItemAddView: Bool = false

    var body: some View {
        let _ = Self._printChanges()
        VStack {
            List {
                ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                    if let unwraped_item = hab_chain_data.items[item_id] {
                        Text(unwraped_item.item_name)
                    }
                }
                .onMove(perform: moveRow)
                .onDelete(perform: removeRow)
            }
            .environment(\.editMode, .constant(.active))
            //.navigationBarItems(trailing: EditButton())
            //EditButton()
            //    .padding()
            Button(action:{
                isShowItemAddView = true
            }) {
                Text("Add")
            }
            .sheet(isPresented: $isShowItemAddView) {
                ItemAddView(hab_chain_data: $hab_chain_data, isShowItemAddView: $isShowItemAddView)
            }

        }
    }

    func moveRow(from source: IndexSet, to destination: Int) {
        hab_chain_data.item_id_list.move(fromOffsets: source, toOffset: destination)
    }

    func removeRow(from source: IndexSet) {
        for idx in source {
            let item_id = hab_chain_data.item_id_list[idx]
            hab_chain_data.removeItem(trgt_item_id: item_id)
        }
    }
}

//struct SettingView_Previews: PreviewProvider {
//    static var previews: some View {
//        SettingView()
//    }
//}
