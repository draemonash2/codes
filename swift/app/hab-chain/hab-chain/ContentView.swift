//
//  ContentView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI

struct ContentView: View {
    @State var hab_chain_data: HabChainData = HabChainData()
    var add_item: Item?
    var body: some View {
        let _ = Self._printChanges()
        NavigationView {
            VStack {
                Text("test!!!")
                    .padding()
                    .onAppear() {
                        hab_chain_data.printAll()
                    }
                List {
                    ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                        if let unwraped_item = hab_chain_data.items[item_id] {
                            Text(unwraped_item.item_name)
                        }
                    }
                    //.onMove(perform: moveRow)
                    //.onDelete(perform: removeRow)
                }
                .environment(\.editMode, .constant(.active))
                Text("test!!!!!!")
                    .padding()
                Button(action: {
                    hab_chain_data.printAll()
                }) {
                    Text("button")
                        .padding()
                }
            }
            .toolbar {
                ToolbarItem(placement: .navigationBarTrailing) {
                    NavigationLink(
                        destination: ItemEditView(hab_chain_data: $hab_chain_data)
                    ) {
                        Text("Edit")
                    }
                }
            }
        }
        //.onAppear {
        //    hab_chain_data.setValueForTest()
        //}
        .navigationViewStyle(StackNavigationViewStyle()) // for iPad
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

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}
