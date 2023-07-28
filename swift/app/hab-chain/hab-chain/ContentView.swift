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
        let color_idx: Int = 5
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
                            HStack {
                                Text(unwraped_item.item_name)
                                
                                Spacer()

                                Button {
                                    print("pressed \(unwraped_item.item_name) 3day before button")
                                } label: {
                                    Text("3")
                                        //.bold()
                                        //.padding()
                                        .frame(width: 40, height: 40)
                                        .foregroundColor(Color.white)
                                        .background(Color("color_red" + String(color_idx)))
                                        .clipShape(Circle())
                                }
                                .buttonStyle(PlainButtonStyle())
                                
                                Button {
                                    print("pressed \(unwraped_item.item_name) 2day before button")
                                } label: {
                                    Text("2")
                                        //.bold()
                                        //.padding()
                                        .frame(width: 40, height: 40)
                                        .foregroundColor(Color.white)
                                        .background(Color.yellow)
                                        .clipShape(Circle())
                                }
                                .buttonStyle(PlainButtonStyle())
                                
                                Button {
                                    print("pressed \(unwraped_item.item_name) 1day before button")
                                } label: {
                                    Text("1")
                                        //.bold()
                                        //.padding()
                                        .frame(width: 40, height: 40)
                                        .foregroundColor(Color.white)
                                        .background(Color.yellow)
                                        .clipShape(Circle())
                                }
                                .buttonStyle(PlainButtonStyle())
                                
                                Button {
                                    print("pressed \(unwraped_item.item_name) today button")
                                } label: {
                                    Text("0")
                                        //.bold()
                                        //.padding()
                                        .frame(width: 40, height: 40)
                                        .foregroundColor(Color.white)
                                        .background(Color.yellow)
                                        .clipShape(Circle())
                                }
                                .buttonStyle(PlainButtonStyle())
                            }
                            .contentShape(Rectangle())
                            .onTapGesture {
                                print("pressed \(unwraped_item.item_name) item")
                            }
                        }
                    }
                    //.onMove(perform: moveRow)
                    //.onDelete(perform: removeRow)
                }
                .frame(height: 300)
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
                        destination: ItemSettingView(hab_chain_data: $hab_chain_data)
                    ) {
                        Text("Edit")
                    }
                }
                ToolbarItem(placement: .navigationBarLeading) {
                    NavigationLink(
                        destination: ItemSettingView(hab_chain_data: $hab_chain_data)
                    ) {
                        Image("setting")
                            .resizable()
                            .aspectRatio(contentMode: .fit)
                            .frame(height: 30)
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
    
    func getColorString(color: Color) -> String {
        switch color {
            case Color.red: return "color_red"
            case Color.blue: return "color_blue"
            case Color.green: return "color_green"
            default: return ""
        }
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}
