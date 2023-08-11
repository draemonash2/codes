//
//  ItemSettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI
import WidgetKit

struct ItemSettingView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Environment(\.dismiss) var dismiss
    @Binding var hab_chain_data: HabChainData
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @State var is_show_item_add_view: Bool = false
    @State var is_show_item_edit_view: Bool = false
    @State var trgt_item_id: String = ""
    @State var trgt_item_name: String = ""

    var body: some View {
        let ICON_SIZE_PX: CGFloat = 20
        let BUTTON_HEIGHT_PX: CGFloat = 50
        
        let _ = Self._printChanges()
        VStack {
            Text("アイテム設定")
                .font(.largeTitle)
            List {
                ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                    if let unwraped_item = hab_chain_data.items[item_id] {
                        HStack {
                            Text(unwraped_item.item_name)
                            
                            Spacer()

                            Button {
                                print("pressed \(unwraped_item.item_name) button")
                                trgt_item_id = item_id
                                trgt_item_name = unwraped_item.item_name
                                is_show_item_edit_view = true
                            } label: {
                                Image(colorScheme == .light ? "pencil_light": "pencil_dark")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: ICON_SIZE_PX)
                            }
                            .buttonStyle(PlainButtonStyle())
                        }
                        .contentShape(Rectangle())
                        .onTapGesture {
                            print("pressed \(unwraped_item.item_name) item")
                        }
                    }
                }
                .onMove(perform: moveRow)
                .onDelete(perform: removeRow)
                Button(action: {
                    is_show_item_add_view = true
                }) {
                    Text("+ add")
                        .foregroundColor(Color.gray)
                }
                .buttonStyle(BorderlessButtonStyle())
            }
            .environment(\.editMode, .constant(.active))
            Button(action:{
                dismiss()
                app_json_string = hab_chain_data.getRawStruct2JsonString()
                WidgetCenter.shared.reloadAllTimelines()
            }) {
                Text("Done")
                    .frame(maxWidth: .infinity)
                    .frame(height: BUTTON_HEIGHT_PX)
                    .multilineTextAlignment(.center)
                    .background(Color.blue)
                    .foregroundColor(Color.white)
            }
            .padding()
            .sheet(isPresented: $is_show_item_add_view) {
                ItemAddView(
                    hab_chain_data: $hab_chain_data,
                    is_show_item_add_view: $is_show_item_add_view
                )
            }
            .sheet(isPresented: $is_show_item_edit_view) {
                ItemEditView(
                    hab_chain_data: $hab_chain_data,
                    is_show_item_edit_view: $is_show_item_edit_view,
                    trgt_item_id: $trgt_item_id
                )
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

struct ItemSettingView_Previews: PreviewProvider {
    @State static var dummy_hab_chain_data: HabChainData = HabChainData()
    static var previews: some View {
        ItemSettingView(hab_chain_data: $dummy_hab_chain_data)
    }
}
