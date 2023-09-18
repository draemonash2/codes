//
//  ItemSettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI
import WidgetKit

struct ItemSettingViewSetting {
    let ICON_SIZE_PX: CGFloat = 17
    let COLOR_INDI_SIZE_PX: CGFloat = 10
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let ARCHIVE_ICON_OPACITY: CGFloat = 0.2
}

struct ItemSettingView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Environment(\.dismiss) var dismiss
    @Binding var hab_chain_data: HabChainData
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @State var is_show_item_add_view: Bool = false
    @State var is_show_item_edit_view: Bool = false
    @State var is_show_item_status_edit_view: Bool = false
    @State var trgt_item_id: String = ""
    @State var trgt_item_name: String = ""
    private let VIEW_SETTING: ItemSettingViewSetting = ItemSettingViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
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
                                print("pressed \(unwraped_item.item_name) color button")
                            } label: {
                                let color_str: String = getColorString(color: unwraped_item.color, continuation_count: 3)
                                Text("")
                                    .font(.caption)
                                    .frame(width: VIEW_SETTING.COLOR_INDI_SIZE_PX, height: VIEW_SETTING.COLOR_INDI_SIZE_PX)
                                    .multilineTextAlignment(.center)
                                    .foregroundColor(Color.white)
                                    .background(Color(color_str))
                                    .clipShape(Circle())
                            }
                            .buttonStyle(PlainButtonStyle())
                            Button {
                                print("pressed \(unwraped_item.item_name) archive button")
                                if unwraped_item.is_archived {
                                    hab_chain_data.items[item_id]!.is_archived = false
                                } else {
                                    hab_chain_data.items[item_id]!.is_archived = true
                                }
                            } label: {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                if unwraped_item.is_archived == true {
                                    Image(systemName: "tray.and.arrow.down")
                                        .resizable()
                                        .aspectRatio(contentMode: .fit)
                                        .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                        .foregroundColor(icon_color)
                                } else {
                                    Image(systemName: "tray.and.arrow.down")
                                        .resizable()
                                        .aspectRatio(contentMode: .fit)
                                        .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                        .foregroundColor(icon_color)
                                        .opacity(VIEW_SETTING.ARCHIVE_ICON_OPACITY)
                                }
                            }
                            .buttonStyle(PlainButtonStyle())
                            Button {
                                print("pressed \(unwraped_item.item_name) cal button")
                                trgt_item_id = item_id
                                is_show_item_status_edit_view = true
                            } label: {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "calendar")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
                            }
                            .buttonStyle(PlainButtonStyle())
                            Button {
                                print("pressed \(unwraped_item.item_name) button")
                                trgt_item_id = item_id
                                trgt_item_name = unwraped_item.item_name
                                is_show_item_edit_view = true
                            } label: {
                                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                Image(systemName: "pencil")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                    .foregroundColor(icon_color)
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
                    .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
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
            .sheet(isPresented: $is_show_item_status_edit_view) {
                ItemStatusEditView(
                    hab_chain_data: $hab_chain_data,
                    is_show_item_status_edit_view: $is_show_item_status_edit_view,
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
