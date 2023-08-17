//
//  ItemStatusEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/08/12.
//

import SwiftUI

struct ItemStatusEditViewSetting {
    let DATE_NUM :Int = 60
    let BUTTON_SIZE_PX :CGFloat? = 40
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 35
    let ICON_SIZE_PX :CGFloat? = 30
    let BUTTON_HEIGHT_PX: CGFloat = 50
}

struct ItemStatusEditView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_status_edit_view: Bool
    @Binding var trgt_item_id: String
    private let VIEW_SETTING: ItemStatusEditViewSetting = ItemStatusEditViewSetting()
    var body: some View {
        if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
            Text(unwrapped_item.item_name)
                .font(.largeTitle)
                .padding(.all)
            List {
                ForEach(0..<VIEW_SETTING.DATE_NUM, id: \.self) { i in
                    let date_offset: Int = -i
                    let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
                    HStack {
                        Text(hab_chain_data.convDateToMmdd(date: date, delimiter: " "))
                        Spacer()
                        Button(action:{
                            hab_chain_data.toggleItemStatus(item_id: trgt_item_id, date: date)
                        }) {
                            let date_str = hab_chain_data.convDateToStr(date: date)
                            if unwrapped_item.status.keys.contains(date_str) {
                                if let unwrapped_item_status = unwrapped_item.status[date_str] {
                                    let icon_name: String = "check_" + unwrapped_item_status.rawValue.lowercased()
                                    Image(colorScheme == .light ? icon_name + "_light": icon_name + "_dark")
                                        .resizable()
                                        .aspectRatio(contentMode: .fit)
                                        .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                }
                            } else {
                                let icon_name: String = "check_notyet"
                                Image(colorScheme == .light ? icon_name + "_light": icon_name + "_dark")
                                    .resizable()
                                    .aspectRatio(contentMode: .fit)
                                    .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                            }
                        }
                    }
                }
            }
            Button(action: {
                is_show_item_status_edit_view = false
            }) {
                Text("Done")
                    .frame(maxWidth: .infinity)
                    .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                    .multilineTextAlignment(.center)
                    .background(Color.blue)
                    .foregroundColor(Color.white)
            }
            .padding()
        }
    }
}

//struct ItemStatusEditView_Previews: PreviewProvider {
//    static var previews: some View {
//        ItemStatusEditView()
//    }
//}