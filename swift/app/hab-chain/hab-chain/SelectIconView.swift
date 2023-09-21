//
//  SelectIconView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/09/02.
//

import SwiftUI

struct SelectIconViewSetting {
    let ICON_SIZE_PX :CGFloat = 25
    let ICON_COLUMN_NUM :Int = 8
}

struct SelectIconView: View {
    @Environment(\ .colorScheme) var colorScheme
    //@Binding var hab_chain_data: HabChainData
    @Binding var is_show_select_icon_view: Bool
    //@Binding var item_id: String
    @Binding var icon_name: String
    private let VIEW_SETTING: SelectIconViewSetting = SelectIconViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        VStack {
            List {
                let symbol_num: Int = sf_symbols.count
                let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                //Text(String(symbol_num))
                ForEach(0..<(symbol_num/VIEW_SETTING.ICON_COLUMN_NUM+1), id: \.self) { row_idx in
                    HStack {
                        Spacer()
                        ForEach(0..<VIEW_SETTING.ICON_COLUMN_NUM, id: \.self) { clm_idx in
                            let idx: Int = (row_idx * VIEW_SETTING.ICON_COLUMN_NUM) + clm_idx
                            if idx < symbol_num {
                                Button(action: {
                                    print("pressed \(sf_symbols[idx]) button")
                                    //hab_chain_data.items[item_id]!.icon_name = sf_symbols[idx]
                                    icon_name = sf_symbols[idx]
                                    is_show_select_icon_view = false
                                }) {
                                    Image(systemName: sf_symbols[idx])
                                        .resizable()
                                        .scaledToFit()
                                        .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                                        .foregroundColor(icon_color)
                                }
                                .buttonStyle(PlainButtonStyle())
                            } else {
                                Button(action: {
                                    print("pressed unknown button")
                                }) {
                                    Image(systemName: sf_symbols[symbol_num-1])
                                        .resizable()
                                        .scaledToFit()
                                        .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                                        .opacity(0)
                                }
                                .buttonStyle(PlainButtonStyle())
                            }
                            Spacer()
                        }
                    }
                }
            }
        }
    }
}

//struct SelectIconView_Previews: PreviewProvider {
//    static var previews: some View {
//        SelectIconView()
//    }
//}
