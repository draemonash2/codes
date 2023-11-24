//
//  SelectIconView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/09/02.
//

import SwiftUI

struct SelectIconViewSetting {
    let ICON_SIZE_PX :CGFloat = 25
    let ICON_COLUMN_NUM :Int = 7
    let ICON_COLUMN_NUM_MIN :Int = 3
    let ICON_PADDING_SIZE_PX: CGFloat = 5
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let LIST_PADDING_PX :CGFloat = 5
    let TITLE_TOP_PADDING_PX :CGFloat = 30
    let BUTTON_PADDING_TOPBOTTOM_PX: CGFloat = 4
}

struct SelectIconView: View {
    //@Environment(\ .colorScheme) var colorScheme
    @Binding var is_show_select_icon_view: Bool
    @Binding var icon_name: String
    @State var selected_icon_name: String
    private let VIEW_SETTING: SelectIconViewSetting = SelectIconViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        GeometryReader { geometry in
            VStack {
                Text("アイコン選択")
                    .font(.title)
                    .padding([.top], VIEW_SETTING.TITLE_TOP_PADDING_PX )
                List {
                    let symbol_num: Int = sf_symbols.count
                    //Text(String(symbol_num))
                    let icon_column_num_tmp: Int = Int((geometry.size.width - VIEW_SETTING.LIST_PADDING_PX*2) / (VIEW_SETTING.ICON_SIZE_PX + VIEW_SETTING.ICON_PADDING_SIZE_PX*2 + 1))
                    let icon_column_num: Int = icon_column_num_tmp > VIEW_SETTING.ICON_COLUMN_NUM_MIN ? icon_column_num_tmp : VIEW_SETTING.ICON_COLUMN_NUM_MIN
                    //Text("\(icon_column_num_tmp)")
                    let icon_row_num: Int = (symbol_num/icon_column_num+1)
                    ForEach(0..<icon_row_num, id: \.self) { row_idx in
                        HStack {
                            Spacer(minLength: 1)
                            ForEach(0..<icon_column_num, id: \.self) { clm_idx in
                                let idx: Int = (row_idx * icon_column_num) + clm_idx
                                SelectIconButton(selected_icon_name: $selected_icon_name, sf_symbols_idx: idx)
                                Spacer(minLength: 1)
                            }
                        }
                        .listRowInsets(EdgeInsets())
                    }
                }
                .listStyle(.plain)
                .environment(\.editMode, .constant(.active))
                .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                Button(action: {
                    icon_name = selected_icon_name
                    is_show_select_icon_view = false
                }) {
                    Text("Done")
                        .frame(maxWidth: .infinity)
                        .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                        .multilineTextAlignment(.center)
                        .background(Color.blue)
                        .foregroundColor(Color.white)
                }
                .padding([.top, .bottom], VIEW_SETTING.BUTTON_PADDING_TOPBOTTOM_PX )
                .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
                Button(action: {
                    is_show_select_icon_view = false
                }) {
                    Text("Cancel")
                        .frame(maxWidth: .infinity)
                        .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                        .multilineTextAlignment(.center)
                        .background(Color.blue)
                        .foregroundColor(Color.white)
                }
                .padding([.top, .bottom], VIEW_SETTING.BUTTON_PADDING_TOPBOTTOM_PX )
                .padding([.leading, .trailing], VIEW_SETTING.LIST_PADDING_PX )
            }
            //.padding(0)
        }
    }
}

struct SelectIconButton: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var selected_icon_name: String
    @State var sf_symbols_idx: Int
    private let VIEW_SETTING: SelectIconViewSetting = SelectIconViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        let symbol_num: Int = sf_symbols.count
        let icon_color :Color = colorScheme == .light ? Color.black: Color.white
        if sf_symbols_idx < symbol_num {
            Button(action: {
                print("pressed \(sf_symbols[sf_symbols_idx]) button")
                selected_icon_name = sf_symbols[sf_symbols_idx]
            }) {
                if selected_icon_name == sf_symbols[sf_symbols_idx] {
                    Image(systemName: sf_symbols[sf_symbols_idx])
                        .resizable()
                        .scaledToFit()
                        .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                        .padding(VIEW_SETTING.ICON_PADDING_SIZE_PX)
                        .foregroundColor(icon_color)
                        .background(Color.blue)
                } else {
                    Image(systemName: sf_symbols[sf_symbols_idx])
                        .resizable()
                        .scaledToFit()
                        .frame(width: VIEW_SETTING.ICON_SIZE_PX, height: VIEW_SETTING.ICON_SIZE_PX)
                        .padding(VIEW_SETTING.ICON_PADDING_SIZE_PX)
                        .foregroundColor(icon_color)
                }
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
                    .padding(VIEW_SETTING.ICON_PADDING_SIZE_PX)
            }
            .buttonStyle(PlainButtonStyle())
        }
    }
}


//struct SelectIconView_Previews: PreviewProvider {
//    static var previews: some View {
//        SelectIconView()
//    }
//}
