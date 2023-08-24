//
//  ItemStatusEditView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/08/12.
//

import SwiftUI

struct ItemStatusEditViewSetting {
    let DATE_NUM :Int = 90
    let BUTTON_SIZE_PX :CGFloat? = 40
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 35
    let ICON_SIZE_PX :CGFloat? = 30
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let STATUS_TEXT_WIDTH_PX: CGFloat = 60
    let ICON_NAME_DIC: Dictionary<ItemStatus, String> = [
        .Done: "checkmark.square",
        .NotYet: "square",
        .Skip: "clear"
    ]
}

struct ItemStatusEditView: View {
    @Environment(\ .colorScheme) var colorScheme
    @Binding var hab_chain_data: HabChainData
    @Binding var is_show_item_status_edit_view: Bool
    @Binding var trgt_item_id: String
    @State var shouldPresentPopUpDialog: Bool = false
    private let VIEW_SETTING: ItemStatusEditViewSetting = ItemStatusEditViewSetting()

    var body: some View {
        ZStack {
            VStack {
                if let unwrapped_item = hab_chain_data.items[trgt_item_id] {
                    HStack {
                        Text(unwrapped_item.item_name)
                            .font(.largeTitle)
                            .padding(.all)
                        #if false
                        Button {
                            withAnimation {
                                shouldPresentPopUpDialog = true
                            }
                        } label: {
                            let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                            Image(systemName: "info.circle")
                                .resizable()
                                .aspectRatio(contentMode: .fit)
                                .frame(height: 15)
                                .foregroundColor(icon_color)
                        }
                        #endif
                    }
                    List {
                        ForEach(0..<VIEW_SETTING.DATE_NUM, id: \.self) { i in
                            let date_offset: Int = -i
                            let date: Date = Calendar.current.date(byAdding: .day,value: date_offset, to: Date())!
                            HStack {
                                Text(hab_chain_data.formatDateMmdd(date: date, delimiter: " "))
                                Spacer()
                                Button(action:{
                                    hab_chain_data.toggleItemStatus(item_id: trgt_item_id, date: date)
                                }) {
                                    let date_str = hab_chain_data.convToStr(date: date)
                                    let icon_color :Color = colorScheme == .light ? Color.black: Color.white
                                    if unwrapped_item.daily_statuses.keys.contains(date_str) {
                                        if let unwrapped_item_status = unwrapped_item.daily_statuses[date_str] {
                                            Image(systemName: VIEW_SETTING.ICON_NAME_DIC[unwrapped_item_status]!)
                                                .resizable()
                                                .aspectRatio(contentMode: .fit)
                                                .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                                .foregroundColor(icon_color)
                                        }
                                    } else {
                                        Image(systemName: VIEW_SETTING.ICON_NAME_DIC[.NotYet]!)
                                            .resizable()
                                            .aspectRatio(contentMode: .fit)
                                            .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                                            .foregroundColor(icon_color)
                                    }
                                }
                                let date_str = hab_chain_data.convToStr(date: date)
                                if unwrapped_item.daily_statuses.keys.contains(date_str) {
                                    if let unwrapped_item_status = unwrapped_item.daily_statuses[date_str] {
                                        Text(hab_chain_data.convToStr(item_status: unwrapped_item_status))
                                            .frame(width: VIEW_SETTING.STATUS_TEXT_WIDTH_PX, alignment: .leading)
                                    }
                                } else {
                                    Text(hab_chain_data.convToStr(item_status: .NotYet))
                                        .frame(width: VIEW_SETTING.STATUS_TEXT_WIDTH_PX, alignment: .leading)
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
            if shouldPresentPopUpDialog {
                PopUpDialogView(isPresented: $shouldPresentPopUpDialog)
            }
        }
    }
}

//struct ItemStatusEditView_Previews: PreviewProvider {
//    static var previews: some View {
//        ItemStatusEditView()
//    }
//}

struct PopUpDialogView: View {

    @Environment(\ .colorScheme) var colorScheme
    @Binding var isPresented: Bool
    let isEnabledToCloseByBackgroundTap: Bool = true

    private let buttonSize: CGFloat = 24
    private let VIEW_SETTING: ItemStatusEditViewSetting = ItemStatusEditViewSetting()

    var body: some View {
        GeometryReader { proxy in
            let dialogWidth = proxy.size.width * 0.75
            ZStack {
                BackgroundView(color: .gray.opacity(0.7))
                    .onTapGesture {
                        if isEnabledToCloseByBackgroundTap {
                            withAnimation {
                                isPresented = false
                            }
                        }
                    }
                VStack (alignment: .leading) {
                    HStack {
                        Image(systemName: VIEW_SETTING.ICON_NAME_DIC[.NotYet]!)
                            .resizable()
                            .aspectRatio(contentMode: .fit)
                            .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                            .foregroundColor(Color.black)
                        Text(": 未完了(NotYet)")
                            .foregroundColor(Color.black)
                    }
                    HStack {
                        Image(systemName: VIEW_SETTING.ICON_NAME_DIC[.Done]!)
                            .resizable()
                            .aspectRatio(contentMode: .fit)
                            .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                            .foregroundColor(Color.black)
                        Text(": 完了(Done)")
                            .foregroundColor(Color.black)
                    }
                    HStack {
                        Image(systemName: VIEW_SETTING.ICON_NAME_DIC[.Skip]!)
                            .resizable()
                            .aspectRatio(contentMode: .fit)
                            .frame(height: VIEW_SETTING.ICON_SIZE_PX)
                            .foregroundColor(Color.black)
                        Text(": スキップ(Skip)")
                            .foregroundColor(Color.black)
                    }
                }
                    .frame(width: dialogWidth)
                    .padding()
                    .padding(.top, buttonSize)
                    .background(.white)
                    .cornerRadius(12)
                    .overlay(alignment: .topTrailing) {
                        CloseButton(fontSize: buttonSize,
                                    weight: .bold,
                                    color: .gray.opacity(0.7)) {
                            withAnimation {
                                isPresented = false
                            }
                        }
                        .padding(4)
                    }
            }
        }
    }
}
//extension PopUpDialogView {
//
//    init(isPresented: Binding<Bool>,
//         isEnabledToCloseByBackgroundTap: Bool = true,
//         @ViewBuilder _ content: () -> Content) {
//        _isPresented = isPresented
//        self.isEnabledToCloseByBackgroundTap = isEnabledToCloseByBackgroundTap
//        self.content = content()
//    }
//}

struct BackgroundView: View {
    let color: Color
    var body: some View {
        Rectangle()
            .fill(color)
            .ignoresSafeArea()
    }
}

struct CloseButton: View {
    let fontSize: CGFloat
    let weight: Font.Weight
    let color: Color
    let action: () -> Void
    var body: some View {
        Button {
            action()
        } label: {
            Image(systemName: "xmark.circle")
        }
        .font(.system(size: fontSize,
                      weight: weight,
                      design: .default))
        .foregroundColor(color)
    }
}
