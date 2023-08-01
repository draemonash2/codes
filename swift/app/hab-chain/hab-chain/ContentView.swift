//
//  ContentView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI

struct ContentView: View {
    @Environment(\ .colorScheme) var colorScheme
    @State var hab_chain_data: HabChainData = HabChainData()
    @State private var is_overlay_presented = false
    @State private var trgt_status: String = ""
    var add_item: Item?
    var body: some View {
        let _ = Self._printChanges()
        NavigationView {
            ZStack {
                VStack {
                    Text("hab-chain")
                        .font(.largeTitle)
                        .onAppear() {
                            //hab_chain_data.printAll()
                            writeJson()
                            readJson()
                            //testJsonDict()
                            //testJsonDict2()
                        }
                        .padding()
                    List {
                        HStack {
                            Spacer()
                            ForEach(-3..<1) { i in
                                let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                Text(hab_chain_data.convDateToMmdd(date: date))
                                    .font(.caption)
                                    .frame(width: 40, height: 40)
                                    .multilineTextAlignment(.center)
                            }
                        }
                        HStack {
                            Spacer()
                            ForEach(-3..<1) { i in
                                let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                let continuation_cnt: Int = hab_chain_data.calcContinuationCountAll(base_date: date)
                                let color_str: String = getColorString(color: Color.red, continuation_count: continuation_cnt)
                                Text(String(continuation_cnt))
                                    .font(.caption)
                                    .frame(width: 40, height: 40)
                                    .multilineTextAlignment(.center)
                                    .foregroundColor(Color.white)
                                    .background(Color(color_str))
                                    .clipShape(Circle())
                            }
                        }
                        ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                            if let unwraped_item = hab_chain_data.items[item_id] {
                                HStack {
                                    Text(unwraped_item.item_name)
                                    
                                    Spacer()

                                    ForEach(-3..<1) { i in
                                        let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                        let date_str: String = hab_chain_data.convDateToStr(date: date)
                                        let continuation_cnt: Int = hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
                                        let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
                                        ZStack {
                                            Button {
                                                print("pressed \(unwraped_item.item_name) \(i) day button")
                                                hab_chain_data.toggleItemStatus(item_id: item_id, date: date)
                                                // output popup message
                                                withAnimation(.easeIn(duration: 0.2)) {
                                                    trgt_status = hab_chain_data.getItemStatusStr(item_id: item_id, date: date)
                                                    is_overlay_presented = true
                                                }
                                                DispatchQueue.main.asyncAfter(deadline: .now() + 1.0) {
                                                    withAnimation(.easeOut(duration: 0.1)) {
                                                        is_overlay_presented = false
                                                    }
                                                }
                                            } label: {
                                                Text(String(continuation_cnt))
                                                    .font(.caption)
                                                    .frame(width: 40, height: 40)
                                                    .multilineTextAlignment(.center)
                                                    .foregroundColor(Color.white)
                                                    .background(Color(color_str))
                                                    .clipShape(Circle())
                                            }
                                            .buttonStyle(PlainButtonStyle())
                                            
                                            if let unwrapped_item_status = unwraped_item.status[date_str] {
                                                if unwrapped_item_status == .Done {
                                                    Circle()
                                                        .stroke(Color.white, lineWidth: 1)
                                                        .frame(width: 35, height: 35)
                                                } else if unwrapped_item_status == .Skip {
                                                    Circle()
                                                        .stroke(Color.white, style: StrokeStyle(lineWidth: 1, dash: [4]))
                                                        .frame(width: 35, height: 35)
                                                } else {
                                                    // Do Nothing
                                                }
                                            }
                                        }
                                    }
                                }
                                .contentShape(Rectangle())
                                .onTapGesture {
                                    print("pressed \(unwraped_item.item_name) item")
                                }
                            }
                        }
                    }
                    .environment(\.editMode, .constant(.active))
                }
                .toolbar {
                    ToolbarItem(placement: .navigationBarTrailing) {
                        NavigationLink(
                            destination: ItemSettingView(hab_chain_data: $hab_chain_data)
                        ) {
                            Image(colorScheme == .light ? "pencil_light": "pencil_dark")
                                .resizable()
                                .aspectRatio(contentMode: .fit)
                                .frame(height: 30)
                        }
                    }
                    ToolbarItem(placement: .navigationBarLeading) {
                        NavigationLink(
                            destination: AppSettingView()
                        ) {
                            Image(colorScheme == .light ? "setting_light": "setting_dark")
                                .resizable()
                                .aspectRatio(contentMode: .fit)
                                .frame(height: 30)
                        }
                    }
                }
                if is_overlay_presented {
                    PopupView(is_presented: $is_overlay_presented, trgt_status: $trgt_status)
                }
            }
        }
        .navigationViewStyle(StackNavigationViewStyle()) // for iPad
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}

struct PopupView: View {
    @Binding var is_presented: Bool
    @Binding var trgt_status: String
    
    var body: some View {
        Text(trgt_status)
            .frame(width: 150, height: 50)
            .foregroundColor(.black)
            .background(Color(red: 0.9, green: 0.9, blue: 0.9, opacity: 0.9))
            .cornerRadius(10)
    }
}
