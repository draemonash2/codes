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
                            hab_chain_data.printAll()
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
                        ForEach(hab_chain_data.item_id_list, id: \.self) { item_id in
                            if let unwraped_item = hab_chain_data.items[item_id] {
                                HStack {
                                    Text(unwraped_item.item_name)
                                    
                                    Spacer()

                                    ForEach(-3..<1) { i in
                                        Button {
                                            print("pressed \(unwraped_item.item_name) \(i) day button")
                                            let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
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
                                            let date: Date = Calendar.current.date(byAdding: .day,value: i, to: Date())!
                                            let continuation_cnt: Int = hab_chain_data.calcContinuationCount(base_date: date, item_id: item_id)
                                            let color_str: String = getColorString(color: unwraped_item.color, continuation_count: continuation_cnt)
                                            Text(String(continuation_cnt))
                                                .font(.caption)
                                                .frame(width: 40, height: 40)
                                                .multilineTextAlignment(.center)
                                                .foregroundColor(Color.white)
                                                .background(Color(color_str))
                                                .clipShape(Circle())
                                        }
                                        .buttonStyle(PlainButtonStyle())
                                    }
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
                    //.frame(height: 300)
                    .environment(\.editMode, .constant(.active))
                    //Text("test!!!!!!")
                    //    .padding()
                    //Button(action: {
                    //    hab_chain_data.printAll()
                    //}) {
                    //    Text("button")
                    //        .padding()
                    //}
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
                    OverlayView(is_presented: $is_overlay_presented, trgt_status: $trgt_status)
                }
            }
        }
        .onAppear {
            //hab_chain_data.setValueForTest()
            //_test_getColorString()
        }
        .navigationViewStyle(StackNavigationViewStyle()) // for iPad
    }
    
    func getColorString(color: Color, continuation_count: Int) -> String {
        var color_name: String = ""
        switch color {
            case Color.red: color_name = "color_red"
            case Color.blue: color_name = "color_blue"
            case Color.green: color_name = "color_green"
            default: return ""
        }
        
        var color_index: Int = 0
        let min: Int = 0
        let max: Int = 5
        if continuation_count < min {
            color_index = min
        } else if min <= continuation_count && continuation_count <= max {
            color_index = continuation_count
        } else {
            color_index = max
        }
        
        return String(color_name) + String(color_index)
    }
    func _test_getColorString() {
        print(getColorString(color: Color.red, continuation_count: 0))
        print(getColorString(color: Color.red, continuation_count: 1))
        print(getColorString(color: Color.red, continuation_count: 3))
        print(getColorString(color: Color.red, continuation_count: 5))
        print(getColorString(color: Color.red, continuation_count: 6))
        print(getColorString(color: Color.blue, continuation_count: 3))
        print(getColorString(color: Color.green, continuation_count: 3))
        print(getColorString(color: Color.white, continuation_count: 3))
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}

struct OverlayView: View {
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
