//
//  ContentView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/25.
//

import SwiftUI

struct ContentView: View {
    var body: some View {
        NavigationView {
            VStack {
                let str = Test2()
                Text(str)
                    .padding()
            }
            .toolbar {
                ToolbarItem(placement: .navigationBarLeading) {
                    NavigationLink(destination: SettingView())
                    {
                        Text("<")
                    }
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
