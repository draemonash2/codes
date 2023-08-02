//
//  AppSettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/29.
//

import SwiftUI

struct AppSettingView: View {
    @Binding var hab_chain_data: HabChainData
    var body: some View {
        Button(action: {
            hab_chain_data.readJson()
        }) {
            Text("Json読み出し")
                .frame(maxWidth: .infinity)
                .frame(height: 50)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .padding()
        Button(action: {
            hab_chain_data.writeJson()
        }) {
            Text("Json書き込み")
                .frame(maxWidth: .infinity)
                .frame(height: 50)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .padding()
    }
}

struct AppSettingView_Previews: PreviewProvider {
    @State static var dummy_hab_chain_data: HabChainData = HabChainData()
    static var previews: some View {
        AppSettingView(hab_chain_data: $dummy_hab_chain_data)
    }
}
