//
//  AppSettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/29.
//

import SwiftUI

struct AppSettingView: View {
    @Binding var hab_chain_data: HabChainData
    @AppStorage("hab_chain_data_jsonstr") var hab_chain_data_jsonstr: String = ""
    @State private var showingAlertBackup = false
    @State private var showingAlertRestore = false
    var body: some View {
        let BUTTON_HEIGHT_PX: CGFloat = 50

        Button(action: {
            showingAlertBackup = true
        }) {
            Text("バックアップ実行")
                .frame(maxWidth: .infinity)
                .frame(height: BUTTON_HEIGHT_PX)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .alert(isPresented: $showingAlertBackup) {
            Alert(
                title: Text("確認"),
                message: Text("バックアップを行います。よろしいですか？"),
                primaryButton: .default(Text("はい"), action: {
                    hab_chain_data.saveJsonString()
                }),
                secondaryButton: .destructive(Text("いいえ"), action: {
                    print("処理を中断します。")
                })
            )
        }
        .padding()
        Button(action: {
            showingAlertRestore = true
        }) {
            Text("バックアップデータ復旧")
                .frame(maxWidth: .infinity)
                .frame(height: BUTTON_HEIGHT_PX)
                .multilineTextAlignment(.center)
                .background(Color.blue)
                .foregroundColor(Color.white)
        }
        .alert(isPresented: $showingAlertRestore) {
            Alert(
                title: Text("確認"),
                message: Text("バックアップデータの復旧を行います。よろしいですか？"),
                primaryButton: .default(Text("はい"), action: {
                    hab_chain_data.loadJsonString()
                }),
                secondaryButton: .destructive(Text("いいえ"), action: {
                    print("処理を中断します。")
                })
            )
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
