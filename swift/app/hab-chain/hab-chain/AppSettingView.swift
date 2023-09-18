//
//  AppSettingView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/07/29.
//

import SwiftUI
import WidgetKit

struct AppSettingViewSetting {
    let BUTTON_WIDTH_PX: CGFloat = 100
    let BUTTON_HEIGHT_PX: CGFloat = 50
    let BUTTON_CORNER_RADIUS: CGFloat = 10
}

struct AppSettingView: View {
    @Binding var hab_chain_data: HabChainData
    @AppStorage("app_json_string", store: UserDefaults(suiteName: "group.hab_chain")) var app_json_string: String = ""
    @Environment(\.dismiss) var dismiss
    @State private var showingAlertBackup = false
    @State private var showingAlertRestore = false
    private let VIEW_SETTING: AppSettingViewSetting = AppSettingViewSetting()
    var body: some View {
        Form {
            Section {
                Picker("", selection: $hab_chain_data.whole_color) {
                    Text("red").tag(Color.red)
                    Text("green").tag(Color.green)
                    Text("blue").tag(Color.blue)
                }
            } header: {
                Text("色")
            }
            Section {
                Toggle(isOn: $hab_chain_data.is_show_status_popup) {
                }
            } header: {
                Text("Popup表示")
            }
            Section {
                HStack {
                    Spacer()
                    
                    Button(action: {
                        showingAlertBackup = true
                    }) {
                        Text("バックアップ")
                            .bold()
                            .padding()
                            .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                            .foregroundColor(Color.white)
                            .background(Color.blue)
                            .cornerRadius(VIEW_SETTING.BUTTON_CORNER_RADIUS)
                    }
                    .buttonStyle(PlainButtonStyle())
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
                    Button(action: {
                        showingAlertRestore = true
                    }) {
                        Text("復旧")
                            .bold()
                            .padding()
                            .frame(height: VIEW_SETTING.BUTTON_HEIGHT_PX)
                            .foregroundColor(Color.white)
                            .background(Color.blue)
                            .cornerRadius(VIEW_SETTING.BUTTON_CORNER_RADIUS)
                    }
                    .buttonStyle(PlainButtonStyle())
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
                }
            } header: {
                Text("バックアップ")
            }
        }
        Button(action: {
            pressDoneButtonAction()
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
    func pressDoneButtonAction() {
        app_json_string = hab_chain_data.getRawStruct2JsonString()
        WidgetCenter.shared.reloadAllTimelines()
        dismiss()
    }
}

struct AppSettingView_Previews: PreviewProvider {
    @State static var dummy_hab_chain_data: HabChainData = HabChainData()
    static var previews: some View {
        AppSettingView(hab_chain_data: $dummy_hab_chain_data)
    }
}
