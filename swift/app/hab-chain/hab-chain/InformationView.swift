//
//  InformationView.swift
//  hab-chain
//
//  Created by Tatsuya Endo on 2023/08/21.
//

import SwiftUI

struct InformationViewSetting {
    let BUTTON_SIZE_PX :CGFloat? = 35
    let BUTTON_INCIRCLE_SIZE_PX :CGFloat? = 30
    let BUTTON_INCIRCLE_LINEWIDTH :CGFloat? = 1
}

struct InformationView: View {
    private let VIEW_SETTING: InformationViewSetting = InformationViewSetting()
    private let FUNC_SETTING: FunctionSetting = FunctionSetting()

    var body: some View {
        if FUNC_SETTING.debug_mode {
            let _ = Self._printChanges()
        }
        VStack (alignment: .leading) {
            //Text("hab-chainについて")
            //    .font(.largeTitle)
            //    .frame(maxWidth: .infinity, alignment: .center)
            //    .padding()
            ScrollView(.vertical, showsIndicators: true) {
                Group {
                    Text("hab-chainとは？")
                        .font(.title)
                        .padding()
                    Text("一つ一つの行動を鎖のようにつないで、その行動が何日間続いているかを記録して習慣化する方法です。\n続いた日数を記録し、鎖のように繋ぐことで習慣化しやすくします。\n")
                    Text("人は一度積み上げたものをゼロに戻すことを嫌います。\nせっかく積み上げてきたのに、リセットしてしまうのは、心理的抵抗を感じてしまいます。\n続ける苦しさよりも途中でやめるほうが苦しさを感じます。\nそういった続けないと苦しいという思いが、ハビットチェーンのポイントです。")
                }
                .frame(alignment: .leading)
                .fixedSize(horizontal: false, vertical: true)
                Group {
                    Text("本アプリについて")
                        .font(.title)
                        .padding()
                    Text("該当する日付のボタンを押すごとに、ステータスが変化します。\n")
                    Text("未完了 (NotYet) -> 完了(Done) →スキップ(Skip) -> 未完了 (NotYet) -> ...")
                    HStack {
                        ZStack {
                            let color_str: String = "color_red3"
                            Text("N")
                                .font(.caption)
                                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                .multilineTextAlignment(.center)
                                .foregroundColor(Color.white)
                                .background(Color(color_str))
                                .clipShape(Circle())
                            Circle()
                                .stroke(Color.white, lineWidth: VIEW_SETTING.BUTTON_INCIRCLE_LINEWIDTH!)
                                .frame(width: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                        }
                        Text(" : 完了 (N=連続達成回数)")
                        Spacer()
                    }
                    .frame(alignment: .leading)
                    HStack {
                        ZStack {
                            let color_str: String = "color_red3"
                            Text("N")
                                .font(.caption)
                                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                .multilineTextAlignment(.center)
                                .foregroundColor(Color.white)
                                .background(Color(color_str))
                                .clipShape(Circle())
                            Circle()
                                .stroke(Color.white, style: StrokeStyle(lineWidth: VIEW_SETTING.BUTTON_INCIRCLE_LINEWIDTH!, dash: [4]))
                                .frame(width: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX, height: VIEW_SETTING.BUTTON_INCIRCLE_SIZE_PX)
                        }
                        Text(" : スキップ (N=連続達成回数)")
                        Spacer()
                    }
                    .frame(alignment: .leading)
                    Text("")
                    Text("習慣が続けば続くほどボタンの色が濃くなり、習慣が切れると色がリセットされます。\n(スキップした場合は色が保持されます)")
                    HStack {
                        let button_num: Int = 5
                        ForEach(1..<(button_num+1), id: \.self) { button_idx in
                            let color_str :String = "color_red" + String(button_idx)
                            Text(String(button_idx))
                                .font(.caption)
                                .frame(width: VIEW_SETTING.BUTTON_SIZE_PX, height: VIEW_SETTING.BUTTON_SIZE_PX)
                                .multilineTextAlignment(.center)
                                .foregroundColor(Color.white)
                                .background(Color(color_str))
                                .clipShape(Circle())
                            if button_idx < button_num {
                                Text("->")
                                    .font(.caption)
                                    .multilineTextAlignment(.center)
                            }
                        }
                    }
                    .padding()
                    Text("色を維持することをモチベーションに、習慣を継続させましょう！\n")
                }
                .frame(alignment: .leading)
                .fixedSize(horizontal: false, vertical: true)
            }
        }
    }
}

struct InformationView_Previews: PreviewProvider {
    static var previews: some View {
        InformationView()
    }
}
