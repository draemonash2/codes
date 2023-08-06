//
//  hab_chain_widget.swift
//  hab-chain-widget
//
//  Created by Tatsuya Endo on 2023/08/04.
//

import WidgetKit
import SwiftUI

struct Provider: TimelineProvider {
    func placeholder(in context: Context) -> SimpleEntry {
        //var hab_chain_data: HabChainData = HabChainData()
        //let new_item: Item = Item()
        //hab_chain_data.addItem(new_item_id: hab_chain_data.generateItemId(), new_item: new_item)
        SimpleEntry(date: Date())
    }

    func getSnapshot(in context: Context, completion: @escaping (SimpleEntry) -> ()) {
        let entry = SimpleEntry(date: Date())
        completion(entry)
    }

    func getTimeline(in context: Context, completion: @escaping (Timeline<Entry>) -> ()) {
        var entries: [SimpleEntry] = []

        // Generate a timeline consisting of five entries an hour apart, starting from the current date.
        let currentDate = Date()
        for hourOffset in 0 ..< 5 {
            let entryDate = Calendar.current.date(byAdding: .hour, value: hourOffset, to: currentDate)!
            let entry = SimpleEntry(date: entryDate)
            entries.append(entry)
        }

        let timeline = Timeline(entries: entries, policy: .atEnd)
        completion(timeline)
    }
}

struct SimpleEntry: TimelineEntry {
    let date: Date
}

struct hab_chain_widgetEntryView : View {
    var entry: Provider.Entry
    @State var hab_chain_data: HabChainData = HabChainData()
    @AppStorage("testval", store: UserDefaults(suiteName: "group.hab_chain")) var testval: String = ""
    //@AppStorage("hab_chain_data_jsonstr") var hab_chain_data_jsonstr: String = ""

    var body: some View {
        Text(entry.date, style: .time)
            .onAppear {
                //print(hab_chain_data_jsonstr)
            }
        Text(testval)
    }
}

@main
struct hab_chain_widget: Widget {
    let kind: String = "hab_chain_widget"

    var body: some WidgetConfiguration {
        StaticConfiguration(kind: kind, provider: Provider()) { entry in
            hab_chain_widgetEntryView(entry: entry)
        }
        .configurationDisplayName("My Widget")
        .description("This is an example widget.")
    }
}

struct hab_chain_widget_Previews: PreviewProvider {
    static var previews: some View {
        hab_chain_widgetEntryView(entry: SimpleEntry(date: Date()))
            .previewContext(WidgetPreviewContext(family: .systemSmall))
    }
}
