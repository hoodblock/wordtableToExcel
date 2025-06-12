//
//  WordToExcelApp.swift
//  WordToExcel
//
//  Created by nan on 2025/5/18.
//

import SwiftUI

@main
struct WordToExcelApp: App {
    let persistenceController = PersistenceController.shared

    var body: some Scene {
        WindowGroup {
            ContentView()
                .frame(width: 500, height: 500)
                .background(.gray.opacity(0.05))
                .cornerRadius(10)
                .environment(\.managedObjectContext, persistenceController.container.viewContext)
                .environmentObject(SettingConfig.shared)
        }
    }
}
