//
//  SavePanel.swift
//  WordToExcel
//
//  Created by nan on 2025/5/18.
//

import SwiftUI
import AppKit
import Foundation

class SavePanelHelper {
    
    static func chooseSaveLocation(suggestedFileName: String = "测试用例.xlsx", completion: @escaping (URL?) -> Void) {
        let panel = NSSavePanel()
        panel.title = "保存 Excel 文件"
        panel.nameFieldStringValue = suggestedFileName
        panel.allowedFileTypes = ["xlsx"]
        panel.begin { response in
            if response == .OK {
                completion(panel.url)
            } else {
                completion(nil)
            }
        }
    }
    
}
