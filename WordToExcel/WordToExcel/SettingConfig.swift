//
//  SettingConfig.swift
//  WordToExcel
//
//  Created by nan on 2025/6/5.
//

import SwiftUI

// 配置全局表格数据样式

class SettingConfig: NSObject, ObservableObject {
        
    static var shared = SettingConfig()

    /// 是否补全...跳号的编号列表
    @Published var shouldFixMissingNumbers: Bool = false
    /// 是否去掉结尾的特殊字符或者结尾的...
    @Published var shouldDeleteEllipsis: Bool = false
    /// 是否去掉编号前缀
    @Published var shouldDeleteNumberPrefix: Bool = false

    override init() {
        super.init()
    }

}
