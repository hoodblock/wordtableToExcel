//
//  Array+Extension.swift
//  WordToExcel
//
//  Created by nan on 2025/5/25.
//

import SwiftUI

extension Array {
    /// 安全访问数组元素：越界返回 nil
    subscript(safe index: Int) -> Element? {
        return indices.contains(index) ? self[index] : nil
    }
}




