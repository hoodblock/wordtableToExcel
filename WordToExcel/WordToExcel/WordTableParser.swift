//
//  WordTableParser.swift
//  WordToExcel
//
//  Created by nan on 2025/5/19.
//

import Foundation
import ZIPFoundation

/// 纯 Swift 解析 .docx，核心是：
/// 解压 .docx 为 XML 文件
/// 解析 XML，提取 <w:tbl> 里的表格数据
/// 推荐用 ZIPFoundation 做解压
/// 用 XMLParser 做 XML 解析
class WordTableParser: NSObject, XMLParserDelegate {
    
    var tables: [[[String]]] = []
    private var currentTable: [[String]] = []
    private var currentRow: [String] = []
    private var currentCellText = ""
    
    private var insideTable = false
    private var insideCell = false
    private var capturingText = false

    func parse(data: Data) {
        let parser = XMLParser(data: data)
        parser.delegate = self
        parser.parse()
    }

    // MARK: - XMLParserDelegate
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        switch elementName {
        case "w:tbl":
            insideTable = true
            currentTable = []
        case "w:tr":
            currentRow = []
        case "w:tc":
            insideCell = true
            currentCellText = ""
        case "w:t":
            capturingText = true
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if insideCell && capturingText {
            currentCellText += string
        }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        switch elementName {
        case "w:t":
            capturingText = false
        case "w:tc":
            insideCell = false
            currentRow.append(currentCellText.trimmingCharacters(in: .whitespacesAndNewlines))
        case "w:tr":
            currentTable.append(currentRow)
        case "w:tbl":
            insideTable = false
            tables.append(currentTable)
        default:
            break
        }
    }
}
