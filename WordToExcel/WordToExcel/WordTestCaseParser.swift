//
//  WordTestCaseParser.swift
//  WordToExcel
//
//  Created by nan on 2025/5/19.
//

import SwiftUI
import Foundation
import ZIPFoundation
import libxlsxwriter


enum WordTestingType: String {
    case functionalTesting = "功能测试"
    case performanceTesting = "性能测试"
    case interfaceTesting = "接口测试"
    case boundaryTesting = "边界测试"
    case dataProcessingTesting = "数据处理测试"
    case humanComputerInteractionTesting = "人机交互界面测试"
    case strengthTesting = "强度测试"
    case resilienceTesting = "恢复性测试"
    case tnstallationTesting = "安装性测试"
    case safetyTesting = "安全性测试"
    case marginTesting = "余量测试"
    case capacityTesting = "容量测试"
    
    func simpleString() -> String {
        switch self {
        case .functionalTesting:
            return "GN"
        case .performanceTesting:
            return "XN"
        case .interfaceTesting:
            return "JK"
        case .boundaryTesting:
            return "BJ"
        case .dataProcessingTesting:
            return "SJ"
        case .humanComputerInteractionTesting:
            return "JM"
        case .strengthTesting:
            return "QD"
        case .resilienceTesting:
            return "HF"
        case .tnstallationTesting:
            return "AZ"
        case .safetyTesting:
            return "AQX"
        case .marginTesting:
            return "YL"
        case .capacityTesting:
            return "RL"
        }
    }
    
    static func testingType(from string: String) -> WordTestingType? {
        return WordTestingType(rawValue: string.trimmingCharacters(in: .whitespacesAndNewlines))
    }
    
    static func incloudTestingType(from string: String) -> Bool {
        return string == WordTestingType.functionalTesting.simpleString() ||
            string == WordTestingType.performanceTesting.simpleString() ||
            string == WordTestingType.interfaceTesting.simpleString() ||
            string == WordTestingType.boundaryTesting.simpleString() ||
            string == WordTestingType.dataProcessingTesting.simpleString() ||
            string == WordTestingType.humanComputerInteractionTesting.simpleString() ||
            string == WordTestingType.strengthTesting.simpleString() ||
            string == WordTestingType.resilienceTesting.simpleString() ||
            string == WordTestingType.tnstallationTesting.simpleString() ||
            string == WordTestingType.safetyTesting.simpleString() ||
            string == WordTestingType.marginTesting.simpleString() ||
            string == WordTestingType.capacityTesting.simpleString()
    }
    
}

enum WordUsecasePropertiesType: String {
    case functional = "功能性"
    case safety = "安全性"
}

enum WordEnvironmentType: String {
    case laboratoryEnvironment = "实验室环境"
    case systemEnvironment = "系统环境"
}

enum WordTestPriorityType: String {
    case high = "高"
    case medium = "中"
    case low = "低"
    case all = "高/中/低"
}

enum WordTestLevelType: String {
    case configurationTesting = "配置项测试"
    case systemTesting = "系统测试"
}

class WordItem: NSObject, ObservableObject {
    /// 用例序号 (累加)
    @Published var id: String = "0"
    /// 被测试软件版本（空）
    @Published var version: String = ""
    /// 测试用例名称
    @Published var name: String = ""
    /// 测试用例标识
    @Published var identifier: [String] = []
    /// 测试追踪(追踪到标识)
    @Published var track: String = ""
    /// 测试说明 （测试充分性要求）
    @Published var explanation: [String] = []
    /// 测试用例初始化
    @Published var useCaseInit: String = "软件正常运行"
    /// 提前与约束
    @Published var constraint: String = "无"
    /// 输入及操作说明1 (测试方法)
    @Published var inputOperation : [String] = []
    /// 期望测试说明1 （通过准则）
    @Published var expectedTest: [String] = []
    /// 评估准则1
    @Published var assessmentCriteria : String = "与预期结果一致"
    /// 实际测试结果1 （空）
    @Published var testResults: String = ""
    /// 回归测试依据(问题单号或更改单号)
    @Published var regressionTest: String = "回归测试依据"
    /// 设计人员
    @Published var resigner: String = ""
    /// 设计日期
    @Published var designDate: String = ""
    /// 执行情况（空）
    @Published var execution: String = ""
    /// 执行结果（空）
    @Published var executionResult: String = ""
    /// 问题标识（空）
    @Published var problemIdentifier: String = ""
    /// 测试人员
    @Published var tester: String = ""
    /// 测试监督员
    @Published var tesSupervisor: String = ""
    /// 测试执行日期
    @Published var testExecutionDate: String = ""
    /// 测试类型
    @Published var testType: [String] = []
    /// 测试需求描述(测试项名称)
    @Published var testRequirementDescription: String = ""
    /// 测试需求状态
    @Published var testRequirementStatus: String = "原始"
    /// 测试用例状态
    @Published var testCaseStatus: String = "原始"
    /// 用例属性
    var usecaseProperties: [WordUsecasePropertiesType] {
        get {
            var tempUsecaseProperties: [WordUsecasePropertiesType] = []
            for item in testType {
                if item == WordTestingType.functionalTesting.simpleString() || item == WordTestingType.boundaryTesting.simpleString() {
                    tempUsecaseProperties.append(.functional)
                } else {
                    tempUsecaseProperties.append(.safety)
                }
            }
            return tempUsecaseProperties
        }
    }
    /// 测试环境名称
    @Published var testEnvironmentName: WordEnvironmentType = .laboratoryEnvironment
    /// 测试结果判别人
    @Published var testResultsJudgeOthers: String = ""
    /// 测试优先级
    @Published var testPriority: WordTestPriorityType = .all
    /// 软件名称
    @Published var softwareName: String = ""
    /// 测试级别
    @Published var testLevel: WordTestLevelType = .configurationTesting
    
    init(id: String, identifier: [String], explanation: [String], inputOperation: [String], expectedTest: [String], testType: [String]) {
        self.id = id
        self.identifier = identifier
        self.explanation = explanation
        self.inputOperation = inputOperation
        self.expectedTest = expectedTest
        self.testType = testType
    }
    
    init(id: String) {
        self.id = id
    }
    
    var wordItemValues: [String] {
           [
               id,
               version,
               name,
               identifier[safe: 0] ?? "",
               track,
               explanation[safe: 0] ?? "",
               useCaseInit,
               constraint,
               inputOperation[safe: 0] ?? "",
               expectedTest[safe: 0] ?? "",
               assessmentCriteria,
               testResults,
               regressionTest,
               resigner,
               designDate,
               execution,
               executionResult,
               problemIdentifier,
               tester,
               tesSupervisor,
               testExecutionDate,
               testType[safe: 0] ?? "",
               testRequirementDescription,
               testRequirementStatus,
               testCaseStatus,
               usecaseProperties[safe: 0]?.rawValue ?? "",
               testEnvironmentName.rawValue,
               testResultsJudgeOthers,
               testPriority.rawValue,
               softwareName,
               testLevel.rawValue
           ]
       }
    
}

class WordTestCaseParser: NSObject, ObservableObject {
        
    private var settingConfig: SettingConfig = SettingConfig.shared

    @Published var header = ExcelHeader()
    @Published var sheetTitles = ["Sheet"]
    @Published var testCaseItems: [WordItem] = []
    @Published var wordTables: [[[String]]] = []
    @Published var shouldStop: Bool = false

    // 配置全局表格
    @Published var testEnvironmentName: WordEnvironmentType = .laboratoryEnvironment
    @Published var testLevel: WordTestLevelType = .configurationTesting

    override init() {
        super.init()
    }
    
    func cleanData() {
        testCaseItems = []
        testEnvironmentName = .laboratoryEnvironment
        testLevel = .configurationTesting
    }
}

// MARK: - 解析word数据，生成元数据
extension WordTestCaseParser {
    
    func extractDocxTables(from fileUrl: URL, completion: @escaping (Bool, String) -> Void) {
        let tempDir = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        do {
            try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
            try FileManager.default.unzipItem(at: fileUrl, to: tempDir)
            
            let documentXML = tempDir.appendingPathComponent("word/document.xml")
            let xmlData = try Data(contentsOf: documentXML)
            
            let parser = WordTableParser()
            parser.parse(data: xmlData)
            if parser.tables.count > 0 {
                wordTables = parser.tables
                completion(true, "✅ 解析完成，有表格存在，正在生成数据模型...")
            } else {
                completion(false, "⚠️ 没有解析到表格数据")
            }
            try FileManager.default.removeItem(at: tempDir)
        } catch {
            completion(false, "❌ 解压或解析失败: \(error)")
        }
    }
}

// MARK: - 根据元数据生成模型数据
extension WordTestCaseParser {
    
    func createTestModel(completion: @escaping (Bool, String) -> Void) {
        for (_, table) in wordTables.enumerated() {
            if shouldStop { return }
            var testItem: WordItem
            if testCaseItems.count > 0 {
                testItem = WordItem(id: String(Int(testCaseItems.last!.id)! + 1))
            } else {
                testItem = WordItem(id: "0")
            }
            for row in table {
                if shouldStop { return }
                for (itemIndex, item) in row.enumerated() {
                    if shouldStop { return }
                    if item == "测试项名称" {
                        testItem.name = row[itemIndex + 1] + "-"
                        testItem.testRequirementDescription = row[itemIndex + 1]
                    } else if item == "测试项标识" {
                        testItem.track = row[itemIndex + 1]
                    } else if item == "追踪关系" {
                        
                    } else if item == "测试项描述" {
                        
                    } else if item == "优先级" || item == "测试优先级" {
                        if row[itemIndex + 1] == "高" {
                            testItem.testPriority = .high
                        } else if row[itemIndex + 1] == "中" {
                            testItem.testPriority = .medium
                        } else if row[itemIndex + 1] == "低" {
                            testItem.testPriority = .low
                        }
                    } else if item == "测试类型" {
                        // 空
                    } else if let type = WordTestingType.testingType(from: item) {
                        // key: 功能测试 value:[测试充分性要求 [1, 2, 3], 测试方法[1, 2, 3], 通过准则[1, 2, 3]]
                        parseStructuredText(row[itemIndex + 1], shouldFunctionKey: false) {[weak self] status, values, error in
                            if !status {
                                completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                return
                            }
                            for itemValue in values {
                                if itemValue.key == "测试充分性要求" {
                                    for (valueIndex, itemKeyValue) in itemValue.values.enumerated() {
                                        testItem.testType.append(type.simpleString())
                                        testItem.identifier.append(testItem.track + "-" + type.simpleString() + "-" + (self?.formatToThreeDigits(valueIndex + 1) ?? ""))
                                        testItem.explanation.append(itemKeyValue)
                                    }
                                } else if itemValue.key == "测试方法" {
                                    testItem.inputOperation.append(contentsOf: itemValue.values)
                                } else if itemValue.key == "通过准则" {
                                    testItem.expectedTest.append(contentsOf: itemValue.values)
                                }
                            }
                        }
                    } else if item == "测试充分性要求" || item == "充分性要求" || item == "测试充分性分析与终止条件"{
                        // 两种情况
                        parseStructuredText(row[itemIndex + 1], shouldFunctionKey: true) {[weak self] status, parseTestValues, error in
                            if !status {
                                completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                return
                            }
                            if parseTestValues.count > 0 {
                                // key: 测试充分性要求 value:[功能测试 [1, 2, 3], 边界测试[1, 2, 3]]
                                for (_, itemKeyValue) in parseTestValues.enumerated() {
                                    if let type = WordTestingType.testingType(from: itemKeyValue.key) {
                                        for (testValueIndex, testItemsValues) in itemKeyValue.values.enumerated() {
                                            testItem.testType.append(type.simpleString())
                                            testItem.identifier.append(testItem.track + "-" + type.simpleString() + "-" + (self?.formatToThreeDigits(testValueIndex + 1) ?? ""))
                                            testItem.explanation.append(testItemsValues)
                                        }
                                    }
                                }
                            } else {
                                // key: 测试充分性要求 value:[1, 2, 3]
                                self?.parseNumberedEntries(row[itemIndex + 1]) {[weak self] status, values, error in
                                    if !status {
                                        completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                        return
                                    }
                                    for (valueIndex, itemKeyValue) in values.enumerated() {
                                        // 这个有可能是单一测试功能表格，提取测试项标识的前缀部分
                                        let componentsString = testItem.track.components(separatedBy: "-")
                                        if WordTestingType.incloudTestingType(from: componentsString[safe: 0] ?? "") {
                                            testItem.testType.append(componentsString[safe: 0]!)
                                            testItem.identifier.append(testItem.track + "-" + componentsString[safe: 0]! + "-" + (self?.formatToThreeDigits(valueIndex + 1) ?? ""))
                                        }
                                        testItem.explanation.append(itemKeyValue)
                                    }
                                }
                            }
                        }
                    } else if item == "测试方法" {
                        // 两种情况
                       parseStructuredText(row[itemIndex + 1], shouldFunctionKey: true) {[weak self] status, parseTestValues, error in
                           if !status {
                               completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                               return
                           }
                           if parseTestValues.count > 0 {
                               // key: 测试方法 value:[功能测试 [1, 2, 3], 边界测试[1, 2, 3]]
                               let flattened = parseTestValues.flatMap { $0.values }
                               testItem.inputOperation.append(contentsOf: flattened)
                           } else {
                               // key: 测试方法 value:[1, 2, 3]
                               // 这里要适配第二种样式， xxx, 查看xxx， 遍历所有的item，不为空的情况下，都包含查看
                               self?.parseNumberedEntries(row[itemIndex + 1]) {[weak self] status, values, error in
                                   if !status {
                                       completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                       return
                                   }
                                   let keywordItems = self?.parseKeywordItems(values)
                                   if keywordItems?.count ?? 0 > 0 {
                                       let keys = keywordItems!.map { $0.first?.key ?? "" }
                                       let values = keywordItems!.map { $0.first?.value ?? "" }
                                       let specialCharactersPattern = #"[\.,。;；,，]+$"#
                                       let keysCleaned = keys.map { line in
                                           line.replacingOccurrences(of: specialCharactersPattern, with: "", options: .regularExpression)
                                       }
                                       let valuesCleaned = values.map { line in
                                           line.replacingOccurrences(of: specialCharactersPattern, with: "", options: .regularExpression)
                                       }
                                       testItem.inputOperation.append(contentsOf: keysCleaned)
                                       testItem.expectedTest.append(contentsOf: valuesCleaned)
                                       for (valueIndex, _) in keysCleaned.enumerated() {
                                           let componentsString = testItem.track.components(separatedBy: "-")
                                           if WordTestingType.incloudTestingType(from: componentsString[safe: 0] ?? "") {
                                               testItem.testType.append(componentsString[safe: 0]!)
                                               testItem.identifier.append(testItem.track + "-" + componentsString[safe: 0]! + "-" + (self?.formatToThreeDigits(valueIndex + 1) ?? ""))
                                           }
                                           testItem.explanation.append("")
                                       }
                                   } else {
                                       testItem.inputOperation.append(contentsOf: values)
                                   }
                               }
                           }
                        }
                    }  else if item == "通过准则" {
                        // 两种情况
                        parseStructuredText(row[itemIndex + 1], shouldFunctionKey: true) {[weak self] status, parseTestValues, error in
                            if !status {
                                completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                return
                            }
                            if parseTestValues.count > 0 {
                                // key: 通过准则 value:[功能测试 [1, 2, 3], 边界测试[1, 2, 3]]
                                // TODO: - 这种貌似有没这个格式，要写死一种准则标砖
                                let flattened = parseTestValues.flatMap { $0.values }
                                testItem.expectedTest.append(contentsOf: flattened)
                            } else {
                                // key: 通过准则 value:[1, 2, 3]
                                self?.parseNumberedEntries(row[itemIndex + 1]) { status, values, error in
                                    if !status {
                                        completion(status, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中 \(error)")
                                        return
                                    }
                                    testItem.expectedTest.append(contentsOf: values)
                                }
                            }
                        }
                    }
                }
            }
            if shouldStop { break }
            if !testItem.name.isEmpty && !testItem.track.isEmpty {
                let checkCounts = [testItem.explanation.count, testItem.expectedTest.count, testItem.inputOperation.count]
                let noZeroCounts = checkCounts.filter { $0 > 0 }
                if !(Set(noZeroCounts).count <= 1) {
                    shouldStop = true
                    completion(false, "❌ 解析到在【 名称 = \(testItem.name), 标识 = \(testItem.track) 】表中的【 测试充分性要求 = \(testItem.explanation.count)条 】，【 测试方法 = \(testItem.expectedTest.count)条 】，【 通过准则 = \(testItem.inputOperation.count)条 】对应关系无法对应")
                    return
                } else {
                    if !settingConfig.shouldFixMissingNumbers {
                        var indicesToRemove: [Int] = []
                        for index in 0..<testItem.explanation.count {
                            var explanation: String = ""
                            var expectedTest: String = ""
                            var inputOperation: String = ""
                            if testItem.explanation.count > index {
                                explanation = testItem.explanation[index].trimmingCharacters(in: .whitespacesAndNewlines)
                            }
                            if testItem.expectedTest.count > index {
                                expectedTest = testItem.expectedTest[index].trimmingCharacters(in: .whitespacesAndNewlines)
                            }
                            if testItem.inputOperation.count > index {
                                inputOperation = testItem.inputOperation[index].trimmingCharacters(in: .whitespacesAndNewlines)
                            }
                            if explanation.isEmpty && expectedTest.isEmpty && inputOperation.isEmpty {
                                indicesToRemove.append(index)
                            }
                        }
                        for index in indicesToRemove.reversed() {
                            if testItem.explanation.count > index {
                                testItem.explanation.remove(at: index)
                            }
                            if testItem.expectedTest.count > index {
                                testItem.expectedTest.remove(at: index)
                            }
                            if testItem.inputOperation.count > index {
                                testItem.inputOperation.remove(at: index)
                            }
                            if testItem.testType.count > index {
                                testItem.testType.remove(at: index)
                            }
                            if testItem.identifier.count > index {
                                testItem.identifier.remove(at: index)
                            }
                        }
                    }
                }
                testItem.testLevel = testLevel
                testItem.testEnvironmentName = testEnvironmentName
                testCaseItems.append(testItem)
            }
        }
        if testCaseItems.count > 0 {
            completion(true, "✅ 解析完成，生成符合条件的Excel数据模型")
        } else {
            completion(false, "⚠️ 没有筛选出符合条件的Word数据表")
        }
    }
    
    /// 数字前面补0
    func formatToThreeDigits(_ number: Int) -> String {
        return String(format: "%03d", number)
    }
    
    /// 样例二，字符串以查看分割
    func parseKeywordItems(_ items: [String]) -> [[String: String]] {
        let keyword = "查看"
        for item in items {
            if !item.isEmpty && !item.contains(keyword) {
                return []
            }
        }
        return items.map { item in
            if item.isEmpty {
                return [:]
            }
            let parts = item.components(separatedBy: keyword)
            if parts.count == 2 {
                return [parts[0]: parts[1]]
            } else {
                return [:]
            }
        }
    }
    
    /// 这种数据中包含"测试充分性要求", "测试方法", "通过准则"的全部内容 -> keys = ["测试充分性要求", "测试方法", "通过准则"]
    /// 以测试类型分为模块 如 key: 功能测试  value  [内容一；内容二；..]."
    func parseStructuredText(_ input: String, shouldFunctionKey: Bool, completion: @escaping (Bool, [(key: String, values: [String])], String) -> Void) {
        var keys: [String] = []
        if shouldFunctionKey {
            keys = [
                        WordTestingType.functionalTesting.rawValue,
                        WordTestingType.performanceTesting.rawValue,
                        WordTestingType.interfaceTesting.rawValue,
                        WordTestingType.boundaryTesting.rawValue,
                        WordTestingType.dataProcessingTesting.rawValue,
                        WordTestingType.humanComputerInteractionTesting.rawValue,
                        WordTestingType.strengthTesting.rawValue,
                        WordTestingType.resilienceTesting.rawValue,
                        WordTestingType.tnstallationTesting.rawValue,
                        WordTestingType.safetyTesting.rawValue,
                        WordTestingType.marginTesting.rawValue,
                        WordTestingType.capacityTesting.rawValue,
            ]
        } else {
            keys = ["测试充分性要求", "测试方法", "通过准则"]
        }
        var result: [(key: String, values: [String])] = []
        let keyPattern = keys.joined(separator: "|")
        let sectionPattern = "(?:(\(keyPattern))：)(.*?)(?=(\(keyPattern)：)|$)"
        let sectionRegex = try! NSRegularExpression(pattern: sectionPattern, options: [.dotMatchesLineSeparators])
        let cleanedInput = input
            .replacingOccurrences(of: "↵", with: "\n")
            .replacingOccurrences(of: "\r", with: "")
        let nsrange = NSRange(cleanedInput.startIndex..<cleanedInput.endIndex, in: cleanedInput)
        let matches = sectionRegex.matches(in: cleanedInput, options: [], range: nsrange)
        for match in matches {
            guard
                let keyRange = Range(match.range(at: 1), in: cleanedInput),
                let valueRange = Range(match.range(at: 2), in: cleanedInput)
            else { continue }

            let key = String(cleanedInput[keyRange])
            let rawValue = String(cleanedInput[valueRange])
            let parsingEntries = parsingInputString(input: rawValue)
            processingParsingText(parsing: parsingEntries) { status, resultItems, errorValue in
                if status {
                    result.append((key: key, values: resultItems))
                } else {
                    let block: [(key: String, values: [String])] = []
                    completion(status, block, errorValue)
                    return
                }
            }
        }
        completion(true, result, "")
    }

    func parseNumberedEntries(_ input: String, completion: @escaping (Bool, [String], String) -> Void) {
        let parsingEntries = parsingInputString(input: input)
        processingParsingText(parsing: parsingEntries) { status, result, errorValue in
            completion(status, result, errorValue)
        }
    }
    
    /// 根据文本以数字分割，解析成数组， 带编号
    func parsingInputString(input: String) -> [String] {
        let cleanedInput = input
            .replacingOccurrences(of: "↵", with: "\n")
            .replacingOccurrences(of: "\r", with: "")
        let entryPattern = #"(?<!\d)(?=\d+(?:\.\d+)*[\)\）\.\．:：,，])"#
        let entryRegex = try! NSRegularExpression(pattern: entryPattern)
        let matches = entryRegex.matches(in: cleanedInput, options: [], range: NSRange(cleanedInput.startIndex..., in: cleanedInput))
        let entries = matches.enumerated().compactMap { (i, match) -> String? in
            guard let start = Range(match.range, in: cleanedInput) else { return nil }
            let startIndex = start.lowerBound
            let endIndex: String.Index = (i + 1 < matches.count)
                ? cleanedInput.index(cleanedInput.startIndex, offsetBy: matches[i + 1].range.location)
                : cleanedInput.endIndex
            let substring = cleanedInput[startIndex..<endIndex].trimmingCharacters(in: .whitespacesAndNewlines)
            return substring.isEmpty ? nil : String(substring)
        }
        return entries
    }
    
    /// 把解析的数据进一步成立成想要的数据
    func processingParsingText(parsing texts: [String], completion: @escaping (Bool, [String], String) -> Void) {
        var cleanResult: [String] = []
        // 第一层清理 去掉【x) 测试内容】的数据
        let pattern = #"^\d+(?:\.\d+)*[\)\）\.\．:：,，、]\s*测试内容$"#
        cleanResult = texts.filter { item in
            return item.range(of: pattern, options: .regularExpression) == nil
        }
        // 判断组合 2）、5）或者2）~ 5）补全
        cleanResult = expandTextRanges(in: cleanResult)
        // 判断编号是否是按顺序的，如果不按顺序，则报错处理
        let result = checkAndFixSequence(cleanResult)
        if result.error.count > 0 {
            completion(false, [], result.error + "\(texts.joined(separator: "\n"))")
            shouldStop = true
            return
        } else {
            cleanResult = result.result
        }
        if settingConfig.shouldDeleteEllipsis {
            let ellipsisPattern = #"(?:[\)\)）:：、.．，,；;]|\s*)\.{3}$"#
            cleanResult = cleanResult.map { line in
                line.replacingOccurrences(of: ellipsisPattern, with: "", options: .regularExpression)
            }
            let specialCharactersPattern = #"[\.,。;；,，]+$"#
            cleanResult = cleanResult.map { line in
                line.replacingOccurrences(of: specialCharactersPattern, with: "", options: .regularExpression)
            }
        }
        if settingConfig.shouldDeleteNumberPrefix {
            let prefixPattern = #"^\s*\d+[\)\)）:：、.．]+[\s]*"#
            cleanResult = cleanResult.map { line in
                line.replacingOccurrences(of: prefixPattern, with: "", options: .regularExpression).trimmingCharacters(in: .whitespaces)
            }
        }
        completion(true, cleanResult, "")
    }

    /// 解析数字组合，分裂成数组
    func expandTextRanges(in items: [String]) -> [String] {
        var result: [String] = []
        var i = 0
        // 提取编号数字
        func extractNumber(from str: String) -> Int? {
            let pattern = #"^(\d+)[\)\）]"#
            guard let regex = try? NSRegularExpression(pattern: pattern),
                  let match = regex.firstMatch(in: str, range: NSRange(str.startIndex..., in: str)),
                  let range = Range(match.range(at: 1), in: str)
            else { return nil }
            return Int(str[range])
        }
        // 提取内容（跳过编号和符号）
        func extractContent(from str: String) -> String {
            guard let idx = str.firstIndex(where: { !"0123456789)）~、".contains($0) }) else {
                return ""
            }
            return str[idx...].trimmingCharacters(in: .whitespacesAndNewlines)
        }
        // 识别 "2）~5）内容" 并展开
        func expandWaveRange(_ str: String) -> [String]? {
            let pattern = #"^(\d+)[\)\）]\s*~\s*(\d+)[\)\）](.*)"#
            guard let regex = try? NSRegularExpression(pattern: pattern),
                  let match = regex.firstMatch(in: str, range: NSRange(str.startIndex..., in: str)),
                  match.numberOfRanges == 4,
                  let startRange = Range(match.range(at: 1), in: str),
                  let endRange = Range(match.range(at: 2), in: str),
                  let contentRange = Range(match.range(at: 3), in: str)
            else {
                return nil
            }
            let start = Int(str[startRange]) ?? 0
            let end = Int(str[endRange]) ?? 0
            let content = str[contentRange].trimmingCharacters(in: .whitespacesAndNewlines)
            guard start <= end else { return nil }

            return (start...end).map { "\($0)）" + content }
        }

        while i < items.count {
            let current = items[i]
            if let expanded = expandWaveRange(current) {
                result.append(contentsOf: expanded)
                i += 1
                continue
            }
            // 区间起点（匹配 "2）、" 或 "2）~"）
            if current.range(of: #"^\d+[\)\）][~、]$"#, options: .regularExpression) != nil {
                guard let startNum = extractNumber(from: current) else {
                    result.append(current)
                    i += 1
                    continue
                }
                var j = i + 1
                var endNum: Int? = nil
                var content: String? = nil
                while j < items.count {
                    let candidate = items[j]
                    if let num = extractNumber(from: candidate),
                       !candidate.hasSuffix("、") && !candidate.hasSuffix("~") {
                        endNum = num
                        content = extractContent(from: candidate)
                        break
                    }
                    j += 1
                }
                if let end = endNum, let text = content {
                    for n in startNum...end {
                        result.append("\(n)）" + text)
                    }
                    i = j + 1
                } else {
                    result.append(current)
                    i += 1
                }
            } else {
                result.append(current)
                i += 1
            }
        }
        return result
    }

    /// 中间某个item以...结尾，并且下一个编号为跳号
    func fixMissingNumbers(from items: [String]) -> [String] {
        var result: [String] = []
        func extractNumber(from str: String) -> Int? {
            let pattern = #"^(\d+)[\)）]"#
            guard let regex = try? NSRegularExpression(pattern: pattern),
                  let match = regex.firstMatch(in: str, range: NSRange(str.startIndex..., in: str)),
                  let range = Range(match.range(at: 1), in: str)
            else { return nil }
            return Int(str[range])
        }
        for i in 0..<items.count {
            let current = items[i]
            result.append(current)
            guard current.hasSuffix("..."),
                  let currentNum = extractNumber(from: current),
                  i + 1 < items.count,
                  let nextNum = extractNumber(from: items[i + 1]),
                  nextNum > currentNum + 1
            else {
                continue
            }
            for missing in (currentNum + 1)..<nextNum {
                result.append("\(missing)）")
            }
        }
        return result
    }
    
    /// 检查数组是否是严格递增，如果遇到跳号，补全跳号 例如[1, 2, 3, 6] 严格递增，并且会补全中间缺失的数据 [4, 5]
    func checkAndFixSequence(_ input: [String]) -> (error: String, result: [String]) {
        let prefixPattern = #"^(\d+)[\)\.\：:、．，]?"#
        let regex = try! NSRegularExpression(pattern: prefixPattern)
        // 提取 (编号, 原始项)
        let numberedItems: [(number: Int, content: String)] = input.compactMap { item in
            guard let match = regex.firstMatch(in: item, range: NSRange(item.startIndex..., in: item)),
                  let range = Range(match.range(at: 1), in: item),
                  let number = Int(item[range]) else {
                return nil
            }
            return (number, item)
        }
        let numbers = numberedItems.map { $0.number }
        if numbers.count > 0 {
            if numbers[0] < 1 {
                return ("列表编号数据顺序不符合规范，编号必须从1开始\n", [])
            }
        }
        let isStrictlyIncreasing = zip(numbers, numbers.dropFirst()).allSatisfy { $0 < $1 }
        if !isStrictlyIncreasing {
            return ("列表编号数据不是严格递增\n", [])
        } else {
            guard let maxNumber = numbers.max() else {
                return ("", input)
            }
            let contentMap = Dictionary(uniqueKeysWithValues: numberedItems.map { ($0.number, $0.content) })
            let result = (1...maxNumber).map { contentMap[$0] ?? "" }
            return ("", result)
        }
    }


}

// MARK: - 使用数据模型生成表格
extension WordTestCaseParser {
    
    func exportToExcel(url: URL, completion: @escaping (Bool, String) -> Void) {
        guard let workbook = workbook_new(url.path) else {
            completion(false, "❌ 无法创建 Excel 文件")
            return
        }
        let boldFormat = workbook_add_format(workbook)
        format_set_bold(boldFormat)
        let normalFormat = workbook_add_format(workbook)
        
        let blueFillFormat = workbook_add_format(workbook)
        format_set_bg_color(blueFillFormat, LXW_COLOR_BLUE.rawValue)
        let emptyFillFormat = workbook_add_format(workbook)
        format_set_bg_color(emptyFillFormat, LXW_COLOR_RED.rawValue)

        for sheetName in sheetTitles {
            let worksheet = workbook_add_worksheet(workbook, sheetName)
            for (col, title) in header.headers.enumerated() {
                worksheet_write_string(worksheet, 0, lxw_col_t(col), title, boldFormat)
                // 所有列默认设置为25宽
                if title == header.explanation || title == header.inputOperation || title == header.expectedTest {
                    worksheet_set_column(worksheet, lxw_col_t(col), lxw_col_t(col), 250, nil)
                } else {
                    worksheet_set_column(worksheet, lxw_col_t(col), lxw_col_t(col), 25, nil)
                }
            }
            var row: Int32 = 2
            for tc in testCaseItems {
                // 认为是多功能测试
                if tc.identifier.count > 0 {
                    for index in 0..<tc.identifier.count {
                        if index == 0 {
                            for column in 0..<header.headers.count {
                                worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.wordItemValues[safe: column], normalFormat)
                            }
                            // 不对应的空值用红色标记
                            if (tc.explanation[safe: index]?.count ?? 0 > 0) && (tc.inputOperation[safe: index]?.count ?? 0 > 0) && (tc.expectedTest[safe: index]?.count ?? 0 > 0) {
                                if let column = header.indexOfHeader(header.explanation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.explanation[safe: index], normalFormat)
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.inputOperation[safe: index], normalFormat)
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.expectedTest[safe: index], normalFormat)
                                }
                            } else if (tc.explanation[safe: index]?.count ?? 0 == 0) && (tc.inputOperation[safe: index]?.count ?? 0 == 0) && (tc.expectedTest[safe: index]?.count ?? 0 == 0) {
                                if let column = header.indexOfHeader(header.explanation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                            } else {
                                if let column = header.indexOfHeader(header.explanation) {
                                    if tc.explanation[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.explanation[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    if tc.inputOperation[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.inputOperation[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    if tc.expectedTest[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.expectedTest[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                            }
                        } else {
                            if let column = header.indexOfHeader(header.identifier) {
                                worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.identifier[safe: index], normalFormat)
                            }
                            // 不对应的空值用红色标记
                            if (tc.explanation[safe: index]?.count ?? 0 > 0) && (tc.inputOperation[safe: index]?.count ?? 0 > 0) && (tc.expectedTest[safe: index]?.count ?? 0 > 0) {
                                if let column = header.indexOfHeader(header.explanation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.explanation[safe: index], normalFormat)
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.inputOperation[safe: index], normalFormat)
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.expectedTest[safe: index], normalFormat)
                                }
                            } else if (tc.explanation[safe: index]?.count ?? 0 == 0) && (tc.inputOperation[safe: index]?.count ?? 0 == 0) && (tc.expectedTest[safe: index]?.count ?? 0 == 0) {
                                if let column = header.indexOfHeader(header.explanation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", normalFormat)
                                }
                            } else {
                                if let column = header.indexOfHeader(header.explanation) {
                                    if tc.explanation[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.explanation[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                                if let column = header.indexOfHeader(header.inputOperation) {
                                    if tc.inputOperation[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.inputOperation[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                                if let column = header.indexOfHeader(header.expectedTest) {
                                    if tc.expectedTest[safe: index]?.count ?? 0 > 0 {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.expectedTest[safe: index], normalFormat)
                                    } else {
                                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), "", emptyFillFormat)
                                    }
                                }
                            }
                            if let column = header.indexOfHeader(header.testType) {
                                worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.testType[safe: index], normalFormat)
                            }
                            if let column = header.indexOfHeader(header.usecaseProperties) {
                                worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.usecaseProperties[safe: index]?.rawValue, normalFormat)
                            }
                        }
                        row += 1
                    }
                } else {
                    // 认为是单一功能测试
                    for column in 0..<header.headers.count {
                        print("_____【 WordTestCaseParser 】_____【 exportToExcel 】_____【 row = \(row) 】_____【 column = \(column) 】_____【 value = \(String(describing: tc.wordItemValues[safe: column])) 】")
                        worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(column), tc.wordItemValues[safe: column], normalFormat)
                    }
                }
                row += 1
                for col in 0..<header.headers.count {
                    worksheet_write_string(worksheet, lxw_row_t(row), lxw_col_t(col), "", blueFillFormat)
                }
                row += 2
            }
        }
        workbook_close(workbook)
        completion(true, "✅ Excel 文件已导出：\(url.path)")
    }
    
    func fillColorForEmptyData() {
        
    }
}

// MARK: - 扩展
extension NSRegularExpression {
    func split(text: String) -> [String] {
        let nsrange = NSRange(text.startIndex..<text.endIndex, in: text)
        let matches = self.matches(in: text, options: [], range: nsrange)

        var results: [String] = []
        var lastIndex = text.startIndex

        for match in matches {
            let range = Range(match.range, in: text)!
            if lastIndex < range.lowerBound {
                let segment = String(text[lastIndex..<range.lowerBound])
                results.append(segment)
            }
            lastIndex = range.lowerBound
        }

        let tail = String(text[lastIndex...])
        if !tail.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
            results.append(tail)
        }

        return results
    }
}
