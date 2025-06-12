//
//  ExcelHeader.swift
//  WordToExcel
//
//  Created by nan on 2025/5/19.
//

import SwiftUI
import Foundation


class ExcelHeader: NSObject, ObservableObject {
    /// 用例序号 (累加)
    @Published var id: String = "用例序号"
    /// 被测试软件版本（空）
    @Published var version: String = "被测试软件版本"
    /// 测试用例名称
    @Published var name: String = "测试用例名称"
    /// 测试用例标识
    @Published var identifier: String = "测试用例标识"
    /// 测试追踪(追踪到标识)
    @Published var track: String = "测试追踪(追踪到标识)"
    /// 测试说明
    @Published var explanation: String = "测试说明"
    /// 测试用例初始化
    @Published var useCaseInit: String = "测试用例初始化"
    /// 提前与约束
    @Published var constraint: String = "提前与约束"
    /// 输入及操作说明1
    @Published var inputOperation : String = "输入及操作说明1"
    /// 期望测试说明1
    @Published var expectedTest: String = "期望测试结果1"
    /// 评估准则1
    @Published var assessmentCriteria : String = "评估准则1"
    /// 实际测试结果1 （空）
    @Published var testResults: String = "实际测试结果1"
    /// 回归测试依据(问题单号或更改单号)
    @Published var regressionTest: String = "回归测试依据(问题单号或更改单号)"
    /// 设计人员
    @Published var resigner: String = "设计人员"
    /// 设计日期
    @Published var designDate: String = "设计日期"
    /// 执行情况（空）
    @Published var execution: String = "执行情况"
    /// 执行结果（空）
    @Published var executionResult: String = "执行结果"
    /// 问题标识（空）
    @Published var problemIdentifier: String = "问题标识"
    /// 测试人员
    @Published var tester: String = "测试人员"
    /// 测试监督员
    @Published var tesSupervisor: String = "测试监督员"
    /// 测试执行日期
    @Published var testExecutionDate: String = "测试执行日期"
    /// 测试类型
    @Published var testType: String = "测试类型"
    /// 测试需求描述(测试项名称)
    @Published var testRequirementDescription: String = "测试需求描述(测试项名称)"
    /// 测试需求状态
    @Published var testRequirementStatus: String = "测试需求状态"
    /// 测试用例状态
    @Published var testCaseStatus: String = "测试用例状态"
    /// 用例属性
    @Published var usecaseProperties: String = "用例属性"
    /// 测试环境名称
    @Published var testEnvironmentName: String = "测试环境名称"
    /// 测试结果判别人
    @Published var testResultsJudgeOthers: String = "测试结果判别人"
    /// 测试优先级
    @Published var testPriority: String = "测试优先级"
    /// 软件名称
    @Published var softwareName: String = "软件名称"
    /// 测试级别
    @Published var testLevel: String = "测试级别"

    var headers: [String] {
           [
               id,
               version,
               name,
               identifier,
               track,
               explanation,
               useCaseInit,
               constraint,
               inputOperation,
               expectedTest,
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
               testType,
               testRequirementDescription,
               testRequirementStatus,
               testCaseStatus,
               usecaseProperties,
               testEnvironmentName,
               testResultsJudgeOthers,
               testPriority,
               softwareName,
               testLevel
           ]
       }
    
    func indexOfHeader(_ field: String) -> Int? {
        return headers.firstIndex(of: field)
    }
}




