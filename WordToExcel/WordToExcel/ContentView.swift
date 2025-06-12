//
//  ContentView.swift
//  WordToExcel
//
//  Created by nan on 2025/5/18.
//

import SwiftUI
import AppKit
import UniformTypeIdentifiers

struct ContentView: View {
    
    @ObservedObject private var testCaseParser: WordTestCaseParser = WordTestCaseParser()
    @EnvironmentObject var settingConfig: SettingConfig
    
    @State private var dragOver = false
    @State private var selectedFileURL: URL?
    @State private var exportStatus: String = ""
    @State private var shouldShowAlert: Bool = false
    @State private var shouldShowSettingView: Bool = false

    var body: some View {
        VStack(spacing: 20) {
            Text("转换工具")
                .font(.title)
            Spacer()
            HStack() {
                Text("选择文件:")
                    .foregroundColor(.secondary)
                Spacer()
                RoundedRectangle(cornerRadius: 5)
                    .stroke(dragOver ? Color.accentColor : Color.gray, style: StrokeStyle(lineWidth: 1, dash: [2]))
                    .background(dragOver ? Color.gray.opacity(0.2) : Color.clear)
                    .frame(width: 280, height: 30)
                    .overlay(
                        Text(selectedFileURL?.pathExtension.count ?? 0 > 0 ? selectedFileURL!.lastPathComponent : "拖拽文件到当前区域或者点击选择文件")
                            .foregroundColor(.secondary)
                            .padding()
                    )
                    .onDrop(of: [.fileURL], isTargeted: $dragOver) { providers in
                        handleDrop(providers: providers)
                    }
            }
            .onTapGesture {
                chooseFile()
            }
            if selectedFileURL != nil {
                VStack () {
                    HStack () {
                        Text(testCaseParser.header.testEnvironmentName + ":")
                            .foregroundColor(.secondary)
                        Spacer()
                        HStack () {
                            Text(WordEnvironmentType.laboratoryEnvironment.rawValue)
                                .padding([.leading, .trailing], 6)
                                .padding([.top, .bottom], 3)
                                .background(testCaseParser.testEnvironmentName == .laboratoryEnvironment ? Color.blue : Color.gray)
                                .cornerRadius(5)
                                .onTapGesture {
                                    testCaseParser.testEnvironmentName = .laboratoryEnvironment
                                }
                            Text(WordEnvironmentType.systemEnvironment.rawValue)
                                .padding([.leading, .trailing], 6)
                                .padding([.top, .bottom], 3)
                                .background(testCaseParser.testEnvironmentName == .systemEnvironment ? Color.blue : Color.gray)
                                .cornerRadius(5)
                                .onTapGesture {
                                    testCaseParser.testEnvironmentName = .systemEnvironment
                                }
                        }
                    }
                    HStack () {
                        Text(testCaseParser.header.testLevel + ":")
                            .foregroundColor(.secondary)
                        Spacer()
                        HStack () {
                            Text(WordTestLevelType.configurationTesting.rawValue)
                                .padding([.leading, .trailing], 6)
                                .padding([.top, .bottom], 3)
                                .background(testCaseParser.testLevel == .configurationTesting ? Color.blue : Color.gray)
                                .cornerRadius(5)
                                .onTapGesture {
                                    testCaseParser.testLevel = .configurationTesting
                                }
                            Text(WordTestLevelType.systemTesting.rawValue)
                                .padding([.leading, .trailing], 6)
                                .padding([.top, .bottom], 3)
                                .background(testCaseParser.testLevel == .systemTesting ? Color.blue : Color.gray)
                                .cornerRadius(5)
                                .onTapGesture {
                                    testCaseParser.testLevel = .systemTesting
                                }
                        }
                    }
                    RoundedRectangle(cornerRadius: 1)
                        .frame(height: 1)
                        .foregroundColor(Color.gray.opacity(0.5))
                    VStack () {
                        HStack () {
                            Text("输出全局配置")
                                .foregroundColor(.secondary)
                            Spacer()
                        }
                        HStack () {
                            HStack () {
                                Toggle("补全跳号", isOn: $settingConfig.shouldFixMissingNumbers)
                                    .toggleStyle(SwitchToggleStyle())
                            }
                            Spacer()
                            HStack () {
                                Toggle("移除结尾特殊字符", isOn: $settingConfig.shouldDeleteEllipsis)
                                    .toggleStyle(SwitchToggleStyle())
                            }
                            Spacer()
                            HStack () {
                                Toggle("移除编号", isOn: $settingConfig.shouldDeleteNumberPrefix)
                                    .toggleStyle(SwitchToggleStyle())
                            }
                        }
                    }
                    RoundedRectangle(cornerRadius: 1)
                        .frame(height: 1)
                        .foregroundColor(Color.gray.opacity(0.5))
                }
                Button("解析数据并导出 Excel") {
                    if let fileUrl = selectedFileURL {
                        testCaseParser.extractDocxTables(from: fileUrl) { status, value  in
                            exportStatus = value
                            testCaseParser.shouldStop = false
                            if status {
                                testCaseParser.createTestModel { excelStatus, excelValue in
                                    exportStatus = excelValue
                                    if excelStatus {
                                        SavePanelHelper.chooseSaveLocation { url in
                                            if let url = url {
                                                testCaseParser.exportToExcel(url: url) { status, value in
                                                    exportStatus = value
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                .padding()
                ScrollView([.horizontal, .vertical]) {
                    Text(exportStatus)
                        .foregroundColor(.primary)
                        .fixedSize(horizontal: true, vertical: true)
                        .padding()
                }
                Spacer()
            }
            Spacer()
        }
        .padding()
        .alert(isPresented: $shouldShowAlert) {
            Alert(title: Text("提示"), message: Text("是否清理上次生成的列表数据，如不清除，数据列表将累加！"),
                  primaryButton: .destructive(Text("清除")) {
                testCaseParser.cleanData()
            },
                  secondaryButton: .cancel(Text("保留")) {
            }
            )
        }
    }
    
    private func chooseFile() {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.allowedContentTypes = [.plainText, .utf8PlainText, .item]
        if panel.runModal() == .OK {
            if let url = panel.url {
                selectedFileURL = url
                exportStatus = ""
                shouldAlertView()
            }
        }
    }

    private func handleDrop(providers: [NSItemProvider]) -> Bool {
        for provider in providers {
            if provider.hasItemConformingToTypeIdentifier(UTType.fileURL.identifier) {
                provider.loadItem(forTypeIdentifier: UTType.fileURL.identifier, options: nil) { item, error in
                    DispatchQueue.main.async {
                        if let data = item as? Data,
                           let url = URL(dataRepresentation: data, relativeTo: nil) {
                            selectedFileURL = url
                            exportStatus = ""
                            shouldAlertView()
                        }
                    }
                }
                return true
            }
        }
        return false
    }

    private func shouldAlertView() {
        if testCaseParser.testCaseItems.count > 0 {
            shouldShowAlert.toggle()
        }
    }
}
