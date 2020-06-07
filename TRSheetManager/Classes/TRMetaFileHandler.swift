/*

TRMetaFileHandler.m

Copyright (c) 2020 Truong Vinh Tran

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/

import Foundation

public struct TRSheetConst {
    static let header = "<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\nxmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n\n"
    
    static let documentProperties = "DocumentProperties"
    static let documentPropertiesConfig = "xmlns=\"urn:schemas-microsoft-com:office:office\""
    static let workbook = "Workbook"
    static let author = "Author"
    static let created = "Created"
    static let worksheet = "Worksheet"
    static let table = "Table"
    static let column = "Column"
    static let row = "Row"
    static let cell = "Cell"
    static let data = "Data"
    
}

/** Handler for managing sheet meta.*/
public class TRMetaFileHandler {
    
    /// sheet author
    private var author = ""
    
    /// sheet status
    private var hasActiveSheetNode = false
    
    /// sheet data
    public var content = ""
    
    /**
     Constructor for string value
     - Parameters:
        - withString: string value
     */
    public init(with author: String) {
        self.author = author
        
        // init
        self.content = TRSheetConst.header
        
        // begin properties
        beginTag(tag: TRSheetConst.documentProperties, config: TRSheetConst.documentPropertiesConfig)
        
        // add author
        beginTag(tag: TRSheetConst.author)
        addContent(content: author)
        endTag(tag: TRSheetConst.author)
        
        // add created at
        beginTag(tag: TRSheetConst.created)
        addContent(content: "\(Date())")
        endTag(tag: TRSheetConst.created)
        
        endTag(tag: TRSheetConst.documentProperties)
    }
    
    // MARK: - General sheet methods
    
    public func beginTag(tag: String) {
        self.content.append("<\(tag)>")
    }
    
    public func beginTag(tag: String, config: String) {
        self.content.append("<\(tag) \(config)>")
    }
    
    public func addContent(content: String) {
        self.content.append(content)
    }
    
    public func endTag(tag: String) {
        self.content.append("</\(tag)>\n")
    }
    
    public func beginAndEndTag(tag: String, config: String) {
        self.content.append("<\(tag) \(config) />\n")
    }
    
    public func resetContent(){
        self.content = ""
    }
    
    // MARK: - Custom sheet methods
    
    public func beginWorkSheet(name: String) {
        self.beginTag(tag: TRSheetConst.worksheet, config: "ss:Name=\"\(name)\">\n")
    }
    
    public func endWorksheet() {
        self.endTag(tag: TRSheetConst.worksheet)
    }
    
    public func beginTable() {
        self.beginTag(tag: TRSheetConst.table)
    }
    
    public func endTable() {
        self.endTag(tag: TRSheetConst.table)
    }
    
    public func appendColumn(autoFitWidth: Int, width: Int) {
        self.beginAndEndTag(tag: TRSheetConst.column, config: "ss:AutoFitWidth=\(autoFitWidth) ss:Width=\"\(width)\"")
    }
    
    public func beginRow() {
        self.beginTag(tag: TRSheetConst.row)
    }
    
    public func endRow() {
        self.endTag(tag: TRSheetConst.row)
    }
    
    public func beginCell(configuration: String) {
        self.beginTag(tag: TRSheetConst.cell, config: configuration)
    }
    
    public func endCell() {
        self.endTag(tag: TRSheetConst.cell)
    }
    
    public func beginAndEndCell() {
        self.beginAndEndTag(tag: TRSheetConst.cell, config: "")
    }
    
    public func beginData(configuration: String) {
        self.beginTag(tag: TRSheetConst.data, config: configuration)
    }
    
    public func endData() {
        self.endTag(tag: TRSheetConst.data)
    }
    
    public func endWorkbook() {
        self.endTag(tag: TRSheetConst.workbook)
    }
    
}
