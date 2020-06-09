/*

TRSheetManager.m

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

// MARK: - TRSheetManager

/** Class for creating Spreadsheets (using XML structure)*/
public class TRSheetManager {
    
    // MARK: - Local Variable
    
    /// Array with all sheet pages
    private var sheetArray: [TRSheetObject]
    
    /// name of the author
    private var authorName: String?
    
    /// default style
    private var defaultStyle: String?
    
    /// options for the work sheet
    private var workSheetOptions: [String: Any]?
    
    // MARK: - Methods
    
    /**
     Constructor for new Sheet Manager instance
        - Parameters:
            - withAuthor: name of the author
            - returns: new instance
    */
    public init(withAuthor author: String?) {
        self.sheetArray = []
        
        self.authorName = author
    }
    
    /**
     Method to set the default style for the whole table
        - Parameters:
            - defaultStyle: dictionary which carries all attributes (size,fontName,color,bold,italic,underline, backgroundColor)
    */
    public func setDefaultFontStyle(defaultStyle: [String: Any]) {
        let styleString = self.parseStyle(styleInfo: defaultStyle, forId: "Default", forName: "ss:Name\"Normal\"")
        
        // remove all defaults
        if styleString.isEmpty {
            self.defaultStyle = nil
            return
        }
        
        self.defaultStyle = styleString
    }
    
    /**
     Method to set worksheet Options
        - Parameters:
            - options: dictionary with more options for  displaying
    */
    public func setSheetOptions(options: [String: Any]) {
        self.workSheetOptions = options
    }
    
    func parseWorkSheetOptions(options: [String: Any]) -> String {
        
        var sheetOptionString = "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><FitToPage />"
        
        if let value = options[TRWorkSheet.zoom] {
            sheetOptionString = "\(sheetOptionString)<Zoom>\(value)</Zoom>"
        }
        
        // optional stuff
        if let value = options[TRWorkSheet.custom] {
            sheetOptionString = "\(sheetOptionString)\(value)"
        }
        
        return "\(sheetOptionString)</WorksheetOptions>"
    }
    
    /**
     Method to add a new sheetpage
        - Parameters:
            - sheetName: new name of the sheet
            - returns: TRSheetObject if the sheet was created successfully
    */
    public func addSheet(sheetName: String) -> TRSheetObject? {

        //check wether sheet name already exist
        for sheet in self.sheetArray where sheet.sheetname == sheetName {
            return nil
        }
        
        // create new sheet
        let sheet = TRSheetObject(name: sheetName)
        self.sheetArray.append(sheet)
        return sheet
    }
    
    /**
        Method to remove an existing sheet page
        - Parameters:
            - sheet: name of the deleting sheet
     */
    public func removeSheet(sheet: TRSheetObject) {
        
        guard let index = self.sheetArray.firstIndex(where: { $0.sheetname == sheet.sheetname })
                else { return }

        self.sheetArray.remove(at: index)
    }
    
    /**
     Method to remove all sheet pages
     */
    public func removeAllSheets() {
        
        if !self.sheetArray.isEmpty {
            for index in 0...self.sheetArray.count - 1 {
                self.sheetArray.remove(at: index)
            }
        }
        
    }
    
    /**
     Method to read all sheets already in the array
        - Parameters:
            - returns: available sheets
    */
    public func getAllSheets() -> [TRSheetObject] {
        return self.sheetArray
    }
    
    /**
      Method to get number of current sheets within document
        - Parameters:
            - returns: number of available sheets
     */
    public func numberOfSheets() -> Int {
        return self.sheetArray.count
    }
    
    /**
     Method to genereate the sheet data
        - Parameters:
            - returns: sheet data
    */
    public func generateSheet() -> Data? {
        
        let sheetHandler: TRMetaFileHandler
        
        // only set author if available
        if let author = self.authorName {
            sheetHandler = TRMetaFileHandler(withAuthor: author)
        } else {
            sheetHandler = TRMetaFileHandler(withAuthor: "")
        }
        
        // add styles formatting rules
        sheetHandler.addContent(content: loadCustomStyles(sheets: self.sheetArray))

        //write every sheet page
        for sheet in self.sheetArray {

            sheetHandler.beginWorkSheet(name: sheet.sheetname)
            
            //start the table
            sheetHandler.beginTable()
            
            //add column size if available
            if let sizeList = sheet.columnSizeList {
                for size in sizeList {
                    sheetHandler.appendColumn(autoFitWidth: 0, width: size)
                }
            }
            
            // append every row
            if !sheet.array.isEmpty {
                for rowIndex in 0...sheet.array.count - 1 {
                    //for rowEntry in sheet.array {
                    let rowEntry = sheet.array[rowIndex]
                    sheetHandler.beginRow()
                    
                    //every cell
                    for cellIndex in 0...rowEntry.count - 1 {
                        //for cellEntry in rowEntry {
                        
                        let cellEntry = rowEntry[cellIndex]
                        var type = cellEntry.contentType()
                        var contentData: Any = ""
                        var cellConfiguration = ""
                        
                        //add cell configuration, if available
                        if sheet.array.count == sheet.configArray.count {
                            
                            if let index = sheet.array.firstIndex(where: { $0 as AnyObject === rowEntry as AnyObject }) {
                                cellConfiguration = sheet.configArray[index]
                            }
                        }
                        
                        // check content type
                        let cellData = cellEntry.formatedData(cellConfiguration: cellConfiguration)
                        contentData = cellData.contentData
                        type = cellData.outputType
                        
                        // format cell
                        if let style = sheet.styleForRow(rowIndex: rowIndex, columnIndex: cellIndex) {
                            cellConfiguration = String(format: "%@ss:StyleID=\"%@\"", cellConfiguration, style)
                        }
                        
                        var isValidString = false
                        
                        if let content = contentData as? String {
                            if content.isEmpty {
                                isValidString = true
                            }
                        }
                        
                        if type == .string && !isValidString {
                            sheetHandler.beginAndEndCell()
                        } else {
                            if let varArg = contentData as? CVarArg {
                                sheetHandler.beginCell(configuration: cellConfiguration)
                                sheetHandler.beginDataWithType(type: type.rawValue)
                                sheetHandler.addContent(content: String(format: "%@", varArg))
                                sheetHandler.endData()
                                sheetHandler.endCell()
                                
                            } else {
                                // default closing
                                sheetHandler.beginAndEndCell()
                            }
                        }
                    }
                    sheetHandler.endRow()
                }
            }

            //end table
            sheetHandler.endTable()
            
            //if wsheet options are available add them
            if let options = self.workSheetOptions {
                sheetHandler.addContent(content: self.parseWorkSheetOptions(options: options))
            }
            sheetHandler.endWorksheet()
        }
        // end tag
        sheetHandler.endWorkbook()
        return sheetHandler.content.data(using: .utf8)
    }
    
    private func getStyleForDict(style: [String: Any]) -> String? {
        return nil
    }
    
    private func loadCustomStyles(sheets: [TRSheetObject]) -> String {
        var dataStream = ""
        
        //add styles for date (week day)
        dataStream.append("\n<Styles>\n")
        
        //add default styles if available
        if let style = self.defaultStyle {
            dataStream.append(style)
        }
        
        for sheet in sheets {
            
            // skip empty
            if sheet.formatArray.isEmpty {
                continue
            }
            
            for formatRow in 0...sheet.formatArray.count - 1 {
                
                let list = sheet.formatArray[formatRow]
                
                if list.isEmpty {
                    for formatCell in 0...list.count - 1 {
                        
                        if let dictionary = list[formatCell] as? [String: Any] {
                            if let style = self.getStyleForDict(style: dictionary) {
                                dataStream.append(style)
                            }
                        }
                    }
                }
            }
        }
        dataStream.append("</Styles>\n\n")
        return dataStream
    }
    
    // MARK: - File access Methods
    
    /**
     Method to write the given sheet to file
        - Parameters:
            - fileName: file name of the sheet
            - returns: true if the sheet was successfully written
    */
    public func writeSheetToFile(fileName: String) -> Bool {
        
        let defaultPath = NSSearchPathForDirectoriesInDomains(.documentDirectory, .userDomainMask, true)[0]
        
        if let sheet = self.generateSheet() {
            
            do {
                 try sheet.write(to: URL(fileURLWithPath: defaultPath).appendingPathComponent(fileName), options: .atomicWrite)
                return true
            } catch {
                print("TRSheetManager#\(#function): Error saving file [\(defaultPath)]")
            }
        }
        
        return false
    }
    
    /**
     Method to write the given sheet into the file with given path
        - Parameters:
            - fileName: file name of the sheet
            - path: location of the file
            - returns: true, if the sheet was successfully written
    */
    public func writeSheetToFile(fileName: String, inPath path: String) -> Bool {
        
        // write to file
        if let sheet = self.generateSheet() {
            do {
                 try sheet.write(to: URL(fileURLWithPath: path).appendingPathComponent(fileName), options: .atomicWrite)
                return true
            } catch {
                print("TRSheetManager#\(#function): Error saving file [\(path)]")
            }
        }
        
        return false
    }
    
    // MARK: - Helper Methods
    
    /**
     Wrapper Method for parseStyle Defaults
        - Parameters:
            - styleInfo: dictionary with all style informations
            - styleID: Style ID tag
            - returns: new string style
    */
    private func parseStyle(styleInfo: [String: Any]?, forID styleID: String) -> String {
        return self.parseStyle(styleInfo: styleInfo, forId: styleID, forName: "")
    }
    /**
     Helper method to parse the styleInformations
        - Parameters:
            - styleInfo: dictionary with all style informations
            - styleID: Style ID tag
            - name: name flag in the style for Default
            - returns: style string
     
    */
    private func parseStyle(styleInfo: [String: Any]?, forId styleID: String, forName name: String) -> String {

        //invalid format will be ignored
        guard let info = styleInfo else {
            return ""
        }
        
        // no keys
        if info.isEmpty {
            return ""
        }
        
        //replace the all with other keys
        if info[TRBorder.all] != nil {
            var dictionary = info
            dictionary[TRBorder.bottom] = info[TRBorder.all]
            dictionary[TRBorder.top] = info[TRBorder.all]
            dictionary[TRBorder.right] = info[TRBorder.all]
            dictionary[TRBorder.left] = info[TRBorder.all]
            dictionary.removeValue(forKey: TRBorder.all)
            return self.parseStyle(styleInfo: dictionary, forId: styleID, forName: name)
        }
        
        // TODO: need to clean up
        //let styleString = "\t<Style ss:ID=\"\(styleID)\"\(name)>\n"
        
        var sizeAttribute = ""
        
        //set the default font size
        if let size = info[TRKey.size] as? NSNumber {
            sizeAttribute = " ss:Size=\"\(size)\""
        } else {
            //try to use default style, if available
            if let style = self.defaultStyle {
                if name.isEmpty {
                    let splitDefaultStyle = style.components(separatedBy: " ")
                    
                    for part in splitDefaultStyle {
                        let substring = part[..<part.index(part.startIndex, offsetBy: 7)]
                        //part.substring(to: part.index(part.startIndex, offsetBy: 7))
                        
                        if substring == "ss:Size" {
                            sizeAttribute = " \(part)"
                        }
                    }
                }
            }
        }
        
        // set default font name
        let fontName = info[TRKey.fontName]
        
        if let name = fontName as? String {
            print("\(sizeAttribute) \(name)")
        }
        
        // TODO: Incomplete implementation
        
        return ""
    }
}
