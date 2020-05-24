/*

TRSheetObject.m

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

// MARK: - TRSheetObject

/** The Sheet object representates a sheet in the spreadsheet.*/
public class TRSheetObject {
    
    // MARK: - Local Variable
    
    /// name of sheet
    var sheetname: String
    
    /// list of cells within sheet
    var array: [[TRValue]]
    
    /// list of cell configurations
    var configArray: [String]
    
    /// list with rows of formatting cells
    var formatArray: [[Any]]
    
    /// column size
    var columnSizeList: [Int]?
    
    /// attribute to count styles for cells
    private var styleCounter: UInt
    
    /// map for remembering available styles
    private var styleMap: [String: Any]
    
    // MARK: - Methods
    
    /**
     Constructor for new Sheet object
        - Parameters:
            - name:  sheet file name
    */
    public init(name: String) {
        self.sheetname = name
        self.array = []
        self.configArray = []
        self.formatArray = []
        
        self.styleMap = [:]
        self.styleCounter = 0
    }
    
    /**
     Method to add a row with entries
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
    */
    public func addRow(entries: [TRValue]) {
        addRow(entries: entries, formatting: [], configurations: [])
    }
    /**
     Method to add a row with entries and formats
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - formatting: contains cell formatting, if it contains one item, then it works for the whole row
    */
    public func addRow(entries: [TRValue], formatting: [Any]?) {
        addRow(entries: entries, formatting: formatting, configurations: [])
    }
    
    /**
     Method to add a row with entries and using configurations
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - configurations: contains links and cell configurations
    */
    public func addRow(entries: [TRValue], configurations: [String]) {
        addRow(entries: entries, formatting: [], configurations: configurations)
    }
    
    public func styleForRow(rowIndex: Int, columnIndex: Int) -> String? {
        let formatArray = self.formatArray
        if formatArray.count > rowIndex {
            let formatRow = formatArray[rowIndex]
            if formatRow.count > columnIndex {
                let formatCell = formatRow[columnIndex]
                if let styleId = getStyleId(mapValue: formatCell) {
                    return styleId
                }
                
            } else {
                print("TRSheetObject#\(#function): Error finding format for cell: \(columnIndex)")
            }
        } else {
            print("TRSheetObject#\(#function): Error finding format for row: \(rowIndex)")
        }
        return nil
    }
    
    private func getStyleId(mapValue: Any) -> String? {

        //invalid format will be ignored
        if let _ = mapValue as? String {
            return ""
        }
        
        if let _ = mapValue as? NSNull {
            return ""
        }
        
//        if let style = mapValue as? Dictionary {
//
//        }
        
        
        return nil
    }
    
    /**
     Method to add a row with entries and using configurations
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - formatting: contains cell formatting, if it contains one item, then it works for the whole row
            - configurations: contains links and cell configurations
    */
    public func addRow(entries: [TRValue], formatting: [Any]?, configurations: [String]?) {
        self.array.append(entries)
        
        if let formattingList = formatting {
            if formattingList.count > 0 {
                
                // copy all formatting and check valid number of entries (match entries)
                var format: [Any] = []
                
                for f in formattingList {
                    format.append(f)
                }
                
                // fill up with last formatting
                while entries.count > format.count {
                    if let entry = entries.last {
                        format.append(entry)
                    } else {
                        break
                    }
                }
                
                self.formatArray.append(format)
            } else {
                self.formatArray.append([])
            }
        } else {
            self.formatArray.append([])
        }
        
        if let configurationList = configurations {
            self.configArray.append(contentsOf: configurationList)
        } else {
            self.configArray.append(contentsOf: [])
        }
    }
    
    /**
     Method to replace a row with new antries
        - Parameters:
            - index:   row position
            - replaceArray:  new content
    */
    public func replaceRow(index: Int, replaceArray: [TRValue]) {
        if index >= self.array.count {
            self.array.append(replaceArray)
        } else {
            self.array[index] = replaceArray
        }
    }
    
    /**
     Method to delete a row and replace with empty entries
        - Parameters:
            - index:   row index which needs to be deleted
    */
    public func deleteRow(index: Int) {
        //try to delete something which doesnt exist
        if index >= self.array.count {
            // ignore it
        } else {
            self.array[index] = []
        }
    }
}
