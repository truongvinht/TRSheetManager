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
class TRSheetObject {
    
    // MARK: - Local Variable
    
    /// name of sheet
    private var sheetname: String
    
    /// list of cells within sheet
    private var array: [[TRValue]]
    
    /// list of cell configurations
    private var configArray: [Any]
    
    /// list with rows of formatting cells
    private var formatArray: [[Any]]
    
    /// column size
    private var columnSizeList: [Int]
    
    // MARK: - Methods
    
    /**
     Constructor for new Sheet object
        - Parameters:
            - name:  sheet file name
    */
    init(name: String) {
        self.sheetname = name
        self.array = []
        self.configArray = []
        self.formatArray = []
        self.columnSizeList = []
    }
    
    /**
        Method for getting sheet name
        - Parameters:
            - returns: sheet name
     */
    func getSheetName() -> String {
        return self.sheetname
    }
    
    func getFormatArray() -> [[Any]] {
        return self.formatArray
    }
    
    /**
     Method to set the width of the column
        - Parameters:
            - list:  list is an array with NSNumbers regarding to its size
    */
    func setColumnSize(list: [Int]) {
        self.columnSizeList = list
    }
    
    /**
     Method to add a row with entries
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
    */
    func addRow(entries: [TRValue]) {
        addRow(entries: entries, formatting: [], configurations: [])
    }
    /**
     Method to add a row with entries and formats
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - formatting: contains cell formatting, if it contains one item, then it works for the whole row
    */
    func addRow(entries: [TRValue], formatting: [Any]?) {
        addRow(entries: entries, formatting: formatting, configurations: [])
    }
    
    /**
     Method to add a row with entries and using configurations
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - configurations: contains links and cell configurations
    */
    func addRow(entries: [TRValue], configurations: [Any]) {
        addRow(entries: entries, formatting: [], configurations: configurations)
    }
    
    /**
     Method to add a row with entries and using configurations
        - Parameters:
            - entries:   an array with NSNumber objects as number and String as string
            - formatting: contains cell formatting, if it contains one item, then it works for the whole row
            - configurations: contains links and cell configurations
    */
    func addRow(entries: [TRValue], formatting: [Any]?, configurations: [Any]?) {
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
            self.configArray.append(configurationList)
        } else {
            self.configArray.append([])
        }
    }
    
    /**
     Method to replace a row with new antries
        - Parameters:
            - index:   row position
            - replaceArray:  new content
    */
    func replaceRow(index: Int, replaceArray: [TRValue]) {
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
    func deleteRow(index: Int) {
        //try to delete something which doesnt exist
        if index >= self.array.count {
            // ignore it
        } else {
            self.array[index] = []
        }
    }
}
