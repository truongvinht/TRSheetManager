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

// MARK: - Constants for style the cell

///font size key
let TR_SIZE_KEY = "size"

///font name key
let TR_FONTNAME_KEY = "fontName"

///font color
let TR_COLOR_KEY = "color"

///font is bold
let TR_BOLD_KEY = "bold"

///font is italic
let TR_ITALIC_KEY = "italic"

///font is underlined
let TR_UNDERLINE_KEY = "underline"

///cell background color in HEX (e.g #000000 = black)
let TR_BGCOLOR_KEY = "backgroundColor"

/// horizontal key
let TR_HORIZONTAL_KEY = "horizontalAlignment"
let TR_HORIZONTAL_AUTOMATIC = "Automatic"
let TR_HORIZONTAL_LEFT = "Left"
let TR_HORIZONTAL_CENTER = "Center"
let TR_HORIZONTAL_RIGHT = "Right"

///vertical key
let TR_VERTICAL_KEY = "verticalAlignment"
let TR_VERTICAL_AUTOMATIC = "Automatic"
let TR_VERTICAL_TOP = "Top"
let TR_VERTICAL_CENTER = "Center"
let TR_VERTICAL_BOTTOM = "Bottom"

///wrap the text within cell
let TR_WRAPTEXT_KEY = "wrapText"

///border attributes
let TR_BORDER_TOP_KEY = "borderTop"
let TR_BORDER_BOTTOM_KEY = "borderBottom"
let TR_BORDER_RIGHT_KEY = "borderRight"
let TR_BORDER_LEFT_KEY = "borderLeft"
let TR_BORDER_ALL_KEY = "borderAll"

///fixed number to x,xx
let TR_NUMBER_FIXED = "fixedNumber"

///add custom format
let TR_CUSTOM_KEY = "custom"

///zoom level for the work sheet window
let TR_WSO_ZOOM_KEY = "zoom"

///custom work sheet setting
let TR_WSO_CUSTOM_KEY = "custom"

// MARK: - TRSheetManager

/** Class for creating Spreadsheets (using XML structure)*/
class TRSheetManager {
    
    // MARK: - Local Variable
    
    /// Array with all sheet pages
    private var sheetArray: [TRSheetObject]
    
    /// name of the author
    private var authorName: String?
    
    /// default style
    private var defaultStyle: String?
    
    /// map for remembering available styles
    private var styleMap: [String: Any]
    
    /// attribute to count styles for cells
    private var styleCounter: UInt
    
    /// options for the work sheet
    private var workSheetOptions: [String: Any]?
    
    // MARK: - Methods
    
    /**
     Constructor for new Sheet Manager instance
        - Parameters:
            - withAuthor: name of the author
            - returns: new instance
    */
    init(withAuthor author: String?) {
        self.sheetArray = []
        
        self.authorName = author
        self.styleMap = [:]
        
        self.styleCounter = 0
    }
    
    /**
     Method to set the default style for the whole table
        - Parameters:
            - defaultStyle: dictionary which carries all attributes (size,fontName,color,bold,italic,underline, backgroundColor)
    */
    func setDefaultFontStyle(defaultStyle: [String: Any]) {
        let styleString = self.parseStyle(styleInfo: defaultStyle, forId: "Default", forName: "ss:Name\"Normal\"")
        
        // remove all defaults
        if styleString.count == 0 {
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
    func setSheetOptions(options: [String: Any]) {
        self.workSheetOptions = options
    }
    
    func parseWorkSheetOptions(options:[String:Any]) -> String {
        
        var sheetOptionString = "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><FitToPage />"
        
        if let value = options[TR_WSO_ZOOM_KEY] {
            sheetOptionString = "\(sheetOptionString)<Zoom>\(value)</Zoom>"
        }
        
        // optional stuff
        if let value = options[TR_WSO_CUSTOM_KEY] {
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
    func addSheet(sheetName: String) -> TRSheetObject? {

        //check wether sheet name already exist
        for sheet in self.sheetArray {
            if sheet.getSheetName() == sheetName {
                return nil
            }
            
        }
        
        // create new sheet
        let sheet = TRSheetObject(name: sheetName)
        self.sheetArray.append(sheet)
        return sheet
    }
    
    /**
     Method to read all sheets already in the array
        - Parameters:
            - returns: available sheets
    */
    func getAllSheets() -> [TRSheetObject] {
        return self.sheetArray
    }
    
    /**
     Method to genereate the sheet data
        - Parameters:
            - returns: sheet data
    */
    func generateSheet() -> Data? {
        
        //reset style counter
        self.styleCounter = 0
        
        var dataStream = "<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\nxmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n\n<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n"
        
        // only set author if available
        if let author = self.authorName {
            dataStream = "\(dataStream)\t<Author>\(author)</Author>\n"
        }
        
        // set created date
        dataStream = "\(dataStream)\t<Created>\(Date())</Created>\n</DocumentProperties>\n"
        
        //add styles for date (week day)
        dataStream.append("\n<Styles>\n")
        
        //add default styles if available
        if let style = self.defaultStyle {
            dataStream.append(style)
        }
        
        for sheet in self.sheetArray {
            for y in 0...sheet.getFormatArray().count-1 {
                
                let list = sheet.getFormatArray()[y]
                
                for z in 0...list.count {
                    
                    if let dictionary = list[z] as? [String: Any] {
                        if let style = self.getStyleForDict(style: dictionary) {
                            dataStream.append(style)
                        }
                    }
                }
            }
        }
        dataStream.append("</Styles>\n\n")

        //write every sheet page
        for sheet in self.sheetArray {

            //TODO: Incomplete
            
        }
        // end tag
        dataStream.append("</Workbook>")
        return dataStream.data(using: .utf8);
    }
    
    private func getStyleForDict(style:[String:Any]) -> String? {
        return nil
    }
    
    // MARK: - File access Methods
    
    
    /**
     Method to write the given sheet to file
        - Parameters:
            - fileName: file name of the sheet
            - returns: true if the sheet was successfully written
    */
    func writeSheetToFile(fileName: String) -> Bool {
        
        let defaultPath = NSSearchPathForDirectoriesInDomains(.documentDirectory, .userDomainMask, true)[0]
        
        if let sheet = self.generateSheet() {
            
            do {
                 try sheet.write(to: URL(fileURLWithPath: defaultPath).appendingPathComponent(fileName), options: .atomicWrite)
                return true
            } catch  {
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
    func writeSheetToFile(fileName: String, inPath path: String) -> Bool {
        
        // write to file
        if let sheet = self.generateSheet() {
            do {
                 try sheet.write(to: URL(fileURLWithPath: path).appendingPathComponent(fileName), options: .atomicWrite)
                return true
            } catch  {
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
    private func parseStyle(styleInfo:[String: Any]?, forId styleID:String, forName name: String) -> String {

        //invalid format will be ignored
        guard let info = styleInfo else {
            return ""
        }
        
        // no keys
        if info.count == 0 {
            return ""
        }
        
        //replace the all with other keys
        if let _ = info[TR_BORDER_ALL_KEY] {
            var dictionary = info;
            dictionary[TR_BORDER_BOTTOM_KEY] = info[TR_BORDER_ALL_KEY]
            dictionary[TR_BORDER_TOP_KEY] = info[TR_BORDER_ALL_KEY]
            dictionary[TR_BORDER_RIGHT_KEY] = info[TR_BORDER_ALL_KEY]
            dictionary[TR_BORDER_LEFT_KEY] = info[TR_BORDER_ALL_KEY]
            dictionary.removeValue(forKey: TR_BORDER_ALL_KEY)
            return self.parseStyle(styleInfo: dictionary, forId: styleID, forName: name)
        }
        
        let styleString = "\t<Style ss:ID=\"\(styleID)\"\(name)>\n"
        
        var sizeAttribute = ""
        
        //set the default font size
        if let size = info[TR_SIZE_KEY] as? NSNumber  {
            sizeAttribute = " ss:Size=\"\(size)\""
        } else {
            //try to use default style, if available
            if let style = self.defaultStyle  {
                if name.count == 0 {
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
        var fontName = info[TR_FONTNAME_KEY]
        
        // TODO: Incomplete implementation
        
        return ""
    }
    
}
