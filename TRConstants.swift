//
//  TRConstants.swift
//  Pods-TRSheetManager_Example
//
//  Created by Truong Vinh Tran on 04.06.20.
//

import Foundation


// MARK: - Constants for style the cell

/// key for mapping configurations
struct TRKey {
    
    /// key for font size
    static let size = "size"
    
    /// key for font name
    static let fontName = "fontName"
    
    /// key for horizontal alignment
    static let horizontal = "horizontalAlignment"
    
    /// key for font color
    static let color = "color"
    
    /// key for bold font
    static let bold = "bold"
    
    /// key for italic font
    static let italic = "italic"
    
    /// key for underlined font
    static let underline = "underline"
    
    /// cell background color in HEX (e.g #000000 = black)
    static let bgcolor = "backgroundColor"
    
    /// key for vertical alignment
    static let vertical = "verticalAlignment"
    
    /// wrap the text within cell
    static let wraptext = "wrapText"
    
    ///fixed number to x,xx
    static let numberFixed = "fixedNumber"
    
    /// custom format
    static let custom = "custom"
}

/// horizontal key aligment
struct TRHorizontal {
    
    /// automatic alignment
    let automatic = "Automatic"
    
    /// left alignment
    let left = "Left"
    
    /// center alignment
    let center = "Center"
    
    /// right alignment
    let right = "right"
}

/// vertical key aligment
struct TRVertical {
    
    /// automatic alignment
    let automatic = "Automatic"
    
    /// top alignment
    let top = "Top"
    
    /// center alignment
    let center = "Center"
    
    /// bottom alignment
    let bottom = "Bottom"
}

/// cell border line
struct TRBorder {
    
    /// top border
    static let top = "borderTop"
    
    /// bottom border
    static let bottom = "borderBottom"
    
    /// right border
    static let right = "borderRight"
    
    /// left border
    static let left = "borderLeft"
    
    /// all border
    static let all = "borderAll"
}

/// worksheet constants
struct TRWorkSheet {
    
    ///zoom level for the work sheet window
    static let zoom = "zoom"
    
    ///custom work sheet setting
    static let custom = "custom"
}
