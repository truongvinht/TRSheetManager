/*

TRValue.m

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

public enum TRValueType: String {
    case numeric = "Number"
    case string = "String"
    case date = "DateTime"
    case undefined = "undef"
}

/** Sheet object for carrying value.*/
public class TRValue {
    
    /// numeric value
    private var numericValue: NSNumber?
    
    /// string value
    private var stringValue: String?
    
    /// date value
    private var dateValue: Date?
    
    /**
     Constructor for numeric value
     - Parameters:
        - withNumber: numeric value
     */
    public init(withNumber number: NSNumber) {
        self.numericValue = number
    }
    
    /**
     Constructor for string value
     - Parameters:
        - withString: string value
     */
    public init(withString string: String) {
        self.stringValue = string
    }
    
    
    /**
     Constructor for date value
     - Parameters:
        - withDate: date value
     */
     public init(withDate date: Date) {
         self.dateValue = date
     }
    
    
    /**
     Get stored value
     - returns: stored value (Number or String), if no data is stored, an empty string is returned
     */
    public func value() -> Any {
        if isValid() {
            
            if let number = self.numericValue {
                return number
            }
            
            if let string = self.stringValue {
                return string
            }
            
            if let date = self.dateValue {
                return date
            }
        }
        
        // empty string as fallback
        return ""
    }
    
    public func isValid() -> Bool {
        
        if contentType() == .undefined {
            return false
        } else {
            return true
        }
    }
    
    
    /**
     Get details whether stored value is numeric or string
     - returns: true, if value is numeric
     */
    public func contentType() -> TRValueType {
        if self.numericValue != nil {
            return .numeric
        }
        
        if self.stringValue != nil {
            return .string
        }
        
        if self.dateValue != nil {
            return .date
        }
        
        return .undefined
    }
}
