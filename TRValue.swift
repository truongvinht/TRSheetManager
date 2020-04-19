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

/** Sheet object for carrying value.*/
class TRValue {
    
    /// numeric value
    private var numericValue: NSNumber?
    
    /// string value
    private var stringValue: String?
    
    /**
     Constructor for numeric value
     - Parameters:
        - withNumber: numeric value
     */
    init(withNumber number: NSNumber) {
        self.numericValue = number
    }
    
    /**
     Constructor for string value
     - Parameters:
        - withString: string value
     */
    init(withString string: String) {
        self.stringValue = string
    }
    
    /**
     Get stored value
     - returns: stored value (Number or String), if no data is stored, an empty string is returned
     */
    func getValue() -> Any {
        if let number = self.numericValue {
            return number
        } else {
            if let string = self.stringValue {
                return string
            }
        }
        // empty string as fallback
        return ""
    }
}
