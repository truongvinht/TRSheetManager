import XCTest
import TRSheetManager

class Tests: XCTestCase {
    
    var sheetManager = TRSheetManager(withAuthor: "Truong Vinh Tran")
    
    override func setUp() {
        super.setUp()
        // Put setup code here. This method is called before the invocation of each test method in the class.

    }
    
    override func tearDown() {
        // Put teardown code here. This method is called after the invocation of each test method in the class.
        super.tearDown()
        
        // write file
        if sheetManager.writeSheetToFile(fileName: "Example.xls") {
            XCTAssert(true, "File written")
        } else {
            XCTAssert(false, "File not written")
        }
        
        //clean up all sheet pages
        sheetManager.removeAllSheets()
    }
    
    func testAddSheet() {
        XCTAssert(sheetManager.numberOfSheets()==0, "Starting with 0 sheets")
        
        if let _ = sheetManager.addSheet(sheetName: "1st page") {
            XCTAssert(sheetManager.numberOfSheets()==1, "Added 1 sheet")
        }
        
        
    }
    
    func testExample() {
        // This is an example of a functional test case.
        
        if let firstPage = sheetManager.addSheet(sheetName: "1st page") {
            firstPage.addRow(entries: [TRValue(withString: "Name"),TRValue(withString: "Hans")])
            firstPage.addRow(entries: [TRValue(withString: "Score"),TRValue(withNumber: NSNumber(value: 1000))])
        }
        
        XCTAssert(true, "Pass")
    }
    
    func testPerformanceExample() {
        // This is an example of a performance test case.
        self.measure() {
            // Put the code you want to measure the time of here.
        }
    }
    
}
