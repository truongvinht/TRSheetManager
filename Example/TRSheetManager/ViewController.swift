//
//  ViewController.swift
//  TRSheetManager
//
//  Created by Truong Vinh Tran on 04/19/2020.
//  Copyright (c) 2020 Truong Vinh Tran. All rights reserved.
//

import UIKit
import TRSheetManager

class ViewController: UIViewController {

    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view, typically from a nib.
        
        let sheetManager = TRSheetManager(withAuthor: "Truong Vinh Tran")
        if let firstPage = sheetManager.addSheet(sheetName: "1st page") {
            firstPage.addRow(entries: [TRValue(withString: "Name"),TRValue(withString: "Hans")])
            firstPage.addRow(entries: [TRValue(withString: "Score"),TRValue(withNumber: NSNumber(value: 1000))])
        }
        if sheetManager.writeSheetToFile(fileName: "Example.xls") {
            print("file saved")
        }
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }

}
