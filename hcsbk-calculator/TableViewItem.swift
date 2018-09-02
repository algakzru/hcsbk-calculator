//
//  TableViewItem.swift
//  hcsbk-calculator
//
//  Created by Apple on 06/04/2018.
//  Copyright © 2018 SF-Express. All rights reserved.
//

import UIKit

class TableViewItem {
    
    //MARK: Properties
    
    var title: String
    var detail: String
    
    //MARK: Initialization
    
    init?(title: String, detail: String) {
        
        // The title must not be empty
        guard !title.isEmpty else {
            return nil
        }
        
        // The detail must not be empty
        guard !detail.isEmpty else {
            return nil
        }
        
        // Initialize stored properties.
        self.title = title
        self.detail = detail
    }
    
}
