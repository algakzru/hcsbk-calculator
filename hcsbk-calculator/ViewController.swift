//
//  ViewController.swift
//  hcsbk-calculator
//
//  Created by Apple on 24/03/2018.
//  Copyright © 2018 SF-Express. All rights reserved.
//

import UIKit
import xlsxwriter
import os.log

class ViewController: UIViewController, UITextFieldDelegate, UIDocumentInteractionControllerDelegate, UIPickerViewDataSource, UIPickerViewDelegate {
    
    //MARK: Properties
    
    @IBOutlet weak var tfSummaCredita: UITextField!
    @IBOutlet weak var tfSrokCredita: UITextField!
    @IBOutlet weak var tfProcentnayaStavka: UITextField!
    @IBOutlet weak var tfDataVydachiCredita: UITextField!
    @IBOutlet weak var tfDataPervogoPlatezha: UITextField!
    @IBOutlet weak var tfPaymentType: UITextField!
    
    let dateFormatter = DateFormatter()
    let numberFormatter = NumberFormatter()
    
    let summa_credita = "summa_credita"
    let srok_credita = "srok_credita"
    let procentnaya_stavka = "procentnaya_stavka"
    let data_vydachi_credita = "data_vydachi_credita"
    let data_pervogo_platezha = "data_pervogo_platezha"
    let payment_type = "payment_type"
    
    let paymentTypes = ["Аннуитетный", "Дифференцированный"]
    let ACCOUNTING_FLOAT = "#,##0.00тг";
    let BORDER_FORMAT = "borderFormat";
    let CENTRE_FORMAT = "centreFromat";
    let PERCENT_FORMAT = "percentFormat";
    let TENGE_FORMAT = "tengeFormat";
    let YELLOW_TENGE_FORMAT = "yellowTengeFormat";
    let GREEN_TENGE_FORMAT = "greenTengeFormat";
    let DATE_FORMAT = "dateFormat";
    let BLUE_TITLE_FORMAT = "blueTitleFormat";
    
    private var documentInteractionController: UIDocumentInteractionController!
    private var activityIndicatorView: UIActivityIndicatorView!
    
    enum Exception: Error {
        case invalidTextField(String)
    }
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        // Set dateFormatter
        dateFormatter.dateFormat = "yyyy-MM-dd"
        
        // Set numberFormatter
        numberFormatter.numberStyle = .decimal
        numberFormatter.locale = Locale(identifier: "en_US")
        numberFormatter.minimumFractionDigits = 2
        numberFormatter.maximumFractionDigits = 2
        
        // Set activityIndicatorView
        activityIndicatorView = UIActivityIndicatorView()
        activityIndicatorView.center = self.view.center
        activityIndicatorView.hidesWhenStopped = true
        activityIndicatorView.activityIndicatorViewStyle = UIActivityIndicatorViewStyle.whiteLarge
        activityIndicatorView.color = .black
        view.addSubview(activityIndicatorView)
        
        // Set thousands separator
        tfSummaCredita.delegate = self
        tfProcentnayaStavka.delegate = self
        
        // Retrieve UserDefaults
        let summaCredita = Double(UserDefaults.standard.string(forKey: summa_credita) ?? "")
        tfSummaCredita.text = numberFormatter.string(for: summaCredita)
        tfSrokCredita.text = UserDefaults.standard.string(forKey: srok_credita)
        let procentnayaStavka = Double(UserDefaults.standard.string(forKey: procentnaya_stavka) ?? "")
        tfProcentnayaStavka.text = numberFormatter.string(for: procentnayaStavka)
        tfDataVydachiCredita.text = UserDefaults.standard.string(forKey: data_vydachi_credita)
        tfDataPervogoPlatezha.text = UserDefaults.standard.string(forKey: data_pervogo_platezha)
        tfPaymentType.text = UserDefaults.standard.string(forKey: payment_type)
        
        // Сумма кредита
//        tfSummaCredita.addTarget(self, action: #selector(summaCreditaDidChange(_:)), for: .editingChanged)
        
        // Срок кредита
        tfSrokCredita.addTarget(self, action: #selector(srokCreditaDidChange(_:)), for: .editingChanged)
        
        // Процентная ставка
//        tfProcentnayaStavka.addTarget(self, action: #selector(procentnayaStavkaDidChange(_:)), for: .editingChanged)
        
        // Дата выдачи кредита
        let dataVydachiCreditaPicker = UIDatePicker()
        dataVydachiCreditaPicker.datePickerMode = .date
        let dataVydachiCredita = dateFormatter.date(from: tfDataVydachiCredita.text!) ?? Date()
        dataVydachiCreditaPicker.setDate(dataVydachiCredita, animated: true)
        dataVydachiCreditaPicker.addTarget(self, action: #selector(ViewController.handleDataVydachiCreditaPicker(sender:)), for: .valueChanged)
        tfDataVydachiCredita.inputView = dataVydachiCreditaPicker
        
        // Дата первого платежа
        let dataPervogoPlatezhaPicker = UIDatePicker()
        dataPervogoPlatezhaPicker.datePickerMode = .date
        let dataPervogoPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!) ?? Date()
        dataPervogoPlatezhaPicker.setDate(dataPervogoPlatezha, animated: true)
        dataPervogoPlatezhaPicker.addTarget(self, action: #selector(ViewController.handleDataPervogoPlatezhaPicker(sender:)), for: .valueChanged)
        tfDataPervogoPlatezha.inputView = dataPervogoPlatezhaPicker
        
        // UIPickerView
        let paymentTypePicker = UIPickerView()
        paymentTypePicker.delegate = self
        paymentTypePicker.dataSource = self
        var defaultRowIndex = paymentTypes.index(of: tfPaymentType.text!)
        if (defaultRowIndex == nil) { defaultRowIndex = 0 }
        paymentTypePicker.selectRow(defaultRowIndex!, inComponent: 0, animated: false)
        tfPaymentType.inputView = paymentTypePicker
        
        //Looks for single or multiple taps.
        let tap: UITapGestureRecognizer = UITapGestureRecognizer(target: self, action: #selector(dismissKeyboard))
        view.addGestureRecognizer(tap)
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    
    //MARK: Navigation
    
    override func prepare(for segue: UIStoryboardSegue, sender: Any?) {
        super.prepare(for: segue, sender: sender)
        
        do {
            
            switch (segue.identifier ?? "") {
            case "Calculate":
                
                guard let tableViewController = segue.destination as? TableViewController else {
                    fatalError("Unexpected destination: \(segue.destination)")
                }
                
                try validateTextField()
                
                //Causes the view (or one of its embedded text fields) to resign the first responder status.
                view.endEditing(true)
                
                // Set the tableViewItems to be passed to TableViewController after the unwind segue.
                var tableViewItems = [TableViewItem]()
                
                // Аннуитетный
                if paymentTypes[0] == tfPaymentType.text {
                    let procentPervogoRaschetnogoPerioda = try getProcentPervogoRaschetnogoPerioda()
                    let procentPoslednegoRaschetnogoPerioda = try getProcentPoslednegoRaschetnogoPerioda()
                    let procentEzhemesiachnyi = try getProcentEzhemesiachnyi()
                    let up = try getUp(procentPervogoRaschetnogoPerioda: procentPervogoRaschetnogoPerioda, procentPoslednegoRaschetnogoPerioda: procentPoslednegoRaschetnogoPerioda, procentEzhemesiachnyi: procentEzhemesiachnyi)
                    let down = try getDown(procentPoslednegoRaschetnogoPerioda: procentPoslednegoRaschetnogoPerioda, procentEzhemesiachnyi: procentEzhemesiachnyi)
                    let annuitet = try getAnnuitet(up: up, down: down)
                    let pereplata = try getPereplata(annuitet: annuitet)
                    
                    tableViewItems.append(TableViewItem(title: "Переплата:", detail: numberFormatter.string(for: pereplata)!)!)
                    tableViewItems.append(TableViewItem(title: "Ежемесячный платёж:", detail: numberFormatter.string(for: annuitet)!)!)
                }
                // Дифференцированный
                else if paymentTypes[1] == tfPaymentType.text {
                    var pereplata = Double(0)
                    let pogashenieOsnovnogoDolga = try getPogashenieOsnovnogoDolga()
                    guard let srokCredita = Int(tfSrokCredita.text!) else {
                        throw Exception.invalidTextField("Неверный формат срока кредита")
                    }
                    guard var ostatokOsnovnogoDolga = Double(trimCommaOfString(string: tfSummaCredita.text!)) else {
                        throw Exception.invalidTextField("Неверный формат суммы кредита")
                    }
                    for i in 1...srokCredita {
                        let procentRaschetnogoPerioda: Double
                        let dataPlatezha: Date
                        // first payment
                        if (1 == i) {
                            procentRaschetnogoPerioda = try getProcentPervogoRaschetnogoPerioda()
                            dataPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!)!
                        }
                        // last payment
                        else if (srokCredita == i) {
                            procentRaschetnogoPerioda = try getProcentPoslednegoRaschetnogoPerioda()
                            guard let dataVydachiCredita = dateFormatter.date(from: tfDataVydachiCredita.text!) else {
                                throw Exception.invalidTextField("Неверный формат даты выдачи кредита")
                            }
                            dataPlatezha = Calendar.current.date(byAdding: .month, value: srokCredita, to: dataVydachiCredita)!
                        }
                        // other payments
                        else {
                            procentRaschetnogoPerioda = try getProcentObichnogoRaschetnogoPerioda(raschiotnyPeriod: i)
                            guard let dataPervogoPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!) else {
                                throw Exception.invalidTextField("Неверный формат даты первого платежа")
                            }
                            dataPlatezha = Calendar.current.date(byAdding: .month, value: i - 1, to: dataPervogoPlatezha)!
                        }
                        let procentBanka = ostatokOsnovnogoDolga * procentRaschetnogoPerioda
                        pereplata += procentBanka
                        tableViewItems.append(TableViewItem(title: "Платёж за \(dateFormatter.string(from: dataPlatezha)):", detail: numberFormatter.string(for: procentBanka + pogashenieOsnovnogoDolga)!)!)
                        ostatokOsnovnogoDolga = ostatokOsnovnogoDolga - pogashenieOsnovnogoDolga;
                    }
                    
                    tableViewItems.insert(TableViewItem(title: "Переплата:", detail: numberFormatter.string(for: pereplata)!)!, at: 0)
                }
                
                tableViewController.tableViewItems = tableViewItems
                
            default:
                fatalError("Unexpected Segue Identifier; \(String(describing: segue.identifier))")
            }
            
        } catch Exception.invalidTextField(let errorMessage) {
            presentAlertController(title: "Ошибка", message: errorMessage)
        } catch {
            print("Unexpected error: \(error).")
            presentAlertController(title: "Ошибка", message: "Unexpected error")
        }
    }
    
    //MARK: Actions
    
    @IBAction func export(_ sender: UIButton) {
        // https://libxlsxwriter.github.io/getting_started.html
        // https://github.com/lrossi/libxlsxwriterCocoaExamples
        
        do {
            try validateTextField()
            
            // Causes the view (or one of its embedded text fields) to resign the first responder status.
            view.endEditing(true)
            
            activityIndicatorView.startAnimating()
            UIApplication.shared.beginIgnoringInteractionEvents()
            
            var fileURL: URL!
            
            DispatchQueue.global(qos:.userInteractive).async {
                let outputFileName = "Жилищный заем"
                let kFileExtension = "xlsx"
                
                if let documentsFolderPath = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first {
                    
                    fileURL = documentsFolderPath.appendingPathComponent(outputFileName).appendingPathExtension(kFileExtension);
                    
                    // Create a new workbook.
                    let workbook = new_workbook((fileURL as NSURL).fileSystemRepresentation)
                    
                    // Add a worksheet with a user defined sheet name.
                    let paramsSheet = workbook_add_worksheet(workbook, "Параметры")
                    let calculateSheet = workbook_add_worksheet(workbook, "Вычисление")
                    
                    // Set selected sheet
                    worksheet_activate(calculateSheet);
                    
                    let formatDictionary = self.createFormatDictionary(workbook: workbook!)
                    
                    // Add items to Параметры sheet
                    self.addItemsToParamsSheet(paramsSheet: paramsSheet!, formatDictionary: formatDictionary)
                    
                    // Add items to Вычисление sheet
                    if self.paymentTypes[0] == self.tfPaymentType.text {
                        // Аннуитетный
                        self.addItemsToAnnuitetCalculateSheet(calculateSheet: calculateSheet!, formatDictionary: formatDictionary)
                    } else if self.paymentTypes[1] == self.tfPaymentType.text {
                        // Дифференцированный
                        self.addItemsToDifferentiatedCalculateSheet(calculateSheet: calculateSheet!, formatDictionary: formatDictionary)
                    }
                    
                    // Close the workbook, save the file and free any memory
                    workbook_close(workbook)
                }
                
                DispatchQueue.main.async {
                    self.activityIndicatorView.stopAnimating()
                    UIApplication.shared.endIgnoringInteractionEvents()
                    
                    // show default options for the output Excel file
                    self.documentInteractionController = UIDocumentInteractionController(url: fileURL)
                    self.documentInteractionController.delegate = self
                    self.documentInteractionController.presentOpenInMenu(from: sender.frame, in: self.view, animated: true)
                }
            }
            
        } catch Exception.invalidTextField(let errorMessage) {
            presentAlertController(title: "Ошибка", message: errorMessage)
        } catch {
            print("Unexpected error: \(error).")
            presentAlertController(title: "Ошибка", message: "Unexpected error")
        }
    }
    
    func getProcentPervogoRaschetnogoPerioda() throws -> Double {
        guard let firstDate = dateFormatter.date(from: tfDataVydachiCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат даты выдачи кредита")
        }
        guard let secondDate = dateFormatter.date(from: tfDataPervogoPlatezha.text!) else {
            throw Exception.invalidTextField("Неверный формат даты первого платежа")
        }
        let yearFraction = calculateYearFraction(firstDate: firstDate, secondDate: secondDate);
        let procentnayaStavka = Double(tfProcentnayaStavka.text!)! / 100
        
        return yearFraction * procentnayaStavka
    }
    
    func getProcentPoslednegoRaschetnogoPerioda() throws -> Double {
        guard let srokCredita = Int(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат срока кредита")
        }
        guard let dataVydachiCredita = dateFormatter.date(from: tfDataVydachiCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат даты выдачи кредита")
        }
        guard let dataPervogoPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!) else {
            throw Exception.invalidTextField("Неверный формат даты первого платежа")
        }
        guard let firstDate = Calendar.current.date(byAdding: .month, value: srokCredita - 2, to: dataPervogoPlatezha) else {
            throw Exception.invalidTextField("firstDate = nil")
        }
        guard let secondDate = Calendar.current.date(byAdding: .month, value: srokCredita, to: dataVydachiCredita) else {
            throw Exception.invalidTextField("secondDate = nil")
        }
        let yearFraction = calculateYearFraction(firstDate: firstDate, secondDate: secondDate)
        let procentnayaStavka = Double(tfProcentnayaStavka.text!)! / 100
        
        return yearFraction * procentnayaStavka
    }
    
    func getProcentEzhemesiachnyi() throws -> Double {
        let procentnayaStavka = Double(tfProcentnayaStavka.text!)! / 100
        return procentnayaStavka / 12
    }
    
    func getProcentObichnogoRaschetnogoPerioda(raschiotnyPeriod: Int) throws -> Double {
        guard let dataPervogoPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!) else {
            throw Exception.invalidTextField("Неверный формат даты первого платежа")
        }
        guard let firstDate = Calendar.current.date(byAdding: .month, value: raschiotnyPeriod - 2, to: dataPervogoPlatezha) else {
            throw Exception.invalidTextField("firstDate = nil")
        }
        guard let secondDate = Calendar.current.date(byAdding: .month, value: raschiotnyPeriod - 1, to: dataPervogoPlatezha) else {
            throw Exception.invalidTextField("firstDate = nil")
        }
        let yearFraction = calculateYearFraction(firstDate: firstDate, secondDate: secondDate)
        let procentnayaStavka = Double(tfProcentnayaStavka.text!)! / 100
        
        return yearFraction * procentnayaStavka;
    }
    
    func getUp(procentPervogoRaschetnogoPerioda: Double, procentPoslednegoRaschetnogoPerioda: Double, procentEzhemesiachnyi: Double) throws -> Double {
        guard let srokCredita = Double(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат срока кредита")
        }
        let power = srokCredita - 2;
        return (1 + procentPervogoRaschetnogoPerioda) * pow(1 + procentEzhemesiachnyi, power) * (1 + procentPoslednegoRaschetnogoPerioda)
    }
    
    func getDown(procentPoslednegoRaschetnogoPerioda: Double, procentEzhemesiachnyi: Double) throws -> Double {
        guard let srokCredita = Int(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат срока кредита")
        }
        var down = Double(1)
        let iMax = srokCredita - 1
        for i in 1...iMax {
            let test = (1 + procentPoslednegoRaschetnogoPerioda) * pow(1 + procentEzhemesiachnyi, Double(i-1))
            down += test
        }
        return down
    }
    
    func getAnnuitet(up: Double, down: Double) throws -> Double {
        guard let summaCredita = Double(trimCommaOfString(string: tfSummaCredita.text!)) else {
            throw Exception.invalidTextField("Неверный формат суммы кредита")
        }
        return summaCredita * up / down
    }
    
    func getPereplata(annuitet: Double) throws -> Double {
        guard let summaCredita = Double(trimCommaOfString(string: tfSummaCredita.text!)) else {
            throw Exception.invalidTextField("Неверный формат суммы кредита")
        }
        guard let srokCredita = Double(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат срока кредита")
        }
        return (annuitet * srokCredita) - summaCredita
    }
    
    func getPogashenieOsnovnogoDolga() throws -> Double {
        guard let summaCredita = Double(trimCommaOfString(string: tfSummaCredita.text!)) else {
            throw Exception.invalidTextField("Неверный формат суммы кредита")
        }
        guard let srokCredita = Double(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный формат срока кредита")
        }
        return summaCredita / srokCredita;
    }
    
    // E thirty day months / 360 ("30E/360")
    // https://stackoverflow.com/questions/28277833/how-to-create-a-bank-calendar-with-30-days-each-month
    // https://github.com/OpenGamma/OG-Commons/blob/master/modules/basics/src/main/java/com/opengamma/basics/date/StandardDayCounts.java#L266-L284
    func calculateYearFraction(firstDate: Date, secondDate: Date) -> Double {
        let calendar = Calendar.current
        var d1 = calendar.component(.day, from: firstDate)
        var d2 = calendar.component(.day, from: secondDate)
        if (d1 == 31) {
            d1 = 30;
        }
        if (d2 == 31) {
            d2 = 30;
        }
        return thirty360(
            y1: calendar.component(.year, from: firstDate), m1: calendar.component(.month, from: firstDate), d1: d1,
            y2: calendar.component(.year, from: secondDate), m2: calendar.component(.month, from: secondDate), d2: d2);
    }
    
    // calculate using the standard 30/360 function - 360(y2 - y1) + 30(m2 - m1) + (d2 - d1)) / 360
    func thirty360(y1: Int, m1: Int, d1: Int, y2: Int, m2: Int, d2: Int) -> Double {
        let a = Double(360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1))
        let b = Double(360)
        return a / b
    }
    
    func addItemsToAnnuitetCalculateSheet(calculateSheet: UnsafeMutablePointer<lxw_worksheet>, formatDictionary: [String : UnsafeMutablePointer<lxw_format>]) {
        let firstRow = 3
        let srokCredita = Int(tfSrokCredita.text!)
        
        // Add blue titles
        do {
            let row1 = firstRow - 3
            let row2 = firstRow - 2
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 0, lxw_row_t(row2), 0, "Расчётный период", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 1, lxw_row_t(row2), 1, "Дата погашения", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 2, lxw_row_t(row2), 2, "Сумма платежа", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 3, lxw_row_t(row1), 4, "Погашение", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_write_string(calculateSheet, lxw_row_t(row2), 3, "Процентов банка", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_write_string(calculateSheet, lxw_row_t(row2), 4, "Основного долга", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 5, lxw_row_t(row2), 5, "Остаток\nосновного долга", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 6, lxw_row_t(row2), 6, "Колличество дней", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 7, lxw_row_t(row2), 7, "Коэффициент", formatDictionary[BLUE_TITLE_FORMAT])
            
            // Set column width
            worksheet_set_column(calculateSheet, 0, 0, Double("Расчётный".count + 2), nil)
            worksheet_set_column(calculateSheet, 1, 1, Double("погашения".count + 2), nil)
            worksheet_set_column(calculateSheet, 2, 2, Double("Сумма платежа".count + 2), nil)
            worksheet_set_column(calculateSheet, 3, 3, Double("Процентов банка".count + 2), nil)
            worksheet_set_column(calculateSheet, 4, 4, Double("Основного долга".count + 2), nil)
            worksheet_set_column(calculateSheet, 5, 5, Double("основного долга".count + 2), nil)
            worksheet_set_column(calculateSheet, 6, 6, Double("Колличество".count + 2), nil)
            worksheet_set_column(calculateSheet, 7, 7, Double("Коэффициент".count + 2), nil)
        }
        
        for i in 1...srokCredita! {
            let currentRow = firstRow - 2 + i
            // first row
            if (1 == i) {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=Параметры!A5", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=Параметры!A1", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(Параметры!A4,Параметры!A5,TRUE)", formatDictionary[CENTRE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 7, "=(1+Параметры!A8)*POWER(1+Параметры!A9,A\(firstRow - 1 + i)-1)", formatDictionary[CENTRE_FORMAT]);
            }
            // last row
            else if (srokCredita == i) {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=DATE(YEAR(Параметры!A4),MONTH(Параметры!A4)+\(srokCredita!),DAY(Параметры!A4))", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=F\(firstRow - 2 + i)-E\(firstRow - 2 + i)", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(B\(currentRow),B\(currentRow+1),TRUE)", formatDictionary[CENTRE_FORMAT]);
            }
            // other rows
            else {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=DATE(YEAR(B\(firstRow)),MONTH(B\(firstRow))+\(i-1),DAY(B\(firstRow)))", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=F\(firstRow - 2 + i)-E\(firstRow - 2 + i)", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(B\(currentRow),B\(currentRow+1),TRUE)", formatDictionary[CENTRE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 7, "=(1+Параметры!A8)*POWER(1+Параметры!A9,A\(firstRow - 1 + i)-1)", formatDictionary[CENTRE_FORMAT]);
            }
            worksheet_write_number(calculateSheet, lxw_row_t(currentRow), 0, Double(i), formatDictionary[CENTRE_FORMAT])
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 2, "=Параметры!A1*((1+Параметры!A7)*(1+Параметры!A8)*POWER(1+Параметры!A9,Параметры!A2-2))/(1+SUM(H\(firstRow):H\(firstRow+srokCredita!-2)))", formatDictionary[GREEN_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 3, "=F\(firstRow - 1 + i)*Параметры!A3*G\(firstRow - 1 + i)/360", formatDictionary[YELLOW_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 4, "=C\(firstRow - 1 + i)-D\(firstRow - 1 + i)", formatDictionary[TENGE_FORMAT]);
        }
        
        // Pereplata row
        do {
            let pereplataRow = firstRow + srokCredita! - 1
            worksheet_merge_range(calculateSheet, lxw_row_t(pereplataRow), 0, lxw_row_t(pereplataRow), 1, "Итого:", formatDictionary[CENTRE_FORMAT])
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 2, "=SUM(C\(firstRow):C\(firstRow+srokCredita!-1))", formatDictionary[GREEN_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 3, "=SUM(D\(firstRow):D\(firstRow+srokCredita!-1))", formatDictionary[YELLOW_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 4, "=SUM(E\(firstRow):E\(firstRow+srokCredita!-1))", formatDictionary[TENGE_FORMAT]);
        }
    }
    
    func addItemsToDifferentiatedCalculateSheet(calculateSheet: UnsafeMutablePointer<lxw_worksheet>, formatDictionary: [String : UnsafeMutablePointer<lxw_format>]) {
        let firstRow = 3
        let srokCredita = Int(tfSrokCredita.text!)
        
        // Add blue titles
        do {
            let row1 = firstRow - 3
            let row2 = firstRow - 2
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 0, lxw_row_t(row2), 0, "Расчётный период", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 1, lxw_row_t(row2), 1, "Дата погашения", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 2, lxw_row_t(row2), 2, "Сумма платежа", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 3, lxw_row_t(row1), 4, "Погашение", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_write_string(calculateSheet, lxw_row_t(row2), 3, "Процентов банка", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_write_string(calculateSheet, lxw_row_t(row2), 4, "Основного долга", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 5, lxw_row_t(row2), 5, "Остаток\nосновного долга", formatDictionary[BLUE_TITLE_FORMAT])
            worksheet_merge_range(calculateSheet, lxw_row_t(row1), 6, lxw_row_t(row2), 6, "Колличество дней", formatDictionary[BLUE_TITLE_FORMAT])
            
            // Set column width
            worksheet_set_column(calculateSheet, 0, 0, Double("Расчётный".count + 2), nil)
            worksheet_set_column(calculateSheet, 1, 1, Double("погашения".count + 2), nil)
            worksheet_set_column(calculateSheet, 2, 2, Double("Сумма платежа".count + 2), nil)
            worksheet_set_column(calculateSheet, 3, 3, Double("Процентов банка".count + 2), nil)
            worksheet_set_column(calculateSheet, 4, 4, Double("Основного долга".count + 2), nil)
            worksheet_set_column(calculateSheet, 5, 5, Double("основного долга".count + 2), nil)
            worksheet_set_column(calculateSheet, 6, 6, Double("Колличество".count + 2), nil)
        }
        
        for i in 1...srokCredita! {
            let currentRow = firstRow - 2 + i
            // first row
            if (1 == i) {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=Параметры!A5", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=Параметры!A1", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(Параметры!A4,Параметры!A5,TRUE)", formatDictionary[CENTRE_FORMAT]);
            }
            // last row
            else if (srokCredita == i) {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=DATE(YEAR(Параметры!A4),MONTH(Параметры!A4)+\(srokCredita!),DAY(Параметры!A4))", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=F\(firstRow - 2 + i)-E\(firstRow - 2 + i)", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(B\(currentRow),B\(currentRow+1),TRUE)", formatDictionary[CENTRE_FORMAT]);
            }
            // other rows
            else {
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 1, "=DATE(YEAR(B\(firstRow)),MONTH(B\(firstRow))+\(i-1),DAY(B\(firstRow)))", formatDictionary[DATE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 5, "=F\(firstRow - 2 + i)-E\(firstRow - 2 + i)", formatDictionary[TENGE_FORMAT]);
                worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 6, "=DAYS360(B\(currentRow),B\(currentRow+1),TRUE)", formatDictionary[CENTRE_FORMAT]);
            }
            worksheet_write_number(calculateSheet, lxw_row_t(currentRow), 0, Double(i), formatDictionary[CENTRE_FORMAT])
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 2, "=D\(firstRow - 1 + i)+E\(firstRow - 1 + i)", formatDictionary[GREEN_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 3, "=F\(firstRow - 1 + i)*Параметры!A3*G\(firstRow - 1 + i)/360", formatDictionary[YELLOW_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(currentRow), 4, "=Параметры!A1/Параметры!A2", formatDictionary[TENGE_FORMAT]);
        }
        
        // Pereplata row
        do {
            let pereplataRow = firstRow + srokCredita! - 1
            worksheet_merge_range(calculateSheet, lxw_row_t(pereplataRow), 0, lxw_row_t(pereplataRow), 1, "Итого:", formatDictionary[CENTRE_FORMAT])
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 2, "=SUM(C\(firstRow):C\(firstRow+srokCredita!-1))", formatDictionary[GREEN_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 3, "=SUM(D\(firstRow):D\(firstRow+srokCredita!-1))", formatDictionary[YELLOW_TENGE_FORMAT]);
            worksheet_write_formula(calculateSheet, lxw_row_t(pereplataRow), 4, "=SUM(E\(firstRow):E\(firstRow+srokCredita!-1))", formatDictionary[TENGE_FORMAT]);
        }
    }
    
    func addItemsToParamsSheet(paramsSheet: UnsafeMutablePointer<lxw_worksheet>, formatDictionary: [String : UnsafeMutablePointer<lxw_format>]) {
        
        let summaCredita = Double(trimCommaOfString(string: tfSummaCredita.text!))
        let srokCredita = Double(tfSrokCredita.text!)
        let procentnayaStavka = Double(tfProcentnayaStavka.text!)
        let paymentTypeText = tfPaymentType.text
        let dataVydachiCredita = dateFormatter.date(from: tfDataVydachiCredita.text!)
        let dataPervogoPlatezha = dateFormatter.date(from: tfDataPervogoPlatezha.text!)
        let calendar = Calendar.current
        
        var dataVydachiCreditaExcel = lxw_datetime(year: Int32(calendar.component(.year, from: dataVydachiCredita!)),
                                                   month: Int32(calendar.component(.month, from: dataVydachiCredita!)),
                                                   day: Int32(calendar.component(.day, from: dataVydachiCredita!)), hour: 0, min: 0, sec: 0)
        
        var dataPervogoPlatezhaExcel = lxw_datetime(year: Int32(calendar.component(.year, from: dataPervogoPlatezha!)),
                                                    month: Int32(calendar.component(.month, from: dataPervogoPlatezha!)),
                                                    day: Int32(calendar.component(.day, from: dataPervogoPlatezha!)), hour: 0, min: 0, sec: 0)
        
        // Write formatted data.
        worksheet_write_number(paramsSheet, 0, 0, summaCredita!, formatDictionary[TENGE_FORMAT])
        worksheet_write_string(paramsSheet, 0, 1, "Сумма кредита", formatDictionary[BORDER_FORMAT])
        worksheet_write_number(paramsSheet, 1, 0, srokCredita!, formatDictionary[CENTRE_FORMAT])
        worksheet_write_string(paramsSheet, 1, 1, "Кол-во расчётных периодов (срок кредита)", formatDictionary[BORDER_FORMAT])
        worksheet_write_number(paramsSheet, 2, 0, procentnayaStavka! / 100, formatDictionary[PERCENT_FORMAT])
        worksheet_write_string(paramsSheet, 2, 1, "Процентная ставка годовая", formatDictionary[BORDER_FORMAT])
        worksheet_write_datetime(paramsSheet, 3, 0, &dataVydachiCreditaExcel, formatDictionary[DATE_FORMAT]);
        worksheet_write_string(paramsSheet, 3, 1, "Дата выдачи кредита", formatDictionary[BORDER_FORMAT])
        worksheet_write_datetime(paramsSheet, 4, 0, &dataPervogoPlatezhaExcel, formatDictionary[DATE_FORMAT]);
        worksheet_write_string(paramsSheet, 4, 1, "Дата первого платежа", formatDictionary[BORDER_FORMAT])
        worksheet_write_string(paramsSheet, 5, 0, paymentTypeText, formatDictionary[CENTRE_FORMAT])
        worksheet_write_string(paramsSheet, 5, 1, "Вид платежа", formatDictionary[BORDER_FORMAT])
        
        // Аннуитетный
        if paymentTypes[0] == tfPaymentType.text {
            let firstRow = 3
            let srokCredita = Int(tfSrokCredita.text!)
            worksheet_write_formula(paramsSheet, 6, 0, "=A3*Вычисление!G\(firstRow)/360", formatDictionary[PERCENT_FORMAT]);
            worksheet_write_string(paramsSheet, 6, 1, "Процентная ставка в первом расчётном периоде", formatDictionary[BORDER_FORMAT])
            worksheet_write_formula(paramsSheet, 7, 0, "=A3*Вычисление!G\(firstRow+srokCredita!-1)/360", formatDictionary[PERCENT_FORMAT])
            worksheet_write_string(paramsSheet, 7, 1, "Процентная ставка в последнем расчётном периоде", formatDictionary[BORDER_FORMAT])
            worksheet_write_formula(paramsSheet, 8, 0, "=A3/12", formatDictionary[PERCENT_FORMAT]);
            worksheet_write_string(paramsSheet, 8, 1, "Процентная ставка в остальных расчётных периодах", formatDictionary[BORDER_FORMAT])
        }
        
        // Set column width
        worksheet_set_column(paramsSheet, 0, 0, Double(paymentTypes[1].count + 2), nil)
        worksheet_set_column(paramsSheet, 1, 2, Double("Процентная ставка в остальных расчётных периодах".count + 2), nil)
    }
    
    func createFormatDictionary(workbook: UnsafeMutablePointer<lxw_workbook>) -> [String: UnsafeMutablePointer<lxw_format>] {
        
        var formatDictionary = [String: UnsafeMutablePointer<lxw_format>]()
        
        // Borderd cell format
        do {
            let format = workbook_add_format(workbook)
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            formatDictionary[BORDER_FORMAT] = format
        }
        
        // Borderd cell format with center alignment
        do {
            let format = workbook_add_format(workbook)
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            formatDictionary[CENTRE_FORMAT] = format
        }
        
        // Percent format
        do {
            let format = workbook_add_format(workbook)
            format_set_num_format(format, "0.00%");
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            formatDictionary[PERCENT_FORMAT] = format
        }
        
        // Accounting format
        do {
            let format = workbook_add_format(workbook)
            format_set_num_format(format, "#,##0.00тг");
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            formatDictionary[TENGE_FORMAT] = format
        }
        
        // Yellow accounting format
        do {
            let format = workbook_add_format(workbook)
            format_set_num_format(format, "#,##0.00тг");
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            format_set_bg_color(format, lxw_color_t(0xffff99));
            formatDictionary[YELLOW_TENGE_FORMAT] = format
        }
        
        // Green accounting format
        do {
            let format = workbook_add_format(workbook)
            format_set_num_format(format, "#,##0.00тг");
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            format_set_bg_color(format, lxw_color_t(0xccffcc));
            formatDictionary[GREEN_TENGE_FORMAT] = format
        }
        
        // Date format
        do {
            let format = workbook_add_format(workbook)
            format_set_num_format(format, dateFormatter.dateFormat);
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            formatDictionary[DATE_FORMAT] = format
        }
        
        // Create a bold font
        do {
            let format = workbook_add_format(workbook)
            format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue));
            format_set_align(format, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue));
            format_set_border(format, UInt8(LXW_BORDER_THIN.rawValue))
            format_set_bg_color(format, lxw_color_t(0x99ccff));
            format_set_text_wrap(format);
            formatDictionary[BLUE_TITLE_FORMAT] = format
        }
        
        return formatDictionary
    }
    
    func validateTextField() throws {
        
        // Валидация суммы кредита
        guard trimCommaOfString(string: tfSummaCredita.text!).isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали сумму кредита")
        }
        guard (Double(trimCommaOfString(string: tfSummaCredita.text!)) != nil) else {
            throw Exception.invalidTextField("Неверный формат суммы кредита")
        }
        
        // Валидация срока кредита
        guard tfSrokCredita.text?.isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали срок кредита")
        }
        guard let srokCredita = Int(tfSrokCredita.text!) else {
            throw Exception.invalidTextField("Неверный срок кредита")
        }
        guard 6 * 12 <= srokCredita else {
            throw Exception.invalidTextField("Срок кредита не может быть меньше 72 месяцев")
        }
        guard 25 * 12 >= srokCredita else {
            throw Exception.invalidTextField("Срок кредита не может быть больше 300 месяцев")
        }
        
        // Валидация процентной ставки
        guard tfProcentnayaStavka.text?.isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали процентную ставку")
        }
        guard (Double(tfProcentnayaStavka.text!) != nil) else {
            throw Exception.invalidTextField("Неверный формат процентной ставки")
        }
        
        // Валидация даты выдачи кредита и даты первого платежа
        guard let dataVydachiCreditaText = tfDataVydachiCredita.text else {
            throw Exception.invalidTextField("Дата выдачи кредита = nil")
        }
        guard let dataPervogoPlatezhaText = tfDataPervogoPlatezha.text else {
            throw Exception.invalidTextField("Дата выдачи кредита = nil")
        }
        guard dataVydachiCreditaText.isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали дату выдачи кредита")
        }
        guard dataPervogoPlatezhaText.isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали дату первого платежа")
        }
        guard let dataVydachiCredita = dateFormatter.date(from: dataVydachiCreditaText) else {
            throw Exception.invalidTextField("Неверный формат даты выдачи кредита")
        }
        guard let dataPervogoPlatezha = dateFormatter.date(from: dataPervogoPlatezhaText) else {
            throw Exception.invalidTextField("Неверный формат даты первого платежа")
        }
        guard dataPervogoPlatezha >= dataVydachiCredita else {
            throw Exception.invalidTextField("Дата первого платежа не может быть раньше даты выдачи кредита")
        }
        guard let dataVydachiCreditaPlusMonth = Calendar.current.date(byAdding: .month, value: 1, to: dataVydachiCredita) else {
            throw Exception.invalidTextField("dataVydachiCreditaPlusMonth = nil")
        }
        guard dataVydachiCreditaPlusMonth >= dataPervogoPlatezha else {
            throw Exception.invalidTextField("Разница от даты выдачи кредита до даты первого платежа не может превышать один месяц")
        }
        
        // Валидация вида платежа
        guard tfPaymentType.text?.isEmpty == false else {
            throw Exception.invalidTextField("Вы не указали вид платежа")
        }
        guard paymentTypes[0] == tfPaymentType.text || paymentTypes[1] == tfPaymentType.text else {
            throw Exception.invalidTextField("Неверный вид платежа")
        }
    }
    
    func presentAlertController(title: String, message: String) {
        let alert = UIAlertController(title: title, message: message, preferredStyle: .alert)
        alert.addAction(UIAlertAction(title: "Ok", style: .default, handler: nil))
        self.present(alert, animated: true, completion: nil)
    }
    
    @objc func srokCreditaDidChange(_ textField: UITextField) {
        UserDefaults.standard.set(textField.text, forKey: srok_credita)
    }
    
    @objc func procentnayaStavkaDidChange(_ textField: UITextField) {
        UserDefaults.standard.set(textField.text, forKey: procentnaya_stavka)
    }
    
    @objc func handleDataVydachiCreditaPicker(sender: UIDatePicker) {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd"
        tfDataVydachiCredita.text = dateFormatter.string(from: sender.date)
        UserDefaults.standard.set(tfDataVydachiCredita.text, forKey: data_vydachi_credita)
    }
    
    @objc func handleDataPervogoPlatezhaPicker(sender: UIDatePicker) {
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "yyyy-MM-dd"
        tfDataPervogoPlatezha.text = dateFormatter.string(from: sender.date)
        UserDefaults.standard.set(tfDataPervogoPlatezha.text, forKey: data_pervogo_platezha)
    }
    
    func numberOfComponents(in pickerView: UIPickerView) -> Int {
        return 1
    }
    
    func pickerView(_ pickerView: UIPickerView, titleForRow row: Int, forComponent component: Int) -> String? {
        return paymentTypes[row]
    }
    
    func pickerView(_ pickerView: UIPickerView, numberOfRowsInComponent component: Int) -> Int {
        return paymentTypes.count
    }
    
    func pickerView(_ pickerView: UIPickerView, didSelectRow row: Int, inComponent component: Int) {
        tfPaymentType.text = paymentTypes[row]
        UserDefaults.standard.set(tfPaymentType.text, forKey: payment_type)
    }
    
    //Calls this function when the tap is recognized.
    @objc func dismissKeyboard() {
        //Causes the view (or one of its embedded text fields) to resign the first responder status.
        view.endEditing(true)
    }
    
    func textField(_ textField: UITextField, shouldChangeCharactersIn range: NSRange, replacementString newString: String) -> Bool {
        var string = newString
        if string == "," {
            string = "."
        }
        if ((string == "0" || string == "") && (textField.text! as NSString).range(of: ".").location < range.location) {
            return true
        }

        // First check whether the replacement string's numeric...
        let cs = NSCharacterSet(charactersIn: "0123456789.").inverted
        let filtered = string.components(separatedBy: cs)
        let component = filtered.joined(separator: "")
        let isNumeric = string == component

        // Then if the replacement string's numeric, or if it's
        // a backspace, or if it's a decimal point and the text
        // field doesn't already contain a decimal point,
        // reformat the new complete number using
        if isNumeric {
            let formatter = NumberFormatter()
            formatter.locale = Locale(identifier: "en_US")
            formatter.numberStyle = .decimal
            formatter.maximumFractionDigits = 2
            // Combine the new text with the old; then remove any
            // commas from the textField before formatting
            let newString = (textField.text! as NSString).replacingCharacters(in: range, with: string)
            let numberWithOutCommas = newString.replacingOccurrences(of: ",", with: "")
            let number = formatter.number(from: numberWithOutCommas)
            if number != nil {
                var formattedString = formatter.string(from: number!)
                // If the last entry was a decimal or a zero after a decimal,
                // re-add it here because the formatter will naturally remove
                // it.
                if string == "." && range.location == textField.text?.count {
                    formattedString = formattedString?.appending(".")
                }
                textField.text = formattedString
                let test = String(format:"%f", Double(truncating: number!))
                if textField == tfSummaCredita {
                    UserDefaults.standard.set(test, forKey: summa_credita)
                }
                if textField == tfProcentnayaStavka {
                    UserDefaults.standard.set(test, forKey: procentnaya_stavka)
                }
            } else {
                textField.text = nil
                if textField == tfSummaCredita {
                    UserDefaults.standard.set(nil, forKey: summa_credita)
                }
                if textField == tfProcentnayaStavka {
                    UserDefaults.standard.set(nil, forKey: procentnaya_stavka)
                }
            }
        }
        return false
    }
    
    func trimCommaOfString(string: String) -> String {
        if string.contains(",") {
            return string.replacingOccurrences(of: ",", with: "")
        } else {
            return string
        }
    }
    
}

