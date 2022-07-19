//
//  ExportService.swift
//  ExcelExportSample
//
//  Created by Yakup Yazar on 11.05.2022.
//

import Foundation
import UIKit
import xlsxwriter

    class DatabaseExcel {
        let fruit1 = Fruit(salePrice: 12.2, purchasePrice: 11.5, saleType: "Kg", name: "Elma") // veri modelinden oluşturulan alt ürün
        let fruit2 = Fruit(salePrice: 20.5, purchasePrice: 16.7, saleType: "Adet", name: "Ananas") // veri modelinden oluşturulan alt ürün
        let fruit3 = Fruit(salePrice: 19.5, purchasePrice: 15.04, saleType: "Kg", name: "Muz") // veri modelinden oluşturulan alt ürün
        let fruit4 = Fruit(salePrice: 25.0, purchasePrice: 20.0, saleType: "Kg", name: "Çilek") // veri modelinden oluşturulan alt ürün
        
        lazy var fruitList: [Fruit] = {
            return [fruit1, fruit2,fruit3,fruit4] // alt ürünlerden oluşturulan ürün listesi
        }()
        
    }
    // END OF FAKE DATA

final class ExportService {
    
    let filename = "/gelis_gidis_fiyatlari.xlsx"
    let cell_width: Double = 64 // hücre genişliği
    let cell_height: Double = 50 // hücre yüksekliği

    var workbook: UnsafeMutablePointer<lxw_workbook>?
    var worksheet: UnsafeMutablePointer<lxw_worksheet>?
    var format_header: UnsafeMutablePointer<lxw_format>?
    var format_1: UnsafeMutablePointer<lxw_format>?
    
    private var writingLine: UInt32 = 0
    private var needWriterPreparation = false
    
    init() {
        prepareXlsWriter()
    }
    
    private func docDirectoryPath() -> String{
        let dirPaths = NSSearchPathForDirectoriesInDomains(.documentDirectory,
                                                           .userDomainMask,
                                                           true)
        return dirPaths[0]
    }
    
    /// Prepare the xlsx objects
    private func prepareXlsWriter() {
        print("open \(docDirectoryPath())")
        var destination_path = docDirectoryPath()
        destination_path.append(filename) // hedef klasör yolunda dosyayı açıyor
        workbook = workbook_new(destination_path)
        worksheet = workbook_add_worksheet(workbook, nil)
        // Add style
        format_header = workbook_add_format(workbook)
        format_set_bold(format_header)
        format_1 = workbook_add_format(workbook)
        format_set_bg_color(format_1, 0xDDDDDD)
        needWriterPreparation = false
        
        
        
    }
    
    private func minRatio(left: (Double, Double), right: (Double, Double)) -> Double {
        min(left.0 / right.0, left.1 / right.1)
    }
    
    /// İlk Satırı Header Olarak Belirledik
    private func buildHeader() {
        writingLine = 0
        let format = format_header
        format_set_bold(format)
        worksheet_write_string(worksheet, writingLine, 0, "Meyve Adı", format)
        worksheet_write_string(worksheet, writingLine, 1, "Satış Tipi", format)
        worksheet_write_string(worksheet, writingLine, 2, "Geliş Fiyatı", format)
        worksheet_write_string(worksheet, writingLine, 3, "Satış Fiyatı", format)
    }
    
    /// For döngüsü ile çağırılan yeni satır yazma işlemi
    private func buildNewLine(fruit: Fruit) {
        writingLine += 1 //her çağırılışında 1 satır alta yazıyor.
        let lineFormat = (writingLine % 2 == 1) ? format_1 : nil // 1 satır gri 1 satır beyaz formatı
        worksheet_write_string(worksheet, writingLine, 0, fruit.name, lineFormat)
        worksheet_write_string(worksheet, writingLine, 1, fruit.saleType, lineFormat)
        worksheet_write_number(worksheet, writingLine, 2, Double(fruit.purchasePrice), lineFormat)
        worksheet_write_number(worksheet, writingLine, 3, Double(fruit.salePrice), lineFormat)
    }
        
    
    func getFileDir()-> String
    {
        return docDirectoryPath()
    }
    
    // Xls dosyası oluşturma ya da yeniden açma işlemleri
    func export() {
        if(needWriterPreparation == true){
            prepareXlsWriter()
        }
        
        buildHeader()
       
        let list = DatabaseExcel().fruitList

        for fruit in list {
            buildNewLine(fruit: fruit)
        }
        workbook_close(workbook)
        needWriterPreparation = true
       
    }

    
    
    
    
    
    
    
}
