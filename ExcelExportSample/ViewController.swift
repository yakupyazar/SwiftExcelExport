//
//  ViewController.swift
//  ExcelExportSample
//
//  Created by Yakup Yazar on 11.05.2022.
//

import UIKit
import MessageUI
class ViewController: UIViewController, MFMailComposeViewControllerDelegate {

    override func viewDidLoad() {
        super.viewDidLoad()
        // İhtiyaca göre bu method farklı yerlerde kullanılabilir, örnek olması açısından viewDidLoad'a eklendi
        ExportService().export()
    
    }
   
    @IBAction func btnSendMail(_ sender: Any) {
       if MFMailComposeViewController.canSendMail() {
          let mail = MFMailComposeViewController()
          mail.setToRecipients(["sample@sample.com"])
          mail.setSubject("Dask Excel")
          mail.setMessageBody("Dosyanız ektedir", isHTML: true)
          mail.mailComposeDelegate = self
     
           let filePath = ExportService().getFileDir()+"/gelis_gidis_fiyatlari.xlsx" //dosyamızın bulunduğu path bilgisini değişkene aktarıyoruz.
          
            let data = NSData(contentsOfFile: filePath)
           mail.addAttachmentData(data as! Data, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName: "olusturulan_excel_dosyasi.xlsx")
           //mail.addAttachmentData(oluşturduğumuz dosyanın yolu gelecek( data değişkeni) as! Data, mimeType: hangi formatta dosya dönderecekseniz, onun mimetype  türünü yazmalısınız  , fileName: burada mailde gözükecek adı belirtiyorsunuz, dosya adı ile aynı olabilir )
          present(mail, animated: true)
       }
       else {
          print("Email gönderilemedi, Cihazda Kurulu Olan Mail Adresinin Doğruluğunu Kontrol Edin")
       }
    }
    
    func mailComposeController(_ controller: MFMailComposeViewController, didFinishWith result:          MFMailComposeResult, error: Error?) {
          if let _ = error {
             self.dismiss(animated: true, completion: nil)
          }
          switch result {
             case .cancelled:
             print("İptal Edildi")
             break
             case .sent:
             print("Mail Başarıyla Gönderildi")
             break
             case .failed:
             print("Mail Gönderim Hatası, Cihazda Kurulu Olan Mail Adresinin Doğruluğunu Kontrol Edin")
             break
             default:
             break
          }
          controller.dismiss(animated: true, completion: nil)
       }
}

