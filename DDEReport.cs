using GrapeCity.Documents.Excel;
using System.IO;
using System.Text.Json;

namespace DDExcelReport
{
    public static class DDEReport
    {
        public static void CreateReport()
        {

            // データソース（JSON）

            var jsonString_p = File.ReadAllText("Data/publisher.json");
            var publisher = JsonSerializer.Deserialize<Publisher>(jsonString_p);
            var jsonString_c = File.ReadAllText("Data/customers.json");
            var customerdata = JsonSerializer.Deserialize<Customer[]>(jsonString_c);

            // テンプレート
            var template_file = "Templates/SimpleInvoiceJP_Template.xlsx";

            // ライセンスキー
            string key = DDLIC.Key_V7.Excel;
            Workbook.SetLicenseKey(key);

            // 新しいワークブックを生成
            var workbook = new Workbook();
            // テンプレートを読み込む
            workbook.Open(template_file);
            Workbook.FontsFolderPath = @"./Fonts";
            // データソースを追加
            workbook.AddDataSource("pubds", publisher); // 発行者データ
            workbook.AddDataSource("ds", customerdata); // 顧客データ 
            workbook.AddDataSource("Env", RuntimeInfo.getEnvironmentInfo()); // 備考欄に付加情報

            // テンプレート処理を呼び出し
            workbook.ProcessTemplate();

            // Excelファイルに保存
            workbook.Save("result.xlsx");
            System.Console.WriteLine("Excelファイル生成完了");

            // PDFファイルに保存
            workbook.Save("result.pdf", SaveFileFormat.Pdf);
            System.Console.WriteLine("PDFファイル生成完了");

            // テンプレートを読み込んで退避
            var temp_workbook = new Workbook();
            temp_workbook.Open(template_file);
            // 退避したテンプレートの各シートを結果ファイルの最後にコピーして追加
            foreach (var item in temp_workbook.Worksheets)
            {
                item.CopyAfter(workbook.Worksheets[workbook.Worksheets.Count - 1]);
            }
            // テンプレートを追加してExcelファイルに保存
            workbook.Save("result_with_template.xlsx");

        }
    }
}
