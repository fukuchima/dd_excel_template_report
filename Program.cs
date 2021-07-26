using GrapeCity.Documents.Excel;
using System.IO;
using System.Text.Json;

namespace DDExcelReport
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello DioDocs!");

            // データソース（JSON）

            var jsonString_p = File.ReadAllText("Data/publisher.json");
            var publisher = JsonSerializer.Deserialize<Publisher>(jsonString_p);
            var jsonString_c = File.ReadAllText("Data/customers.json");
            var customerdata = JsonSerializer.Deserialize<Customer[]>(jsonString_c);

            // テンプレート
            var template_file = "Templates/SimpleInvoiceJP_Template.xlsx";

            // ライセンスキー
            // string key = DDLIC.Key_V4.Excel;
            // Workbook.SetLicenseKey(key);

            // 新しいワークブックを生成
            var workbook = new Workbook();
            // テンプレートを読み込む
            workbook.Open(template_file);
            
            // データソースを追加
            workbook.AddDataSource("pubds", publisher);
            workbook.AddDataSource("ds", customerdata);

            // テンプレート処理を呼び出し
            workbook.ProcessTemplate();

            // Excelファイルに保存
            workbook.Save("result.xlsx");

            // PDFファイルに保存
            workbook.Save("result.pdf", SaveFileFormat.Pdf);
            

            // テンプレートを読み込んで退避
            var temp_workbook = new Workbook();
            temp_workbook.Open(template_file);
            // 退避したテンプレートの各シートを結果ファイルの最後にコピーして追加
            foreach (var item in temp_workbook.Worksheets)
            {
                item.CopyAfter(workbook.Worksheets[workbook.Worksheets.Count-1]);
            }
            // テンプレートを追加してExcelファイルに保存
            workbook.Save("result_with_template.xlsx");

        }
    }
}
