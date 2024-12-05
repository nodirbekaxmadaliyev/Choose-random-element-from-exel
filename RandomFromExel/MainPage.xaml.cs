using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;

namespace RandomFromExel
{
    public partial class MainPage : ContentPage
    {
        private List<string> words = new List<string>();
        public MainPage()
        {
            InitializeComponent();
        }
        private void Kirit(object sender, EventArgs e)
        {
            words.Clear();
            var filePath = inputEntry.Text;

             // Excel faylini yuklash
            using (var workbook = new XLWorkbook(filePath))
            {
                // Birinchi varaqni olish
                var worksheet = workbook.Worksheet(1);

                // Faqat to'liq ishlatilgan qatorlardan foydalanish
                foreach (var row in worksheet.RowsUsed())
                {
                    // Har bir qatorning birinchi ustunidan ma'lumotni olish
                    var cellValue = row.Cell(1).GetString(); // 1-ustun
                    words.Add(cellValue);
                }
            }
        }
        private void Tanla(object sender, EventArgs e)
        {
            int n = words.Count;
            Random random = new Random();

            int randomInt = random.Next();
            resultLabel.Text = "Natija: " + words[randomInt % n];
        }
    }

}
