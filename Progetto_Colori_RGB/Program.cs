using ClosedXML.Excel;
class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/Progetto_Colori_RGB.xlsx";

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Colore RGB");

            int r, g, b;
            do
            {
                Console.WriteLine("inserire il la quantità di rosso, 0-255");
                r = Convert.ToInt16(Console.ReadLine());
                if (r < 0 || r > 255)
                {
                    Console.WriteLine("valore non valido,");
                }
            }while (r < 0 || r > 255);
            do
            {
                Console.WriteLine("inserire il la quantità di verde, 0-255");
                g = Convert.ToInt16(Console.ReadLine());
                if (g < 0 || g > 255)
                {
                    Console.WriteLine("valore non valido,");
                }
            } while (g < 0 || g > 255);
            do
            {
                Console.WriteLine("inserire il la quantità di blu, 0-255");
                b = Convert.ToInt16(Console.ReadLine());
                if (b < 0 || b > 255)
                {
                    Console.WriteLine("valore non valido,");
                }
            } while (b < 0 || b > 255);

            worksheet.Cell(1, 1).Value = $"{r}, {g}, {b}";
            worksheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);

            workbook.SaveAs(filePath);
            Console.WriteLine($"\nIl File '{filePath}' e' stato creato con successo!\n");
        }
    }
}
