using Bytescout.Spreadsheet;
using Bytescout.Spreadsheet.Constants;

namespace Report
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            //Application.Run(new Form1());


            int today1 = 0;

            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile("C:\\Users\\ITSkopje\\OneDrive - GFG Alliance\\Desktop\\Izvestaj2024.xlsx");
            Worksheet worksheet2 = document.Workbook.Worksheets[0];

            Spreadsheet documentIn = new Spreadsheet();
            documentIn.LoadFromFile("C:\\Users\\ITSkopje\\OneDrive - GFG Alliance\\Desktop\\Deliveries.xlsx");
            Worksheet worksheet1 = documentIn.Workbook.Worksheets[0];
            
            double[] suma = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            //DateTime den;

            int lastRow = worksheet1.UsedRangeRowMax;
            for (int row = 1; row <= lastRow; row++) {
                if (Convert.ToString(worksheet1.Cell(row, 10).Value) == "521") {
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "01.12.2024")
                    {
                        suma[0] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "02.12.2024") {
                        suma[1] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);  
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "03.12.2024")
                    {
                        suma[2] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "04.12.2024")
                    {
                        suma[3] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "05.12.2024")
                    {
                        suma[4] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "06.12.2024")
                    {
                        suma[5] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "07.12.2024")
                    {
                        suma[6] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "08.12.2024")
                    {
                        suma[7] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "09.12.2024")
                    {
                        suma[8] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "10.12.2024")
                    {
                        suma[9] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "11.12.2024")
                    {
                        suma[10] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "12.12.2024")
                    {
                        suma[11] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "13.12.2024")
                    {
                        suma[12] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "14.12.2024")
                    {
                        suma[13] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "15.12.2024")
                    {
                        suma[14] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "16.12.2024")
                    {
                        suma[15] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "17.12.2024")
                    {
                        suma[16] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "18.12.2024")
                    {
                        suma[17] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "19.12.2024")
                    {
                        suma[18] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "20.12.2024")
                    {
                        suma[19] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "21.12.2024")
                    {
                        suma[20] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "22.12.2024")
                    {
                        suma[21] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "23.12.2024")
                    {
                        suma[22] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "24.12.2024")
                    {
                        suma[23] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "25.12.2024")
                    {
                        suma[24] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "26.12.2024")
                    {
                        suma[25] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "27.12.2024")
                    {
                        suma[26] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "28.12.2024")
                    {
                        suma[27] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "29.12.2024")
                    {
                        suma[28] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "30.12.2024")
                    {
                        suma[29] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                    if (Convert.ToString(worksheet1.Cell(row, 51).Value) == "31.12.2024")
                    {
                        suma[30] += Convert.ToDouble(worksheet1.Cell(row, 19).Value);
                    }
                }
            }
            
          



            worksheet2.Cell(1, 1).Value = "Den";
            worksheet2.Cell(2, 1).Value = "Primen materijal";
            worksheet2.Cell(3, 1).Value = "Otpremen materijal";
            for (int boja = 1; boja <= 3; boja++)
            {

                Color headerColor = Color.FromArgb(75, 172, 198);
                Color contentColor = Color.FromArgb(141, 180, 227);
                worksheet2.Cell(boja, 1).FillPattern = PatternStyle.Solid;
                worksheet2.Cell(boja, 1).FillPatternForeColor = headerColor;

                worksheet2.Cell(boja, 1).RightBorderStyle = LineStyle.Medium;
                worksheet2.Cell(boja, 1).LeftBorderStyle = LineStyle.Medium;
                worksheet2.Cell(boja, 1).BottomBorderStyle = LineStyle.Medium;
                worksheet2.Cell(boja, 1).TopBorderStyle = LineStyle.Medium;
                worksheet2.Cell(boja, 1).AlignmentHorizontal = AlignmentHorizontal.Centered;


            }
            //worksheet2.Cell(28,5).ValueAsDateTime = DateTime.Now.Day;
            //worksheet2.Cell(30,5).NumberFormatString = "dd";
            //today1 = Convert.ToInt32(worksheet2.Cell(28, 5).Value);

            today1 = DateTime.Today.Day;
            for (int i = 1; i < today1; i++) {
                worksheet2.Cell(1, i + 1).Value = i;
                worksheet2.Cell(1, i+1).AlignmentHorizontal = AlignmentHorizontal.Centered;
                worksheet2.Cell(1, i+1).AlignmentVertical = AlignmentVertical.Centered;
                worksheet2.Cell(1, i).RightBorderStyle = LineStyle.Medium;
                worksheet2.Cell(1, i).LeftBorderStyle = LineStyle.Medium;
                worksheet2.Cell(1, i).BottomBorderStyle = LineStyle.Medium;
                worksheet2.Cell(1, i).TopBorderStyle = LineStyle.Medium;
                for (int j = 1; j <= 3; j++) {
                    worksheet2.Cell(j, i+1).RightBorderStyle = LineStyle.Medium;
                    worksheet2.Cell(j, i+1).LeftBorderStyle = LineStyle.Medium;
                    worksheet2.Cell(j, i+1).BottomBorderStyle = LineStyle.Medium;
                    worksheet2.Cell(j, i+1).TopBorderStyle = LineStyle.Medium;
                }
                worksheet2.Cell(1, i+1).Font = new Font("Arial", 10, FontStyle.Bold | FontStyle.Italic);
                worksheet2.Cell(2, i+1).Font = new Font("Arial", 10, FontStyle.Bold | FontStyle.Italic);
                worksheet2.Cell(3, i+1).Font = new Font("Arial", 10, FontStyle.Bold | FontStyle.Italic);
            }
            //worksheet2.Cell(29, 4).Value = today1;
            worksheet2.Columns[1].Width = 150;
            for (int h = 1; h <= 3; h++) { 

              worksheet2.Rows[h].Height = 50;
              worksheet2.Cell(h, 1).AlignmentHorizontal = AlignmentHorizontal.Centered;
                worksheet2.Cell(h, 1).AlignmentVertical = AlignmentVertical.Centered;
            }

            for (int i = 1; i < today1; i++)
            {
                if (i < today1)
                {
                    worksheet2.Cell(2, i + 1).Value = suma[i - 1];
                }
                else break;
                
            }

            document.SaveAs("C:\\Users\\ITSkopje\\OneDrive - GFG Alliance\\Desktop\\Izvestaj2024.xlsx");
            document.Close();




        }
    }
}