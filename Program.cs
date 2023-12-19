// See https://aka.ms/new-console-template for more information
using Xceed.Document.NET;
using Xceed.Words.NET;


while (true)
{
    Console.Write("請輸入年月份（例如：202312）：");

    // 輸入年月份  
    string yearMonth = Console.ReadLine();

    // 判斷是否輸入 'q' 以退出程式
    if (yearMonth.ToLower() == "q")
    {
        break;
    }

    

    // 設定檔案路徑
    string filePath = $"E:\\001. TQE\\004. 備份規劃\\ISMS-B-OO-OO備份紀錄檢核表.docx";
    string outputFileName = $"E:\\001. TQE\\004. 備份規劃\\ISMS-B-OO-OO備份紀錄檢核表_{yearMonth}.docx";

    // 讀取 DOCX 檔案
    using (DocX document = DocX.Load(filePath))
    {
        // 獲取表格
        Table table = document.Tables[0];

        // 設定當月的第一天
        DateTime firstDayOfMonth = new DateTime(int.Parse(yearMonth.Substring(0, 4)), int.Parse(yearMonth.Substring(4, 2)), 1);

        // 定義灰色背景色
        System.Drawing.Color grayColor = System.Drawing.ColorTranslator.FromHtml("#D9D9D9");

        // 填入日期和星期
        for (int i = 1; i < table.RowCount; i++)
        {
            DateTime nowDate = firstDayOfMonth;
            // 填入日期
            table.Rows[i].Cells[0].Paragraphs[0].Append(nowDate.Day.ToString());


            // 填入星期
            table.Rows[i].Cells[1].Paragraphs[0].Append(GetDayOfWeekChinese(nowDate.DayOfWeek.ToString()));

            // 如果是星期六或星期日，設定整列背景色為灰色
            if (IsWeekend(firstDayOfMonth.DayOfWeek))
            {
                foreach (var cell in table.Rows[i].Cells)
                {
                    cell.FillColor = grayColor;
                }
            }

            if (i >= 2)
            {
                firstDayOfMonth = firstDayOfMonth.AddDays(1);
            }


        }

        // 儲存修改後的檔案
        document.SaveAs(outputFileName);
    }

    Console.WriteLine($"已生成檔案：{outputFileName}");
}

// 將英文星期轉換為中文星期
static string GetDayOfWeekChinese(string dayOfWeek)
{
    switch (dayOfWeek.ToLower())
    {
        case "monday":
            return "一";
        case "tuesday":
            return "二";
        case "wednesday":
            return "三";
        case "thursday":
            return "四";
        case "friday":
            return "五";
        case "saturday":
            return "六";
        case "sunday":
            return "日";
        default:
            return "";
    }
}


// 判斷是否為週末
static bool IsWeekend(DayOfWeek dayOfWeek)
{
    return dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday;
}