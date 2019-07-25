using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace ReportParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var docx = @"C:\Users\jacky_rxwx1ok\source\repos\ReportParser\ReportParser\bin\Debug\A.docx";
            try
            {
                //建立開啟Word工具
                Word.Application app = new Word.Application();
                //取得文件
                Word.Document dx = app.Documents.Open(docx);

                //Word.WdStatistic是一群列舉常數
                /*
                 包含：
                    wdStatisticCharacters	3	字元數。
                    wdStatisticCharactersWithSpaces	5	含空格的字元數。
                    wdStatisticFarEastCharacters	6	亞洲語言的字元數。
                    wdStatisticLines	1	行數。
                    wdStatisticPages	2	頁數。
                    wdStatisticParagraphs	4	段落數。
                    wdStatisticWords	0	字數。
                 */
                //取得本文件總頁數
                object Miss = System.Reflection.Missing.Value;
                Word.WdStatistic PagesCountStat = Word.WdStatistic.wdStatisticPages;
                int PageCount = dx.ComputeStatistics(PagesCountStat, Miss);
                Console.Out.WriteLine("Word page count : " + PageCount);

                //取得段落總數
                Word.WdStatistic PagesParagraphsStat = Word.WdStatistic.wdStatisticParagraphs;
                int PageParagraphsCount = dx.ComputeStatistics(PagesParagraphsStat, Miss);
                Console.Out.WriteLine("Paragraphs count : " + PageParagraphsCount);

                //取得Table，並檢查Table結束於哪個頁面，以及Row位於哪個頁面
                Console.Out.WriteLine("Tables count : " + dx.Tables.Count);

                foreach(Word.Table tb in dx.Tables)
                {
                    //檢查Table第3個欄位數的Cell總數，如果有合併將顯示合併後數量
                    Console.Out.WriteLine("Column 3 Cells : " + tb.Columns[3].Cells.Count);
                    
                    Console.Out.WriteLine(" OK！ ");

                    //Tabel.Uniform 屬性用來判斷是否有合併儲存格，如果有就False否則True
                    if (tb.Uniform)
                    {
                        Console.Out.WriteLine("###該表格沒有合併儲存格，所以尋列讀取！###");
                        int start = tb.Rows[1].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
                        int end = tb.Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
                        Console.Out.WriteLine("Table Start Page : " + start.ToString() + " - End Page : " + end.ToString());
                        foreach (Word.Row row in tb.Rows)
                        {
                            int endpage = row.Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
                            Console.Out.WriteLine("Row index : " + row.Index + " At Page : " + endpage.ToString());
                        }
                    }
                    else
                    {
                        Console.Out.WriteLine("###該表格有合併儲存格，所以將Table轉Cell依序讀取！###");
                        //完整頁數
                        int[] fullPages = new int[] { };
                        //計算時暫存用頁數
                        int[] tmPage = new int[] { };
                        //如果Table有合併儲存格讀取方式要將Table轉Cell依序讀取
                        Word.Range ra = tb.Range;
                        for (int i = 1; i <= ra.Cells.Count; i++)
                        {
                            try
                            {
                                Word.Cell cell = ra.Cells[i]; //取得Cell
                                int numPage = cell.Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber]; //取得頁數
                                int[] st = new int[] { numPage }; //頁數準備合併前的暫存陣列物件
                                Console.Out.WriteLine("[ " + cell.RowIndex.ToString() + " - " + cell.ColumnIndex.ToString() + " ] - [ " + numPage + " ]");
                                fullPages = tmPage.Union(st).ToArray<int>(); //合併
                            }
                            catch (Exception inex)
                            {
                                Console.Out.WriteLine(inex.Message);
                            }
                        }
                        Console.Out.WriteLine("Page Num : " + fullPages.Count().ToString());
                    }
                }

                //依序移動頁面，並取得頁面文字陣列
                List<string> Pages = new List<string>();
                object What = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;
                object Which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute;
                object Start;
                object End;
                object CurrentPageNumber;
                object NextPageNumber;

                for (int Index = 1; Index < PageCount + 1; Index++)
                {
                    CurrentPageNumber = (Convert.ToInt32(Index.ToString()));
                    NextPageNumber = (Convert.ToInt32((Index + 1).ToString()));

                    // Get start position of current page
                    Start = app.Selection.GoTo(ref What, ref Which, ref CurrentPageNumber, ref Miss).Start;

                    // Get end position of current page                                
                    End = app.Selection.GoTo(ref What, ref Which, ref NextPageNumber, ref Miss).End;

                    // Get text
                    if (Convert.ToInt32(Start.ToString()) != Convert.ToInt32(End.ToString()))
                        Pages.Add(dx.Range(ref Start, ref End).Text);
                    else
                        Pages.Add(dx.Range(ref Start).Text);
                }

                //結束關閉檔案
                dx.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Console.Out.WriteLine(" OK！ ");
            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message);
            }
            finally
            {
                //釋放執行緒
                Process myProcess = new Process();
                Process[] wordProcess = Process.GetProcessesByName("winword");
                Console.Out.WriteLine("Word process count : " + wordProcess.Count());
                foreach (Process pro in wordProcess)
                {
                    pro.Kill();
                }
            }

            string x = Console.ReadLine();
        }
    }
}
