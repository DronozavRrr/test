using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            // Проверяем, что пользователь передал параметры
            if (args.Length < 2)
            {
                Console.WriteLine("Использование: ExcelMerger.exe <input_file1> <input_file2> ... <output_file>");
                return;
            }

            // Последний параметр - выходной файл, остальные - входные файлы
            string outputFile = args.Last();
            string[] inputFiles = args.Take(args.Length - 1).ToArray();

            MergeExcelFiles(inputFiles, outputFile);
        }

        static void MergeExcelFiles(string[] inputFiles, string outputFile)
        {
            Application excelApp = new Application();
            Workbook outputWorkbook = excelApp.Workbooks.Add();
            Worksheet summarySheet = (Worksheet)outputWorkbook.Worksheets[1];
            summarySheet.Name = "Summary";

            int currentRow = 1;  // Счетчик для строк в листе "Summary"
            HashSet<string> uniqueWallets = new HashSet<string>();
            Dictionary<string, string> sheetNameMap = new Dictionary<string, string>(); // Оригинальные имена и фактические имена листов

            // Сначала копируем все листы, кроме "Summary"
            foreach (string file in inputFiles)
            {
                Workbook workbook = excelApp.Workbooks.Open(file);

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = (Worksheet)workbook.Worksheets[i];
                    string originalSheetName = sheet.Name;

                    if (originalSheetName != "Summary" && !sheetNameMap.ContainsKey(originalSheetName))
                    {
                        // Копируем лист в выходной файл
                        sheet.Copy(After: outputWorkbook.Sheets[outputWorkbook.Sheets.Count]);

                        // Получаем фактическое имя листа после копирования
                        Worksheet copiedSheet = (Worksheet)outputWorkbook.Sheets[outputWorkbook.Sheets.Count];
                        string actualSheetName = copiedSheet.Name;

                        // Добавляем в словарь для сопоставления оригинального имени с фактическим именем
                        sheetNameMap[originalSheetName] = actualSheetName;

                        // Выводим отладочное сообщение
                        Console.WriteLine($"Скопирован лист: Оригинальное имя = '{originalSheetName}', Фактическое имя = '{actualSheetName}'");
                    }
                }

                workbook.Close(false);  // Закрываем исходный файл без сохранения изменений
            }

            // Вывод всех листов в итоговом файле для отладки
            Console.WriteLine("Список всех листов в итоговом файле:");
            foreach (Worksheet ws in outputWorkbook.Worksheets)
            {
                Console.WriteLine($"Лист: {ws.Name}");
            }

            // Затем обрабатываем листы "Summary" и добавляем ссылки
            foreach (string file in inputFiles)
            {
                Workbook workbook = excelApp.Workbooks.Open(file);
                Worksheet summaryWorksheet = (Worksheet)workbook.Worksheets["Summary"];
                Range usedRange = summaryWorksheet.UsedRange;

                // Копирование заголовков (выполняем только один раз)
                if (currentRow == 1)
                {
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        summarySheet.Cells[currentRow, col].Value = usedRange.Cells[1, col].Value;
                    }
                    currentRow++; // Переходим на следующую строку после заголовков
                }

                for (int row = 2; row <= usedRange.Rows.Count; row++)
                {
                    Range walletCell = summaryWorksheet.Cells[row, 1];
                    string wallet = walletCell.Text.ToString();

                    if (uniqueWallets.Add(wallet))  // Добавляем только уникальные кошельки
                    {
                        for (int col = 1; col <= usedRange.Columns.Count; col++)
                        {
                            Range sourceCell = summaryWorksheet.Cells[row, col];
                            Range targetCell = summarySheet.Cells[currentRow, col];

                            // Копируем значение ячейки
                            targetCell.Value = sourceCell.Value;

                            // Если это первая колонка (кошельки), добавляем гиперссылку на лист с соответствующим именем
                            if (col == 1 && !string.IsNullOrEmpty(wallet))
                            {
                                try
                                {
                                    // Проверяем, существует ли фактическое имя листа для этого кошелька
                                    string originalSheetName = $"Wallet_{wallet}";
                                    if (sheetNameMap.TryGetValue(originalSheetName, out string actualSheetName))
                                    {
                                        string subAddress = $"'{actualSheetName}'!A1"; // Используем фактическое имя листа

                                        targetCell.Hyperlinks.Add(
                                            Anchor: targetCell,
                                            Address: "", // Пустой адрес, так как это внутренняя ссылка на лист
                                            SubAddress: subAddress,
                                            TextToDisplay: wallet
                                        );
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Лист с именем '{originalSheetName}' не найден в итоговом файле.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Ошибка при добавлении гиперссылки: {ex.Message}");
                                }
                            }

                            // Копируем цвет фона и шрифта
                            targetCell.Interior.Color = sourceCell.Interior.Color;
                            targetCell.Font.Color = sourceCell.Font.Color;
                            targetCell.Font.Bold = sourceCell.Font.Bold;
                            targetCell.Font.Italic = sourceCell.Font.Italic;
                        }
                        currentRow++;
                    }
                }

                workbook.Close(false);  // Закрываем исходный файл без сохранения изменений
            }

            // Сохраняем и закрываем итоговый файл
            outputWorkbook.SaveAs(outputFile);
            outputWorkbook.Close(false);
            excelApp.Quit();

            Console.WriteLine($"Результат сохранен в файл {outputFile}");
        }
    }
}
