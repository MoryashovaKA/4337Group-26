using ClosedXML.Excel;
using System.Data;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Group4337
{
    public partial class _4337_Moryashova : Window
    {
        public _4337_Moryashova()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx",
                FileName = "2.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        int rows = range.RowCount();
                        MessageBox.Show($"Импортировано строк: {rows - 1}", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = GetTestData();
                var grouped = dt.AsEnumerable()
                    .GroupBy(row => row.Field<string>("Статус"));

                using (var workbook = new XLWorkbook())
                {
                    foreach (var group in grouped)
                    {
                        string status = group.Key ?? "Unknown";
                        var ws = workbook.Worksheets.Add(status);

                        ws.Cell(1, 1).Value = "Id";
                        ws.Cell(1, 2).Value = "Код заказа";
                        ws.Cell(1, 3).Value = "Дата создания";
                        ws.Cell(1, 4).Value = "Код клиента";
                        ws.Cell(1, 5).Value = "Услуги";

                        int row = 2;
                        foreach (var dataRow in group)
                        {
                            ws.Cell(row, 1).Value = dataRow.Field<int>("Id");
                            ws.Cell(row, 2).Value = dataRow.Field<string>("Код заказа") ?? "";
                            ws.Cell(row, 3).Value = dataRow.Field<DateTime>("Дата создания");
                            ws.Cell(row, 4).Value = dataRow.Field<string>("Код клиента") ?? "";
                            ws.Cell(row, 5).Value = dataRow.Field<string>("Услуги") ?? "";
                            row++;
                        }
                        ws.Columns().AdjustToContents();
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel files|*.xlsx",
                        FileName = "Export_ByStatus.xlsx"
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Экспорт завершён!", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DataTable GetTestData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Код заказа", typeof(string));
            dt.Columns.Add("Дата создания", typeof(DateTime));
            dt.Columns.Add("Код клиента", typeof(string));
            dt.Columns.Add("Услуги", typeof(string));
            dt.Columns.Add("Статус", typeof(string));

            dt.Rows.Add(1, "ORD-001", DateTime.Parse("2025-01-10"), "CL-100", "Ремонт ПК", "Новый");
            dt.Rows.Add(2, "ORD-002", DateTime.Parse("2025-01-11"), "CL-101", "Установка ПО", "В работе");
            dt.Rows.Add(3, "ORD-003", DateTime.Parse("2025-01-12"), "CL-100", "Замена экрана", "Завершён");
            dt.Rows.Add(4, "ORD-004", DateTime.Parse("2025-01-13"), "CL-102", "Диагностика", "Новый");
            dt.Rows.Add(5, "ORD-005", DateTime.Parse("2025-01-14"), "CL-101", "Чистка", "В работе");
            dt.Rows.Add(6, "ORD-006", DateTime.Parse("2025-01-15"), "CL-103", "Апгрейд", "Завершён");
            dt.Rows.Add(7, "ORD-007", DateTime.Parse("2025-01-16"), "CL-100", "Установка Windows", "Новый");
            dt.Rows.Add(8, "ORD-008", DateTime.Parse("2025-01-17"), "CL-104", "Настройка сети", "В работе");
            dt.Rows.Add(9, "ORD-009", DateTime.Parse("2025-01-18"), "CL-102", "Замена батареи", "Завершён");
            dt.Rows.Add(10, "ORD-010", DateTime.Parse("2025-01-19"), "CL-105", "Консультация", "Отменён");

            return dt;
        }
    }
}