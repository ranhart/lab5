using lab5.objects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Globalization;

namespace lab5
{
    internal class Database
    {
        private Logger _logger;
        private string _filePath;
        private List<Category> _category;
        private List<Product> _product;
        private List<ProductMovement> _productMovement;
        private List<Shop> _shop;

        public Database(string filePath, Logger logger)
        {
            if (File.Exists(filePath))
            {
                _logger = logger;
                _filePath = filePath;
                _category = new List<Category>();
                _product = new List<Product>();
                _productMovement = new List<ProductMovement>();
                _shop = new List<Shop>();
            }
            else { throw new FileNotFoundException("Файл excel не найден"); }

        }
        public void Read()
        {
            ReadProductMovement();
            ReadProduct();
            ReadCategory();
            ReadShop();
            _logger.Log("Чтение файла завершено");
        }
        public void ReadProductMovement()
        {
            using (var package = new ExcelPackage(_filePath))
            {
                try
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rows = worksheet.Dimension.Rows;
                    for (int i = 2; i <= rows; i++)
                    {
                        _productMovement.Add(new ProductMovement(
                            int.Parse(worksheet.Cells[i, 1].Text),
                            DateOnly.ParseExact(worksheet.Cells[i, 2].Text, "dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU")),
                            worksheet.Cells[i, 3].Text,
                            int.Parse(worksheet.Cells[i, 4].Text),
                            worksheet.Cells[i, 5].Text,
                            int.Parse(worksheet.Cells[i, 6].Text),
                            worksheet.Cells[i, 7].Text
                            ));
                    }
                    _logger.Log("Лист \"Движение товаров\" успешно считан.");
                }
                catch (Exception e) { _logger.Log("Ошибка  чтения листа \"Движение товаров\": " + e); }
            }
        }

        public void ReadProduct()
        {
            using (var package = new ExcelPackage(_filePath))
            {
                try
                {
                    var worksheet = package.Workbook.Worksheets[1];
                    int rows = worksheet.Dimension.Rows;
                    for (int i = 2; i <= rows; i++)
                    {
                        _product.Add(new Product(
                            int.Parse(worksheet.Cells[i, 1].Text),
                            int.Parse(worksheet.Cells[i, 2].Text),
                            worksheet.Cells[i, 3].Text,
                            int.Parse(worksheet.Cells[i, 4].Text),
                            int.Parse(worksheet.Cells[i, 5].Text),
                            worksheet.Cells[i, 6].Text
                            ));
                    }
                    _logger.Log("Лист \"Товар\" успешно считан.");
                }
                catch (Exception e) { _logger.Log("Ошибка  чтения листа \"Товар\": " + e); }
            }
        }

        public void ReadCategory()
        {
            using (var package = new ExcelPackage(_filePath))
            {
                try
                {
                    var worksheet = package.Workbook.Worksheets[2];
                    int rows = worksheet.Dimension.Rows;
                    for (int i = 2; i <= rows; i++)
                    {
                        _category.Add(new Category(
                            int.Parse(worksheet.Cells[i, 1].Text),
                            worksheet.Cells[i, 2].Text,
                            worksheet.Cells[i, 3].Text
                            ));
                    }
                    _logger.Log("Лист \"Категория\" успешно считан.");
                }
                catch (Exception e) { _logger.Log("Ошибка  чтения листа \"Категория\": " + e); }
            }
        }

        public void ReadShop()
        {
            using (var package = new ExcelPackage(_filePath))
            {
                try
                {
                    var worksheet = package.Workbook.Worksheets[3];
                    int rows = worksheet.Dimension.Rows;
                    for (int i = 2; i <= rows; i++)
                    {
                        _shop.Add(new Shop(
                            worksheet.Cells[i, 1].Text,
                            worksheet.Cells[i, 2].Text,
                            worksheet.Cells[i, 3].Text
                            ));
                    }
                    _logger.Log("Лист \"Магазин\" успешно считан.");
                }
                catch (Exception e) { _logger.Log("Ошибка  чтения листа \"Магазин\": " + e); }
            }
        }
    }
}
