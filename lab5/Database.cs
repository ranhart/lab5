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
using System.Runtime.InteropServices;
using System.ComponentModel.DataAnnotations;
using Microsoft.Office.Interop.Excel;
using System.Transactions;
using System.Data.Common;
using static System.Formats.Asn1.AsnWriter;

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
                    for (int i = 2; i <= rows - 1; i++)
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

        public void Display()
        {
            DisplayProductMovement();
            Console.WriteLine();
            DisplayProduct();
            Console.WriteLine();
            DisplayCategory();
            Console.WriteLine();
            DisplayShop();
        }

        public void DisplayProductMovement()
        {
            Console.WriteLine("Таблица \"Движение товаров\":");
            if (_productMovement != null && _productMovement.Any())
            {
                Console.WriteLine($"{"OperationID",-20}{"Date",-15}{"ShopID",-10}{"Article",-15}{"OperationType",-20}{"ItemQuantity",-20}{"Card",-10}");
                Console.WriteLine(new string('-', 150));
                foreach (var productMovement in _productMovement)
                {
                    Console.WriteLine($"{productMovement.OperationID,-20}{productMovement.Date,-15}{productMovement.ShopID,-10}{productMovement.Article,-15}{productMovement.OperationType,-20}{productMovement.ItemsQuantity,-20}{productMovement.Card,-10}");
                }
                _logger.Log("Лист \"Движение товаров\" успешно выведен в консоль.");
            }
            else
            {
                Console.WriteLine("Данные таблицы \"Движение товаров\" отсутствуют.");
                _logger.Log("Данные таблицы \"Движение товаров\" отсутствуют.");
            }
        }
        public void DisplayProduct()
        {
            Console.WriteLine("Таблица \"Товар\":");
            if (_product != null && _product.Any())
            {
                Console.WriteLine($"{"Article",-10}{"CategoryID",-15}{"ItemName",-40}{"FirstPrice",-15}{"SecondPrice",-15}{"Discount",-15}");
                Console.WriteLine(new string('-', 150));
                foreach (var product in _product)
                {
                    Console.WriteLine($"{product.Article,-10}{product.CategoryID,-15}{product.ItemName,-40}{product.FirstPrice,-15}{product.SecondPrice,-15}{product.Discount,-15}");
                }
                _logger.Log("Лист \"Товар\" успешно выведен в консоль.");
            }
            else
            {
                Console.WriteLine("Данные таблицы \"Товар\" отсутствуют.");
                _logger.Log("Данные таблицы \"Товар\" отсутствуют.");
            }
        }

        public void DisplayCategory()
        {
            Console.WriteLine("Таблица \"Категория\":");
            if (_category != null && _category.Any())
            {
                Console.WriteLine($"{"ID",-10}{"Name",-40}{"AgeRestriction",-20}");
                Console.WriteLine(new string('-', 150));
                foreach (var category in _category)
                {
                    Console.WriteLine($"{category.ID,-10}{category.Name,-40}{category.AgeRestriction,-20}");
                }
                _logger.Log("Лист \"Категория\" успешно выведен в консоль.");
            }
            else
            {
                Console.WriteLine("Данные таблицы \"Категория\" отсутствуют.");
                _logger.Log("Данные таблицы \"Категория\" отсутствуют.");
            }
        }

        public void DisplayShop()
        {
            Console.WriteLine("Таблица \"Магазин\":");
            if (_shop != null && _shop.Any())
            {
                Console.WriteLine($"{"ID",-10}{"Area",-15}{"Adress",-20}");
                Console.WriteLine(new string('-', 150));
                foreach (var shop in _shop)
                {
                    Console.WriteLine($"{shop.ID,-10}{shop.Area,-15}{shop.Adress,-20}");
                }
                _logger.Log("Лист \"Магазин\" успешно выведен в консоль.");
            }
            else
            {
                Console.WriteLine("Данные таблицы \"Магазин\" отсутствуют.");
                _logger.Log("Данные таблицы \"Магазин\" отсутствуют.");
            }
        }

        public void Save()
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    var productMovementWorksheet = package.Workbook.Worksheets[0];
                    for (int i = 0; i < _productMovement.Count; i++)
                    {
                        var productMovement = _productMovement[i];
                        DateTime date = productMovement.Date.ToDateTime(new TimeOnly(0, 0));
                        productMovementWorksheet.Cells[i + 2, 1].Value = productMovement.OperationID;
                        productMovementWorksheet.Cells[i + 2, 2].Value = date.ToOADate();
                        productMovementWorksheet.Cells[i + 2, 3].Value = productMovement.ShopID;
                        productMovementWorksheet.Cells[i + 2, 4].Value = productMovement.Article;
                        productMovementWorksheet.Cells[i + 2, 5].Value = productMovement.OperationID;
                        productMovementWorksheet.Cells[i + 2, 6].Value = productMovement.ItemsQuantity;
                        productMovementWorksheet.Cells[i + 2, 7].Value = productMovement.Card;
                    }

                    var productWorksheet = package.Workbook.Worksheets[1];
                    for (int i = 0; i < _product.Count; i++)
                    {
                        var product = _product[i];

                        productWorksheet.Cells[i + 2, 1].Value = product.Article;
                        productWorksheet.Cells[i + 2, 2].Value = product.CategoryID;
                        productWorksheet.Cells[i + 2, 3].Value = product.ItemName;
                        productWorksheet.Cells[i + 2, 4].Value = product.FirstPrice;
                        productWorksheet.Cells[i + 2, 5].Value = product.SecondPrice;
                        productWorksheet.Cells[i + 2, 6].Value = product.Discount;
                    }

                    var categoryWorksheet = package.Workbook.Worksheets[2];
                    for (int i = 0; i < _category.Count; i++)
                    {
                        var category = _category[i];
                        categoryWorksheet.Cells[i + 2, 1].Value = category.ID;
                        categoryWorksheet.Cells[i + 2, 2].Value = category.Name;
                        categoryWorksheet.Cells[i + 2, 3].Value = category.AgeRestriction;
                    }

                    var shopWorksheet = package.Workbook.Worksheets[3];
                    for (int i = 0; i < _shop.Count; i++)
                    {
                        var shop = _shop[i];
                        shopWorksheet.Cells[i + 2, 1].Value = shop.ID;
                        shopWorksheet.Cells[i + 2, 2].Value = shop.Area;
                        shopWorksheet.Cells[i + 2, 3].Value = shop.Adress;
                    }

                    package.Save();

                    _logger.Log("Изменения успешно сохранены в Excel файл.");
                }
            }
            catch (Exception e)
            {
                _logger.Log("Ошибка при сохранении изменений: " + e.Message);
            }
        }

        public void DeleteProductMovement(int operationID)
        {
            var movementToRemove = (from ProductMovement in _productMovement where ProductMovement.OperationID == operationID select ProductMovement).FirstOrDefault();

            if (movementToRemove != null)
            {
                _productMovement.Remove(movementToRemove);
                _logger.Log($"Движение товара с OperationID {operationID} удалено.");
            }
            else
            {
                _logger.Log($"Движение товара с OperationID {operationID} не найдено.");
            }
        }

        public void DeleteProduct(int article)
        {
            var productToRemove = (from Product in _product where Product.Article == article select Product).FirstOrDefault();

            if (productToRemove != null)
            {
                _product.Remove(productToRemove);
                _logger.Log($"Товар с артикулом {article} удален.");
            }
            else
            {
                _logger.Log($"Товар с артикулом {article} не найден.");
            }
        }

        public void DeleteCategory(int id)
        {
            var categoryToRemove = (from Category in _category where Category.ID == id select Category).FirstOrDefault();

            if (categoryToRemove != null)
            {
                _category.Remove(categoryToRemove);
                _logger.Log($"Категория с ID {id} удалена.");
            }
            else
            {
                _logger.Log($"Категория с ID {id} не найдена.");
            }
        }

        public void DeleteShop(string id)
        {
            var shopToRemove = (from Shop in _shop where Shop.ID == id select Shop).FirstOrDefault();

            if (shopToRemove != null)
            {
                _shop.Remove(shopToRemove);
                _logger.Log($"Магазин с ID {id} удален.");
            }
            else
            {
                _logger.Log($"Магазин с ID {id} не найден.");
            }
        }

        public void ChangeProductMovement(int operationID, ProductMovement newMovement)
        {
            var oldMovement = _productMovement.FirstOrDefault(pm => pm.OperationID == operationID);

            if (oldMovement != null)
            {
                try
                {
                    var index = _productMovement.IndexOf(oldMovement);
                    var sameId = _productMovement.FirstOrDefault(pm => pm.OperationID == newMovement.OperationID && pm.OperationID != operationID);

                    if (sameId == null)
                    {
                        _productMovement[index] = newMovement;
                        _logger.Log("Запись \"Движение товаров\" успешно изменена.");
                    }
                    else
                    {
                        _logger.Log($"Движение товаров с OperationID {newMovement.OperationID} уже существует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log("Ошибка при изменении записи \"Движение товаров\": " + e.Message);
                }
            }
            else
            {
                _logger.Log($"Движение товаров с OperationID {operationID} не найдено.");
            }
        }

        public void ChangeProduct(int article, Product newProduct)
        {
            var oldProduct = _product.FirstOrDefault(p => p.Article == article);

            if (oldProduct != null)
            {
                try
                {
                    if (!_product.Any(p => p.Article == newProduct.Article && p.Article != article))
                    {
                        var index = _product.IndexOf(oldProduct);
                        _product[index] = newProduct;
                        _logger.Log($"Товар с артикулом {article} успешно обновлен.");
                    }
                    else
                    {
                        _logger.Log($"Товар с артикулом {newProduct.Article} уже существует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log($"Ошибка при изменении товара: {e.Message}");
                }
            }
            else
            {
                _logger.Log($"Товар с артикулом {article} не найден.");
            }
        }

        public void ChangeCategory(int id, Category newCategory)
        {
            var oldCategory = _category.FirstOrDefault(c => c.ID == id);

            if (oldCategory != null)
            {
                try
                {
                    if (!_category.Any(s => s.ID == newCategory.ID && s.ID != id))
                    {
                        var index = _category.IndexOf(oldCategory);
                        _category[index] = newCategory;
                        _logger.Log($"Категория с ID {id} успешно обновлена.");
                    }
                    else
                    {
                        _logger.Log($"Категория с ID {newCategory.ID} уже существует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log($"Ошибка при изменении категории: {e.Message}");
                }
            }
            else
            {
                _logger.Log($"категория с ID {id} не найдена.");
            }
        }

        public void ChangeShop(string id, Shop newShop)
        {
            var oldShop = _shop.FirstOrDefault(s => s.ID == id);

            if (oldShop != null)
            {
                try
                {
                    if (!_shop.Any(s => s.ID == newShop.ID && s.ID != id))
                    {
                        var index = _shop.IndexOf(oldShop);
                        _shop[index] = newShop;
                        _logger.Log($"Магазин с ID {id} успешно обновлен.");
                    }
                    else
                    {
                        _logger.Log($"Магазин с ID {newShop.ID} уже существует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log($"Ошибка при изменении магазина: {e.Message}");
                }
            }
            else
            {
                _logger.Log($"Магазин с ID {id} не найден.");
            }
        }

        public void Add(Object obj)
        {
            if (obj is ProductMovement)
            {
                try
                {
                    var productMovement = obj as ProductMovement;
                    var id = productMovement.OperationID;
                    var sameId = (from ProductMovement in _productMovement
                                  where (ProductMovement.OperationID == id)
                                  select ProductMovement.OperationID).FirstOrDefault(-1);
                    if (sameId != -1)
                    {
                        _productMovement.Add(productMovement);
                    }
                    else
                    {
                        _logger.Log($"Движение товара с артикулом {id} уже сущесвтует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log("Ошибка при добавлении движения товара: " + e.Message);
                }
            }
            else if (obj is Product)
            {
                try
                {
                    var product = obj as Product;
                    var id = product.Article;
                    var sameId = (from Product in _product
                                  where (Product.Article == id)
                                  select Product.Article).FirstOrDefault(-1);
                    if (sameId != -1)
                    {
                        _product.Add(product);
                    }
                    else
                    {
                        _logger.Log($"Товар с артикулом {id} уже сущесвтует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log("Ошибка при добавлении товара: " + e.Message);
                }
            }
            else if (obj is Category)
            {
                try
                {
                    var category = obj as Category;
                    var id = category.ID;
                    var sameId = (from Category in _category
                                  where (Category.ID == id)
                                  select category.ID).FirstOrDefault(-1);
                    if (sameId != -1)
                    {
                        _category.Add(category);
                    }
                    else
                    {
                        _logger.Log($"Категория с id {id} уже существует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log("Ошибка при добавлении категории: " + e.Message);
                }
            }
            else if (obj is Shop)
            {
                try
                {
                    var shop = obj as Shop;
                    var id = shop.ID;
                    var sameId = (from Shop in _shop
                                  where (Shop.ID == id)
                                  select Shop.ID).FirstOrDefault("");
                    if (sameId != "")
                    {
                        _shop.Add(shop);
                    }
                    else
                    {
                        _logger.Log($"Магазин с id {id} уже сущесвтует.");
                    }
                }
                catch (Exception e)
                {
                    _logger.Log("Ошибка при добавлении магазина : " + e.Message);
                }
            }
            else
            {
                throw new ArgumentException("Переданный объект не представлен в базе данных.");
            }
        }
    }
}
