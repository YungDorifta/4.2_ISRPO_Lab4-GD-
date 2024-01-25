using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Лаб4AJAX.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Лаб4AJAX.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index(string pattern, string name, string surname)
        {
            List<SearchResultLine> result;
            using (db_BelashevEntities2 db = new db_BelashevEntities2())
            {
                /*db.Books.Add(new Books()
                {
                    ID_book = 125,
                    Book_name = "ghgfhfh",

                });*/
                db.SaveChanges();
                result = db.Readers.Join(
                        db.BooksReaders,
                        reader => reader.ID_reader,
                        bookreader => bookreader.ID_reader,
                        (reader, bookreader) => new
                        {
                            Name = reader.Name,
                            Surname = reader.Surname,
                            id_reader_book = bookreader.ID,
                            id_book = bookreader.ID_book
                        }).Join(db.Books,
                        bookreader => bookreader.id_book,
                        book => book.ID_book,
                        (bookreader, book) => new SearchResultLine()
                        {
                            Name = bookreader.Name,
                            Surname = bookreader.Surname,
                            Book_name = book.Book_name
                        }).ToList();

                if (pattern == null)
                {
                    ViewBag.SearchData = result;
                    return View();
                }
                else
                {
                    result = result.Where((p) => p.Name.Contains(pattern)).ToList();
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
            }
        }
       
        // Создание и заполнение файла Word в загрузки
        public FileStreamResult GetWord()
        {
            // Извлечь данные из БД
            List<SearchResultLine> result;
            using (db_BelashevEntities2 db = new db_BelashevEntities2())
            {
                result = db.Readers.Join(
                        db.BooksReaders,
                        reader => reader.ID_reader,
                        bookreader => bookreader.ID_reader,
                        (reader, bookreader) => new
                        {
                            Name = reader.Name,
                            Surname = reader.Surname,
                            id_reader_book = bookreader.ID,
                            id_book = bookreader.ID_book
                        }).Join(db.Books,
                        bookreader => bookreader.id_book,
                        book => book.ID_book,
                        (bookreader, book) => new SearchResultLine()
                        {
                            Name = bookreader.Name,
                            Surname = bookreader.Surname,
                            Book_name = book.Book_name
                        }).ToList();
            }

            // Создать массив ячеек из БД
            string[,] tableData = new string[5, 3];
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    switch (i)
                    {
                        case (0):
                            {
                                tableData[j, i] = result[j].Name.ToString();
                                break;
                            }
                        case (1):
                            {
                                tableData[j, i] = result[j].Surname.ToString();
                                break;
                            }
                        case (2):
                            {
                                tableData[j, i] = result[j].Book_name.ToString();
                                break;
                            }
                        default: break;
                    }
                }
            }

            // Создать матрицу
            string[,] data = new string[3, 5];
            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 5; j++)
                    data[i, j] = (i + j).ToString();

            // Выгрузить файл в загрузки
            MemoryStream memoryStream = GenerateWord(tableData);
            return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = "demo.docx"
            };
        }

        // Структура файла Word
        private MemoryStream GenerateWord(string[,] data)
        {
            MemoryStream mStream = new MemoryStream();
            // Создаем документ
            using (WordprocessingDocument document = WordprocessingDocument.Create(mStream, WordprocessingDocumentType.Document, true))
            {
                // Добавляется главная часть документа. 
                MainDocumentPart mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                // Создаем таблицу. 
                Table table = new Table();
                body.AppendChild(table);

                // Устанавливаем свойства таблицы(границы и размер).
                TableProperties props = new TableProperties(
                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));


                // Назначаем свойства props объекту table
                table.AppendChild<TableProperties>(props);

                // Заполняем ячейки таблицы.
                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        // размер колонок определяется автоматически.
                        tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tr.Append(tc);
                    }
                    table.Append(tr);
                }

                mainPart.Document.Save();
            }
            mStream.Position = 0;
            return mStream;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        
        public string addSQL(string name, string surname)
        {
            string sqlrequest = "INSERT INTO [db_Belashev].[dbo].[Readers](Name, Surname) VALUES ('" + name + "', '" + surname + "')";
            return sqlrequest;
        }

        public string delSQL(string name, string surname)
        {
            string sqlrequest = "DELETE FROM [db_Belashev].[dbo].[Readers] WHERE (Name = '" + name + "') AND (Surname = '" + surname + "')";
            return sqlrequest;
        }
    }
}