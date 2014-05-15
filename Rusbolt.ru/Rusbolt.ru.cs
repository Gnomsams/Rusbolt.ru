using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using Microsoft.Office.Interop.Excel;

namespace Rusbolt.ru
{
    public class Rusbolt
    {
        private Application xlApplication;
        private Workbook xlWorkbook;
        private Worksheet xlWorksheet;
        private IWebDriver driver;

        public Rusbolt()
        {
            driver = new FirefoxDriver();
            xlApplication = new Application();
            xlWorkbook = xlApplication.Workbooks.Add();
            xlWorksheet = xlWorkbook.Sheets.Add(); ;

            xlWorksheet.Cells[1, 1] = "Тип метизов";
            xlWorksheet.Cells[1, 2] = "Стандарт";
            xlWorksheet.Cells[1, 3] = "Диаметр";
            xlWorksheet.Cells[1, 4] = "Длина";
            xlWorksheet.Cells[1, 5] = "Тип покрытия";
            xlWorksheet.Cells[1, 6] = "Класс прочности";
            xlWorksheet.Cells[1, 7] = "Кол-во, шт.";
            xlWorksheet.Cells[1, 8] = "Вес, кг.";
        }

        private int CountOfOptionElements(string id)
        {
            return driver.FindElement(By.Id(id)).FindElements(By.TagName("option")).Count;
        }

        private ReadOnlyCollection<IWebElement> OptionsElements(string id)
        {
            return driver.FindElement(By.Id(id)).FindElements(By.TagName("option"));
        }

        private IWebElement OptionElement(string id, int i)
        {
            IWebElement element;
            try
            {
                element = OptionsElements(id)[i];
            }
            catch (Exception)
            {
                Thread.Sleep(4000);
                element = OptionsElements(id)[i];
            }
            return element;
        }

        //++Болт
        //++Гайка
        //++Шайба
        //++Гровер
        //++Винт
        //++Гвоздь
        //++Ось мебельная
        //++Анкеры
        //++Шуруп
        //++Шплинт
        //++Заклепка
        //++Шпилька
        //++Такелаж
        //++Рым-крепеж
        //++Саморез
        //++Круг отрезной
        //++Проволока


        public void Parse(string boltType)
        {          
            driver.Navigate().GoToUrl("http://www.rusbolt.ru/ves/");

            var totalCount = 1;
            string _type, _standart, _diametr, _length, _coverType, _classStrength, _weight;

            var types = OptionsElements("type");
            foreach (var type in types.Where(type => type.Text == boltType))
            {
                _type = type.Text;
                type.Click();

                Thread.Sleep(800);
                var standarts_count = CountOfOptionElements("din");
                var standarts_count_true = standarts_count;
                if (standarts_count_true == 0)
                    standarts_count++;

                for (var i1 = 0; i1 < standarts_count; i1++)
                {
                    if (standarts_count_true != 0)
                    {
                        IWebElement standart = OptionElement("din", i1);
                        _standart = standart.Text;
                        standart.Click();
                    }
                    else _standart = "";

                    Thread.Sleep(800);
                    var diametrs_count = CountOfOptionElements("diam");
                    var diametrs_count_true = diametrs_count;
                    if (diametrs_count_true == 0)
                        diametrs_count++;

                    for (var i2 = 0; i2 < diametrs_count; i2++)
                    {
                        if (diametrs_count_true != 0)
                        {
                            IWebElement diametr = OptionElement("diam", i2);
                            _diametr = diametr.Text;
                            diametr.Click();
                        }
                        else _diametr = "";

                        Thread.Sleep(800);
                        var lengths_count = CountOfOptionElements("length");
                        var lengths_count_true = lengths_count;
                        if (lengths_count_true == 0)
                            lengths_count++;

                        for (var i3 = 0; i3 < lengths_count; i3++)
                        {
                            if (lengths_count_true != 0)
                            {
                                IWebElement length = OptionElement("length", i3);
                                _length = length.Text;
                                length.Click();
                            }
                            else _length = "";

                            Thread.Sleep(800);
                            var coverTypes_count = CountOfOptionElements("typeCover");
                            var coverTypes_count_true = coverTypes_count;
                            if (coverTypes_count_true == 0)
                                coverTypes_count++;

                            for (var i4 = 0; i4 < coverTypes_count; i4++)
                            {
                                if (coverTypes_count_true != 0)
                                {
                                    IWebElement coverType = OptionElement("typeCover", i4);
                                    _coverType = coverType.Text;
                                    coverType.Click();
                                }
                                else _coverType = "";

                                Thread.Sleep(800);
                                var classStrengths_count = CountOfOptionElements("classStrength");
                                var classStrengths_count_true = classStrengths_count;
                                if (classStrengths_count_true == 0)
                                    classStrengths_count++;

                                for (var i5 = 0; i5 < classStrengths_count; i5++)
                                {
                                    if (classStrengths_count_true != 0)
                                    {
                                        IWebElement classStrength = OptionElement("classStrength", i5);
                                        _classStrength = classStrength.Text;
                                        classStrength.Click();
                                    }
                                    else _classStrength = "";

                                    _weight = driver.FindElement(By.Id("ves")).GetAttribute("value");

                                    totalCount++;

                                    FillRow(totalCount, _type, _standart, _diametr, _length, _coverType,
                                        _classStrength, _weight);
                                }
                            }
                        }
                    }
                }

                SaveAndCloseExcel(type.Text);
            }
        }

        public void FillRow(int x, string type, string standart, string diametr, string length, string coverType,
            string classStrength, string weight)
        {
            for (var i = 1; i < 7; i++)
            {
                ((Range)xlWorksheet.Cells[x, i]).NumberFormat = "@";
            }
            
            xlWorksheet.Cells[x, 1] = type;
            xlWorksheet.Cells[x, 2] = standart;
            xlWorksheet.Cells[x, 3] = diametr;
            xlWorksheet.Cells[x, 4] = length;
            xlWorksheet.Cells[x, 5] = coverType;
            xlWorksheet.Cells[x, 6] = classStrength;
            xlWorksheet.Cells[x, 7] = "1000";
            xlWorksheet.Cells[x, 8] = weight;
        }

        public void SaveAndCloseExcel(string fileName)
        {
            xlWorkbook.SaveAs(Environment.CurrentDirectory + @"\" + fileName + ".xls");
            xlWorkbook.Close(true);
            xlApplication.Quit();
        }
    }
}