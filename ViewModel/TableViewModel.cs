using DBClasses.Model.Enums;
using DBClasses.Module;
using DevExpress.Mvvm;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;



namespace DBClasses.ViewModel
{
    public class TableViewModel : ViewModelBase
    {
        private ObservableCollection<RowBroadcast> _tableBroadcast = null!;
        private ObservableCollection<RowChannel> _tableChannel = null!;
        private ObservableCollection<RowShow> _tableShow = null!;

        private object _currentTable = null!;

        public TableViewModel()
        {
            FillTables();
            CurrentTable = _tableShow;
        }
        public object CurrentTable
        {
            get => _currentTable;
            set
            {
                _currentTable = value;
                RaisePropertyChanged(() => CurrentTable); // обнова змінної при кожному set
            }
        }

        #region Tables and buttons

        public ICommand TvShowTableCommand => new DelegateCommand(() => CurrentTable = _tableShow);
        public ICommand TvChannelTableCommand => new DelegateCommand(() => CurrentTable = _tableChannel);
        public ICommand TvBroadcastTableCommand => new DelegateCommand(() => CurrentTable = _tableBroadcast);

        public CommandBase AddCommand
        {
            get => new DelegateCommand<RowShow>(row =>
            {
                RowShow rowShow = row;
                row.IdShow++;
                _tableShow.Add(row);
            });
        }
        public CommandBase RemoveCommand => new DelegateCommand<RowShow>(row => { _tableShow.Remove(row); });

        #endregion

        public ICommand SaveJsonCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Json file (*.json)|*.json";

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        JsonSerializerOptions jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                        string showsJson = JsonSerializer.Serialize(_tableShow, jsonOptions);
                        File.WriteAllText(saveFileDialog.FileName, showsJson);
                    }
                    _tableShow.Clear();
                });
            }
        }
        public ICommand LoadJsonCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Json file (*.json)|*.json";

                    if (openFileDialog.ShowDialog() == true)
                    {
                        string tableJson = File.ReadAllText(openFileDialog.FileName);
                        ObservableCollection<RowShow>? tableShow = JsonSerializer.Deserialize<ObservableCollection<RowShow>>(tableJson);
                        foreach (RowShow row in tableShow)
                        {
                            if (!_tableShow.Contains(row))
                                _tableShow.Add(row);
                        }
                    }


                });
            }
        }
        public ICommand SaveXmlCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Xml file (*.xml)|*.xml";

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        var writer = new XmlSerializer(typeof(ObservableCollection<RowShow>));
                        XmlWriter xmlWriter = new XmlTextWriter(saveFileDialog.FileName, Encoding.Default);
                        writer.Serialize( xmlWriter, _tableShow);
                    }
                    _tableShow.Clear();
                });
            }
        }
        public ICommand LoadXmlCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Xml file (*.xml)|*.xml";

                    if (openFileDialog.ShowDialog() == true)
                    {
                        var reader = new XmlSerializer(typeof(ObservableCollection<RowShow>));
                        XmlTextReader xmlReader = new XmlTextReader(openFileDialog.FileName);
                        var r = reader.Deserialize(xmlReader);
                        foreach (var row in (r as ObservableCollection<RowShow>)!)
                            if (!_tableShow.Contains(row))
                                _tableShow.Add(row);
                    }
                });
            }
        }
        public ICommand SaveExcelCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    var excel = new Excel.Application();
                    var wb = excel.Workbooks.Add();
                    var ws = (Excel.Worksheet)wb.ActiveSheet;
                    ws.Columns.AutoFit();
                    ws.Columns.EntireColumn.ColumnWidth = 25;

                    //now the list
                    int counter = 1;
                    foreach (var item in _tableShow)
                    {
                        ws.Cells[counter, 1] = item.IdShow.ToString();
                        ws.Cells[counter, 2] = item.Name;
                        ws.Cells[counter, 3] = item.TypeShow.ToString();
                        ws.Cells[counter, 4] = item.Duration.ToString();
                        ws.Cells[counter, 5] = item.ShowCategory.ToString();
                        ++counter;
                    }


                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        
                        wb.SaveAs(saveFileDialog.FileName);
                        wb.Close();
                        excel.Quit();
                        Marshal.FinalReleaseComObject(excel);
                        GC.Collect(); //collect
                        GC.WaitForPendingFinalizers(); //end it
                    }
                    _tableShow.Clear();
                    
                    
                });
            }
        }
        public ICommand LoadExcelCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    try
                    {
                        _tableShow.Clear();
                        OpenFileDialog openFileDialog = new OpenFileDialog();
                        openFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

                        if (openFileDialog.ShowDialog() != true)
                            throw new ApplicationException("Could not open");
                        string filePath = openFileDialog.FileName;
                        //string filePath = @"C:\Users\paloc\OneDrive\Рабочий стол\test.xlsx";

                        Excel.Application excel = new Excel.Application();
                        Excel.Workbook wb = excel.Workbooks.Open(filePath);
                        Excel.Worksheet ws = wb.Worksheets[1];

                        int counter = 1;
                        for (int i = 0; i < 12; i++)
                        {
                            _tableShow.Add(new RowShow());
                            RowShow item = _tableShow[i];
                            item.IdShow = int.Parse(ws.Cells[counter, 1].Value.ToString());
                            item.Name = ws.Cells[counter, 2].Value.ToString();
                            item.TypeShow = Enum.Parse(typeof(TypeShow), ws.Cells[counter, 3].Value.ToString());
                            item.Duration = uint.Parse(ws.Cells[counter, 4].Value.ToString());
                            item.ShowCategory = Enum.Parse(typeof(CategoryShow), ws.Cells[counter, 5].Value.ToString());

                            ++counter;
                        }
                        wb.Close();
                        excel.Quit();
                        Marshal.FinalReleaseComObject(excel);
                        GC.Collect(); //collect
                        GC.WaitForPendingFinalizers(); //end it
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                });
            }
        }
        public ICommand SaveWordCommand
        {
            get
            {
                return new DelegateCommand(() =>
                {
                    
                    Word.Application word = new Word.Application();
                    Word.Document doc = word.Documents.Add();

                    try
                    {
                        doc.Paragraphs.Add();

                        for (int row = 1; row < _tableShow.Count+1; row++)
                        {
                            doc.Paragraphs.Add();

                            string s = _tableShow[row - 1].IdShow.ToString();
                            s = string.Concat(s, "\t", _tableShow[row - 1].Name);
                            s = string.Concat(s, "\t", _tableShow[row - 1].TypeShow.ToString());
                            s = string.Concat(s, "\t", _tableShow[row - 1].Duration.ToString());
                            s = string.Concat(s, "\t", _tableShow[row - 1].ShowCategory.ToString());

                            doc.Paragraphs[row].Range.Text = s; //add s to range.
                        }
                        var r = word.ActiveDocument.Content;
                        r.Select();
                        r.ConvertToTable(Word.WdTableFieldSeparator.wdSeparateByTabs,_tableShow.Count - 1, 5);
                        word.ActiveDocument.Tables[1].Borders.Enable = 1;

                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.Filter = "Word file (*.docx)|*.docx";
                        
                        if (saveFileDialog.ShowDialog() != true) return;
                        doc.SaveAs(saveFileDialog.FileName);

                        GC.Collect(); //collect
                        GC.WaitForPendingFinalizers(); //end it
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        doc.Close();
                        word.Quit();
                        Marshal.FinalReleaseComObject(doc); //корректное взаимодествие 
                        Marshal.FinalReleaseComObject(word);  //корректное взаимодествие 
                        _tableShow.Clear();
                    }
                });
            }
        }



        private static void SerializeObjectToXml<T>(T item, string filePath)
        {
            XmlSerializer xs = new XmlSerializer(typeof(T));
            using (StreamWriter wr = new StreamWriter(filePath))
            {
                xs.Serialize(wr, item);
            }
        }
        private void FillTables()
        {
            _tableShow = new ObservableCollection<RowShow>
            {
                new RowShow("Хто зверху?", TypeShow.РозважальнаПередача, 90, CategoryShow.Червоний),
                new RowShow("Життя зірок", TypeShow.РозважальнаПередача, 120, CategoryShow.Зелений),
                new RowShow("Єдині. Головне", TypeShow.НауковоПопулярнаПередача, 90, CategoryShow.Зелений),
                new RowShow("Мауглі", TypeShow.Мультфільм, 129, CategoryShow.Жовтий),
                new RowShow("Хіти Non-Stop", TypeShow.РозважальнаПередача, 55, CategoryShow.Зелений),
                new RowShow("Леді Баг і Супер-Кіт", TypeShow.Мультфільм, 105, CategoryShow.Зелений),
                new RowShow("Інформаційний марафон", TypeShow.НауковоПопулярнаПередача, 130, CategoryShow.Жовтий),
                new RowShow("МайстерШеф. Професіонали", TypeShow.РозважальнаПередача, 90, CategoryShow.Червоний),
                new RowShow("Слідство ведуть екстрасенси", TypeShow.РозважальнаПередача, 120, CategoryShow.Жовтий),
                new RowShow("Місія Блейк", TypeShow.Мультфільм, 140, CategoryShow.Зелений),
                new RowShow("Оггі та кукарачі", TypeShow.Мультфільм, 120, CategoryShow.Зелений),
                new RowShow("Супер мама", TypeShow.РозважальнаПередача, 60, CategoryShow.Жовтий)
            }; // fill tableShow
            _tableChannel = new ObservableCollection<RowChannel>
            {
                new RowChannel("Про Все", TypeChannel.Супутниковий, 6543.2),
                new RowChannel("СТБ", TypeChannel.Національний, 1567.8),
                new RowChannel("1+1", TypeChannel.Національний, 2695.8),
                new RowChannel("Рада", TypeChannel.Національний, 9564.4),
                new RowChannel("Starlight Media", TypeChannel.Супутниковий, 9956.1),
                new RowChannel("Суспільне Культура", TypeChannel.Національний, 9563.4),
                new RowChannel("Інтер", TypeChannel.Національний, 2265.1),
                new RowChannel("Сонце", TypeChannel.Супутниковий, 9564.1),
                new RowChannel("Вільні", TypeChannel.Супутниковий, 7564.4),
                new RowChannel("ТЕТ", TypeChannel.Національний, 2649.7),
                new RowChannel("МЕГА", TypeChannel.Національний, 5554.7),
                new RowChannel("НТН", TypeChannel.Супутниковий, 9546.5)
            }; // fill tableChannel
            _tableBroadcast = new ObservableCollection<RowBroadcast>
            {
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 20, 00, 00),
                    new DateTime(2023, 01, 13, 21, 30, 00),
                    2,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 21, 00, 00),
                    new DateTime(2023, 01, 13, 23, 20, 00),
                    3,
                    2),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 07, 00, 00),
                    new DateTime(2023, 01, 13, 09, 10, 00),
                    3,
                    3),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 09, 30, 00),
                    new DateTime(2023, 01, 13, 10, 30, 00),
                    2,
                    4),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 10, 00, 00),
                    new DateTime(2023, 01, 13, 11, 20, 00),
                    1,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 08, 00, 00),
                    new DateTime(2023, 01, 13, 09, 30, 00),
                    1,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 00, 00),
                    new DateTime(2023, 01, 13, 14, 00, 00),
                    6,
                    7),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 20, 00),
                    new DateTime(2023, 01, 13, 15, 00, 00),
                    8,
                    8),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 18, 00, 00),
                    new DateTime(2023, 01, 13, 20, 30, 00),
                    11,
                    9),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 19, 00, 00),
                    new DateTime(2023, 01, 13, 22, 10, 00),
                    8,
                    10),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 06, 00, 00),
                    new DateTime(2023, 01, 13, 08, 00, 00),
                    10,
                    11),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 00, 00),
                    new DateTime(2023, 01, 13, 14, 00, 00),
                    9,
                    7),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 14, 30, 00),
                    new DateTime(2023, 01, 13, 16, 00, 00),
                    5,
                    8),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 18, 00, 00),
                    new DateTime(2023, 01, 13, 20, 30, 00),
                    9,
                    9),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 03, 20, 00),
                    new DateTime(2023, 01, 13, 05, 40, 00),
                    5,
                    10),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 06, 00, 00),
                    new DateTime(2023, 01, 13, 08, 00, 00),
                    5,
                    11),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 10, 00, 00),
                    new DateTime(2023, 01, 13, 11, 20, 00),
                    4,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 08, 00, 00),
                    new DateTime(2023, 01, 13, 09, 30, 00),
                    4,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 00, 00),
                    new DateTime(2023, 01, 13, 14, 00, 00),
                    4,
                    7),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 20, 00),
                    new DateTime(2023, 01, 13, 15, 00, 00),
                    7,
                    8),
                new RowBroadcast(
                    new DateTime(2023, 02, 24, 10, 30, 00),
                    new DateTime(2023, 02, 24, 12, 00, 00),
                    5,
                    11),
                new RowBroadcast(
                    new DateTime(2023, 02, 24, 15, 00, 00),
                    new DateTime(2023, 02, 24, 17, 00, 00),
                    5,
                    8),
                new RowBroadcast(
                    new DateTime(2023, 02, 24, 13, 30, 00),
                    new DateTime(2023, 02, 24, 15, 20, 00),
                    3,
                    4),
                new RowBroadcast(
                    new DateTime(2023, 02, 24, 21, 20, 00),
                    new DateTime(2023, 02, 24, 23, 30, 00),
                    3,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 02, 24, 11, 30, 00),
                    new DateTime(2023, 02, 24, 13, 30, 00),
                    5,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 15, 07, 00, 00),
                    new DateTime(2023, 01, 15, 08, 30, 00),
                    8,
                    10),
                new RowBroadcast(
                    new DateTime(2023, 01, 15, 10, 00, 00),
                    new DateTime(2023, 01, 15, 12, 20, 00),
                    4,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 15, 10, 00, 00),
                    new DateTime(2023, 01, 15, 12, 20, 00),
                    5,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 15, 10, 00, 00),
                    new DateTime(2023, 01, 15, 12, 20, 00),
                    1,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 15, 07, 00, 00),
                    new DateTime(2023, 01, 15, 08, 30, 00),
                    2,
                    10),
                new RowBroadcast(
                    new DateTime(2023, 01, 17, 07, 00, 00),
                    new DateTime(2023, 01, 17, 08, 30, 00),
                    9,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 17, 07, 00, 00),
                    new DateTime(2023, 01, 17, 08, 30, 00),
                    4,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 17, 07, 00, 00),
                    new DateTime(2023, 01, 17, 08, 30, 00),
                    2,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 17, 07, 00, 00),
                    new DateTime(2023, 01, 17, 08, 30, 00),
                    3,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 19, 07, 00, 00),
                    new DateTime(2023, 01, 19, 08, 30, 00),
                    2,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 19, 07, 00, 00),
                    new DateTime(2023, 01, 19, 08, 30, 00),
                    3,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 19, 07, 00, 00),
                    new DateTime(2023, 01, 19, 08, 30, 00),
                    1,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 19, 07, 00, 00),
                    new DateTime(2023, 01, 19, 08, 30, 00),
                    9,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 20, 21, 00, 00),
                    new DateTime(2023, 01, 20, 23, 30, 00),
                    3,
                    2),
                new RowBroadcast(
                    new DateTime(2023, 01, 20, 21, 00, 00),
                    new DateTime(2023, 01, 20, 23, 30, 00),
                    2,
                    3),
                new RowBroadcast(
                    new DateTime(2023, 01, 20, 21, 00, 00),
                    new DateTime(2023, 01, 20, 23, 30, 00),
                    1,
                    2),
                new RowBroadcast(
                    new DateTime(2023, 01, 20, 21, 00, 00),
                    new DateTime(2023, 01, 20, 23, 30, 00),
                    8,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 20, 21, 00, 00),
                    new DateTime(2023, 01, 20, 23, 30, 00),
                    11,
                    4),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 08, 20, 00),
                    new DateTime(2023, 01, 13, 09, 20, 00),
                    5,
                    5),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 09, 40, 00),
                    new DateTime(2023, 01, 13, 11, 40, 00),
                    5,
                    2),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 12, 00, 00),
                    new DateTime(2023, 01, 13, 14, 10, 00),
                    5,
                    7),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 16, 20, 00),
                    new DateTime(2023, 01, 13, 17, 50, 00),
                    5,
                    1),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 18, 10, 00),
                    new DateTime(2023, 01, 13, 19, 50, 00),
                    5,
                    6),
                new RowBroadcast(
                    new DateTime(2023, 01, 13, 20, 10, 00),
                    new DateTime(2023, 01, 13, 21, 40, 00),
                    5,
                    3)
            }; // fill tableBroadcast
        }
    }
}
