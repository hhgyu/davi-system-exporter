using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace DAVISystemExporter
{
    class ExcelWorker : IExcelWorker
    {
        public event EventHandler? Stoped;

        private class ReceivedData
        {
            public DateTime Date = DateTime.Now;
            public string[] Values = new string[] { };
        }

        private readonly int MaxDataLine = 100;
        private Application? _app = null;
        private Workbook? _workbook = null;
        private Worksheet? _worksheetData = null;
        private Worksheet? _worksheetLastData = null;

        private List<string[]> _logExcelDatas = new List<string[]>();
        private List<string[]> _currentExcelDatas = new List<string[]>();
        private bool _isNewData = false;

        private int _value_rows = -1;
        public int ValueRows
        {
            get { return _value_rows; }
        }

        private Thread? _thread = null;
        private bool _isStop = false;

        public ExcelWorker(int data_count) { MaxDataLine = data_count; }

        public void Enqueue(DateTime time, string[] values)
        {
            if (_isStop)
            {
                return;
            }

            if(_value_rows  == -1)
            {
                _value_rows = values.Length;
            }

            if(ValueRows != values.Length)
            {
                return;
            }

            lock(_currentExcelDatas)
            {
                _isNewData = true;

                string[] tmpValues = new string[ValueRows + 1];
                tmpValues[0] = time.ToString("HH:mm:ss.FFF");
                for (int i = 0; i < ValueRows; i++)
                {
                    tmpValues[i + 1] = values[i];
                }

                _currentExcelDatas.Add(tmpValues);
                if (_currentExcelDatas.Count > MaxDataLine)
                {
                    _currentExcelDatas.RemoveRange(0, _currentExcelDatas.Count - MaxDataLine);
                }

                _logExcelDatas.Add(tmpValues);
                if (_logExcelDatas.Count > 100000)
                {
                    _logExcelDatas.RemoveRange(0, _logExcelDatas.Count - 100000);
                }
            }
        }

        private void Run()
        {
            try
            {
                while (!_isStop)
                {
                    if(_isNewData)
                    {
                        _isNewData= false;

                        try
                        {
                            string lastTime = "";

                            var worksheetLastData = _worksheetLastData;
                            var worksheetData = _worksheetData;

                            if (worksheetLastData != null && worksheetData != null)
                            {
                                int count = _currentExcelDatas.Count;
                                while(count > 0)
                                {
                                    object[,] receivedDatas = new object[count, 1 + ValueRows];
                                    object[,] lastDatas = new object[1, 1 + ValueRows];

                                    lock (_currentExcelDatas)
                                    {
                                        if (count > 0)
                                        {
                                            count = _currentExcelDatas.Count;

                                            for (int i = 0; i < count; i++)
                                            {
                                                for (int j = 0; j < 1 + ValueRows; j++)
                                                {
                                                    receivedDatas[i, j] = _currentExcelDatas[i][j];
                                                }
                                            }

                                            for (int j = 0; j < 1 + ValueRows; j++)
                                            {
                                                lastDatas[0, j] = _currentExcelDatas[count - 1][j];
                                            }

                                            lastTime = _currentExcelDatas.Last().First();
                                        }
                                    }

                                    if(count <= 0)
                                    {
                                        break;
                                    }

                                    {
                                        int startRow = 2;
                                        var copyRange = worksheetLastData.Range[worksheetLastData.Cells[startRow, 1], worksheetLastData.Cells[startRow, 1 + ValueRows]];

                                        copyRange.Value2 = lastDatas;
                                    }

                                    {
                                        int startRow = 2;
                                        int endRow = 2 + Math.Min(count - 1, MaxDataLine);
                                        var copyRange = worksheetData.Range[worksheetData.Cells[startRow, 1], worksheetData.Cells[endRow, 1 + ValueRows]];

                                        copyRange.Value2 = receivedDatas;
                                    }

                                    break;
                                }
                            }
                        }
                        catch (ThreadInterruptedException) { throw; }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(ex);
                            if (!IsOpened(_app, _workbook))
                            {
                                System.Diagnostics.Debug.WriteLine("죽음");
                                break;
                            }
                        }
                    }

                    Thread.Sleep(100);
                }
            }
            catch (ThreadInterruptedException) { }
            finally
            {
                lock (_currentExcelDatas)
                {
                    _isStop = true;

                    _currentExcelDatas.Clear();
                }

                Stoped?.Invoke(this, new EventArgs());
            }
        }

        public void Create()
        {
            if (_app == null)
            {
                try
                {
                    _app = new Application();
                }
                catch (Exception)
                {
                    return;
                }
            }

            var app = _app;
            if (app != null)
            {
                _workbook = app.Workbooks.Add();

                var worksheetData = app.Worksheets[1] as Worksheet;
                if (worksheetData != null)
                {
                    _worksheetData = worksheetData;

                    worksheetData.Name = "데이터";

                    worksheetData.Cells[1, 1] = "시간";
                    worksheetData.Cells[1, 2] = "데이터1";
                }

                var worksheetLastData = app.Worksheets.Add() as Worksheet;
                if (worksheetLastData != null)
                {
                    _worksheetLastData = worksheetLastData;

                    worksheetLastData.Name = "최근데이터";

                    worksheetLastData.Cells[1, 1] = "시간";
                    worksheetLastData.Cells[1, 2] = "데이터1";
                }
                
                worksheetData?.Activate();

                app.Visible = true;
            }
        }

        public bool Open(string excelPath, out string errorMessage)
        {
            if (_app == null)
            {
                try
                {
                    _app = new Application();
                }
                catch (Exception)
                {
                    errorMessage = "어플리케이션 초기화에 실패했습니다.";
                    return false;
                }
            }

            var app = _app;
            if (app != null)
            {
                _workbook = app.Workbooks.Open(excelPath);

                if (app.Worksheets.Count >= 2)
                {
                    foreach (Worksheet sheet in app.Worksheets)
                    {
                        if (sheet.Name == "데이터")
                        {
                            _worksheetData = sheet;
                        }

                        if (sheet.Name == "최근데이터")
                        {
                            _worksheetLastData = sheet;
                        }
                    }
                }

                if (_worksheetData == null)
                {
                    errorMessage = "데이터 시트가 없습니다.";
                    return false;
                }

                if (_worksheetLastData == null)
                {
                    errorMessage = "최근데이터 시트가 없습니다.";
                    return false;
                }
            }
            else
            {
                errorMessage = "어플리케이션이 비었습니다.";
                return false;
            }

            var worksheetData = _worksheetData;
            if (worksheetData != null)
            {
                if (worksheetData.UsedRange != null)
                {
                    var range = worksheetData.UsedRange;
                    var lastCell = range.SpecialCells(XlCellType.xlCellTypeLastCell);
                    if (lastCell != null)
                    {
                        var lastRow = lastCell.Row;
                        var lastColumn = lastCell.Column;
                        if (lastRow > 3)
                        {
                            worksheetData.Range[worksheetData.Cells[range.Row + 2, range.Column], worksheetData.Cells[lastRow, lastColumn]].Cells.Delete();
                        }

                        if (lastRow > 2)
                        {
                            worksheetData.Range[worksheetData.Cells[range.Row + 1, range.Column], worksheetData.Cells[range.Row + 1, lastColumn]].Cells.ClearContents();
                        }
                    }

                    if (range.Column != 1 || range.Row != 1)
                    {
                        errorMessage = "데이터 시트의 데이터가 이상합니다.";
                        return false;
                    }
                }
            }

            app.Visible = true;
            errorMessage = "";
            return true;
        }

        public bool Check()
        {
            if (_app == null)
            {
                try
                {
                    var app = new Application();

                    try
                    {
                        app.Quit();
                    }
                    catch (Exception) { }

                    Marshal.FinalReleaseComObject(app);
                }
                catch (Exception) { }
            }

            return _app != null;
        }

        public void Dispose()
        {
            Stoped = null;
            Stop();

            lock (_currentExcelDatas)
            {
                _currentExcelDatas.Clear();

                _logExcelDatas.Clear();
            }
            
            var workbook = _workbook;
            var worksheetData = _worksheetData;
            var worksheetLastData = _worksheetLastData;
            var app = _app;

            if (app != null)
            {
                if (worksheetData != null)
                {
                    Marshal.FinalReleaseComObject(worksheetData);
                }

                if (worksheetLastData != null)
                {
                    Marshal.FinalReleaseComObject(worksheetLastData);
                }

                if (workbook != null)
                {
                    /* No App Close
                    try
                    {
                        workbook.Close();
                    }
                    catch (Exception) { }
                    */
                    Marshal.FinalReleaseComObject(workbook);
                }

                /* No App Close
                try
                {
                    app.Quit();
                }
                catch (Exception) { }
                */

                Marshal.FinalReleaseComObject(app);
            }

            _worksheetData = null;
            _worksheetLastData = null;
            _workbook = null;
            _app = null;
        }

        public void Start()
        {
            if(_thread != null)
            {
                Stop();
            }

            _value_rows = -1;
            _isStop = false;

            _thread = new Thread(Run);
            _thread.Start();
        }

        public void Stop()
        {
            lock (_currentExcelDatas)
            {
                _isStop = true;
            }

            var t = _thread;
            if (t?.IsAlive == false)
            {
                return;
            }

            _thread?.Interrupt();

            while (_thread?.IsAlive == true)
            {
                Thread.Sleep(100);
            }

            _thread = null;
        }

        public void Export()
        {
            try
            {
                var app = _app;

                if (app != null)
                {
                    var workbookSheet = app.Worksheets.Add() as Worksheet;

                    if (workbookSheet != null)
                    {
                        workbookSheet.Name = $"로그 {DateTime.Now.ToString("yyyyMMddHHmmss")}";

                        workbookSheet.Cells[1, 1] = "시간";
                        for(int i=0; i< ValueRows; i++)
                        {
                            workbookSheet.Cells[1, 2 + i] = $"데이터{i +1}";
                        }

                        int count = _logExcelDatas.Count;
                        while (count > 0)
                        {
                            object[,] receivedDatas = new object[count, 1 + ValueRows];

                            lock (_currentExcelDatas)
                            {
                                if (count > 0)
                                {
                                    count = _logExcelDatas.Count;

                                    for (int i = 0; i < count; i++)
                                    {
                                        for (int j = 0; j < 1 + ValueRows; j++)
                                        {
                                            receivedDatas[i, j] = _logExcelDatas[i][j];
                                        }
                                    }
                                }
                            }

                            if (count <= 0)
                            {
                                break;
                            }

                            {
                                int startRow = 2;
                                int endRow = 2 + count - 1;
                                var copyRange = workbookSheet.Range[workbookSheet.Cells[startRow, 1], workbookSheet.Cells[endRow, 1 + ValueRows]];

                                copyRange.Value2 = receivedDatas;
                            }

                            break;
                        }
                    }
                }
            }
            catch (ThreadInterruptedException) { throw; }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
                if (!IsOpened(_app, _workbook))
                {
                    System.Diagnostics.Debug.WriteLine("죽음");
                }
            }
        }

        private static bool IsOpened(Application? app, Workbook? workbook)
        {
            try
            {
                //return app != null && app.Workbooks.Cast<Workbook>().FirstOrDefault(i => i == workbook) != workbook;
                return true;
            }
            catch { }
            return false;
        }
    }

    interface IExcelWorker : IDisposable
    {
        event EventHandler? Stoped;

        int ValueRows
        {
            get;
        }

        bool Check();
        void Enqueue(DateTime time, string[] values);
        void Create();
        bool Open(string excelPath, out string errorMessage);

        void Start();
        void Stop();

        void Export();
    }
}
