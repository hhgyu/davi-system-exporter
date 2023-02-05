using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace worksheet_data_generate
{
    class ExcelWorker : IExcelWorker
    {
        private class ReceivedData
        {
            public DateTime Date;
            public string[] Values;
        }

        private const int MaxDataLine = 100;
        private Application? _app = null;
        private Workbook? _workbook = null;
        private Worksheet? _worksheetData = null;
        private Worksheet? _worksheetLastData = null;

        private List<ReceivedData> _datas = new List<ReceivedData>();

        private int currentRow = 2;
        private int valueRows = 1;

        private Thread? _thread = null;
        private bool _isStop = false;

        public void Enqueue(DateTime time, string[] values)
        {
            if (_isStop)
            {
                return;
            }

            lock (_datas)
            {
                valueRows = values.Length;
                _datas.Add(new ReceivedData
                {
                    Date = time,
                    Values = values
                });

                if (_datas.Count > MaxDataLine * 10)
                {
                    _datas.RemoveRange(0, _datas.Count - (MaxDataLine * 10));
                }
            }
        }

        private void Run()
        {
            try
            {
                while (!_isStop)
                {
                    List<ReceivedData> receivedDatas = new List<ReceivedData>();

                    lock (_datas)
                    {
                        for (int i = 0; i < _datas.Count; i++)
                        {
                            receivedDatas.Add(_datas[i]);
                        }
                    }

                    try
                    {
                        if (receivedDatas.Count > 0)
                        {
                            if (receivedDatas.Count > MaxDataLine)
                            {
                                int removeCount = receivedDatas.Count - MaxDataLine;
                                _datas.RemoveRange(0, removeCount);
                                receivedDatas.RemoveRange(0, removeCount);
                            }

                            var worksheetLastData = _worksheetLastData;
                            var worksheetData = _worksheetData;
                            if (worksheetLastData != null && worksheetData != null)
                            {
                                {
                                    ReceivedData data = receivedDatas.Last();

                                    var range = worksheetData.Range[worksheetData.Cells[2, 1], worksheetData.Cells[2, data.Values.Length + 1]];

                                    object[] values = new object[data.Values.Length + 1];
                                    values[0] = data.Date;
                                    for(int i = 0; i < data.Values.Length; i++)
                                    {
                                        values[i + 1] = data.Values[i];
                                    }

                                    range.Value = values;
                                }

                                int startRow = 2;
                                int endRow = 2 + MaxDataLine;
                                if (currentRow < endRow)
                                {
                                    int maxLoop = Math.Min(MaxDataLine - (currentRow - startRow), receivedDatas.Count);
                                    while (maxLoop-- > 0)
                                    {
                                        ReceivedData data = receivedDatas[0];
                                        receivedDatas.RemoveAt(0);
                                        lock (_datas)
                                        {
                                            _datas.RemoveAt(0);
                                        }

                                        var range = worksheetData.Range[worksheetData.Cells[currentRow, 1], worksheetData.Cells[currentRow, data.Values.Length + 1]];

                                        object[] values = new object[data.Values.Length + 1];
                                        values[0] = data.Date;
                                        for (int i = 0; i < data.Values.Length; i++)
                                        {
                                            values[i + 1] = data.Values[i];
                                        }

                                        range.Value = values;

                                        currentRow++;
                                    }
                                }
                                else
                                {
                                    int newDataRows = receivedDatas.Count - (MaxDataLine - (currentRow - startRow));
                                    if (newDataRows > 0)
                                    {
                                        if (MaxDataLine != newDataRows)
                                        {
                                            var copyRange = worksheetData.Range[worksheetData.Cells[startRow + newDataRows, 1], worksheetData.Cells[endRow - 1, 1 + valueRows]];
                                            var pasteRange = worksheetData.Range[worksheetData.Cells[startRow, 1], worksheetData.Cells[endRow - newDataRows - 1, 1 + valueRows]];
                                            copyRange.Copy(pasteRange);
                                        }
                                    }
                                    

                                    currentRow -= newDataRows;

                                    while (newDataRows-- > 0)
                                    {
                                        ReceivedData data = receivedDatas[0];
                                        receivedDatas.RemoveAt(0);
                                        lock (_datas)
                                        {
                                            _datas.RemoveAt(0);
                                        }

                                        var range = worksheetData.Range[worksheetData.Cells[currentRow, 1], worksheetData.Cells[currentRow, data.Values.Length + 1]];

                                        object[] values = new object[data.Values.Length + 1];
                                        values[0] = data.Date;
                                        for (int i = 0; i < data.Values.Length; i++)
                                        {
                                            values[i + 1] = data.Values[i];
                                        }

                                        range.Value = values;

                                        currentRow++;
                                    }
                                }
                            }
                        }

                    }
                    catch (ThreadInterruptedException) { throw; }
                    catch (Exception ex) { 
                        System.Diagnostics.Debug.WriteLine(ex);
                        if(!IsOpened(_app, _workbook))
                        {
                            System.Diagnostics.Debug.WriteLine("죽음");
                            break;
                        }
                    }

                    Thread.Sleep(1000);
                }
            }
            catch (ThreadInterruptedException) { }

            lock(_datas)
            {
                _isStop = true;

                _datas.Clear();
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

                currentRow = 2;
                
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

            currentRow = 2;
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

            Stop();
        }

        public void Start()
        {
            if(_thread != null)
            {
                Stop();
            }

            _isStop = false;

            _thread = new Thread(Run);
            _thread.Start();
        }

        public void Stop()
        {
            lock (_datas)
            {
                _isStop = true;

                _datas.Clear();
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

        private static bool IsOpened(Application app, Workbook workbook)
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
        bool Check();
        void Enqueue(DateTime time, string[] values);
        void Create();
        bool Open(string excelPath, out string errorMessage);

        void Start();
        void Stop();
    }
}
