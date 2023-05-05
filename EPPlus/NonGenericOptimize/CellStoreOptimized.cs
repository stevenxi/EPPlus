using System;
using System.Collections.Generic;

namespace OfficeOpenXml.NonGenericOptimize
{
    internal class CellStoreOptimized : IDisposable// : IEnumerable<ulong>, IEnumerator<ulong>
    {
        /**** Size constants ****/
        internal const int pageBits = 10;   //13bits=8192  Note: Maximum is 13 bits since short is used (PageMax=16K)
        internal const int PageSize = 1 << pageBits;
        internal const int PageSizeMin = 1 << 10;
        internal const int PageSizeMax = PageSize << 1; //Double page size
        internal const int ColSizeMin = 32;
        internal const int PagesPerColumnMin = 32;

        private List<ExcelCoreValue> _valuesGeneric = new List<ExcelCoreValue>();
        private List<string> _valuesString = new List<string>();
        private List<decimal> _valuesDecimal = new List<decimal>();

        internal ColumnIndex[] _columnIndex;
        internal IndexBase _searchIx = new IndexBase();
        internal IndexItem _searchItem = new IndexItem();
        internal int ColumnCount;
        public CellStoreOptimized()
        {
            _columnIndex = new ColumnIndex[ColSizeMin];
        }
        ~CellStoreOptimized()
        {
            if (_valuesGeneric != null)
            {
                _valuesGeneric.Clear();
                _valuesGeneric = null;
            }

            if (_valuesString != null)
            {
                _valuesString.Clear();
                _valuesString = null;
            }

            if (_valuesDecimal != null)
            {
                _valuesDecimal.Clear();
                _valuesDecimal = null;
            }
            _columnIndex = null;
        }
        internal int GetPosition(int Column)
        {
            if (Column < ColumnCount && _columnIndex[Column].Index == Column)      //Check if th column is lesser than
            {
                return Column;
            }
            else
            {
                _searchIx.Index = (short)Column;
                return Array.BinarySearch(_columnIndex, 0, ColumnCount, _searchIx);
            }
        }
  
        internal bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
        {
            if (ColumnCount == 0)
            {
                fromRow = fromCol = toRow = toCol = 0;
                return false;
            }
            else
            {
                fromCol = _columnIndex[0].Index;
                var fromIndex = 0;
                if (fromCol <= 0 && ColumnCount > 1)
                {
                    fromCol = _columnIndex[1].Index;
                    fromIndex = 1;
                }
                else if (ColumnCount == 1 && fromCol <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                var col = ColumnCount - 1;
                while (col > 0)
                {
                    if (_columnIndex[col].PageCount == 0 || _columnIndex[col]._pages[0].RowCount > 1 || _columnIndex[col]._pages[0].Rows[0].Index > 0)
                    {
                        break;
                    }
                    col--;
                }
                toCol = _columnIndex[col].Index;
                if (toCol == 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                fromRow = toRow = 0;

                for (int c = fromIndex; c < ColumnCount; c++)
                {
                    int first, last;
                    if (_columnIndex[c].PageCount == 0) continue;
                    if (_columnIndex[c]._pages[0].RowCount > 0 && _columnIndex[c]._pages[0].Rows[0].Index > 0)
                    {
                        first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[0].Index;
                    }
                    else
                    {
                        if (_columnIndex[c]._pages[0].RowCount > 1)
                        {
                            first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[1].Index;
                        }
                        else if (_columnIndex[c].PageCount > 1)
                        {
                            first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[1].Rows[0].Index;
                        }
                        else
                        {
                            first = 0;
                        }
                    }
                    var lp = _columnIndex[c].PageCount - 1;
                    while (_columnIndex[c]._pages[lp].RowCount == 0 && lp != 0)
                    {
                        lp--;
                    }
                    var p = _columnIndex[c]._pages[lp];
                    if (p.RowCount > 0)
                    {
                        last = p.IndexOffset + p.Rows[p.RowCount - 1].Index;
                    }
                    else
                    {
                        last = first;
                    }
                    if (first > 0 && (first < fromRow || fromRow == 0))
                    {
                        fromRow = first;
                    }
                    if (first > 0 && (last > toRow || toRow == 0))
                    {
                        toRow = last;
                    }
                }
                if (fromRow <= 0 || toRow <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
     
        internal ExcelCoreValue GetValue(int Row, int Column)
        {
            var i = GetPointer(Row, Column);
            if (i.Index >= 0)
            {
                switch(i.Type)
                {
                    case ComposePositionType.String:
                        return new ExcelCoreValue { _value = _valuesString[i.Index] };
                    case ComposePositionType.Decimal:
                        return new ExcelCoreValue { _value = _valuesDecimal[i.Index] };
                }

                return new ExcelCoreValue { _value = _valuesGeneric[i.Index] };
            }

            return default;
        }

        ComposePosition GetPointer(int Row, int Column)
        {
            var col = GetPosition(Column);
            if (col >= 0)
            {
                var pos = _columnIndex[col].GetPosition(Row);
                if (pos >= 0 && pos < _columnIndex[col].PageCount)
                {
                    var pageItem = _columnIndex[col]._pages[pos];
                    if (pageItem.MinIndex > Row)
                    {
                        pos--;
                        if (pos < 0)
                            return new ComposePosition {  Index = -1, Type = ComposePositionType.Generic};

                        pageItem = _columnIndex[col]._pages[pos];
                    }
                    short ix = (short)(Row - pageItem.IndexOffset);
                    _searchItem.Index = ix;
                    var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
                    if (cellPos >= 0)
                    {
                        return pageItem.Rows[cellPos].IndexPointer;
                    }

                    //Cell does not exist
                    return new ComposePosition {  Index = -1, Type = ComposePositionType.Generic};
                }

                //Page does not exist
                return new ComposePosition { Index = -1, Type = ComposePositionType.Generic };
            }

            //Column does not exist
            return new ComposePosition { Index = -1, Type = ComposePositionType.Generic };
        }
     
        internal void SetValue(int Row, int Column, ExcelCoreValue Value)
        {
            lock (_columnIndex)
            {
                var col = GetPosition(Column);          //Array.BinarySearch(_columnIndex, 0, ColumnCount, new IndexBase() { Index = (short)(Column) });
                var page = (short)(Row >> pageBits);
                if (col >= 0)
                {
                    //var pos = Array.BinarySearch(_columnIndex[col].Pages, 0, _columnIndex[col].Count, new IndexBase() { Index = page });
                    var pos = _columnIndex[col].GetPosition(Row);
                    if (pos < 0)
                    {
                        pos = ~pos;
                        if (pos - 1 < 0 || _columnIndex[col]._pages[pos - 1].IndexOffset + PageSize - 1 < Row)
                        {
                            AddPage(_columnIndex[col], pos, page);
                        }
                        else
                        {
                            pos--;
                        }
                    }
                    if (pos >= _columnIndex[col].PageCount)
                    {
                        AddPage(_columnIndex[col], pos, page);
                    }
                    var pageItem = _columnIndex[col]._pages[pos];
                    if (!(pageItem.MinIndex <= Row && pageItem.MaxIndex >= Row) && pageItem.IndexExpanded > Row)   //TODO: Fix issue
                    {
                        pos--;
                        page--;
                        if (pos < 0)
                        {
                            throw (new Exception("Unexpected error when setting value"));
                        }
                        pageItem = _columnIndex[col]._pages[pos];
                    }

                    short ix = (short)(Row - ((pageItem.Index << pageBits) + pageItem.Offset));
                    _searchItem.Index = ix;
                    var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
                    if (cellPos < 0)
                    {
                        cellPos = ~cellPos;
                        AddCell(_columnIndex[col], pos, cellPos, ix, Value);
                    }
                    else
                    {
                        var pointer = pageItem.Rows[cellPos].IndexPointer;
                        switch (pointer.Type)
                        {
                            case ComposePositionType.String:
                                if (Value._value is string s)
                                {
                                    _valuesString[pointer.Index] = s;
                                }
                                else
                                {
                                    pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(Value);
                                }
                                break;
                            case ComposePositionType.Decimal:
                                if (Value._value is decimal d)
                                {
                                    _valuesDecimal[pointer.Index] = d;
                                }
                                else
                                {
                                    pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(Value);
                                }
                                break;
                            default:
                                _valuesGeneric[pointer.Index] = Value;
                                break;
                        }
                    }
                }
                else //Column does not exist
                {
                    col = ~col;
                    AddColumn(col, Column);
                    AddPage(_columnIndex[col], 0, page);
                    short ix = (short)(Row - (page << pageBits));
                    AddCell(_columnIndex[col], 0, 0, ix, Value);
                }
            }
        }

        internal delegate ExcelCoreValue? GetRangeValueDelegate(int row, int column, object value, ExcelCoreValue oldVal);
        /// <summary>
        /// Set Value for Range
        /// </summary>
        /// <param name="fromRow"></param>
        /// <param name="fromColumn"></param>
        /// <param name="toRow"></param>
        /// <param name="toColumn"></param>
        /// <param name="Updater"></param>
        /// <param name="Value"></param>
        internal void SetRangeValueSpecial(int fromRow, int fromColumn, int toRow, int toColumn, GetRangeValueDelegate Updater, object Value)
        {
            lock (_columnIndex)
            {
                // split row to page groups (pageIndex to RowNo List)
                Dictionary<short, List<int>> pages = new Dictionary<short, List<int>>();
                for (int rowIx = fromRow; rowIx <= toRow; rowIx++)
                {
                    var pageIx = (short)(rowIx >> pageBits);
                    if (!pages.ContainsKey(pageIx)) pages.Add(pageIx, new List<int>());
                    pages[pageIx].Add(rowIx);
                }

                for (int colIx = fromColumn; colIx <= toColumn; colIx++)
                {
                    //var col = Array.BinarySearch(_columnIndex, 0, ColumnCount, new IndexBase() { Index = (short)(colIx) });
                    var col = GetPosition(colIx);

                    foreach (var pair in pages)
                    {
                        short page = pair.Key;
                        foreach (var rowIx in pair.Value)
                        {
                            if (col >= 0)
                            {
                                //var pos = Array.BinarySearch(_columnIndex[col].Pages, 0, _columnIndex[col].Count, new IndexBase() { Index = page });
                                var pos = _columnIndex[col].GetPosition(rowIx);
                                if (pos < 0)
                                {
                                    pos = ~pos;
                                    if (pos - 1 < 0 || _columnIndex[col]._pages[pos - 1].IndexOffset + PageSize - 1 < rowIx)
                                    {
                                        AddPage(_columnIndex[col], pos, page);
                                    }
                                    else
                                    {
                                        pos--;
                                    }
                                }
                                if (pos >= _columnIndex[col].PageCount)
                                {
                                    AddPage(_columnIndex[col], pos, page);
                                }
                                var pageItem = _columnIndex[col]._pages[pos];
                                if (pageItem.IndexOffset > rowIx)
                                {
                                    pos--;
                                    page--;
                                    if (pos < 0)
                                    {
                                        throw (new Exception("Unexpected error when setting value"));
                                    }
                                    pageItem = _columnIndex[col]._pages[pos];
                                }

                                short ix = (short)(rowIx - ((pageItem.Index << pageBits) + pageItem.Offset));
                                _searchItem.Index = ix;
                                var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);

                                ComposePosition pointer;
                                ExcelCoreValue? updatedVal;
                                if (cellPos < 0)
                                {
                                    cellPos = ~cellPos;
                                    AddCell(_columnIndex[col], pos, cellPos, ix, default(ExcelCoreValue));

                                    pointer = pageItem.Rows[cellPos].IndexPointer;
                                    ExcelCoreValue currentVal;
                                    switch (pointer.Type)
                                    {
                                        case ComposePositionType.String:
                                            currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                                            break;
                                        case ComposePositionType.Decimal:
                                            currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                                            break;
                                        default:
                                            currentVal = _valuesGeneric[pointer.Index];
                                            break;
                                    }
                                    updatedVal = Updater(rowIx, colIx, Value, currentVal);
                                }
                                else
                                {
                                    pointer = pageItem.Rows[cellPos].IndexPointer;
                                    ExcelCoreValue currentVal;
                                    switch (pointer.Type)
                                    {
                                        case ComposePositionType.String:
                                            currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                                            break;
                                        case ComposePositionType.Decimal:
                                            currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                                            break;
                                        default:
                                            currentVal = _valuesGeneric[pointer.Index];
                                            break;
                                    }
                                    updatedVal = Updater(rowIx, colIx, Value, currentVal);
                                }

                                if (updatedVal.HasValue)
                                {
                                    switch (pointer.Type)
                                    {
                                        case ComposePositionType.String:
                                            if (updatedVal.Value._value is string s && updatedVal.Value._styleId == default)
                                            {
                                                _valuesString[pointer.Index] = s;
                                            }
                                            else
                                            {
                                                pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                                _valuesGeneric.Add(updatedVal.Value);
                                            }
                                            break;
                                        case ComposePositionType.Decimal:
                                            if (updatedVal.Value._value is decimal d && updatedVal.Value._styleId == default)
                                            {
                                                _valuesDecimal[pointer.Index] = d;
                                            }
                                            else
                                            {
                                                pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                                _valuesGeneric.Add(updatedVal.Value);
                                            }
                                            break;
                                        default:
                                            _valuesGeneric[pointer.Index] = updatedVal.Value;
                                            break;
                                    }
                                }
                            }
                            else //Column does not exist
                            {
                                col = ~col;
                                AddColumn(col, colIx);
                                AddPage(_columnIndex[col], 0, page);
                                short ix = (short)(rowIx - (page << pageBits));
                                AddCell(_columnIndex[col], 0, 0, ix, default(ExcelCoreValue));
                                var pointer = _columnIndex[col]._pages[0].Rows[0].IndexPointer;

                                ExcelCoreValue currentVal;
                                switch (pointer.Type)
                                {
                                    case ComposePositionType.String:
                                        currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                                        break;
                                    case ComposePositionType.Decimal:
                                        currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                                        break;
                                    default:
                                        currentVal = _valuesGeneric[pointer.Index];
                                        break;
                                }

                                var updatedVal = Updater(rowIx, colIx, Value, currentVal);

                                if (updatedVal.HasValue)
                                {
                                    switch (pointer.Type)
                                    {
                                        case ComposePositionType.String:
                                            if (updatedVal.Value._value is string s && updatedVal.Value._styleId == default)
                                            {
                                                _valuesString[pointer.Index] = s;
                                            }
                                            else
                                            {
                                                _columnIndex[col]._pages[0].Rows[0].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                                _valuesGeneric.Add(updatedVal.Value);
                                            }
                                            break;
                                        case ComposePositionType.Decimal:
                                            if (updatedVal.Value._value is decimal d && updatedVal.Value._styleId == default)
                                            {
                                                _valuesDecimal[pointer.Index] = d;
                                            }
                                            else
                                            {
                                                _columnIndex[col]._pages[0].Rows[0].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                                _valuesGeneric.Add(updatedVal.Value);
                                            }
                                            break;
                                        default:
                                            _valuesGeneric[pointer.Index] = updatedVal.Value;
                                            break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal delegate ExcelCoreValue? SetValueDelegate(object value, ExcelCoreValue oldVal);

        // Set object's property atomically
        internal void SetValueSpecial(int Row, int Column, SetValueDelegate Updater, object Value)
        {
            lock (_columnIndex)
            {
                //var col = Array.BinarySearch(_columnIndex, 0, ColumnCount, new IndexBase() { Index = (short)(Column) });
                var col = GetPosition(Column);
                var page = (short)(Row >> pageBits);
                if (col >= 0)
                {
                    //var pos = Array.BinarySearch(_columnIndex[col].Pages, 0, _columnIndex[col].Count, new IndexBase() { Index = page });
                    var pos = _columnIndex[col].GetPosition(Row);
                    if (pos < 0)
                    {
                        pos = ~pos;
                        if (pos - 1 < 0 || _columnIndex[col]._pages[pos - 1].IndexOffset + PageSize - 1 < Row)
                        {
                            AddPage(_columnIndex[col], pos, page);
                        }
                        else
                        {
                            pos--;
                        }
                    }
                    if (pos >= _columnIndex[col].PageCount)
                    {
                        AddPage(_columnIndex[col], pos, page);
                    }
                    var pageItem = _columnIndex[col]._pages[pos];
                    if (pageItem.IndexOffset > Row)
                    {
                        pos--;
                        page--;
                        if (pos < 0)
                        {
                            throw (new Exception("Unexpected error when setting value"));
                        }
                        pageItem = _columnIndex[col]._pages[pos];
                    }

                    short ix = (short)(Row - ((pageItem.Index << pageBits) + pageItem.Offset));
                    _searchItem.Index = ix;
                    var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);


                    ComposePosition pointer;
                    ExcelCoreValue? updatedVal;
                    if (cellPos < 0)
                    {
                        cellPos = ~cellPos;
                        AddCell(_columnIndex[col], pos, cellPos, ix, new ExcelCoreValue {  _value = Value});

                        pointer = pageItem.Rows[cellPos].IndexPointer;


                        ExcelCoreValue currentVal;
                        switch (pointer.Type)
                        {
                            case ComposePositionType.String:
                                currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                                break;
                            case ComposePositionType.Decimal:
                                currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                                break;
                            default:
                                currentVal = _valuesGeneric[pointer.Index];
                                break;
                        }

                        updatedVal = Updater(Value, currentVal);
                    }
                    else
                    {
                        pointer = pageItem.Rows[cellPos].IndexPointer;
                        ExcelCoreValue currentVal;
                        switch (pointer.Type)
                        {
                            case ComposePositionType.String:
                                currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                                break;
                            case ComposePositionType.Decimal:
                                currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                                break;
                            default:
                                currentVal = _valuesGeneric[pointer.Index];
                                break;
                        }
                        updatedVal = Updater(Value, currentVal);
                    }

                    if (updatedVal.HasValue)
                    {
                        switch (pointer.Type)
                        {
                            case ComposePositionType.String:
                                if (updatedVal.Value._value is string s && updatedVal.Value._styleId == default)
                                {
                                    _valuesString[pointer.Index] = s;
                                }
                                else
                                {
                                    pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(updatedVal.Value);
                                }
                                break;
                            case ComposePositionType.Decimal:
                                if (updatedVal.Value._value is decimal d && updatedVal.Value._styleId == default)
                                {
                                    _valuesDecimal[pointer.Index] = d;
                                }
                                else
                                {
                                    pageItem.Rows[cellPos].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(updatedVal.Value);
                                }
                                break;
                            default:
                                _valuesGeneric[pointer.Index] = updatedVal.Value;
                                break;
                        }
                    }
                }
                else //Column does not exist
                {
                    col = ~col;
                    AddColumn(col, Column);
                    AddPage(_columnIndex[col], 0, page);
                    short ix = (short)(Row - (page << pageBits));
                    AddCell(_columnIndex[col], 0, 0, ix, new ExcelCoreValue { _value = Value });
                    var pointer = _columnIndex[col]._pages[0].Rows[0].IndexPointer;

                    ExcelCoreValue currentVal;
                    switch (pointer.Type)
                    {
                        case ComposePositionType.String:
                            currentVal = new ExcelCoreValue { _value = _valuesString[pointer.Index] };
                            break;
                        case ComposePositionType.Decimal:
                            currentVal = new ExcelCoreValue { _value = _valuesDecimal[pointer.Index] };
                            break;
                        default:
                            currentVal = _valuesGeneric[pointer.Index];
                            break;
                    }

                    var updatedVal = Updater(Value, currentVal);
                    if (updatedVal.HasValue)
                    {
                        switch (pointer.Type)
                        {
                            case ComposePositionType.String:
                                if (updatedVal.Value._value is string s && updatedVal.Value._styleId == default)
                                {
                                    _valuesString[pointer.Index] = s;
                                }
                                else
                                {
                                    _columnIndex[col]._pages[0].Rows[0].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(updatedVal.Value);
                                }
                                break;
                            case ComposePositionType.Decimal:
                                if (updatedVal.Value._value is decimal d && updatedVal.Value._styleId == default)
                                {
                                    _valuesDecimal[pointer.Index] = d;
                                }
                                else
                                {
                                    _columnIndex[col]._pages[0].Rows[0].IndexPointer = new ComposePosition { Index = _valuesGeneric.Count, Type = ComposePositionType.Generic };
                                    _valuesGeneric.Add(updatedVal.Value);
                                }
                                break;
                            default:
                                _valuesGeneric[pointer.Index] = updatedVal.Value;
                                break;
                        }
                    }
                }
            }
        }

        internal void Insert(int fromRow, int fromCol, int rows, int columns)
        {
            lock (_columnIndex)
            {

                if (columns > 0)
                {
                    var col = GetPosition(fromCol);
                    if (col < 0)
                    {
                        col = ~col;
                    }
                    for (var c = col; c < ColumnCount; c++)
                    {
                        _columnIndex[c].Index += (short)columns;
                    }
                }
                else
                {
                    var page = fromRow >> pageBits;
                    for (int c = 0; c < ColumnCount; c++)
                    {
                        var column = _columnIndex[c];
                        var pagePos = column.GetPosition(fromRow);
                        if (pagePos >= 0)
                        {
                            if (fromRow >= column._pages[pagePos].MinIndex && fromRow <= column._pages[pagePos].MaxIndex) //The row is inside the page
                            {
                                int offset = fromRow - column._pages[pagePos].IndexOffset;
                                var rowPos = column._pages[pagePos].GetPosition(offset);
                                if (rowPos < 0)
                                {
                                    rowPos = ~rowPos;
                                }
                                UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
                            }
                            else if (column._pages[pagePos].MinIndex > fromRow - 1 && pagePos > 0) //The row is on the page before.
                            {
                                int offset = fromRow - ((page - 1) << pageBits);
                                var rowPos = column._pages[pagePos - 1].GetPosition(offset);
                                if (rowPos > 0 && pagePos > 0)
                                {
                                    UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow, rows);
                                }
                            }
                            else if (column.PageCount >= pagePos + 1)
                            {
                                int offset = fromRow - column._pages[pagePos].IndexOffset;
                                var rowPos = column._pages[pagePos].GetPosition(offset);
                                if (rowPos < 0)
                                {
                                    rowPos = ~rowPos;
                                }
                                if (column._pages[pagePos].RowCount > rowPos)
                                {
                                    UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
                                }
                                else
                                {
                                    UpdateIndexOffset(column, pagePos + 1, 0, fromRow, rows);
                                }
                            }
                        }
                        else
                        {
                            UpdateIndexOffset(column, ~pagePos, 0, fromRow, rows);
                        }
                    }
                }
            }
        }
        internal void Clear(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, false);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, true);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
        {
            lock (_columnIndex)
            {
                if (columns > 0 && fromRow == 0 && rows >= ExcelPackage.MaxRows)
                {
                    DeleteColumns(fromCol, columns, shift);
                }
                else
                {
                    var toCol = fromCol + columns - 1;
                    var pageFromRow = fromRow >> pageBits;
                    for (int c = 0; c < ColumnCount; c++)
                    {
                        var column = _columnIndex[c];
                        if (column.Index >= fromCol)
                        {
                            if (column.Index > toCol) break;
                            var pagePos = column.GetPosition(fromRow);
                            if (pagePos < 0) pagePos = ~pagePos;
                            if (pagePos < column.PageCount)
                            {
                                var page = column._pages[pagePos];
                                if (shift && page.RowCount > 0 && page.MinIndex > fromRow && page.MaxIndex >= fromRow + rows)
                                {
                                    var o = page.MinIndex - fromRow;
                                    if (o < rows)
                                    {
                                        rows -= o;
                                        page.Offset -= o;
                                        UpdatePageOffset(column, pagePos, o);
                                    }
                                    else
                                    {
                                        page.Offset -= rows;
                                        UpdatePageOffset(column, pagePos, rows);
                                        continue;
                                    }
                                }
                                if (page.RowCount > 0 && page.MinIndex <= fromRow + rows - 1 && page.MaxIndex >= fromRow) //The row is inside the page
                                {
                                    var endRow = fromRow + rows;
                                    var delEndRow = DeleteCells(column._pages[pagePos], fromRow, endRow, shift);
                                    if (shift && delEndRow != fromRow) UpdatePageOffset(column, pagePos, delEndRow - fromRow);
                                    if (endRow > delEndRow && pagePos < column.PageCount && column._pages[pagePos].MinIndex < endRow)
                                    {
                                        pagePos = (delEndRow == fromRow ? pagePos : pagePos + 1);
                                        var rowsLeft = DeletePage(shift ? fromRow : delEndRow, endRow - delEndRow, column, pagePos, shift);
                                        //if (shift) UpdatePageOffset(column, pagePos, endRow - fromRow - rowsLeft);
                                        if (rowsLeft > 0)
                                        {
                                            var fr = shift ? fromRow : endRow - rowsLeft;
                                            pagePos = column.GetPosition(fr);
                                            delEndRow = DeleteCells(column._pages[pagePos], fr, shift ? fr + rowsLeft : endRow, shift);
                                            if (shift) UpdatePageOffset(column, pagePos, rowsLeft);
                                        }
                                    }
                                }
                                else if (pagePos > 0 && column._pages[pagePos].IndexOffset > fromRow) //The row is on the page before.
                                {
                                    int offset = fromRow + rows - 1 - ((pageFromRow - 1) << pageBits);
                                    var rowPos = column._pages[pagePos - 1].GetPosition(offset);
                                    if (rowPos > 0 && pagePos > 0)
                                    {
                                        if (shift) UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow + rows - 1, -rows);
                                    }
                                }
                                else
                                {
                                    if (shift && pagePos + 1 < column.PageCount) UpdateIndexOffset(column, pagePos + 1, 0, column._pages[pagePos + 1].MinIndex, -rows);
                                }
                            }
                        }
                    }
                }
            }
        }
        private void UpdatePageOffset(ColumnIndex column, int pagePos, int rows)
        {
            //Update Pageoffset

            if (++pagePos < column.PageCount)
            {
                for (int p = pagePos; p < column.PageCount; p++)
                {
                    if (column._pages[p].Offset - rows <= -PageSize)
                    {
                        column._pages[p].Index--;
                        column._pages[p].Offset -= rows - PageSize;
                    }
                    else
                    {
                        column._pages[p].Offset -= rows;
                    }
                }

                if (Math.Abs(column._pages[pagePos].Offset) > PageSize ||
                    Math.Abs(column._pages[pagePos].Rows[column._pages[pagePos].RowCount - 1].Index) > PageSizeMax) //Split or Merge???
                {
                    rows = ResetPageOffset(column, pagePos, rows);
                    ////MergePages
                    //if (column.Pages[pagePos - 1].Index + 1 == column.Pages[pagePos].Index)
                    //{
                    //    if (column.Pages[pagePos].IndexOffset + column.Pages[pagePos].Rows[column.Pages[pagePos].RowCount - 1].Index + rows -
                    //        column.Pages[pagePos - 1].IndexOffset + column.Pages[pagePos - 1].Rows[0].Index <= PageSize)
                    //    {
                    //        //Merge
                    //        MergePage(column, pagePos - 1, -rows);
                    //    }
                    //    else
                    //    {
                    //        //Split
                    //    }
                    //}
                    //rows -= PageSize;
                    //for (int p = pagePos; p < column.PageCount; p++)
                    //{                            
                    //    column.Pages[p].Index -= 1;
                    //}
                    return;
                }
            }
        }

        private int ResetPageOffset(ColumnIndex column, int pagePos, int rows)
        {
            PageIndex fromPage = column._pages[pagePos];
            PageIndex toPage;
            short pageAdd = 0;
            if (fromPage.Offset < -PageSize)
            {
                toPage = column._pages[pagePos - 1];
                pageAdd = -1;
                if (fromPage.Index - 1 == toPage.Index)
                {
                    if (fromPage.IndexOffset + fromPage.Rows[fromPage.RowCount - 1].Index -
                        toPage.IndexOffset + toPage.Rows[0].Index <= PageSizeMax)
                    {
                        MergePage(column, pagePos - 1);
                        //var newPage = new PageIndex(toPage, 0, GetSize(fromPage.RowCount + toPage.RowCount));
                        //newPage.RowCount = fromPage.RowCount + fromPage.RowCount;
                        //Array.Copy(toPage.Rows, 0, newPage.Rows, 0, toPage.RowCount);
                        //Array.Copy(fromPage.Rows, 0, newPage.Rows, toPage.RowCount, fromPage.RowCount);
                        //for (int r = toPage.RowCount; r < newPage.RowCount; r++)
                        //{
                        //    newPage.Rows[r].Index += (short)(fromPage.IndexOffset - toPage.IndexOffset);
                        //}

                    }
                }
                else //No page after 
                {
                    fromPage.Index -= pageAdd;
                    fromPage.Offset += PageSize;
                }
            }
            else if (fromPage.Offset > PageSize)
            {
                toPage = column._pages[pagePos + 1];
                pageAdd = 1;
                if (fromPage.Index + 1 == toPage.Index)
                {

                }
                else
                {
                    fromPage.Index += pageAdd;
                    fromPage.Offset += PageSize;
                }
            }
            return rows;
        }

        private int DeletePage(int fromRow, int rows, ColumnIndex column, int pagePos, bool shift)
        {
            PageIndex page = column._pages[pagePos];
            var startRows = rows;
            while (page != null && page.MinIndex >= fromRow && ((shift && page.MaxIndex < fromRow + rows) || (!shift && page.MaxIndex < fromRow + startRows)))
            {
                //Delete entire page.
                var delSize = page.MaxIndex - page.MinIndex + 1;
                rows -= delSize;
                var prevOffset = page.Offset;
                Array.Copy(column._pages, pagePos + 1, column._pages, pagePos, column.PageCount - pagePos + 1);
                column.PageCount--;
                if (column.PageCount == 0)
                {
                    return 0;
                }
                if (shift)
                {
                    for (int i = pagePos; i < column.PageCount; i++)
                    {
                        column._pages[i].Offset -= delSize;
                        if (column._pages[i].Offset <= -PageSize)
                        {
                            column._pages[i].Index--;
                            column._pages[i].Offset += PageSize;
                        }
                    }
                }
                if (column.PageCount > pagePos)
                {
                    page = column._pages[pagePos];
                    //page.Offset = pagePos == 0 ? 1 : prevOffset;  //First page can only reference to rows starting from Index == 1
                }
                else
                {
                    //No more pages, return 0
                    return 0;
                }
            }
            return rows;
        }
        ///
        private int DeleteCells(PageIndex page, int fromRow, int toRow, bool shift)
        {
            var fromPos = page.GetPosition(fromRow - (page.IndexOffset));
            if (fromPos < 0)
            {
                fromPos = ~fromPos;
            }
            var maxRow = page.MaxIndex;
            var offset = toRow - page.IndexOffset;
            if (offset > PageSizeMax) offset = PageSizeMax;
            var toPos = page.GetPosition(offset);
            if (toPos < 0)
            {
                toPos = ~toPos;
            }

            if (fromPos <= toPos && fromPos < page.RowCount && page.GetIndex(fromPos) < toRow)
            {
                if (toRow > page.MaxIndex)
                {
                    if (fromRow == page.MinIndex) //Delete entire page, late in the page delete method
                    {
                        return fromRow;
                    }
                    var r = page.MaxIndex;
                    var deletedRow = page.RowCount - fromPos;
                    page.RowCount -= deletedRow;
                    return r + 1;
                }
                else
                {
                    var rows = toRow - fromRow;
                    if (shift) UpdateRowIndex(page, toPos, rows);
                    Array.Copy(page.Rows, toPos, page.Rows, fromPos, page.RowCount - toPos);
                    page.RowCount -= toPos - fromPos;

                    return toRow;
                }
            }
            else if (shift)
            {
                UpdateRowIndex(page, toPos, toRow - fromRow);
            }
            return toRow < maxRow ? toRow : maxRow;
        }

        private static void UpdateRowIndex(PageIndex page, int toPos, int rows)
        {
            for (int r = toPos; r < page.RowCount; r++)
            {
                page.Rows[r].Index -= (short)rows;
            }
        }

        private void DeleteColumns(int fromCol, int columns, bool shift)
        {
            var fPos = GetPosition(fromCol);
            if (fPos < 0)
            {
                fPos = ~fPos;
            }
            int tPos = fPos;
            for (var c = fPos; c <= ColumnCount; c++)
            {
                tPos = c;
                if (tPos == ColumnCount || _columnIndex[c].Index >= fromCol + columns)
                {
                    break;
                }
            }

            if (ColumnCount <= fPos)
            {
                return;
            }

            if (_columnIndex[fPos].Index >= fromCol && _columnIndex[fPos].Index <= fromCol + columns)
            {
                //if (_columnIndex[fPos].Index < ColumnCount)
                //{
                if (tPos < ColumnCount)
                {
                    Array.Copy(_columnIndex, tPos, _columnIndex, fPos, ColumnCount - tPos);
                }
                ColumnCount -= (tPos - fPos);
                //}
            }
            if (shift)
            {
                for (var c = fPos; c < ColumnCount; c++)
                {
                    _columnIndex[c].Index -= (short)columns;
                }
            }
        }

        private void UpdateIndexOffset(ColumnIndex column, int pagePos, int rowPos, int row, int rows)
        {
            if (pagePos >= column.PageCount) return;    //A page after last cell.
            var page = column._pages[pagePos];
            if (rows > PageSize)
            {
                short addPages = (short)(rows >> pageBits);
                int offset = +(int)(rows - (PageSize * addPages));
                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    if (column._pages[p].Offset + offset > PageSize)
                    {
                        column._pages[p].Index += (short)(addPages + 1);
                        column._pages[p].Offset += offset - PageSize;
                    }
                    else
                    {
                        column._pages[p].Index += addPages;
                        column._pages[p].Offset += offset;
                    }

                }

                var size = page.RowCount - rowPos;
                if (page.RowCount > rowPos)
                {
                    if (column.PageCount - 1 == pagePos) //No page after, create a new one.
                    {
                        //Copy rows to next page.
                        var newPage = CopyNew(page, rowPos, size);
                        newPage.Index = (short)((row + rows) >> pageBits);
                        newPage.Offset = row + rows - (newPage.Index * PageSize) - newPage.Rows[0].Index;
                        if (newPage.Offset > PageSize)
                        {
                            newPage.Index++;
                            newPage.Offset -= PageSize;
                        }
                        AddPage(column, pagePos + 1, newPage);
                        page.RowCount = rowPos;
                    }
                    else
                    {
                        if (column._pages[pagePos + 1].RowCount + size > PageSizeMax) //Split Page
                        {
                            SplitPageInsert(column, pagePos, rowPos, rows, size, addPages);
                        }
                        else //Copy Page.
                        {
                            CopyMergePage(page, rowPos, rows, size, column._pages[pagePos + 1]);
                        }
                    }
                }
            }
            else
            {
                //Add to Pages.
                for (int r = rowPos; r < page.RowCount; r++)
                {
                    page.Rows[r].Index += (short)rows;
                }
                if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)   //Can not be larger than the max size of the page.
                {
                    AdjustIndex(column, pagePos);
                    if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)
                    {
                        pagePos = SplitPage(column, pagePos);
                    }
                    //IndexItem[] newRows = new IndexItem[GetSize(page.RowCount - page.Rows[r].Index)];
                    //var newPage = new PageIndex(newRows, r);
                    //newPage.Index = (short)(pagePos + 1);
                    //TODO: MoveRows to next page.
                }

                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    if (column._pages[p].Offset + rows < PageSize)
                    {
                        column._pages[p].Offset += rows;
                    }
                    else
                    {
                        column._pages[p].Index++;
                        column._pages[p].Offset = (column._pages[p].Offset + rows) % PageSize;
                    }
                }
            }
        }

        private void SplitPageInsert(ColumnIndex column, int pagePos, int rowPos, int rows, int size, int addPages)
        {
            var page = column._pages[pagePos];

            var rStart = -1;
            for (int r = rowPos; r < page.RowCount; r++)
            {
                if (page.IndexExpanded - (page.Rows[r].Index + rows) > PageSize)
                {
                    rStart = r;
                    break;
                }
                else
                {
                    page.Rows[r].Index += (short)rows;
                }
            }
            var rc = page.RowCount - rStart;
            page.RowCount = rStart;
            if (rc > 0)
            {
                //Copy to a new page
                var row = page.IndexOffset;
                var newPage = CopyNew(page, rStart, rc);
                var ix = (short)(page.Index + addPages);
                var offset = page.IndexOffset + rows - (ix * PageSize);
                if (offset > PageSize)
                {
                    ix += (short)(offset / PageSize);
                    offset %= PageSize;
                }
                newPage.Index = ix;
                newPage.Offset = offset;
                AddPage(column, pagePos + 1, newPage);
            }

            //Copy from next Row
        }

        private void CopyMergePage(PageIndex page, int rowPos, int rows, int size, PageIndex ToPage)
        {
            var startRow = page.IndexOffset + page.Rows[rowPos].Index + rows;
            var newRows = new IndexItem[GetSize(ToPage.RowCount + size)];
            page.RowCount -= size;
            Array.Copy(page.Rows, rowPos, newRows, 0, size);
            for (int r = 0; r < size; r++)
            {
                newRows[r].Index += (short)(page.IndexOffset + rows - ToPage.IndexOffset);
            }

            Array.Copy(ToPage.Rows, 0, newRows, size, ToPage.RowCount);
            ToPage.Rows = newRows;
            ToPage.RowCount += size;
        }
        private void MergePage(ColumnIndex column, int pagePos)
        {
            PageIndex Page1 = column._pages[pagePos];
            PageIndex Page2 = column._pages[pagePos + 1];

            var newPage = new PageIndex(Page1, 0, Page1.RowCount + Page2.RowCount);
            newPage.RowCount = Page1.RowCount + Page2.RowCount;
            Array.Copy(Page1.Rows, 0, newPage.Rows, 0, Page1.RowCount);
            Array.Copy(Page2.Rows, 0, newPage.Rows, Page1.RowCount, Page2.RowCount);
            for (int r = Page1.RowCount; r < newPage.RowCount; r++)
            {
                newPage.Rows[r].Index += (short)(Page2.IndexOffset - Page1.IndexOffset);
            }

            column._pages[pagePos] = newPage;
            column.PageCount--;

            if (column.PageCount > (pagePos + 1))
            {
                Array.Copy(column._pages, pagePos + 2, column._pages, pagePos + 1, column.PageCount - (pagePos + 1));
                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    column._pages[p].Index--;
                    column._pages[p].Offset += PageSize;
                }
            }
        }

        private PageIndex CopyNew(PageIndex pageFrom, int rowPos, int size)
        {
            IndexItem[] newRows = new IndexItem[GetSize(size)];
            Array.Copy(pageFrom.Rows, rowPos, newRows, 0, size);
            return new PageIndex(newRows, size);
        }

        internal static int GetSize(int size)
        {
            var newSize = 256;
            while (newSize < size)
            {
                newSize <<= 1;
            }
            return newSize;
        }
        private void AddCell(ColumnIndex columnIndex, int pagePos, int pos, short ix, ExcelCoreValue value)
        {
            var pageItem = columnIndex._pages[pagePos];
            if (pageItem.RowCount == pageItem.Rows.Length)
            {
                if (pageItem.RowCount == PageSizeMax) //Max size-->Split
                {
                    pagePos = SplitPage(columnIndex, pagePos);
                    if (columnIndex._pages[pagePos - 1].RowCount > pos)
                    {
                        pagePos--;
                    }
                    else
                    {
                        pos -= columnIndex._pages[pagePos - 1].RowCount;
                    }
                    pageItem = columnIndex._pages[pagePos];
                }
                else //Expand to double size.
                {
                    var rowsTmp = new IndexItem[pageItem.Rows.Length << 1];
                    Array.Copy(pageItem.Rows, 0, rowsTmp, 0, pageItem.RowCount);
                    pageItem.Rows = rowsTmp;
                }
            }
            if (pos < pageItem.RowCount)
            {
                Array.Copy(pageItem.Rows, pos, pageItem.Rows, pos + 1, pageItem.RowCount - pos);
            }

            if (value._value is string s && value._styleId == default)
            {
                pageItem.Rows[pos] = new IndexItem { Index = ix, IndexPointer = new ComposePosition { Index = _valuesString.Count, Type = ComposePositionType.String } };
                _valuesString.Add(s);
            }
            else if (value._value is decimal d && value._styleId == default)
            {
                pageItem.Rows[pos] = new IndexItem { Index = ix, IndexPointer = new ComposePosition { Index = _valuesDecimal.Count, Type = ComposePositionType.Decimal } };
                _valuesDecimal.Add(d);
            }
            else
            {
                pageItem.Rows[pos] = new IndexItem { Index = ix, IndexPointer = new ComposePosition { Index =  _valuesGeneric.Count, Type = ComposePositionType.Generic } };
                _valuesGeneric.Add(value);
            }

            


            pageItem.RowCount++;
        }

        private int SplitPage(ColumnIndex columnIndex, int pagePos)
        {
            var page = columnIndex._pages[pagePos];
            if (page.Offset != 0)
            {
                var offset = page.Offset;
                page.Offset = 0;
                for (int r = 0; r < page.RowCount; r++)
                {
                    page.Rows[r].Index -= (short)offset;
                }
            }
            //Find Split pos
            int splitPos = 0;
            for (int r = 0; r < page.RowCount; r++)
            {
                if (page.Rows[r].Index > PageSize)
                {
                    splitPos = r;
                    break;
                }
            }
            var newPage = new PageIndex(page, 0, splitPos);
            var nextPage = new PageIndex(page, splitPos, page.RowCount - splitPos, (short)(page.Index + 1), page.Offset);

            for (int r = 0; r < nextPage.RowCount; r++)
            {
                nextPage.Rows[r].Index = (short)(nextPage.Rows[r].Index - PageSize);
            }

            columnIndex._pages[pagePos] = newPage;
            if (columnIndex.PageCount + 1 > columnIndex._pages.Length)
            {
                var pageTmp = new PageIndex[columnIndex._pages.Length << 1];
                Array.Copy(columnIndex._pages, 0, pageTmp, 0, columnIndex.PageCount);
                columnIndex._pages = pageTmp;
            }
            Array.Copy(columnIndex._pages, pagePos + 1, columnIndex._pages, pagePos + 2, columnIndex.PageCount - pagePos - 1);
            columnIndex._pages[pagePos + 1] = nextPage;
            page = nextPage;
            //pos -= PageSize;
            columnIndex.PageCount++;
            return pagePos + 1;
        }

        private void AdjustIndex(ColumnIndex columnIndex, int pagePos)
        {
            var page = columnIndex._pages[pagePos];
            //First Adjust indexes
            if (page.Offset + page.Rows[0].Index >= PageSize ||
                page.Offset >= PageSize ||
                page.Rows[0].Index >= PageSize)
            {
                page.Index++;
                page.Offset -= PageSize;
            }
            else if (page.Offset + page.Rows[0].Index <= -PageSize ||
                     page.Offset <= -PageSize ||
                     page.Rows[0].Index <= -PageSize)
            {
                page.Index--;
                page.Offset += PageSize;
            }
        }

    
        private void AddPage(ColumnIndex column, int pos, short index)
        {
            AddPage(column, pos);
            column._pages[pos] = new PageIndex() { Index = index };
            if (pos > 0)
            {
                var pp = column._pages[pos - 1];
                if (pp.RowCount > 0 && pp.Rows[pp.RowCount - 1].Index > PageSize)
                {
                    column._pages[pos].Offset = pp.Rows[pp.RowCount - 1].Index - PageSize;
                }
            }
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        /// <param name="page">The new page object to add</param>
        private void AddPage(ColumnIndex column, int pos, PageIndex page)
        {
            AddPage(column, pos);
            column._pages[pos] = page;
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        private void AddPage(ColumnIndex column, int pos)
        {
            if (column.PageCount == column._pages.Length)
            {
                var pageTmp = new PageIndex[column._pages.Length * 2];
                Array.Copy(column._pages, 0, pageTmp, 0, column.PageCount);
                column._pages = pageTmp;
            }
            if (pos < column.PageCount)
            {
                Array.Copy(column._pages, pos, column._pages, pos + 1, column.PageCount - pos);
            }
            column.PageCount++;
        }
        private void AddColumn(int pos, int Column)
        {
            if (ColumnCount == _columnIndex.Length)
            {
                var colTmp = new ColumnIndex[_columnIndex.Length * 2];
                Array.Copy(_columnIndex, 0, colTmp, 0, ColumnCount);
                _columnIndex = colTmp;
            }
            if (pos < ColumnCount)
            {
                Array.Copy(_columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
            }
            _columnIndex[pos] = new ColumnIndex() { Index = (short)(Column) };
            ColumnCount++;
        }
        int _colPos = -1, _row;
        public ulong Current
        {
            get
            {
                return ((ulong)_row << 32) | (ulong)(long)(_columnIndex[_colPos].Index);
            }
        }

        public void Dispose()
        {

            if (_valuesGeneric != null)
                _valuesGeneric.Clear();

            if (_valuesString != null)
                _valuesString.Clear();

            if (_valuesDecimal != null)
                _valuesDecimal.Clear();
            
            for (var c = 0; c < ColumnCount; c++)
            {
                if (_columnIndex[c] != null)
                {
                    ((IDisposable)_columnIndex[c]).Dispose();
                }
            }

            _valuesGeneric = null;
            _valuesString = null;
            _valuesDecimal = null;

            _columnIndex = null;
        }

     
        public bool MoveNext()
        {
            return GetNextCell(ref _row, ref _colPos, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool NextCell(ref int row, ref int col)
        {

            return NextCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool NextCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount - 1;
            }
            var c = GetPosition(col);
            if (c >= 0)
            {
                if (c > maxColPos)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col);
                }
                else
                {
                    var r = GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
                    col = _columnIndex[c].Index;
                    return r;
                }
            }
            else
            {
                c = ~c;
                if (c >= ColumnCount) c = ColumnCount - 1;
                if (col > _columnIndex[c].Index)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    var r = GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
                    col = _columnIndex[c].Index;
                    return r;
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (++colPos < ColumnCount && colPos <= endColPos)
                {
                    var r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        var c = colPos + 1;
                        while (c < ColumnCount && c <= endColPos)
                        {
                            r = _columnIndex[c].GetNextRow(row);
                            if (r == row) //Exists next Row
                            {
                                colPos = c;
                                return true;
                            }
                            if (r > row && r < minRow)
                            {
                                minRow = r;
                                minCol = c;
                            }
                            c++;
                        }
                        c = startColPos;
                        if (row < endRow)
                        {
                            row++;
                            while (c < colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && (r < minRow || (r == minRow && c < minCol)) && r <= endRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c++;
                            }
                        }

                        if (minRow == int.MaxValue || minRow > endRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    if (colPos <= startColPos || row >= endRow)
                    {
                        return false;
                    }
                    colPos = startColPos - 1;
                    row++;
                    return GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos, ref int[] pagePos, ref int[] cellPos)
        {
            if (colPos == endColPos)
            {
                colPos = startColPos;
                row++;
            }
            else
            {
                colPos++;
            }

            if (pagePos[colPos] < 0)
            {
                if (pagePos[colPos] == -1)
                {
                    pagePos[colPos] = _columnIndex[colPos].GetPosition(row);
                }
            }
            else if (_columnIndex[colPos]._pages[pagePos[colPos]].RowCount <= row)
            {
                if (_columnIndex[colPos].PageCount > pagePos[colPos])
                    pagePos[colPos]++;
                else
                {
                    pagePos[colPos] = -2;
                }
            }

            var r = _columnIndex[colPos]._pages[pagePos[colPos]].IndexOffset + _columnIndex[colPos]._pages[pagePos[colPos]].Rows[cellPos[colPos]].Index;
            if (r == row)
            {
                row = r;
            }
            else
            {
            }
            return true;
        }
        internal bool PrevCell(ref int row, ref int col)
        {
            return PrevCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount - 1;
            }
            var c = GetPosition(col);
            if (c >= 0)
            {
                if (c == 0)
                {
                    if (col >= maxColPos)
                    {
                        return false;
                    }
                    if (row == minRow)
                    {
                        return false;
                    }
                    row--;
                    col = maxColPos;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    var ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
            else
            {
                c = ~c;
                if (c == 0)
                {
                    if (col >= maxColPos || row <= 0)
                    {
                        return false;
                    }
                    col = maxColPos;
                    row--;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    var ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
        }
        internal bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (--colPos >= startColPos)
                //                if (++colPos < ColumnCount && colPos <= endColPos)
                {
                    var r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row && r >= startRow)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        var c = colPos - 1;
                        if (c >= startColPos)
                        {
                            while (c >= startColPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (row > startRow)
                        {
                            c = endColPos;
                            row--;
                            while (c > colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (minRow == int.MaxValue || startRow < minRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    colPos = ColumnCount;
                    row--;
                    if (row < startRow)
                    {
                        Reset();
                        return false;
                    }
                    else
                    {
                        return GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
                    }
                }
            }
        }
        public void Reset()
        {
            _colPos = -1;
            _row = 0;
        }
     
    }
}
