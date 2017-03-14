using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerializer.CellWriters
{
    class ObjectListWriter
    {
        Dictionary<string, ObjectWrapper> ObjectLists { get; set; }
        ExcelWorksheet Sheet { get; set; }
        List<int> UpdatedColumns { get; set; }

        public int Row { get; set; }
        public int Column { get; set; }

        public ObjectListWriter ( Dictionary<string, ObjectWrapper> objectLists, ExcelWorksheet sheet, int row, int column )
        {
            ObjectLists = objectLists;
            Sheet = sheet;

            Row = row;
            Column = column;

            UpdatedColumns = new List<int>();
        }

        public void Run ()
        {
            AllValuesAreOfSameType();

            foreach ( var _object in ObjectLists )
            {
                var objectList = _object.Value.Value as List<object>;

                if ( !objectList.Any() ) continue;

                var mergeDef = ExcelSerializer.GetRange(Row, Column, Row + objectList.Count, Column);
                Console.WriteLine(mergeDef);

                Sheet.Cells[mergeDef].Merge = true;
                Sheet.Cells[Row, Column].Value = _object.Key;
                Sheet.Cells[Row, Column].Style.Font.Bold = true;
                Sheet.Cells[Row, Column].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                foreach ( var item in objectList )
                {


                    var dvw = new DirectValueWriter( as Dictionary<string, ObjectWrapper>, Sheet, Row, Column + 1, false);
                    dvw.Run();
                }


                Row += objectList.Count;
                Row += 1;

                UpdatedColumns.Add(Column);
            }

            foreach ( var updatedColumnIndex in UpdatedColumns )
                Sheet.Column(updatedColumnIndex).AutoFit();
        }

        void AllValuesAreOfSameType ()
        {
            if ( !ObjectLists.All(dv => dv.Value.ObjectType == ObjectTypes.ObjectList) )
                throw new Exception("some fields are not objectList");
        }
    }
}
