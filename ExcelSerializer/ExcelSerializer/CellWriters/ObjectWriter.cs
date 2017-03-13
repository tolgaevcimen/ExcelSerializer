using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerializer.CellWriters
{
    class ObjectWriter
    {
        Dictionary<string, ObjectWrapper> Objects { get; set; }
        ExcelWorksheet Sheet { get; set; }
        List<int> UpdatedColumns { get; set; }

        public int Row { get; set; }
        public int Column { get; set; }

        public ObjectWriter ( Dictionary<string, ObjectWrapper> objects, ExcelWorksheet sheet, int row, int column )
        {
            Objects = objects;
            Sheet = sheet;

            Row = row;
            Column = column;

            UpdatedColumns = new List<int>();
        }

        public void Run ()
        {
            AllValuesAreOfSameType();

            //// set headers
            //Sheet.Row(Row).Style.Font.Bold = true;


            foreach ( var _object in Objects )
            {
                var mergeDef = ExcelSerializer.GetRange(Row, Column, Row + 1, Column);
                Sheet.Cells[mergeDef].Merge = true;
                Sheet.Cells[Row, Column].Value = _object.Key;
                Sheet.Cells[Row, Column].Style.Font.Bold = true;
                Sheet.Cells[Row, Column].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                var dvw = new DirectValueWriter(_object.Value.Value as Dictionary<string, ObjectWrapper>, Sheet, Row, Column + 1);
                dvw.Run();

                Row += 2;
                UpdatedColumns.Add(Column);
            }

            foreach ( var updatedColumnIndex in UpdatedColumns )
                Sheet.Column(updatedColumnIndex).AutoFit();
        }

        void AllValuesAreOfSameType ()
        {
            if ( !Objects.All(dv => dv.Value.ObjectType == ObjectTypes.Object) )
                throw new Exception("some fields are not direct values");
        }
    }
}
