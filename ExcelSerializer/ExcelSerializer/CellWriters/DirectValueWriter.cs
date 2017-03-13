using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerializer.CellWriters
{
    class DirectValueWriter
    {
        Dictionary<string, ObjectWrapper> DirectValues { get; set; }
        ExcelWorksheet Sheet { get; set; }
        List<int> UpdatedColumns { get; set; }
        bool WriteHeader { get; set; }

        public int Row { get; set; }
        public int Column { get; set; }

        public DirectValueWriter ( Dictionary<string, ObjectWrapper> directValues, ExcelWorksheet sheet, int row, int column, bool writeHeader = true )
        {
            DirectValues = directValues;
            Sheet = sheet;

            Row = row;
            Column = column;

            WriteHeader = writeHeader;

            UpdatedColumns = new List<int>();
        }

        public void Run ()
        {
            AllValuesAreDirectValues();

            // set headers
            Sheet.Row(Row).Style.Font.Bold = true;

            foreach ( var directValue in DirectValues )
            {
                if ( WriteHeader )
                    Sheet.Cells[Row, Column].Value = directValue.Key;

                Sheet.Cells[Row + 1, Column].Value = directValue.Value.Value;

                UpdatedColumns.Add(Column++);
            }

            foreach ( var updatedColumnIndex in UpdatedColumns )
                Sheet.Column(updatedColumnIndex).AutoFit();
        }

        void AllValuesAreDirectValues ()
        {
            if ( !DirectValues.All(dv => dv.Value.ObjectType == ObjectTypes.DirectValue) )
                throw new Exception("some fields are not direct values");
        }
    }
}
