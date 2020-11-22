using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Base
{
	public class ExcelWriter : IDisposable
	{

		#region Inner Types

		public class RangeIndex
		{
			public int Initial;

			public int Final;

			public RangeIndex(int initial, int final)
			{
				this.Initial = initial;
				this.Final = final;
			}
		}

		#endregion

		#region Fields

		private Application app;

		private Workbook workbook;

		private Worksheet worksheet;

		private string directoryPath;

		private string filename;

		private int sheetCount;

		#endregion

		#region Constructors

		public ExcelWriter(string directoryPath, string filename)
		{
			this.directoryPath = directoryPath;
			this.filename = filename;
			app = new Application();
			workbook = app.Workbooks.Add();			
			sheetCount = 1;
		}

		#endregion

		#region Methods
		
		public void AddWorkSheet(string name)
		{
			if(sheetCount != workbook.Worksheets.Count)
			{
				workbook.Sheets.Add(After: workbook.Sheets[workbook.Worksheets.Count]);
			}
			
			worksheet = (Worksheet)workbook.Worksheets[sheetCount];
			worksheet.Name = name;				
		}

		public void MergeCells(RangeIndex initial, RangeIndex final, string text = null)
		{
			worksheet.Range[worksheet.Cells[initial.Initial, initial.Final], worksheet.Cells[final.Initial, final.Final]].Merge();
			
			if (!string.IsNullOrEmpty(text))
			{
				worksheet.Cells[initial.Initial, initial.Final] = text;
			}

		}

		public void WriteCell(int i, int j, string text)
		{
			worksheet.Cells[i, j] = text;		
		}

		public void WriteCell(int i, int j, decimal number)
		{
			worksheet.Cells[i, j] = number;
		}

		public void FlushWorksheet()
		{
			worksheet.UsedRange.EntireColumn.AutoFit();			
			worksheet = null;
			sheetCount++;
		}


		private void FlushWorkbook()
		{
			workbook.SaveAs(@directoryPath + "\\" + filename +".xlsx");
			workbook.Close();
			Marshal.ReleaseComObject(workbook);
			workbook = null;
		}
		

		public void Dispose()
		{
			FlushWorkbook();
			app.Quit();
			Marshal.ReleaseComObject(app);

		}

		#endregion
	}
}
