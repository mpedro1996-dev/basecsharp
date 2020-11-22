using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Base
{
	public class ExcelReader : IDisposable
	{

		#region Fields

		private Application xlApp;

		private Workbook xlWorkbook;

		private int startColumn;

		public int startRow;

		#endregion

		#region Properties

		public Sheets Worksheets { get; private set; }

		public string[,] ProcessedSheet { get; private set; }

		public int RowCount { get; private set; }

		public int ColCount { get; private set; }


		#endregion

		#region Constructor

		public ExcelReader(string path, int startRow = 1, int startColumn = 1)
		{
			try
			{
				xlApp = new Application();
				xlWorkbook = xlApp.Workbooks.Open(@path);
				Worksheets = xlWorkbook.Sheets;
				this.startRow = startRow;
				this.startColumn = startColumn;
			}
			catch(Exception ex)
			{
				MessageBox.Show($"Erro ao abrir a planilha: {ex.Message}","Erro",MessageBoxButtons.OK,MessageBoxIcon.Error);
			}
			
		}

		#endregion

		#region Methods		

		#region Process

		public void ProcessSheet(_Worksheet worksheet)
		{
			var range = worksheet.UsedRange;

			RowCount = range.Rows.Count;
			ColCount = range.Columns.Count;

			ProcessedSheet = new string[RowCount, ColCount];

			for (var i = startRow; i <= RowCount; i++)
			{
				for (var j = startColumn; j <= ColCount; j++)
				{
					ProcessedSheet[(i-1), (j-1)] = GetValue(range, i, j);
				}
			}

			FinishSheet(range, worksheet);
		}		

		#endregion

		private string GetValue(Range range, int i, int j)
		{
			if(range[i,j] != null && range[i, j].Value2 !=null)
			{
				return range[i, j].Value2.ToString();
			}

			return null;
		}
		

		private void FinishSheet(Range range, _Worksheet worksheet)
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();

			Marshal.ReleaseComObject(range);
			Marshal.ReleaseComObject(worksheet);

			range = null;
			worksheet = null;			
		}

		public void Dispose()
		{
			xlWorkbook.Close();
			xlApp.Quit();

			Marshal.ReleaseComObject(xlWorkbook);
			Marshal.ReleaseComObject(xlApp);

			xlWorkbook = null;
			xlApp = null;
		}

		#endregion
	}
}
