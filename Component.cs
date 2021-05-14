using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace house_uid_generator
{
	public partial class Component : MetroFramework.Forms.MetroForm
	{
		public Component()
		{
			InitializeComponent();
		}

		private void Generate_UID(object sender, EventArgs e)
		{
			Console.WriteLine("Generate_UID start");

			object[,] cell;
			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
			Workbooks workbooks = app.Workbooks;
			Workbook workbook = workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\4개아파트_15분데이터(180501-190430).xlsx");
			Sheets sheets = workbook.Worksheets;
			Worksheet worksheet = sheets.get_Item(1) as Worksheet;
			Range range = worksheet.UsedRange;

			try
			{
				cell = worksheet.UsedRange.Value;
			}
			catch
			{
				Console.WriteLine("Error!");
				return;
			}

			int row = cell.GetLength(0);
			int column = cell.Length / row;
			int user, dataStartColumn;

			user = column - 7;
			dataStartColumn = 8;

			List<string> uidList = new List<string>();
			for(int c = dataStartColumn; c < column + 1; c++)
			{
				uidList.Add(cell[3, c] + "-" + cell[4, c] + "-" + cell[5, c]);
			}

			workbook.Close();
			workbooks.Close();
			app.Quit();
			object[] obj = new object[] { range, worksheet, sheets, workbook, workbooks, app };
			for (int o = 0; o < obj.Length; o++)
				try
				{
					if (obj[o] != null)
					{
						Marshal.ReleaseComObject(obj[o]);
						obj[o] = null;
					}
				}
				catch (Exception ex)
				{
					obj[o] = null;
					throw ex;
				}
				finally
				{
					GC.Collect();
				}

			Console.WriteLine(uidList.Count);
			Thread.Sleep(5000);
			uidList.ForEach((uid) =>
			{
				Console.WriteLine(uid);
			});

			Console.WriteLine("To CSV");
			string fileName = "household_uid.csv";
			StreamWriter sw;
			sw = new StreamWriter(System.Windows.Forms.Application.StartupPath + @"\data\" + fileName, false, Encoding.GetEncoding("EUC-KR"));
			uidList.ForEach((uid) =>
			{
				sw.WriteLine(uid);
			});
			sw.Close();

			Console.WriteLine("UID Generate Success");
		}
	}
}
