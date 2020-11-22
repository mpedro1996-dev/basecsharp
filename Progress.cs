using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Base
{
	public static class Progress
	{
		private static Form form;
		
		public static void Start(Action action)
		{
			form = new ProgressForm(action);
			form.ShowDialog();					
		}

	
	}
}
