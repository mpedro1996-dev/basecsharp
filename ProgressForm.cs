using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Base
{
	public partial class ProgressForm : Form
	{
		private Action action;
		public ProgressForm(Action action)
		{
			InitializeComponent();
			this.action = action;
		}


		private async Task<bool> Start()
		{
			timer.Start();
			await Task.Run(action);
			timer.Stop();
			return true;
			
		}

		private void Finish()
		{			
			timer.Dispose();
			DialogResult = DialogResult.OK;
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			progressBar.Value += 10;
			if (progressBar.Value == 100)
				progressBar.Value = 0;
		}

		private async void ProgressForm_Load(object sender, EventArgs e)
		{			
			var start = await Start();
			if (start) Finish();

		}
	}
}
