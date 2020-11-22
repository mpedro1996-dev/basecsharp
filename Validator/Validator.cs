using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Base.Validator
{
	public static class Validator
	{
		#region Properties

		public static IWin32Window Owner { get; private set; }

		public static bool IsValid { get; set; }

		#endregion

		#region Methods

		public static void Begin(IWin32Window owner)
		{
			Owner = owner;
		}
		public static void End()
		{
			Owner = null;
		}


		public static void Validate(bool condition, string message = null, Label label = null)
		{
			ResetLabel(label);

			if (!condition)
			{
				if (!string.IsNullOrEmpty(message))
				{
					ShowMessage(message);
				}

				if(label != null)
				{
					ChangeColorLabel(label);
				}

			}

			IsValid = condition;
		}

		public static void Validate(bool condition, string message = null, Label[] labels = null)
		{
			ResetLabel(labels);

			if (!condition)
			{
				if (string.IsNullOrEmpty(message))
				{
					ShowMessage(message);
				}

				if (labels != null && labels.Length > 0)
				{
					foreach(var label in labels)
					{
						ChangeColorLabel(label);
					}
					
				}

			}

			IsValid = condition;
		}


		private static void ChangeColorLabel(Label label)
		{
			label.ForeColor = Color.Red;
		}

		private static void ResetLabel(Label label)
		{
			label.ForeColor = Color.Black;
		}

		private static void ResetLabel(Label[] labels)
		{
			foreach( var label in labels)
			{
				ResetLabel(label);
			}
		}


		private static void ShowMessage(string message)
		{
			MessageBox.Show(Owner, message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		#endregion
	}
}
