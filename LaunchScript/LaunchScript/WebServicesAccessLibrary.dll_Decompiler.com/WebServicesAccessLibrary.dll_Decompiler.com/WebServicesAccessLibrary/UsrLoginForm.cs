using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

namespace WebServicesAccessLibrary;

[DesignerGenerated]
public class UsrLoginForm : Form
{
	public bool QuitByUser;

	public string Password;

	private string _UserfullName;

	[CompilerGenerated]
	[AccessedThroughProperty("PasswordTextBox")]
	private TextBox _PasswordTextBox;

	[CompilerGenerated]
	[AccessedThroughProperty("OK")]
	private Button _OK;

	[CompilerGenerated]
	[AccessedThroughProperty("Cancel")]
	private Button _Cancel;

	private IContainer components;

	[field: AccessedThroughProperty("PasswordLabel")]
	internal virtual Label PasswordLabel
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	internal virtual TextBox PasswordTextBox
	{
		[CompilerGenerated]
		get
		{
			return _PasswordTextBox;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			KeyEventHandler val = new KeyEventHandler(PasswordTextBox_KeyDown);
			TextBox passwordTextBox = _PasswordTextBox;
			if (passwordTextBox != null)
			{
				((Control)passwordTextBox).KeyDown -= val;
			}
			_PasswordTextBox = value;
			passwordTextBox = _PasswordTextBox;
			if (passwordTextBox != null)
			{
				((Control)passwordTextBox).KeyDown += val;
			}
		}
	}

	internal virtual Button OK
	{
		[CompilerGenerated]
		get
		{
			return _OK;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler eventHandler = OK_Click;
			Button oK = _OK;
			if (oK != null)
			{
				((Control)oK).Click -= eventHandler;
			}
			_OK = value;
			oK = _OK;
			if (oK != null)
			{
				((Control)oK).Click += eventHandler;
			}
		}
	}

	internal virtual Button Cancel
	{
		[CompilerGenerated]
		get
		{
			return _Cancel;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler eventHandler = Cancel_Click;
			Button cancel = _Cancel;
			if (cancel != null)
			{
				((Control)cancel).Click -= eventHandler;
			}
			_Cancel = value;
			cancel = _Cancel;
			if (cancel != null)
			{
				((Control)cancel).Click += eventHandler;
			}
		}
	}

	[field: AccessedThroughProperty("usernameTextBox")]
	internal virtual TextBox usernameTextBox
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("Label1")]
	internal virtual Label Label1
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
		}
		finally
		{
			((Form)this).Dispose(disposing);
		}
	}

	public UsrLoginForm(string UserName, int frmleft, int frmtop)
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_002a: Expected O, but got Unknown
		((Form)this).Shown += UsrLoginForm_Shown;
		((Form)this).FormClosing += new FormClosingEventHandler(UsrLoginForm_FormClosing);
		InitializeComponent();
		if (frmleft != 0)
		{
			((Control)this).Left = frmleft;
		}
		if (frmtop != 0)
		{
			((Control)this).Top = frmtop;
		}
		if (frmleft != 0 && frmtop != 0)
		{
			((Form)this).StartPosition = (FormStartPosition)0;
		}
		else
		{
			((Form)this).StartPosition = (FormStartPosition)1;
		}
		Password = "";
		((Form)this).TopMost = true;
		_UserfullName = UserName;
		usernameTextBox.Text = UserName;
	}

	[DebuggerStepThrough]
	private void InitializeComponent()
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Expected O, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Expected O, but got Unknown
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(UsrLoginForm));
		PasswordLabel = new Label();
		PasswordTextBox = new TextBox();
		OK = new Button();
		Cancel = new Button();
		usernameTextBox = new TextBox();
		Label1 = new Label();
		((Control)this).SuspendLayout();
		componentResourceManager.ApplyResources(PasswordLabel, "PasswordLabel");
		((Control)PasswordLabel).Name = "PasswordLabel";
		componentResourceManager.ApplyResources(PasswordTextBox, "PasswordTextBox");
		((Control)PasswordTextBox).Name = "PasswordTextBox";
		componentResourceManager.ApplyResources(OK, "OK");
		((Control)OK).Name = "OK";
		Cancel.DialogResult = (DialogResult)2;
		componentResourceManager.ApplyResources(Cancel, "Cancel");
		((Control)Cancel).Name = "Cancel";
		componentResourceManager.ApplyResources(usernameTextBox, "usernameTextBox");
		((Control)usernameTextBox).Name = "usernameTextBox";
		componentResourceManager.ApplyResources(Label1, "Label1");
		((Control)Label1).Name = "Label1";
		((Form)this).AcceptButton = (IButtonControl)(object)OK;
		componentResourceManager.ApplyResources(this, "$this");
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).CancelButton = (IButtonControl)(object)Cancel;
		((Control)this).Controls.Add((Control)(object)Label1);
		((Control)this).Controls.Add((Control)(object)usernameTextBox);
		((Control)this).Controls.Add((Control)(object)Cancel);
		((Control)this).Controls.Add((Control)(object)OK);
		((Control)this).Controls.Add((Control)(object)PasswordTextBox);
		((Control)this).Controls.Add((Control)(object)PasswordLabel);
		((Form)this).FormBorderStyle = (FormBorderStyle)3;
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "UsrLoginForm";
		((Form)this).SizeGripStyle = (SizeGripStyle)2;
		((Form)this).TopMost = true;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	private void PasswordTextBox_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
		if ((int)e.KeyCode == 13)
		{
			OK.PerformClick();
		}
	}

	private void UsrLoginForm_Shown(object sender, EventArgs e)
	{
	}

	private void OK_Click(object sender, EventArgs e)
	{
		Password = PasswordTextBox.Text;
		((Control)this).Hide();
	}

	private void Cancel_Click(object sender, EventArgs e)
	{
		QuitByUser = true;
		((Form)this).Close();
	}

	private void UsrLoginForm_FormClosing(object sender, FormClosingEventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Invalid comparison between Unknown and I4
		if ((int)e.CloseReason == 3)
		{
			QuitByUser = true;
		}
	}
}
