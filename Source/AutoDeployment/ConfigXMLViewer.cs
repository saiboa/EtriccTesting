using System.Windows.Forms;
using System.Xml;

namespace Epia3Deployment
{
	/// <summary>
	/// Summary description for ConfigXMLViewer.
	/// </summary>
	public class ConfigXMLViewer : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TreeView tvXML;

		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Enums/Constants
		#endregion // —— Enums/Constants ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Structs/Classes
		#endregion // —— Structs/Classes ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Fields
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		#endregion // —— Fields ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Delegates/Events
		#endregion //  —— Delegates/Events ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Constructors/Destructors/Cleanup
		/// <summary>
		/// Default constructor.
		/// </summary>
		public ConfigXMLViewer()
		{
			// Required for Windows Form Designer support
			InitializeComponent();

			// TODO: Add any constructor code after InitializeComponent call
		}

		public ConfigXMLViewer(string xmlfile)
		{
			// Required for Windows Form Designer support
			InitializeComponent();
			OpenXmlDoc(xmlfile);
			// TODO: Add any constructor code after InitializeComponent call
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if ( disposing )
			{
				if (components != null )
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		#endregion // —— Constructors/Destructors/Cleanup ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Properties
		#endregion // —— Properties ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?


		// ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
		#region Methods
		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tvXML = new System.Windows.Forms.TreeView();
			this.SuspendLayout();
			// 
			// tvXML
			// 
			this.tvXML.ImageIndex = -1;
			this.tvXML.Location = new System.Drawing.Point(32, 16);
			this.tvXML.Name = "tvXML";
			this.tvXML.SelectedImageIndex = -1;
			this.tvXML.Size = new System.Drawing.Size(384, 408);
			this.tvXML.TabIndex = 0;
			// 
			// ConfigXMLViewer
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(460, 466);
			this.Controls.Add(this.tvXML);
			this.Name = "ConfigXMLViewer";
			this.Text = "ConfigXMLViewer";
			this.ResumeLayout(false);

		}
		#endregion // Windows Form Designer generated code
		public void OpenXmlDoc(string xmlFile)
		{
			tvXML.Nodes.Clear();

			//Initailize XML objects
			XmlTextReader xmlR = new XmlTextReader(xmlFile);
			XmlDocument xmlDoc = new XmlDocument();
			xmlDoc.Load(xmlR);
			xmlR.Close();

			//navigate inside the XLM document & populate the tree
			TreeNode tnXML = new TreeNode("Configuration");
			tvXML.Nodes.Add(tnXML);

			XmlNode xnGuiNode = xmlDoc.DocumentElement;
			XMLRecursion(xnGuiNode, tnXML);

			tvXML.ExpandAll();
			
		}

		private void XMLRecursion(XmlNode xnGuiNode, TreeNode tnXML)
		{
			TreeNode tmpTN = new TreeNode(xnGuiNode.Name + " " + xnGuiNode.Value);

			if (xnGuiNode.Value == "false")
			{
				tmpTN.ForeColor = System.Drawing.Color.Red;
			}

			tnXML.Nodes.Add(tmpTN);

			//preparing recursive call
			if (xnGuiNode.HasChildNodes)
			{
				XmlNode tmpXN = xnGuiNode.FirstChild;
				while (tmpXN != null)
				{
					XMLRecursion(tmpXN, tmpTN);
					tmpXN = tmpXN.NextSibling;
				}
			}
		}
	
		#endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

	} // class
} // namespace
