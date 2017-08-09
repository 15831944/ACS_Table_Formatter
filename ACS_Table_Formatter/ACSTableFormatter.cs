using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

// [assembly: ExtensionApplication(null)]
[assembly: CommandClass(typeof(ACS_Table_Formatter.ACSTableFormatter))]

namespace ACS_Table_Formatter
{
	public class ACSTableFormatter
	{

		[CommandMethod("acs_TblFrmt")]
		public void ACS_TblFrmt()
		{
			try
			{
				ObjectId tableID = ObjectId.Null;
				do
				{
					tableID = GetACS_Table();
				}
				while (tableID == null);
				if (tableID != null)
				{
					ModifyTable(tableID);
				}
			}
			catch (Autodesk.AutoCAD.Runtime.Exception ex)
			{
				Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog($"Exception: {ex}");
			}
		}

		private ObjectId GetACS_Table()
		{
			// Get the current document editor
			Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
			Database db = Application.DocumentManager.MdiActiveDocument.Database;

			ObjectId tableId = ObjectId.Null;

			//Seleciton options, with single selection
			PromptSelectionOptions Options = new PromptSelectionOptions();
			Options.SingleOnly = true;
			Options.SinglePickInSpace = true;
			Options.MessageForAdding = "Select A \"ACS\" Piece Mark Table; ";

			// Create a TypedValue array to define the filter criteria
			TypedValue[] acTypValAr = new TypedValue[2];
			acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "ACAD_TABLE"), 0);
			acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Z-ANNO-TABL"), 1);

			// Assign the filter criteria to a SelectionFilter object
			SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

			// Request for objects to be selected in the drawing area
			PromptSelectionResult acSSPrompt;
			acSSPrompt = acDocEd.GetSelection(Options, acSelFtr);

			// If the prompt status is OK, objects were selected
			if (acSSPrompt.Status == PromptStatus.OK)
			{
				SelectionSet acSSet = acSSPrompt.Value;
				foreach (ObjectId objID in acSSet.GetObjectIds())
				{
					tableId = objID;
				}
			}
			else
			{
				Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Not an \"ACS\" Piece Mark Table...");
			}
			return tableId;
		}

		private void ModifyTable(ObjectId tableID)
		{
			// Get the current document editor
			Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
			Database db = Application.DocumentManager.MdiActiveDocument.Database;

			using (Transaction myT = db.TransactionManager.StartTransaction())
			{
				Table tbl = db.TransactionManager.GetObject(tableID, OpenMode.ForRead, true) as Table;
				if (tbl.Columns.Count == 4 || tbl.Columns.Count == 10 || tbl.Columns.Count == 13)
				{
					// Upgrade table to write
					tbl.UpgradeOpen();
					switch (tbl.Columns.Count)
					{
						case 4:
							tbl.Columns[0].Width = 1.375;
							tbl.Columns[1].Width = 0.53125;
							tbl.Columns[2].Width = 1.03125;
							tbl.Columns[3].Width = 0.5625;
							break;
						case 10:
							tbl.Columns[0].Width = 2.15625;
							tbl.Columns[1].Width = 1.000;
							tbl.Columns[2].Width = 0.90625;
							tbl.Columns[3].Width = 1.125;
							tbl.Columns[4].Width = 0.90625;
							tbl.Columns[5].Width = 1.21875;
							tbl.Columns[6].Width = 1.000;
							tbl.Columns[7].Width = 1.46875;
							tbl.Columns[8].Width = 1.3125;
							tbl.Columns[9].Width = 1.625;
							break;
						case 13:
							tbl.Columns[0].Width = 0.71875;
							tbl.Columns[1].Width = 0.375;
							tbl.Columns[2].Width = 0.3125;
							tbl.Columns[3].Width = 0.5625;
							tbl.Columns[4].Width = 0.46875;
							tbl.Columns[5].Width = 1.3125;
							tbl.Columns[6].Width = 1.3125;
							tbl.Columns[7].Width = 1.375;
							tbl.Columns[8].Width = 1.3125;
							tbl.Columns[9].Width = 1.3125;
							tbl.Columns[10].Width = 0.4375;
							tbl.Columns[11].Width = 0.4375;
							tbl.Columns[12].Width = 1.3125;
							break;
					}
					myT.Commit();
				}
				else
				{
					myT.Abort();
					Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog($"Selected Table Number of Columns: {tbl.Columns.Count}\nMust be 4,10 or 13\nNothing Modifed...");
				}
			}
		}
	}
}
