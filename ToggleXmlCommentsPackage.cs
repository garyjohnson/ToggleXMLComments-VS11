using System;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Company.VSPackage1;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace ToggleXmlComments {
	[PackageRegistration(UseManagedResourcesOnly = true)]
	[InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(GuidList.guidVSPackage1PkgString)]
	public sealed class ToggleXmlCommentsPackage : Package {
		private const string VisualBasic = "{B5E9BD33-6D3E-4B5D-925E-8A43B79820B4}";
		private const string ToggleExpansionCommand = "Edit.ToggleOutliningExpansion";

		private DTE2 applicationObject;

		protected override void Initialize() {
			base.Initialize();

			applicationObject = (DTE2)GetService(typeof(SDTE));

			// Add our command handlers for menu (commands must exist in the .vsct file)
			OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if(null != mcs) {
				// Create the command for the menu item.
				CommandID menuCommandID = new CommandID(GuidList.guidVSPackage1CmdSet, (int)PkgCmdIDList.cmdidMyCommand);
				MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				mcs.AddCommand(menuItem);
			}
		}

		private void MenuItemCallback(object sender, EventArgs e) {
			CollapseXmlComments();
		}

		private void CollapseXmlComments() {
			applicationObject.UndoContext.Open("Collapse XML comments");

			try {
				foreach(CodeElement2 ce in applicationObject.ActiveDocument.ProjectItem.FileCodeModel.CodeElements) {
					ToggleSubmembers(ce);
				}
			}
			finally {
				applicationObject.UndoContext.Close();
			}
		}

		private void ToggleSubmembers(CodeElement2 codeElement) {
			try {
				EditPoint memberStart = codeElement.GetStartPoint().CreateEditPoint();
				ToggleCommentsAboveMember(memberStart, GetCommentString());
			}
			catch(Exception) { }

			if(codeElement.IsCodeType || IsNamespace(codeElement)) {
				dynamic codeOrNamespace = codeElement;
				foreach(CodeElement2 childCodeElement in codeOrNamespace.Members) {
					ToggleSubmembers(childCodeElement);
				}
			}
		}

		private void ToggleCommentsAboveMember(EditPoint editPoint, string commentPrefix) {
			try {
				int? firstLineOfComment = null;
				editPoint.StartOfLine();
				editPoint.CharLeft();
				while(!editPoint.AtStartOfDocument) {
					String line = editPoint.GetLines(editPoint.Line, editPoint.Line + 1).Trim();
					if(line.Length == 0 || line.StartsWith(commentPrefix)) {
						if(line.Length > 0) {
							firstLineOfComment = editPoint.Line;
						} else if(firstLineOfComment.HasValue) {
							ToggleExpansionAtLine(editPoint.CreateEditPoint(), firstLineOfComment.Value, commentPrefix);
							firstLineOfComment = null;
						}

						editPoint.StartOfLine();
						editPoint.CharLeft();
					} else {
						break;
					}

				}

				if(firstLineOfComment.HasValue) {
					ToggleExpansionAtLine(editPoint.CreateEditPoint(), firstLineOfComment.Value, commentPrefix);
				}

			}
			catch(Exception) { }

		}

		private void ToggleExpansionAtLine(EditPoint commentStartingPoint, int firstLineOfComment, string commentPrefix) {
			commentStartingPoint.MoveToLineAndOffset(firstLineOfComment, 1);
			while(commentStartingPoint.GetText(commentPrefix.Length) != commentPrefix) {
				commentStartingPoint.CharRight();
			}

			ToggleExpansionAtPoint(commentStartingPoint);
		}

		private static bool IsNamespace(CodeElement2 codeElement) {
			return codeElement.Kind == vsCMElement.vsCMElementNamespace;
		}

		private void ToggleExpansionAtPoint(EditPoint commentStart) {
			if(commentStart == null)
				return;

			((TextSelection)applicationObject.ActiveDocument.Selection).MoveToPoint(commentStart);
			applicationObject.ExecuteCommand(ToggleExpansionCommand);
		}

		private string GetCommentString() {
			if(applicationObject.ActiveDocument.Language == VisualBasic) {
				return "''";
			}

			return "//";
		}
	}
}
