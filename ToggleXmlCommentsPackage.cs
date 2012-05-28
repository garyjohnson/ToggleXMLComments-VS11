using System;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Company.VSPackage1;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace ToggleXmlComments
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVSPackage1PkgString)]
    public sealed class ToggleXmlCommentsPackage : Package {
	
    	private DTE2 _applicationObject;

        protected override void Initialize()
        {
            base.Initialize();

			_applicationObject = (EnvDTE80.DTE2)this.GetService(typeof(SDTE));
            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidVSPackage1CmdSet, (int)PkgCmdIDList.cmdidMyCommand);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
			CollapseXmlComments();
        }

		private void CollapseXmlComments() {
			try {
				_applicationObject.UndoContext.Open("Collapse XML comments");

				foreach(CodeElement2 ce in _applicationObject.ActiveDocument.ProjectItem.FileCodeModel.CodeElements) {
					CollapseSubmembers(ce, false);
				}
				_applicationObject.UndoContext.Close();
			}
			catch(Exception ex) {
				_applicationObject.UndoContext.Close();
			}
		}

		private void CollapseSubmembers(CodeElement2 ce, bool toggle = true) {
			EditPoint memberStart;
			EditPoint commentStart;
			EditPoint commentEnd;
			String comChars;

			switch(_applicationObject.ActiveDocument.ProjectItem.FileCodeModel.Language) {
				case "{B5E9BD33-6D3E-4B5D-925E-8A43B79820B4}": {
						comChars = "'''";
						break;
					}
				default: {
						comChars = "///";
						break;
					}
			}

			try {
				memberStart = ce.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint();
				commentStart = GetCommentStart(memberStart.CreateEditPoint(), comChars);
				commentEnd = GetCommentEnd(commentStart.CreateEditPoint(), comChars);
				if(toggle) {
					((TextSelection)_applicationObject.ActiveDocument.Selection).MoveToPoint(commentStart);
					_applicationObject.ExecuteCommand("Edit.ToggleOutliningExpansion");
				}
				else {
					commentStart.OutlineSection(commentEnd);
				}
			}
			catch(Exception) {
			}

			if(ce.IsCodeType) {
				foreach(CodeElement2 ce2 in ((CodeType)ce).Members) {
					CollapseSubmembers(ce2);
				}
			}
			else if(ce.Kind == vsCMElement.vsCMElementNamespace) {
				foreach(CodeElement2 ce2 in ((CodeNamespace)ce).Members) {
					CollapseSubmembers(ce2);
				}
			}
		}

		private EditPoint GetCommentStart(EditPoint ep, string comChars) {
			try {
				String line;
				int lastCommentLine = 0;
				ep.StartOfLine();
				ep.CharLeft();
				while(!ep.AtStartOfDocument) {
					line = ep.GetLines(ep.Line, ep.Line + 1).Trim();
					if(line.Length == 0 || line.StartsWith(comChars)) {
						if(line.Length > 0) {
							lastCommentLine = ep.Line;
						}

						ep.StartOfLine();
						ep.CharLeft();
					}
					else {
						break;
					}
				}

				ep.MoveToLineAndOffset(lastCommentLine, 1);
				while(ep.GetText(comChars.Length) != comChars) {
					ep.CharRight();
				}

				return ep.CreateEditPoint();
			}
			catch(Exception) {
			}

			return null;
		}

		private EditPoint GetCommentEnd(EditPoint ep, string comChars) {
			try {
				String line;
				EditPoint lastCommentPoint = ep.CreateEditPoint();

				ep.EndOfLine();
				ep.CharRight();
				while(!ep.AtEndOfDocument) {
					line = ep.GetLines(ep.Line, ep.Line + 1).Trim();
					if(line.StartsWith(comChars)) {
						lastCommentPoint = ep.CreateEditPoint();
						ep.EndOfLine();
						ep.CharRight();

					}
					else {
						break;
					}
				}

				lastCommentPoint.EndOfLine();
				return lastCommentPoint;
			}
			catch(Exception) {
			}

			return null;
		}
    }
}
