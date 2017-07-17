//------------------------------------------------------------------------------
// <copyright file="NewGuidVSExt.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Collections.Generic;
using EnvDTE80;
using EnvDTE;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;

namespace MBFVSolutions.NewGuidVSExt
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class NewGuidVSExt
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int FileCommandId = 0x0100;
        public const int SelectionCommandId = 0x0101;
        public const string GuidPattern = @"[0-9A-Fa-f]{8}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{12}";

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("6117e7c1-7a4f-4239-ad04-d05a0de80a93");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private Regex GuidRegex = new Regex(GuidPattern);

        /// <summary>
        /// Initializes a new instance of the <see cref="NewGuidVSExt"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private NewGuidVSExt(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));

                var fileMenuCommandId = new CommandID(CommandSet, FileCommandId);
                var fileMenuItem = new OleMenuCommand(this.MenuItemFileCallback, fileMenuCommandId);
                commandService.AddCommand(fileMenuItem);

                var selectionMenuCommandId = new CommandID(CommandSet, SelectionCommandId);
                var selectionMenuItem = new OleMenuCommand(this.MenuItemSelectionCallback, selectionMenuCommandId);
                selectionMenuItem.BeforeQueryStatus += (objSender, evt) =>
                {
                    var sender = (OleMenuCommand)objSender;

                    var doc = (TextDocument)dte.ActiveDocument?.Object("TextDocument");
                    var selection = doc?.Selection;
                    var empty = selection?.IsEmpty ?? true;
                    if (!empty && !dte.ActiveDocument.ReadOnly)
                    {
                        sender.Visible = selection.TextRanges.Cast<TextRange>().Any(range => GuidRegex.IsMatch(range.StartPoint.GetText(range.EndPoint)));
                    }
                    else
                    {
                        sender.Visible = false;
                    }
                };
                commandService.AddCommand(selectionMenuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static NewGuidVSExt Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new NewGuidVSExt(package);
        }

        private string ReplaceGuids(string contents)
        {
            return GuidRegex.Replace(contents, (match) =>
            {
                var guid = Guid.NewGuid().ToString();
                // match case of old guid
                if (match.Value.ToUpperInvariant() == match.Value)
                {
                    guid = guid.ToUpperInvariant();
                }
                return guid;
            });
        }

        private void MenuItemSelectionCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
            if (dte.ActiveDocument.ReadOnly)
            {
                return;
            }
            var doc = (TextDocument)dte.ActiveDocument?.Object("TextDocument");
            var selection = doc?.Selection;
            if (selection == null)
            {
                return;
            }

            // We want to replace the whole range covered by the selection, but only update selected regions.
            // This will ensure only one undo point will be created while still handling disjointed selections
            // (like block regions) correctly.
            StringBuilder replacedText = new StringBuilder();
            EditPoint previousRangeEnd = null;
            foreach (var range in selection.TextRanges.Cast<TextRange>())
            {
                if (previousRangeEnd != null)
                {
                    replacedText.Append(previousRangeEnd.GetText(range.StartPoint));
                }
                var text = range.StartPoint.GetText(range.EndPoint);
                var replaced = ReplaceGuids(text);
                replacedText.Append(replaced);

                previousRangeEnd = range.EndPoint;
            }
            // TextRanges is apparently 1-indexed
            var startPoint = selection.TextRanges.Item(1).StartPoint;
            var endPoint = selection.TextRanges.Item(selection.TextRanges.Count).EndPoint;

            startPoint.ReplaceText(endPoint, replacedText.ToString(), (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemFileCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
            List<string> fileList;

            if( FileSelected(dte, out fileList))
            {
                foreach(var file in fileList)
                {
                    string fileContents;
                    using (System.IO.StreamReader sr = new System.IO.StreamReader(file))
                    {
                        fileContents = sr.ReadToEnd();
                    }

                    fileContents = ReplaceGuids(fileContents);

                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(file, false))
                    {
                        sw.WriteLine(fileContents);
                    }
                }
            }
        }

        public static bool FileSelected(DTE2 dte, out List<string> fileList)
        {
            var items = GetSelectedFiles(dte);

            fileList = items.ToList();

            return ((fileList != null) && (fileList.Count > 0));
        }

        public static IEnumerable<string> GetSelectedFiles(DTE2 dte)
        {
            var items = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;

            return from item in items.Cast<UIHierarchyItem>()
                   let pi = item.Object as ProjectItem
                   select pi.FileNames[1];

        }
    }
}
