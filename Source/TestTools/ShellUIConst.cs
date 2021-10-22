using System.Windows.Automation;

namespace TestTools
{
    public static class ShellUIConst
    {
        // common condition
        public static Condition cButton = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                  );

        public static Condition cCheckBox = new AndCondition(
                     new PropertyCondition(AutomationElement.ClassNameProperty, "CheckBox"),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)
                   );

        public static Condition cCheckBoxEnabled = new AndCondition(
                     new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)
                   );

        public static Condition cComboBox = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ComboBox"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox)
                        );

        public static Condition cCustom = new AndCondition(
                     new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                 );

        public static Condition cDataGrid = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ListView"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid)
                        );

        public static Condition cDataItem = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ListViewItem"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem)
                        );

        public static Condition cEdit = new AndCondition(
                 new PropertyCondition(AutomationElement.ClassNameProperty, "TextBox"),
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)
                   );

        public static Condition cGroup = new AndCondition(
                       new PropertyCondition(AutomationElement.ClassNameProperty, "GroupBox"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group)
                       );

        public static Condition cGroupElement = new AndCondition(
                       new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group)
                       );

        public static Condition cList = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ListView"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                    );

        public static Condition cListBox = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ListBox"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                    );

        public static Condition cListView = new AndCondition(
                                      new PropertyCondition(AutomationElement.ClassNameProperty, "ListView"),
                                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                                          );

        public static Condition cImage = new AndCondition(
                     new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image)
                   );

        public static Condition cListItem = new AndCondition(
                     new PropertyCondition(AutomationElement.ClassNameProperty, "ListBoxItem"),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)
                   );

        public static Condition cListItemElement = new AndCondition(
                     new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)
                   );


        public static Condition cMenu = new AndCondition(
                        new PropertyCondition(AutomationElement.ClassNameProperty, "ContextMenu"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                        );

        public static Condition CMenuByClassName(string Name)
        {
            Condition cMenu = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, Name),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                             );
            return cMenu;

        }

        public static Condition cElement = new AndCondition(
                       new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                       new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                       );


        public static Condition cMenuItem = new AndCondition(
                       new PropertyCondition(AutomationElement.ClassNameProperty, "MenuItem"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                       );

        public static Condition cMenuItemElement = new AndCondition(
                       new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                       );

        public static Condition cTab = new AndCondition(
                       new PropertyCondition(AutomationElement.ClassNameProperty, "TabControl"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab)
                       );

        public static Condition cTabItem = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, "TabItem"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)
                      );

        public static Condition cTable = new AndCondition(
                       new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Table)
                       );


        public static Condition cText = new AndCondition(
                        new PropertyCondition(AutomationElement.ClassNameProperty, "TextBlock"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)
                        );

        public static Condition cTextElement = new AndCondition(
                         new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)
                         );


        public static Condition cTextIsOnScreen = new AndCondition(
                       new PropertyCondition(AutomationElement.IsOffscreenProperty, false),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)
                   );

        public static Condition cScrollBar = new AndCondition(
                   new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ScrollBar)
               );

        public static Condition cTextIsOffScreen = new AndCondition(
                     new PropertyCondition(AutomationElement.IsOffscreenProperty, true),
                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)
                 );

        public static Condition cToolBar = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "ToolBar"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
                  );

        public static Condition cTreeItemIsOnScreen = new AndCondition(
                       new PropertyCondition(AutomationElement.IsOffscreenProperty, false),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem)
                   );

        public static Condition cTreeItem = new AndCondition(
                       new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem)
                   );

        // 
        public static Condition cButtonConfirm = new AndCondition(
                               new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                               new PropertyCondition(AutomationElement.AutomationIdProperty, "confirmBtn"),
                               new PropertyCondition(AutomationElement.NameProperty, "Confirm")
                               );

        public static Condition cButtonCancel = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                             new PropertyCondition(AutomationElement.AutomationIdProperty, "cancelBtn"),
                             new PropertyCondition(AutomationElement.NameProperty, "Cancel")
                             );
        // IO Emulation
        public static Condition cRadioButtonOn = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "RadioButton"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                             new PropertyCondition(AutomationElement.NameProperty, "On")
                             );

        public static Condition cRadioButtonOff = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "RadioButton"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                             new PropertyCondition(AutomationElement.NameProperty, "Off")
                             );

        public static Condition cRadioButtonHigh = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "RadioButton"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                             new PropertyCondition(AutomationElement.NameProperty, "High")
                             );

        public static Condition cRadioButtonLow = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "RadioButton"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                             new PropertyCondition(AutomationElement.NameProperty, "Low")
                             );

        public static Condition cWindow = new AndCondition(
                             new PropertyCondition(AutomationElement.IsControlElementProperty, true),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                         );

        /// <summary>
        ///     IO Management Condition
        ///     1)
        ///     2)
        ///     3)
        ///     
        ///     IO Management pop-up Window Condition
        ///     (1) cNewWindowCustom: pop-up window for make new bit
        ///     
        /// </summary>
        // under details group
        public static Condition cBkBusCouplerView = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "BkBusCouplerView"),
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                         );

        public static Condition cListViewTerminals = new AndCondition(
                             new PropertyCondition(AutomationElement.ClassNameProperty, "ListView"),
                                            new PropertyCondition(AutomationElement.AutomationIdProperty, "Terminals"),
                                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                                                );

        // under bit configuration Tab 
        // VirtualTerminal belong to VirtualCouplerView
        /*
        public static  Condition cCustomVirtualTerminalView = new AndCondition(
         new PropertyCondition(AutomationElement.ClassNameProperty, "VirtualTerminalView"),
         new PropertyCondition(AutomationElement.AutomationIdProperty, "Terminal"),
         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
       );
        */
        // BKVirtualTerminal belong to Test  Virtual CouplerView
        public static Condition cCustomBkDigIOTerminalView = new AndCondition(
         new PropertyCondition(AutomationElement.ClassNameProperty, "BkDigIOTerminalView"),
         new PropertyCondition(AutomationElement.AutomationIdProperty, "Terminal"),
         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
       );

        // under Details Group CustomVirtualBitView belong to "VirtualTerminalView 
        public static Condition cCustomVirtualBitView = new AndCondition(
                   new PropertyCondition(AutomationElement.ClassNameProperty, "VirtualBitView"),
                   new PropertyCondition(AutomationElement.AutomationIdProperty, "Widget"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                 );

        // or under Details Group CustomBkBitView belong to "BkDigIOTerminalView 
        public static Condition cCustomBkBitView = new AndCondition(
                   new PropertyCondition(AutomationElement.ClassNameProperty, "BkBitView"),
                   new PropertyCondition(AutomationElement.AutomationIdProperty, "Widget"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                 );




        /// <summary>
        ///     new pop-up window custom
        ///         CreateEditDialogViewCustom: pop-up window create and edit content with confirm and cancel button
        ///         ConfirmationDialogViewCustom : pop-up window only with confirm and cancel button
        /// </summary>
        public static Condition cCustomCreateEditDialogView = new AndCondition(
                                  new PropertyCondition(AutomationElement.ClassNameProperty, "CreateEditDialogView"),
                                  new PropertyCondition(AutomationElement.AutomationIdProperty, "dialog"),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                                );


        public static Condition cConfirmationDialogViewCustom = new AndCondition(
                                  new PropertyCondition(AutomationElement.ClassNameProperty, "ConfirmationDialogView"),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                                );


        /// <summary>
        /// 
        /// ErrorDialog window
        ///     1) cErrorDialogViewCustom
        ///     2) cButtonOK
        /// </summary>
        public static Condition cErrorDialogViewCustom = new AndCondition(
                                 new PropertyCondition(AutomationElement.ClassNameProperty, "ErrorDialogView"),
                                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                               );

        public static Condition cButtonOK = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                      new PropertyCondition(AutomationElement.AutomationIdProperty, "okBtn"),
                      new PropertyCondition(AutomationElement.NameProperty, "OK")
                  );


        // Condition Function
        public static Condition CUIByName(string Name)
        {
            Condition cButton = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                              );
            return cButton;
        }



        public static Condition CButtonByName(string Name)
        {
            Condition cButton = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                              );
            return cButton;
        }

        public static Condition CButtonByID(string ID)
        {
            Condition cButton = new AndCondition(
                              new PropertyCondition(AutomationElement.AutomationIdProperty, ID),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                              );
            return cButton;
        }

        public static Condition CButtonByClassName(string ClassName)
        {
            Condition cButton = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, ClassName),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                              );
            return cButton;
        }


        public static Condition CButtonByNameAndId(string Name, string Id)
        {
            Condition cButton = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.AutomationIdProperty, Id),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                              );
            return cButton;
        }

        public static Condition CCheckBoxByName(string Name)
        {
            Condition cCheckBox = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, Name),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)
                  );
            return cCheckBox;
        }



        public static Condition CComboBoxByName(string Name)
        {
            Condition cComboBox = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox)
                              );
            return cComboBox;
        }

        public static Condition CDataGridByClassName(string className)
        {
            Condition cDataGrid = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, className),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid)
                              );
            return cDataGrid;

        }

        public static Condition CDataItemByClassName(string className)
        {
            Condition cDataItem = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, className),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem)
                              );
            return cDataItem;

        }

        public static Condition CDocumentByClassName(string className)
        {
            Condition cEdit = new AndCondition(
               new PropertyCondition(AutomationElement.ClassNameProperty, className),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document)
                 );
            return cEdit;
        }

        public static Condition CEditByName(string name)
        {
            Condition cEdit = new AndCondition(
               new PropertyCondition(AutomationElement.NameProperty, name),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)
                 );
            return cEdit;
        }

        public static Condition CEditByClassName(string className)
        {
            Condition cEdit = new AndCondition(
               new PropertyCondition(AutomationElement.ClassNameProperty, className),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)
                 );
            return cEdit;
        }

        public static Condition CEditByID(string Id)
        {
            Condition cEdit = new AndCondition(
               new PropertyCondition(AutomationElement.AutomationIdProperty, Id),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)
                 );
            return cEdit;
        }

        public static Condition CElementByName(string Name)
        {
            Condition cElement = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.IsControlElementProperty, true)
                              );
            return cElement;
        }

        public static Condition CElementById(string Id)
        {
            Condition cElement = new AndCondition(
                              new PropertyCondition(AutomationElement.AutomationIdProperty, Id),
                              new PropertyCondition(AutomationElement.IsControlElementProperty, true)
                              );
            return cElement;
        }

        public static Condition CElementByClassName(string ClassName)
        {
            Condition cPane = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, ClassName),
                              new PropertyCondition(AutomationElement.IsControlElementProperty, true)
                              );
            return cPane;
        }

        public static Condition CMenuItemByName(string Name)
        {
            Condition cElement = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                              );
            return cElement;
        }

        public static Condition CPaneByName(string Name)
        {
            Condition cPane = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, Name),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                              );
            return cPane;
        }

        public static Condition CPaneByClassName(string ClassName)
        {
            Condition cPane = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, ClassName),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                              );
            return cPane;
        }

        public static Condition CProgressBarByClassName(string className)
        {
            Condition cProgressBar = new AndCondition(
                              new PropertyCondition(AutomationElement.ClassNameProperty, className),
                              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ProgressBar)
                              );
            return cProgressBar;
        }

        public static Condition CCalendarByClassName(string className)
        {
            Condition cCalendar = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, className),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Calendar)
                  );
            return cCalendar;
        }

        public static Condition CCustomById(string Id)
        {
            Condition cCustom = new AndCondition(
                      new PropertyCondition(AutomationElement.AutomationIdProperty, Id),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                  );
            return cCustom;
        }

        public static Condition CCustomByClassName(string className)
        {
            Condition cCustom = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, className),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                  );
            return cCustom;
        }

        public static Condition CCustomByName(string name)
        {
            Condition cCustom = new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, name),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                  );
            return cCustom;
        }

        public static Condition CRadioButtonByName(string Name)
        {
            Condition cRadioButton = new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, Name),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton)
                  );
            return cRadioButton;
        }

        public static Condition CTextByClassName(string Name)
        {
            Condition cTest = new AndCondition(
                      new PropertyCondition(AutomationElement.ClassNameProperty, Name),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)
                  );
            return cTest;
        }


        public static Condition CToolBarByClassName(string className)
        {
            Condition cToolBar = new AndCondition(
               new PropertyCondition(AutomationElement.ClassNameProperty, className),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
                 );
            return cToolBar;
        }

        public static Condition CWindowByName(string Name)
        {
            Condition cWindow = new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, Name),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                  );
            return cWindow;
        }

        public static Condition CWindowById(string Id)
        {
            Condition cWindow = new AndCondition(
                      new PropertyCondition(AutomationElement.AutomationIdProperty, Id),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                  );
            return cWindow;
        }

        public static Condition CGroupByClassName(string ClassName)
        {
            Condition cGroup = new AndCondition(
                       //new PropertyCondition(AutomationElement.ClassNameProperty, "GroupBox"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group),
                       new PropertyCondition(AutomationElement.ClassNameProperty, ClassName)
                       );
            return cGroup;
        }


        public static Condition CGroupByName(string Name)
        {
            Condition cGroup = new AndCondition(
                       //new PropertyCondition(AutomationElement.ClassNameProperty, "GroupBox"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group),
                       new PropertyCondition(AutomationElement.NameProperty, Name)
                       );
            return cGroup;
        }

        public static Condition CListById(string Id)
        {
            Condition cList = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.AutomationIdProperty, Id)
                );
            return cList;
        }

        public static Condition CListByClassName(string ClassName)
        {
            Condition cList = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ClassNameProperty, ClassName)
                );
            return cList;
        }

        public static Condition CListViewByID(string ID)
        {
            Condition cListView = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ClassNameProperty, "ListView"),
                        new PropertyCondition(AutomationElement.AutomationIdProperty, ID)
                );
            return cListView;
        }

        public static Condition CListViewByClassName(string classname)
        {
            Condition cListView = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ClassNameProperty, classname)

                );
            return cListView;
        }

        public static Condition CListItemByName(string Name)
        {
            Condition cListItem = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.NameProperty, Name)
                );
            return cListItem;
        }

        public static Condition CListItemByClassName(string className)
        {
            Condition cListItem = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.ClassNameProperty, className)
                );
            return cListItem;
        }

        public static Condition CPaneByID(string ID)
        {
            Condition cPane = new AndCondition(
                   //new PropertyCondition(AutomationElement.ClassNameProperty, "ScrollViewer"),
                   new PropertyCondition(AutomationElement.AutomationIdProperty, ID),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                );
            return cPane;
        }





        public static Condition CTabItemByName(string Name)
        {
            Condition cTabItem = new AndCondition(
                       new PropertyCondition(AutomationElement.ClassNameProperty, "TabItem"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                       new PropertyCondition(AutomationElement.NameProperty, Name)
                       );

            return cTabItem;
        }

        public static Condition CTextByName(string Name)
        {
            Condition cText = new AndCondition(
                                  //new PropertyCondition(AutomationElement.ClassNameProperty, "TextBlock"),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                                  new PropertyCondition(AutomationElement.NameProperty, Name)
                            );
            return cText;
        }

        public static Condition CTextByClassNameAndName(string ClassName, string Name)
        {
            Condition cText = new AndCondition(
                                  new PropertyCondition(AutomationElement.ClassNameProperty, ClassName),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                                  new PropertyCondition(AutomationElement.NameProperty, Name)
                            );
            return cText;
        }

        public static Condition CTreeByName(string Name)
        {
            Condition cTree = new AndCondition(
                                  new PropertyCondition(AutomationElement.NameProperty, Name),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree)
                            );
            return cTree;
        }

        public static Condition CTreeByClassName(string className)
        {
            Condition cTree = new AndCondition(
                                  new PropertyCondition(AutomationElement.ClassNameProperty, className),
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree)
                            );
            return cTree;
        }

        public static Condition CTreeItemByClassName(string className)
        {
            Condition cTreeItem = new AndCondition(
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
                                  new PropertyCondition(AutomationElement.ClassNameProperty, className)
                            );
            return cTreeItem;
        }

        public static Condition CTreeItemByName(string Name)
        {
            Condition cTreeItem = new AndCondition(
                                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
                                  new PropertyCondition(AutomationElement.NameProperty, Name)
                            );
            return cTreeItem;
        }

        public static Condition CWindowByClassName(string className)
        {
            Condition cWindow = new AndCondition(
               new PropertyCondition(AutomationElement.ClassNameProperty, className),
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                 );
            return cWindow;
        }

    }
}
