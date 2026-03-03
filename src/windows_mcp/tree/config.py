INTERACTIVE_CONTROL_TYPE_NAMES = set(
    [
        "ButtonControl",
        "ListItemControl",
        "MenuItemControl",
        "EditControl",
        "CheckBoxControl",
        "RadioButtonControl",
        "ComboBoxControl",
        "HyperlinkControl",
        "SplitButtonControl",
        "TabItemControl",
        "TreeItemControl",
        "DataItemControl",
        "HeaderItemControl",
        "TextBoxControl",
        "SpinnerControl",
        "ScrollBarControl",
    ]
)

INTERACTIVE_ROLES = {
    # Buttons
    "PushButton",
    "SplitButton",
    "ButtonDropDown",
    "ButtonMenu",
    "ButtonDropDownGrid",
    "OutlineButton",
    # Links
    "Link",
    # Inputs & Selection
    "Text",
    "IpAddress",
    "HotkeyField",
    "ComboBox",
    "DropList",
    "CheckButton",
    "RadioButton",
    # Menus & Tabs
    "MenuItem",
    "ListItem",
    "PageTab",
    # Trees
    "OutlineItem",
    # Values
    "Slider",
    "SpinButton",
    "Dial",
    "ScrollBar",
    "Grip",
    # Grids
    "ColumnHeader",
    "RowHeader",
    "Cell",
}

DOCUMENT_CONTROL_TYPE_NAMES = set(["DocumentControl"])

STRUCTURAL_CONTROL_TYPE_NAMES = set(["PaneControl", "GroupControl", "CustomControl"])

INFORMATIVE_CONTROL_TYPE_NAMES = set(
    [
        "TextControl",
        "ImageControl",
        "StatusBarControl",
        # 'ProgressBarControl',
        # 'ToolTipControl',
        # 'TitleBarControl',
        # 'SeparatorControl',
        # 'HeaderControl',
        # 'HeaderItemControl',
    ]
)

DEFAULT_ACTIONS = set(["Click", "Press", "Jump", "Check", "Uncheck", "Double Click"])

import os

THREAD_MAX_RETRIES = int(os.getenv("WMCP_MAX_RETRIES", "1"))
THREAD_MAX_WORKERS = int(os.getenv("WMCP_MAX_WORKERS", "4"))
MAX_TREE_DEPTH = int(os.getenv("WMCP_MAX_TREE_DEPTH", "15"))
MAX_ELEMENTS = int(os.getenv("WMCP_MAX_ELEMENTS", "100"))
WATCHDOG_ENABLED = os.getenv("WMCP_WATCHDOG_ENABLED", "true").lower() != "false"
