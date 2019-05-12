Attribute VB_Name = "m_Version"
'Updates & Improvements List

'11/2/2016 Update

    'Added Paste option to directions list to copy input value
    'to the clipboard and paste it to the activecell.  The undo
    'stack is NOT cleared and the user can undo changes.
    
    'Settings for the options menu and input direction are
    'now saved to the registry so the options will remain
    'when the user opens Excel and the add-in again in the future.
    
    'Added enhancements for Excel Tables.  When the activecell is in
    'a Table and the cell does not contain validation, a unique list
    'of values will be loaded and exclude the Table headers and total row.
    
    'Add Copy List feature that copies the contents of the
    'drop-down list to the clipboard.  This feature is used
    'to create a list of unique values from a column/table
    'when the activecell does not contain validation.  Also
    'works when the list is filtered with a search term to
    'only copy filtered results.


'12/4/2016 Update

    'Updated CopyClipboard code to work with 64-bit Excel.
    
'4/26/2017 Update

    'Now works with data validation created by formulas(OFFSET & INDEX) and
    'comma separated lists.
    
    'Updated Escape Key behavior.  If there text in the search box, then
    'Escape clears the search box.  If the search box is empty, then Escape
    'closes the form.
    
    'Added Auto Open feature to automatically open the form when a cell that
    'contains data validation is selected.  This option can be toggled on/off
    'with the ToggleButtonAutoOpen button.  The user's setting is stored in the
    'registry.
    
    'Input Value can be filled to all selected cells.  A msgbox prompts the user
    'to fill the all selected cells with the input value when multiple cells are
    'selected and the form is opened.
    
'10/23/2018 Update

    'Fixed issue with horizontal data validation lists Search_List macro.
    
    'Fixed issue with merged cells prompting to input value to all cells.
    'Warning will not be displayed when a single merged cell is selected.
    'If multiple merged or regular cells are selected then warning will be displayed.
