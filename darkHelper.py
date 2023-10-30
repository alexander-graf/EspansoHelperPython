import wx
import ctypes
import os
import pyperclip
import subprocess
import win32clipboard
import re

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass


file_path = os.path.expandvars("%APPDATA%/espanso/match/own.yml")
if not os.path.exists(file_path):
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    data = "matches:\n"

    with open(file_path, "w") as file:
        file.write(data)

else:
    print("File already exists")



def create_menu(frame):
    menubar = wx.MenuBar()
    fileMenu = wx.Menu()

    # Create "Open Folder" menu item
    openFolderItem = fileMenu.Append(wx.ID_ANY, 'Open Folder', 'Open current folder')

    fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')

    menubar.Append(fileMenu, '&File')
    frame.SetMenuBar(menubar)

    frame.Bind(wx.EVT_MENU, on_open_folder, openFolderItem)
    frame.Bind(wx.EVT_MENU, on_quit, fileItem)

def on_open_folder(event):
    folder_path = os.path.expandvars(r'%APPDATA%\espanso\match')

    if os.path.exists(folder_path):
        if wx.Platform == '__WXMSW__':
            subprocess.Popen(r'explorer "{}"'.format(folder_path))
        elif wx.Platform == '__WXMAC__':
            subprocess.Popen(['open', folder_path])
        elif wx.Platform == '__WXGTK__':
            subprocess.Popen(['xdg-open', folder_path])

def on_quit(event):
    frame.Close()





def strip_rich_text_formats():
    # Open the clipboard
    win32clipboard.OpenClipboard()

    # Get the clipboard data
    clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)

    # Close the clipboard
    win32clipboard.CloseClipboard()

    # Strip rich text formatting using regex
    plain_text = re.sub(r"<[^>]+>", "", clipboard_data)

    # Set the modified plain text back to the clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(plain_text)
    win32clipboard.CloseClipboard()







def on_clipboard(event):
    trigger_entry = event.GetEventObject().GetParent().GetChildren()[0]

    # Call the function to strip rich text formats from the clipboard
    strip_rich_text_formats()

    # Get the clipboard value
    clipboard_value = pyperclip.paste()



    # Set the clipboard contents with plain text format
    # to get rid of rich text
    pyperclip.copy(clipboard_value)

    # Convert the clipboard value to UTF-8
    utf8_value = clipboard_value.encode('utf-8')

    trigger = trigger_entry.GetValue()
    replace = utf8_value.decode('utf-8')   # Text übertragen, bereinigt

    indented_text = "     " + replace.replace("\n", "\n     ")

    if trigger.strip() == "":
        print("Trigger is empty")
        print(clipboard_value)
        print(utf8_value.decode('utf-8'))
        return

    print(indented_text)
    print("Trigger:", trigger)

    data = "- trigger: {}\n  replace: |\n{}\n".format(trigger, indented_text)

    with open(file_path, "a", encoding="utf-8") as file:
        file.write(data)
def on_append_multi(event):
    trigger_entry = event.GetEventObject().GetParent().GetChildren()[0]
    replace_entry = event.GetEventObject().GetParent().GetChildren()[2]
    print("type:", type(replace_entry.GetValue()))
    trigger = trigger_entry.GetValue()
    replace = replace_entry.GetValue()

    indented_text = "     " + replace.replace("\n", "\n     ")

    print(indented_text)

    data = "- trigger: {}\n  replace: |\n{}\n".format(trigger, indented_text)

    with open(file_path, "a", encoding="utf-8") as file:
        file.write(data)



def copy_to_clipboard(contents):
    # Copy the contents to the clipboard
    pyperclip.copy(contents)
    print("Contents copied to clipboard")
    print(contents)


def on_to_clipboard(event):
    # Get the contents of the multi-field
    contents = multi_entry.GetValue()

    # Call the copy_to_clipboard function
    copy_to_clipboard(contents)


def on_copy_and_print(event):
    # Get the contents of the multi-field
    contents = multi_entry.GetValue()

    # Call the copy_to_clipboard function
    copy_to_clipboard(contents)





def on_append_single(event):
    trigger_entry = event.GetEventObject().GetParent().GetChildren()[0]
    replace_entry = event.GetEventObject().GetParent().GetChildren()[1]

    trigger = trigger_entry.GetValue()
    replace = replace_entry.GetValue()

    if trigger.strip() == "" or replace.strip() == "":
        return  # Do nothing if either field is empty.

    data = "- trigger: {}\n  replace: {}\n".format(trigger, replace)

    with open(file_path, "a", encoding="utf-8") as file:
        file.write(data)


app = wx.App()

# Create the main frame
frame = wx.Frame(None, title="Espanso Helper", size=(440, 450))

# Set the background color for the frame
frame.SetBackgroundColour("#2b2b2b")

# Create a panel to hold the widgets
panel = wx.Panel(frame)

# Set the frame icon
icon = wx.Icon('C:/icons/animals-bear.ico', wx.BITMAP_TYPE_ICO)
frame.SetIcon(icon)

# Set the background color for the panel
panel.SetBackgroundColour("#2b2b2b")

# Create a sizer to organize the widgets
sizer = wx.GridBagSizer(5, 5)

# Create the text entry widgets
trigger_entry = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
replace_entry = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
multi_entry = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
trigger_entry.SetHint("Enter Trigger")
replace_entry.SetHint("Enter Replace")
multi_entry.SetHint("Enter Multi-line Text")

# Set the foreground and background colors for the text entry widgets
trigger_entry.SetForegroundColour("#ffffff")
trigger_entry.SetBackgroundColour("#333333")
replace_entry.SetForegroundColour("#ffffff")
replace_entry.SetBackgroundColour("#333333")
multi_entry.SetForegroundColour("#ffffff")
multi_entry.SetBackgroundColour("#333333")

# Create the buttons
append_single_button = wx.Button(panel, label="Append Single")
append_multi_button = wx.Button(panel, label="Append Multi")
clipboard_button = wx.Button(panel, label="From Clipboard")



# Create a "To Clipboard" button
to_clipboard_btn = wx.Button(panel, label="To Clipboard")
to_clipboard_btn.Bind(wx.EVT_BUTTON, on_to_clipboard)

# Set tooltips for the buttons
append_single_button.SetToolTip(wx.ToolTip("Append single entry to file"))
append_multi_button.SetToolTip(wx.ToolTip("Append multiple entries to file"))
clipboard_button.SetToolTip(wx.ToolTip("Copy contents to clipboard"))
to_clipboard_btn.SetToolTip(wx.ToolTip("Copy contents from multi-field to clipboard"))




# Bind the event handlers to the buttons
append_single_button.Bind(wx.EVT_BUTTON, on_append_single)
append_multi_button.Bind(wx.EVT_BUTTON, on_append_multi)
clipboard_button.Bind(wx.EVT_BUTTON, on_clipboard)

# Set the foreground and background colors for the buttons
append_single_button.SetForegroundColour("#bbee22")
append_single_button.SetBackgroundColour("#555555")
append_multi_button.SetForegroundColour("#bbee22")
append_multi_button.SetBackgroundColour("#555555")
clipboard_button.SetForegroundColour("#bbee22")
clipboard_button.SetBackgroundColour("#555555")
to_clipboard_btn.SetForegroundColour("#bbee22")
to_clipboard_btn.SetBackgroundColour("#555555")


# Add the widgets to the sizer
sizer.Add(trigger_entry, pos=(0, 0), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(replace_entry, pos=(1, 0), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(multi_entry, pos=(2, 0), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(append_single_button, pos=(0, 1), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(append_multi_button, pos=(2, 1), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(clipboard_button, pos=(4, 1), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)
sizer.Add(to_clipboard_btn, pos=(3, 1), span=(1, 1), flag=wx.EXPAND|wx.ALL, border=12)

# Set the sizer for the panel
panel.SetSizer(sizer)

# Show the frame
frame.Show()

# Start the event loop

if __name__ == '__main__':
    create_menu(frame)
    app.MainLoop()