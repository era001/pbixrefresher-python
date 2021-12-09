import time
import os
import sys
import argparse
import psutil
from pywinauto.application import Application
from pywinauto import timings


def type_keys(string, element):
    """Type a string char by char to Element window"""
    for char in string:
        element.type_keys(char)

def refresh_pbix(workbook, pbi_server, folder1, folder2="", init_wait=15, refresh_timeout=30000):
    """Open power bi file, refresh data, pubish to power bi report server in designated folder."""
    
    timings.after_clickinput_wait = 1
    
    # Kill running PBI
    PROCNAME = "PBIDesktop.exe"
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == PROCNAME:
            proc.kill()
    time.sleep(3)

    # Start PBI and open the workbook
    print("Starting Power BI")
    os.system('start "" "' + workbook + '"')
    print("Waiting ",init_wait,"sec")
    time.sleep(init_wait)

    # Connect pywinauto - new from https://github.com/DarknessTech/pbixrefresher2020/blob/master/pbixrefresher.py
    print("Identifying Power BI window")
    app = Application(backend = 'uia').connect(path = PROCNAME)
    win = app.window(title_re = '.*Power BI Desktop')
    time.sleep(5)
    win.set_focus()

    # Refresh data
    print("Refresh Data and Save Report")
    win.Refresh.click_input()
    time.sleep(5)
    print("Waiting for refresh end (timeout in ", refresh_timeout,"sec)")
    win.wait("enabled", timeout = refresh_timeout)

    # Save local file - not working
    print("Saving?")
##	win.file.wait("visible")
##	win.file.click_input()
##	print("clicked file")
    win.save.wait("visible")
    win.save.click_input()
    print("clicked save")
    win.save.click_input() # clicking save twice makes this work, but then have to click file twice below
    print("clicked save again")
    print("Waiting for save (timeout in 300 sec)")
    win.wait("enabled", timeout=300)
    time.sleep(5)
####	#type_keys("%1", win)
####	#type_keys("^S", win)
##	win.Save.wait("enabled", timeout = 300)
##	#wait_win_ready(win)
##	time.sleep(5)
##	win.wait("enabled", timeout = refresh_timeout)

    # file > save as > power bi report server
    print("Save to Report Server")
    win.file.wait("visible")
    win.file.click_input()
    print("clicked file")
    win.file.click_input()
    print("clicked file again") # click twice or save button is in the way and it will not work
    win.saveas.wait("visible")
    win.saveas.click_input()
    print("clicked save as")
    time.sleep(5)
    win.powerbireportserver.wait("visible")
    win.powerbireportserver.click_input()
    print("clicked power bi report server")
    time.sleep(5)
    # choose report server
    publish_dialog = win.child_window(auto_id="modalDialog")
    publish_dialog.child_window(title = pbi_server).click_input()
    print(f"clicked workspace {pbi_server}")
    time.sleep(5)
    publish_dialog.ok.click_input()
    print("clicked ok")
    time.sleep(5)
    # choose folder to save in
    publish_dialog = win.child_window(auto_id="modalDialog")
    publish_dialog.child_window(title=folder1).double_click_input()
    print(f"selected {folder1} folder")
    time.sleep(5)
    publish_dialog.ok.click_input()
    print("clicked ok")
    time.sleep(5)
    # optionally, choose another folder to save in
    if folder2 != "":
        publish_dialog = win.child_window(auto_id="modalDialog")
        publish_dialog.child_window(title=folder2).double_click_input()
        print(f"selected {folder2} folder")
        time.sleep(5)
        publish_dialog.ok.click_input()
        print("clicked ok")
        time.sleep(5)
    # confirm overwrite
    overwrite_dialog = win.child_window(title="Confirm overwrite", auto_id="MessageDialog")
    overwrite_dialog.yes.click_input()
    print("clicked yes to overwrite")
    time.sleep(300)

    #Close
    print("Exiting")
    win.close()

    # Force close
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()

def main():
    import json
    
    # import data from config file
    config_file = r"config.json"
    with open(config_file) as f:
      data = json.load(f)

    workbook = data["workbook_path"]
    pbi_server = data["report_server"]
    folder1 = data["folder1"]
    folder2 = data["folder2"]
    init_wait = data["init_wait"]
    refresh_timeout = data["refresh_timeout"]

    refresh_pbix(workbook, pbi_server, folder1, folder2, init_wait, refresh_timeout)

    
if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)
        sys.exit(1)








