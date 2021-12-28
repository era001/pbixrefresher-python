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

def main():   
	# Parse arguments from cmd
	parser = argparse.ArgumentParser()
	parser.add_argument("workbook", help="Path to .pbix file")
	parser.add_argument("pbi_server", help="Power BI Report Server to publish on")
	parser.add_argument("--folder1", help="name of folder to publish in", default="")
	parser.add_argument("--folder2", help="name of nested folder to publish in", default="")
	parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
	parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
	parser.add_argument("--init-wait", help = "initial wait time on startup", default = 15, type = int)

	args = parser.parse_args()

	timings.after_clickinput_wait = 1
	WORKBOOK = args.workbook
	PBI_SERVER = args.pbi_server
	FOLDER1 = args.folder1
	FOLDER2 = args.folder2
	INIT_WAIT = args.init_wait
	REFRESH_TIMEOUT = args.refresh_timeout

	# Kill running PBI
	PROCNAME = "PBIDesktop.exe"
	for proc in psutil.process_iter():
		# check whether the process name matches
		if proc.name() == PROCNAME:
			proc.kill()
	time.sleep(3)

	# Start PBI and open the workbook
	print("Starting Power BI")
	os.system('start "" "' + WORKBOOK + '"')
	print("Waiting ",INIT_WAIT,"sec")
	time.sleep(INIT_WAIT)

	# Connect pywinauto
	print("Identifying Power BI window")
	app = Application(backend='uia').connect(path=PROCNAME)
	win = app.window(title_re = '.*Power BI Desktop')
	time.sleep(5)
	# win.wait("enabled", timeout = 300)
	# win.Save.wait("enabled", timeout = 300)
	win.set_focus()
	# win.Home.click_input()
	# win.Save.wait("enabled", timeout = 300)
	# win.wait("enabled", timeout = 300)

	# Refresh
	print("Refreshing")
	win.RefreshButton.click_input()
	#wait_win_ready(win)
	time.sleep(5)
	print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT, "sec)")
	win.wait("enabled", timeout=REFRESH_TIMEOUT)
	

	# Save - not working
	print("Saving")
	win.SaveButton.click_input()
	win.wait("enabled", timeout=REFRESH_TIMEOUT)
	win.set_focus()
	time.sleep(5)

	# Publish
	if args.publish:
		print("Publish")
		time.sleep(5)
		# win.file.wait("visible")
		win.file.click_input()
		time.sleep(5)
		# win.saveas.wait("visible")
		win.saveas.click_input()
		time.sleep(5)
		# win.powerbireportserver.wait("visible")
		win.powerbireportserver.click_input()
		time.sleep(5)
		publish_dialog = win.child_window(auto_id="modalDialog")
		publish_dialog.child_window(title=PBI_SERVER).click_input()
		time.sleep(5)
		publish_dialog.ok.click()
		time.sleep(5)
		if FOLDER1 != "":
			publish_dialog = win.child_window(auto_id="modalDialog")
			publish_dialog.child_window(title=FOLDER1).double_click_input()
			print(f"Selected {FOLDER1} folder")
			time.sleep(5)
			if FOLDER2 != "":
				publish_dialog = win.child_window(auto_id="modalDialog")
				publish_dialog.child_window(title=FOLDER2).double_click_input()
				print(f"Selected {FOLDER2} folder")
				publish_dialog.ok.click_input()
				time.sleep(5)
			else:
				publish_dialog.ok.click_input()
				time.sleep(5)

		# confirm overwrite
		overwrite_dialog = win.child_window(title="Confirm overwrite", auto_id="MessageDialog")
		overwrite_dialog.yes.click_input()
		time.sleep(30)

	#Close
	print("Exiting")
	win.close()

	# Force close
	for proc in psutil.process_iter():
		if proc.name() == PROCNAME:
			proc.kill()

		
if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print(e)
		sys.exit(1)








