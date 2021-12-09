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
##	# Parse arguments from cmd
##	parser = argparse.ArgumentParser()
##	parser.add_argument("workbook", help = "Path to .pbix file")
##	parser.add_argument("--workspace", help = "name of online Power BI service work space to publish in", default = "My workspace")
##	parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
##	parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
##	parser.add_argument("--init-wait", help = "initial wait time on startup", default = 15, type = int)
##	args = parser.parse_args()

	timings.after_clickinput_wait = 1
	#WORKBOOK = args.workbook
	#WORKBOOK = r"C:\Git\pbixrefresher-python\sample.pbix"
	#WORKBOOK = r"C:\Git\hfd-public-outreach-survey\power_bi\Public-Outreach.pbix"
	WORKBOOK = r"C:\Git\hfd-special-operations-survey\report\Survey-Special-Operations.pbix"
	#WORKSPACE = args.workspace
	#WORKSPACE = "My workspace"
	WORKSPACE = "http://wit932/ReportsPowerBI"
	FOLDER = "GIS"
	#INIT_WAIT = args.init_wait
	INIT_WAIT = 15
	#REFRESH_TIMEOUT = args.refresh_timeout
	REFRESH_TIMEOUT = 30000

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

	# Connect pywinauto - new from https://github.com/DarknessTech/pbixrefresher2020/blob/master/pbixrefresher.py
	print("Identifying Power BI window")
	app = Application(backend = 'uia').connect(path = PROCNAME)
	win = app.window(title_re = '.*Power BI Desktop')
	time.sleep(5)
	win.set_focus()

	# Refresh
	print("Refresh Data and Save Report")
	win.Refresh.click_input()
	time.sleep(5)
	print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT,"sec)")
	win.wait("enabled", timeout = REFRESH_TIMEOUT)

	# Save - not working
	print("Saving?")
##	win.file.wait("visible")
##	win.file.click_input()
##	print("clicked file")
	win.save.wait("visible")
	win.save.click_input()
	print("clicked save")
	print("Waiting for save (timeout in 300 sec)")
	win.wait("enabled", timeout=300)
	#time.sleep(5)
####	#type_keys("%1", win)
####	#type_keys("^S", win)
##	win.Save.wait("enabled", timeout = 300)
##	#wait_win_ready(win)
##	time.sleep(5)
##	win.wait("enabled", timeout = REFRESH_TIMEOUT)

	# Publish
##	if args.publish:
	if True == True:
		print("Save to Report Server")
		win.file.wait("visible")
		win.file.click_input()
		print("clicked file")
		win.saveas.wait("visible")
		win.saveas.click_input()
		print("clicked save as")
		time.sleep(5)
		win.powerbireportserver.wait("visible")
		win.powerbireportserver.click_input()
		print("clicked power bi report server")
		time.sleep(5)
		publish_dialog = win.child_window(auto_id="modalDialog")
		publish_dialog.child_window(title = WORKSPACE).click_input()
		print("clicked dev workspace")
		time.sleep(5)
		publish_dialog.ok.click_input()
		print("clicked ok")
		time.sleep(5)
		# need to select GIS folder
		publish_dialog = win.child_window(auto_id="modalDialog")
		publish_dialog.child_window(title=FOLDER).double_click_input()
		print("selected gis folder")
		time.sleep(5)
		publish_dialog.ok.click_input()
		print("clicked ok")
		time.sleep(5)
		# need to confirm overwrite
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

		
if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print(e)
		sys.exit(1)








