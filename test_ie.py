"""
System requirements:
OS: Windows Server 2012 (any edition)
Browser: Internet Explorer ver 11.xxx
Java (to execute java applets) JDK ver 8u128 or higher, don't need for this example
Screen resolution: 1920x1020
Python ver 3.7
to install all libraries run:
pip install -r requirements.txt
execute script:
python test_ie.py
This script should be executed using Jenkins
"""
import os
import win32gui
import win32ui
import win32con
import win32api
import time
import unittest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


class PythonOrgSearch(unittest.TestCase):
    def delete_screenshots(filename):
        if os.path.isfile(filename):
            os.remove(filename)

    def select_driver(browser):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            if browser == "ie32":
                driver = webdriver.Ie(current_dir + "\\drivers\\IEDriverServer32.exe")
                return driver
            elif browser == "ie64":
                driver = webdriver.Ie(current_dir + "\\drivers\\IEDriverServer64.exe")
                return driver
            elif browser == "Firefox32":
                Driver_Path = current_dir + "\\drivers\\geckodriver32.exe"
                driver = webdriver.Firefox(executable_path=Driver_Path)
                print(type(driver))
                return driver
            elif browser == "Firefox64":
                Driver_Path = current_dir + "\\drivers\\geckodriver64.exe"
                driver = webdriver.Firefox(executable_path=Driver_Path)
                print(type(driver))
                return driver
            elif browser == "EDGE":
                Driver_Path = current_dir + "\\drivers\\msedgedriver.exe"
                driver = webdriver.Edge(executable_path=Driver_Path)
                print(type(driver))
                return driver
            else:
                Driver_Path = current_dir + "\\drivers\\chromedriver.exe"
                driver = webdriver.Chrome(executable_path=Driver_Path)
                print(type(driver))
                return driver
        except Exception as e:
            print("Driver is not applied: ", str(e))


    def test_execute_small_test(browser):
            width = 1920
            height = 1020
            screenshot_file = "results.png"
            screenshot2 = "results2.png"
            os_screenfilename1 = "os_filename1.bmp"
            os_screenfilename2 = "os_filename2.bmp"
            PythonOrgSearch.delete_screenshots(screenshot_file)
            PythonOrgSearch.delete_screenshots(screenshot2)
            PythonOrgSearch.delete_screenshots(os_screenfilename1)
            PythonOrgSearch.delete_screenshots(os_screenfilename2)
            browser = "ie32"
            drv = PythonOrgSearch.select_driver(browser)
            drv.maximize_window()
            drv.set_window_size(width, height)
            drv.get('https://google.com')
            drv.implicitly_wait(10)
            drv.find_element_by_name('q').send_keys("Python github", Keys.ENTER)
            time.sleep(3)
            get_screenshot(os_screenfilename1)
            try:
                drv.get_screenshot_as_file(screenshot_file)
            except Exception as e:
                print("Screenshot is not saved", str(e))
            try:
                matched_elements = drv.find_elements_by_xpath('//a[starts-with(@href, "https://github.com/PyGithub/PyGithub")]')
                print("Type:", type(matched_elements ), " \n Value: ", matched_elements[0] )
                time.sleep(5)
                if matched_elements:
                    matched_elements[0].click()
                    time.sleep(5)
            except Exception as e:
                print("Unable to find element: ", str(e))

            try:
                get_screenshot(os_screenfilename2)
                drv.get_screenshot_as_file(screenshot2)
            except Exception as e:
                print("Screenshot is not saved", str(e))
            drv.quit()

def get_screenshot(filename):
    # grab a handle to the main desktop window
    hdesktop = win32gui.GetDesktopWindow()

    # determine the size of all monitors in pixels
    width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

    # create a device context
    desktop_dc = win32gui.GetWindowDC(hdesktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)

    # create a memory based device context
    mem_dc = img_dc.CreateCompatibleDC()

    # create a bitmap object
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    mem_dc.SelectObject(screenshot)

    # copy the screen into our memory device context
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)

    # save the bitmap to a file
    screenshot.SaveBitmapFile(mem_dc, filename)
    # free our objects
    mem_dc.DeleteDC()
    win32gui.DeleteObject(screenshot.GetHandle())

if __name__ == "__main__":
    unittest.main()


