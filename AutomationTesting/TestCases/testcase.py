import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from Data.testdata import BrowserData, HomePageData
from Locators.locator import HomePage, OrangeHrmPage
from webdriver_manager.chrome import ChromeDriverManager
from TestCases import XLUtils

class TestApplication:

    def __init__(self):
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        self.driver.maximize_window()

    def test_sanity(self):
        """
            This method / function helps to do sanity testing.It will check all the basic functionality of the webpage
        """
        self.driver.get(BrowserData.url)
        self.driver.implicitly_wait(50)
        self.driver.find_element(By.XPATH, HomePage.name).send_keys(HomePageData.name)
        self.driver.find_element(By.XPATH, HomePage.email).send_keys(HomePageData.email)
        self.driver.find_element(By.XPATH, HomePage.phone).send_keys(HomePageData.phone)
        self.driver.find_element(By.XPATH, HomePage.address).send_keys(HomePageData.address)
        self.driver.find_element(By.XPATH, HomePage.address).send_keys(Keys.ENTER)
        self.driver.find_element(By.XPATH, HomePage.address).send_keys(HomePageData.pincode)
        self.driver.get_screenshot_as_file("..\\Screenshots\\personaldata.png")
        radio = self.driver.find_element(By.XPATH, HomePage.gender_male).is_selected()
        if radio:
            pass
        else:
            self.driver.find_element(By.XPATH, HomePage.gender_male).click()
            print("Gender selected")
        self.driver.find_element(By.XPATH, HomePage.day_saturday).click()
        element = self.driver.find_element(By.XPATH, HomePage.country)
        self.driver.execute_script("arguments[0].scrollIntoView();", element)
        drpdown = Select(element)
        drpdown.select_by_visible_text(HomePageData.country)
        print("selected county : " + HomePageData.country)
        self.driver.execute_script("window.scrollBy(0,document.body.scrollHeight)")
        colours = self.driver.find_element(By.XPATH, HomePage.colours)
        colour = Select(colours)
        colour.select_by_value(HomePageData.colour)
        print("selected colour : " + HomePageData.colour)
        self.driver.find_element(By.XPATH, HomePage.date).click()
        self.driver.get_screenshot_as_file("..\\Screenshots\\personalchoices.png")
        alldates = self.driver.find_elements(By.XPATH, HomePage.datepicker)
        for dates in alldates:
            date = dates.text
            if date == str(HomePageData.date):
                dates.click()
                print("selected date : " + dates.text)
                break
        self.driver.implicitly_wait(20)
        self.driver.find_element(By.XPATH, HomePage.alert).click()
        al = self.driver.switch_to.alert
        alert_text = al.text
        time.sleep(5)
        try:
            assert alert_text == HomePageData.alertext
        except AssertionError:
            print("displayed text inside alert box is wrong")
            self.driver.get_screenshot_as_file("..\\Screenshots\\alertfailed.png")
        al.accept()
        element = self.driver.find_element(By.XPATH, HomePage.copy_text)
        action = ActionChains(self.driver)
        action.double_click(element).perform()
        element = self.driver.find_element(By.XPATH, HomePage.scrolltodrop)
        self.driver.execute_script("arguments[0].scrollIntoView();", element)
        action.drag_and_drop_by_offset(element, 0, 100).perform()
        time.sleep(2)
        self.driver.get_screenshot_as_file("..\\Screenshots\\slideranddoubleclick.png")
        source = self.driver.find_element(By.XPATH, HomePage.source)
        target = self.driver.find_element(By.XPATH, HomePage.target)
        action.drag_and_drop(source, target).perform()
        parent_window_id = self.driver.current_window_handle
        self.driver.find_element(By.XPATH, HomePage.newwindow).click()
        windows = self.driver.window_handles
        for w in windows:
            self.driver.switch_to.window(w)
            if self.driver.title.__eq__("Your Store"):
                break
        self.driver.implicitly_wait(25)
        element = self.driver.find_element(By.XPATH, HomePage.item)
        self.driver.execute_script("arguments[0].scrollIntoView();", element)
        time.sleep(5)
        self.driver.find_element(By.XPATH, HomePage.item).click()
        time.sleep(15)
        self.driver.execute_script("window.scrollBy(0,800)", "")
        self.driver.find_element(By.XPATH, HomePage.addtocart).click()
        time.sleep(3)
        self.driver.get_screenshot_as_file("..\\Screenshots\\addtocart.png")
        wait = WebDriverWait(self.driver, 15)
        wait.until(EC.visibility_of_element_located((By.XPATH, HomePage.successmessage)))
        message = self.driver.find_element(By.XPATH, HomePage.successmessage).text
        assert message == HomePageData.expectedmesage
        print(message)
        self.driver.switch_to.window(parent_window_id)
        wait.until(EC.element_to_be_clickable((By.XPATH, HomePage.alert)))
        result = self.driver.find_element(By.XPATH, HomePage.confirmationbox).is_displayed()
        assert result.__eq__(True)
        print("Testcase passed ")

        def test_login(self):
        self.driver.get(BrowserData.url)
        self.driver.find_element(By.XPATH, "//a[normalize-space()='orange HRM']").click()
        time.sleep(10)
        path = "..\\Data\\testexceldata.xlsx"
        rows = XLUtils.getrowcount(path, 'Sheet1')

        for r in range(2, rows + 1):
            username = XLUtils.readdata(path, "Sheet1", r, 1)
            password = XLUtils.readdata(path, "Sheet1", r, 2)
            time.sleep(5)
            self.driver.implicitly_wait(5)
            cookie_before = self.driver.get_cookies()[0]['value']
            self.driver.find_element(By.XPATH, OrangeHrmPage.username).clear()
            self.driver.find_element(By.XPATH, OrangeHrmPage.username).send_keys(username)
            self.driver.find_element(By.XPATH, OrangeHrmPage.password).clear()
            self.driver.find_element(By.XPATH, OrangeHrmPage.password).send_keys(password)
            self.driver.find_element(By.XPATH,
                                     "//button[normalize-space(@class) = 'oxd-button oxd-button--medium oxd-button--main orangehrm-login-button']").click()
            time.sleep(5)
            self.driver.get_screenshot_as_file("..\\Screenshots\\loginvalidation"+str(r)+".png")
            cookie_after = self.driver.get_cookies()[0]['value']
            if cookie_before != cookie_after:
                print("Login successful")
                XLUtils.writedata(path, "Sheet1", r, 3, "Login Successful")
            else:
                print("Login failed")
                XLUtils.writedata(path, "Sheet1", r, 3, "Login Failed")
            self.driver.delete_all_cookies()
            self.driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
            time.sleep(5)
        self.driver.quit()


test = TestApplication()
test.test_sanity()
test.test_login()

