from RPA.Browser.Selenium import Selenium
from RPA.Dialogs import Dialogs
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Robocorp.Vault import FileSecrets
from RPA.FileSystem import FileSystem


class Certificate_II:
    browser = Selenium()
    pdf = PDF()
    excel = Files()
    dialogs = Dialogs()
    archive = Archive()
    secret = FileSecrets()
    filesys = FileSystem()

    def make_zip(self):
        self.archive.archive_folder_with_zip("output", "output/orders")
        self.dialogs.add_file("output/orders")
        self.dialogs.run_dialog()

    def read_orders(self):
        user = self.secret.get_secret("credentials")
        self.dialogs.add_text(f"Hi {user['username']}!")
        self.dialogs.add_file_input(name="orders")
        user = self.dialogs.run_dialog()
        if self.filesys.get_file_extension(user.orders[0]) == ".csv":
            file_data = self.filesys.read_file(user.orders[0])
            real_data = []
            data = file_data.split("\n")
            for item in data:
                test_data = item.split(",")
                real_data.append({
                    "Order number": test_data[0],
                    "Head": test_data[1],
                    "Body": test_data[2],
                    "Legs": test_data[3],
                    "Address": test_data[4],
                })
            real_data.pop(0)
            return real_data
        else:
            self.excel.open_workbook(path=user.orders[0])
            data = self.excel.read_worksheet(header=True)
            return data

    def place_orders(self, data):
        links = self.secret.get_secret("links")
        self.browser.open_available_browser(links["order_link"], maximized=True)
        for item in data:
            try:
                self.browser.wait_until_page_contains("Order")
                self.browser.click_button("OK")
                self.browser.select_from_list_by_value("head", str(item["Head"]))
                self.browser.select_radio_button("body", str(item["Body"]))
                self.browser.input_text("class:form-control", str(item["Legs"]))
                self.browser.input_text("address", str(item["Address"]))
                self.browser.click_button("Preview")
                self.browser.click_button("Order")
                while True:
                    try:
                        self.browser.find_element("id:receipt")
                        break
                    except:
                        self.browser.click_button("Order")
                self.browser.wait_until_element_is_visible('//*[@id="robot-preview-image"]/img[3]')
                self.browser.wait_until_element_is_visible('//*[@id="robot-preview-image"]/img[2]')
                self.browser.wait_until_element_is_visible('//*[@id="robot-preview-image"]/img[1]')
                recipt = self.browser.get_element_attribute("id:receipt", "outerHTML")
                self.pdf.html_to_pdf(recipt, f'output/order{item["Order number"]}.pdf')
                self.browser.screenshot("robot-preview-image", f'images/order{item["Order number"]}.png')
                self.pdf.add_watermark_image_to_pdf(
                    image_path=f'images/order{item["Order number"]}.png',
                    source_path=f'output/order{item["Order number"]}.pdf',
                    output_path=f'output/order{item["Order number"]}.pdf',
                )
                orders.append(str(item["Order number"]))
                self.browser.click_button("id:order-another")
            except:
                self.browser.close_all_browsers()
                return None


obj = Certificate_II()
orders = obj.read_orders()
obj.place_orders(orders)
obj.make_zip()
