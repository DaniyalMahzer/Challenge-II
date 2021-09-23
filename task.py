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

    def read_orders(self):
        user = self.secret.get_secret("credentials")
        self.dialogs.add_text(f"Hi {user['username']}!")
        self.dialogs.add_file_input(name="orders")
        user = self.dialogs.run_dialog()
        self.filesys.change_file_extension(path=user.orders[0], extension=".xlsx")
        print(user.orders[0])
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
