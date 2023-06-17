# pyinstaller --onefile --hidden-import selenium --add-binary "./drivers/chromedriver;./drivers/" amazon.py
import csv
from random import randint
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
import os
import time
import openai
# openai.api_key = "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"
model_engine = "text-curie-001"

# Project Start From here.
def get_chrome_web_driver(options):
    return webdriver.Chrome("./drivers/chromedriver.exe", chrome_options=options)


def get_web_driver_options():
    chrome_options = webdriver.ChromeOptions()
    return chrome_options


def set_ignore_certificate_error(options):
    options.add_argument('--ignore-certificate-errors')


def set_browser_as_incognito(options):
    options.add_argument('--incognito')


class Amazon_API:
    def __init__(self):
        options = get_web_driver_options()
        set_browser_as_incognito(options)
        # set_automation_as_head_less(options)
        set_ignore_certificate_error(options)
        self.driver = get_chrome_web_driver(options)

    
    def GPU(self):
        gui = Tk()
        gui.geometry("300x300")
        gui.title("Ali-Express")

        
        self.folderPath = StringVar()
        a = Label(gui, text="API KEY : ")
        a.grid(row=0, column=0)
        E = Entry(gui, textvariable=self.folderPath)
        E.grid(row=0, column=1)

        c = ttk.Button(gui, text="Find", command=self.run)
        c.grid(row=6, column=1)
       
        gui.mainloop()
        gui.quit()

    
    def run(self):
        file_path = self.file_path()
        openai.api_key = self.folderPath.get()
        print(self.folderPath.get())
        print("Start Scripting...")

        # Creating New Excel File
        fieldnames = ['Vendor Name', 'Title', 'SKU', 'Brand', 'Condition', 'UPC', 'Description', 'Price', 'Sell_price', 'Cost', 'Quantity',
                      'Color', "Size", 'Feature', "Model", 'Packaging', 'Warranty', 'Weight', 'In_box_details', 'Product_link', 'Image_link_1', 'Image_link_2', 'Image_link_3', 'Image_link_4', 'Image_link_5', 'Image_link_6', 'Image_link_7', 'Image_link_8', 'Image_link_9', 'Image_link_10', 'Image_link_11', 'Image_link_12']
        with open(f'aliexpress.csv', "w", newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

        # Open Excel file to get the links
        with open(file_path, "r") as f:
            reader = csv.reader(f, delimiter=',')
            links = []
            for i, row in enumerate(reader):
                # if i == 0:
                #     links.append(row[0])
                links.append(row[0])

        links = [x.split('?')[0] for x in links]
        
        # links = ["https://www.aliexpress.com/item/1005005206354317.html?pdp_ext_f=%7B%22ship_from%22:%22CN%22,%22sku_id%22:%2212000032158922324%22%7D&sourceType=561&&scm=1007.28480.318308.0&scm_id=1007.28480.318308.0&scm-url=1007.28480.318308.0&pvid=04ab9f0b-7377-48ec-952e-ab34d40bb270&utparam=%257B%2522process_id%2522%253A%2522sd-topn-main-stream-1%2522%252C%2522x_object_type%2522%253A%2522product%2522%252C%2522pvid%2522%253A%252204ab9f0b-7377-48ec-952e-ab34d40bb270%2522%252C%2522belongs%2522%253A%255B%257B%2522id%2522%253A%252232001259%2522%252C%2522type%2522%253A%2522dataset%2522%257D%255D%252C%2522pageSize%2522%253A%252220%2522%252C%2522language%2522%253A%2522en%2522%252C%2522scm%2522%253A%25221007.28480.318308.0%2522%252C%2522countryId%2522%253A%2522CA%2522%252C%2522scene%2522%253A%2522SD-Waterfall%2522%252C%2522tpp_buckets%2522%253A%252221669%25230%2523265320%252350_21669%25234190%252319165%2523778_18480%25230%2523318308%25230_18480%25233885%252317677%25238%2522%252C%2522x_object_id%2522%253A%25221005005206354317%2522%257D&spm=a2g0o.tm1000001392.waterfallfortopN.tab_0_category-1_product28&aecmd=true"]
        number = 1
        # Iterate Each link to get the details of the product
        for link in links:
            print(number)
            try:
                self.driver.get(link)
                self.driver.execute_script("window.stop();")
                title = self.get_title()
                asin = self.get_asin()
                brand = "Generic"
                condition = "New"
                UPC = None
                description = self.get_description(title)
                price = self.get_price()
                sell_price = None
                cost = None
                quantity = 500
                temp_color = self.get_color(link)
                color = temp_color['color']
                size = temp_color['size']
                feature = description
                model = None
                packaging = None
                warranty = "30 Days"
                weight = self.get_weight()
                in_box_details = None
                image_links = temp_color['image_links']

                if len(size) > 1:
                    vary = "Multivariate"
                else:
                    vary = "Standard"

                # Making Columns for new excel file
                product_info = {
                    'Vendor Name': 'ETC BUYS INC.',
                    'Title': title,
                    'SKU': asin,
                    'Brand': brand,
                    'Condition': condition,
                    'UPC': None,
                    'Description': description,
                    'Price': price,
                    'Sell_price': None,
                    'Cost': None,
                    'Quantity': quantity,
                    'Color': color,
                    'Size': size,
                    'Feature': description,
                    'Model': asin,
                    'Packaging': 'Opp',
                    'Warranty': warranty,
                    'Weight': '1 Pound',
                    'In_box_details': None,
                    'Product_link': link,
                    'Image_link_1': "",
                    'Image_link_2': "",
                    'Image_link_3': "",
                    'Image_link_4': "",
                    'Image_link_5': "",
                    'Image_link_6': "",
                    'Image_link_7': "",
                    'Image_link_8': "",
                    'Image_link_9': "",
                    'Image_link_10': "",
                    'Image_link_11': "",
                    'Image_link_12': "",

                }

                with open(f'aliexpress.csv', "a", newline='', encoding="utf-8") as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    for iii, v in enumerate(zip(temp_color['color'], temp_color['image_links'])):
                        image_link_1, image_link_2, image_link_3, image_link_4, image_link_5, image_link_6, image_link_7, image_link_8, image_link_9, image_link_10, image_link_11, image_link_12 = None, None, None, None, None, None, None, None, None, None, None, None
                        product_info['Color'] = v[0]
                        if iii == 0:
                            try:
                                for i, image in enumerate(v[1]):
                                    if i == 0:
                                        image_link_1 = image
                                    if i == 1:
                                        image_link_2 = image
                                    if i == 2:
                                        image_link_3 = image
                                    if i == 3:
                                        image_link_4 = image
                                    if i == 4:
                                        image_link_5 = image
                                    if i == 5:
                                        image_link_6 = image
                                    if i == 6:
                                        image_link_7 = image
                                    if i == 7:
                                        image_link_8 = image
                                    if i == 8:
                                        image_link_9 = image
                                    if i == 9:
                                        image_link_10 = image
                                    if i == 10:
                                        image_link_11 = image
                                    if i == 11:
                                        image_link_12 = image

                                    product_info['Image_link_1'] = image_link_1
                                    product_info['Image_link_2'] = image_link_2
                                    product_info['Image_link_3'] = image_link_3
                                    product_info['Image_link_4'] = image_link_4
                                    product_info['Image_link_5'] = image_link_5
                                    product_info['Image_link_6'] = image_link_6
                                    product_info['Image_link_7'] = image_link_7
                                    product_info['Image_link_8'] = image_link_8
                                    product_info['Image_link_9'] = image_link_9
                                    product_info['Image_link_10'] = image_link_10
                                    product_info['Image_link_11'] = image_link_11
                                    product_info['Image_link_12'] = image_link_12

                            except Exception as e:
                                product_info['Image_link_1'] = v[1]
                        else:
                            product_info['Image_link_1'] = v[1]
                            product_info['Image_link_2'] = image_link_2
                            product_info['Image_link_3'] = image_link_3
                            product_info['Image_link_4'] = image_link_4
                            product_info['Image_link_5'] = image_link_5
                            product_info['Image_link_6'] = image_link_6
                            product_info['Image_link_7'] = image_link_7
                            product_info['Image_link_8'] = image_link_8
                            product_info['Image_link_9'] = image_link_9
                            product_info['Image_link_10'] = image_link_10
                            product_info['Image_link_11'] = image_link_11
                            product_info['Image_link_12'] = image_link_12

                        if temp_color['size'] == None or temp_color['size'] == []:
                            product_info['Size'] = None
                            sku_id = ""
                            for x in range(9):
                                # Generate a random Uppercase letter (based on ASCII code)
                                upperCaseLetter = chr(randint(65, 90))
                                numbers = str(randint(0, 9))
                                if x not in [3, 4, 5, 6]:
                                    sku_id = sku_id + upperCaseLetter
                                else:
                                    sku_id = sku_id + numbers
                            product_info['SKU'] = sku_id
                            product_info['Model'] = sku_id
                            writer.writerow(product_info)

                        else:
                            for j, v2 in enumerate(temp_color['size']):
                                product_info['Size'] = v2
                                sku_id = ""
                                for x in range(9):
                                    # Generate a random Uppercase letter (based on ASCII code)
                                    upperCaseLetter = chr(randint(65, 90))
                                    numbers = str(randint(0, 9))
                                    if x not in [3, 4, 5, 6]:
                                        sku_id = sku_id + upperCaseLetter
                                    else:
                                        sku_id = sku_id + numbers
                                product_info['SKU'] = sku_id
                                product_info['Model'] = sku_id

                                writer.writerow(product_info)

            except Exception as e:
                # print(e)
                if number != 0:
                    print("Link Doesn't Work")
                pass
            number = number + 1

        self.driver.quit()
        print("Done")

    def file_path(self):
        file = filedialog.askopenfile(mode='r')
        if file:
            filepath = os.path.abspath(file.name)
        else:
            filepath = ""
            print("Error in Selection of File!! Select the file Again !!")
        return filepath

    def get_title(self):
        try:
            return self.driver.find_element("class name", 'product-title-text').text
        except Exception as e:
            # print("Didn't get the title !!")
            return None

    def get_asin(self):
        try:
            numberPlate = ""
            for x in range(9):
                # Generate a random Uppercase letter (based on ASCII code)
                upperCaseLetter = chr(randint(65, 90))
                numbers = str(randint(0, 9))
                if x not in [3, 4, 5, 6]:
                    numberPlate = numberPlate + upperCaseLetter
                else:
                    numberPlate = numberPlate + numbers

            return str(numberPlate)
        except Exception as e:
            # print("Didn't get the ASIN !!")
            return None

    def get_description(self, title):
        try:
            self.driver.implicitly_wait(2)
            prompt = f"Please provide me with a brief summary consisting of four sentences that describe this title '{title}'"

            # Generate a response
            completion = openai.Completion.create(
                engine=model_engine,
                prompt=prompt,
                max_tokens=100,
                n=1,
                stop=None,
                temperature=0.5,
            )

            desc = completion.choices[0].text
            desc = desc.replace('\n', '')
            # print(desc)
            return desc
        except Exception as e:
            print(e)
            # print("Didn't get the Description !!")
            return None

    def get_price(self):
        try:
            price = self.driver.find_element("class name",
                                             "uniform-banner-box-price").text
            if "$" in price:
                price = price.split('$')[1]
                price = f"${price}"
            return price
        except Exception as e:
            try:
                price = self.driver.find_element("xpath",
                                                 "//div[contains(@class, 'product-price-current')]").text
                if "$" in price:
                    price = price.split("$")[1]
                    price = f"${price}"
                return price
            except Exception as e:
                pass
                # print(e)
                # print("You didn't Catch the Price")
                return None

    def get_color(self, link):
        asin_color = f"str(link).split('item/')[1].split('.html')[0]"
        try:
            temp = self.driver.find_element("class name", "sku-wrap")
            temp1 = temp.find_elements("class name", "sku-property")
            color = []
            size = []
            images_links = []
            ii = 1
            for x in temp1:
                try:
                    if "color" in x.find_element("class name", "sku-title").text.lower():
                        color1 = x.find_element(
                            "class name", "sku-property-list")
                        color2 = color1.find_elements("tag name", "img")
                        color = [y.get_attribute("alt") for y in color2]
                except:
                    color = []
                try:
                    if "size" in x.find_element("class name", "sku-title").text.lower():
                        size1 = x.find_element(
                            "class name", "sku-property-list")
                        size2 = size1.find_elements("tag name", "span")
                        size = [y.text for y in size2]
                except:
                    size = None
                try:
                    if "color" in x.find_element("class name", "sku-title").text.lower():
                        images_links = []

                        # For first color with all images on display
                        temp_imgs = self.driver.find_element(
                            'class name', 'images-view-list')
                        t_img1 = temp_imgs.find_elements("tag name", "img")
                        t_image_link = [y.get_attribute("src") for y in t_img1]
                        tt_image_link = []
                        for i in t_image_link:
                            new_link = i.split(".jpg")
                            new_link[1] = "_1500x1500"
                            new = ".jpg".join(new_link)
                            tt_image_link.append(str(new))

                        # For other colors
                        img = x.find_element(
                            "class name", "sku-property-list")
                        img1 = img.find_elements("tag name", "img")
                        image_link = [y.get_attribute("src") for y in img1]
                        r_image_link = []
                        for i in image_link[1:]:
                            new_link = i.split(".jpg")
                            new_link[1] = "_1500x1500"
                            new = ".jpg".join(new_link)
                            r_image_link.append(str(new))
                        # Merged all the images in one list
                        all_images_links = [tt_image_link] + r_image_link
                except Exception as e:
                    all_images_links = []

            colors_size = {
                "color": color,
                "size": size,
                "image_links": all_images_links,
            }
            return colors_size
        except Exception as e:
            pass

    def get_weight(self):
        try:
            weight = "1 Pound"
            return weight
        except Exception as e:
            weight = "1 Pound"
            # print("You didn't Catch the Weight")
            return weight

if __name__ == "__main__":
    am = Amazon_API()
    am.GPU()
    # am.run()
