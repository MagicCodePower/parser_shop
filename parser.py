import requests
from bs4 import BeautifulSoup
import xlsxwriter
import os

class parse:
    def __init__(self):
        pass

    def get_url(self, url):
        self.session = requests.Session()
        headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.100 Safari/537.36',
        }
        response = self.session.get(url, headers=headers)
        content = BeautifulSoup(response.text, 'html.parser')
        return content

    def get_info(self, html_page, block_name=None, block_atr=None, rel=None):

        block = html_page.findAll(block_name, block_atr, rel=rel)

        return block

    def del_file(self):
        try:
            os.remove('new_data_rock.xlsx')
        except:
            pass

    def create_file(self):
        self.workbook = xlsxwriter.Workbook('new_data_rock.xlsx')
        self.worksheet = self.workbook.add_worksheet('rock')

    def get_name(self,name,poz):

        name = name.replace('/','')
        name = name.replace('"','')
        folder = r'D:\git\shop_parser\img '
        return folder + name+ str(poz) +'.jpg'

    def save_image(self,name, file_object):
        try:
            with open(name, 'bw') as f:
                for chunk in file_object.iter_content(8192):
                    f.write(chunk)
            self.name_img = name
        except:
            # self.name_img = 'D:\Programs\python\parse_shop\img\ TPU+PC чехол ROCK Origin Pro Series Black для iPhone XXS4.jpg'
            pass

    def into_excel(self, items):
        for elem in range(len(self.all_data)):
            print(elem)
            self.worksheet.write(elem, 0, items[elem][0])
            self.worksheet.write(elem, 1, items[elem][1])
            self.worksheet.write(elem, 2, items[elem][2])
            for i in range(3,3+len(items[elem][3:])):
                try:
                    our_img = requests.get(items[elem][i])
                    if our_img.status_code == 200:
                        self.save_image(self, self.get_name(self,items[elem][0],i), our_img)
                        self.worksheet.insert_image(elem,i,self.name_img,   {'x_scale': 0.1, 'y_scale': 0.1})
                    else:
                        continue
                except:
                    pass


    def main(self):
        self.all_data = []
        for i in range(1,10):

            html =  self.get_url(self, 'https://ilounge.ua/brands/rock?page={}'.format(i))
            # print(self.get_info(self,html,'a','product_click')[0].get('href'))
            all_items = self.get_info(self,html,'a','product_click')
            for item in all_items[::2]:
                data = []
                item_url = 'https://ilounge.ua/' + item.get('href')
                self.item_html = self.get_url(self,item_url)
                print(self.get_info(self,self.item_html,'h1')[0].text)
                data.append(self.get_info(self,self.item_html,'h1')[0].text)
                data.append(self.get_info(self, self.item_html,'span', 'product-price')[0].text)
                data.append((self.get_info(self, self.item_html, 'div','productdesc')[0].text).strip())
                try:
                    img = self.get_info(self, self.get_info(self, self.item_html, 'ul', 'vertslider')[0],'li')
                    for i in img:
                        data.append(i.a['href'])
                except:
                    try:
                        img = self.get_info(self,self.item_html,'div','image')[0].a['href']
                        data.append(img)
                    except:
                        data.append(None)
                print(data)

                self.all_data.append(data)
        self.create_file(self)
        try:
            self.into_excel(self,self.all_data)
        except:
            pass
        self.workbook.close()

if __name__ == '__main__':
    parse.main(parse)